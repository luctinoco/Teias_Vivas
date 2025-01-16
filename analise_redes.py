import sys
import pandas as pd  # Biblioteca para manipulação de dados tabulares
import networkx as nx  # Biblioteca para criação e análise de redes complexas
import matplotlib.pyplot as plt  # Biblioteca para visualização gráfica
from networkx.algorithms.community import greedy_modularity_communities, label_propagation_communities

# ======== OPÇÃO PARA SALVAR PRINTS EM UM ARQUIVO DE TEXTO ======== #
# Defina o caminho do seu arquivo de log
log_path = "C:/Users/lucas/Analise_redes/logs.txt"

# Abre o arquivo em modo de escrita
log_file = open(log_path, "w")
# Salva o stdout original para restaurar depois
original_stdout = sys.stdout
# Redireciona todos os prints para o arquivo log_file
sys.stdout = log_file
# ================================================================== #

try:
    # =======================================================
    # 1) Carregar o arquivo Excel
    # =======================================================
    file_path = 'C:/Users/lucas/Analise_redes/Redes_Mestrado.xlsx'  # Caminho para o arquivo Excel
    data = pd.ExcelFile(file_path)  # Abre o arquivo Excel

    # =======================================================
    # 2) Verificar as planilhas disponíveis no arquivo
    # =======================================================
    sheet_names = data.sheet_names  # Lista os nomes das planilhas
    print("Planilhas disponíveis:", sheet_names)

    # Verificar se as planilhas '18S' e '12S' existem no arquivo
    if '18S' not in sheet_names:
        raise ValueError("A planilha '18S' não foi encontrada no arquivo Excel.")
    if '12S' not in sheet_names:
        raise ValueError("A planilha '12S' não foi encontrada no arquivo Excel.")

    # =======================================================
    # 3) Carregar os dados das planilhas 18S e 12S
    # =======================================================
    data_18s = data.parse(sheet_name='18S')  # Carrega os dados da planilha 18S
    data_12s = data.parse(sheet_name='12S')  # Carrega os dados da planilha 12S

    # =======================================================
    # 4) Normalizar os identificadores de indivíduos
    # =======================================================
    data_18s.iloc[:, 0] = data_18s.iloc[:, 0].str.strip().str.lower()
    data_12s.iloc[:, 0] = data_12s.iloc[:, 0].str.strip().str.lower()

    # =======================================================
    # 5) Verificar se as colunas das planilhas (exceto a primeira) são binárias (0 ou 1)
    # =======================================================
    for col in data_18s.columns[1:]:
        valores_unicos = set(data_18s[col].dropna().unique())
        if not valores_unicos.issubset({0, 1}):
            raise ValueError(f"A coluna '{col}' em 18S contém valores não binários: {valores_unicos}")

    for col in data_12s.columns[1:]:
        valores_unicos = set(data_12s[col].dropna().unique())
        if not valores_unicos.issubset({0, 1}):
            raise ValueError(f"A coluna '{col}' em 12S contém valores não binários: {valores_unicos}")

    # =======================================================
    # 6) Criar listas de arestas para cada rede
    # =======================================================
    edges_18s = []  # Lista para as conexões (arestas) da rede 18S
    edges_12s = []  # Lista para as conexões (arestas) da rede 12S

    # =======================================================
    # 7) Processar a rede 18S (Indivíduos → Infecções)
    # =======================================================
    for index, row in data_18s.iterrows():
        individual = row.iloc[0]  # Nome do indivíduo (primeira coluna normalizada)
        for infection in data_18s.columns[1:]:  # Colunas das infecções
            if row[infection] == 1:  # Verifica se há infecção presente
                edges_18s.append((individual, infection))  # Adiciona uma aresta

    # =======================================================
    # 8) Processar a rede 12S (Indivíduos → Fontes Alimentares)
    # =======================================================
    for index, row in data_12s.iterrows():
        individual = row.iloc[0]  # Nome do indivíduo (primeira coluna normalizada)
        for source in data_12s.columns[1:]:  # Colunas das fontes alimentares
            if row[source] == 1:  # Verifica se há fonte alimentar presente
                edges_12s.append((individual, source))

    # =======================================================
    # 9) Criar o grafo combinado
    # =======================================================
    G_combined = nx.Graph()  # Grafo não direcionado

    # Adicionar as arestas da rede 18S
    for edge in edges_18s:
        G_combined.add_edge(edge[0], edge[1], source='18S')

    # Adicionar as arestas da rede 12S
    for edge in edges_12s:
        G_combined.add_edge(edge[0], edge[1], source='12S')

    # =======================================================
    # 10) Detectar comunidades (duas abordagens para comparação)
    # =======================================================
    # Greedy Modularity
    communities_greedy = list(greedy_modularity_communities(G_combined))
    print(f"Comunidades detectadas (Greedy): {len(communities_greedy)}")

    # Label Propagation (exemplo adicional)
    communities_label_propagation = list(label_propagation_communities(G_combined))
    print(f"Comunidades detectadas (Label Propagation): {len(communities_label_propagation)}")

    # Escolher o método de comunidades principal
    communities = communities_greedy

    # =======================================================
    # 11) Calcular métricas para cada nó
    # =======================================================
    node_metrics = []

    # Mapear cada nó para sua comunidade
    node_to_community = {}
    for i, community in enumerate(communities):
        for node in community:
            node_to_community[node] = f"Comunidade {i + 1}"

    # Calcular várias centralidades
    degree_centrality = nx.degree_centrality(G_combined)
    closeness_centrality = nx.closeness_centrality(G_combined)
    betweenness_centrality = nx.betweenness_centrality(G_combined)

    # Se o grafo for desconectado, a eigenvector pode dar erro.
    # Aqui assumimos que ele é conectável ou aceitamos o risco.
    eigenvector_centrality = nx.eigenvector_centrality(G_combined, max_iter=1000)

    for node in G_combined.nodes():
        # Verificar o tipo do nó
        if node in data_18s.iloc[:, 0].values or node in data_12s.iloc[:, 0].values:
            tipo = 'Indivíduo'
        elif node in data_18s.columns[1:]:
            tipo = 'Infecção'
        elif node in data_12s.columns[1:]:
            tipo = 'Fonte Alimentar'
        else:
            tipo = 'Desconhecido'

        node_metrics.append({
            'Nó': node,
            'Tipo': tipo,
            'Comunidade': node_to_community.get(node, 'Desconhecido'),
            'Grau': G_combined.degree[node],
            'Centralidade de Grau': degree_centrality[node],
            'Centralidade de Proximidade': closeness_centrality[node],
            'Centralidade de Intermediação': betweenness_centrality[node],
            'Centralidade de Autovetor (Eigenvector)': eigenvector_centrality[node]
        })

    node_metrics_df = pd.DataFrame(node_metrics)

    # =======================================================
    # 12) Calcular métricas para cada comunidade
    # =======================================================
    community_metrics = []

    for i, community in enumerate(communities):
        subgraph = G_combined.subgraph(community)
        sub_degree_centrality = nx.degree_centrality(subgraph)
        sub_closeness_centrality = nx.closeness_centrality(subgraph)
        sub_betweenness_centrality = nx.betweenness_centrality(subgraph)

        community_data = {
            'Comunidade': f'Comunidade {i + 1}',
            'Tamanho': len(community),
            'Grau Médio': sum(dict(subgraph.degree()).values()) / len(community),
            'Densidade': nx.density(subgraph),
            'Clustering Médio': nx.average_clustering(subgraph),
            'Centralidade Média (Degree)': sum(sub_degree_centrality.values()) / len(community),
            'Centralidade Média (Closeness)': sum(sub_closeness_centrality.values()) / len(community),
            'Centralidade Média (Betweenness)': sum(sub_betweenness_centrality.values()) / len(community),
        }

        community_metrics.append(community_data)

    community_metrics_df = pd.DataFrame(community_metrics)

    # =======================================================
    # 13) Criar DataFrame detalhado com indivíduos, infecções e fontes alimentares por comunidade
    # =======================================================
    detailed_community_data = []
    for i, community in enumerate(communities):
        for node in community:
            detailed_community_data.append({
                'Comunidade': f'Comunidade {i + 1}',
                'Nó': node,
                'Tipo': (
                    'Indivíduo' if node in data_18s.iloc[:, 0].values or node in data_12s.iloc[:, 0].values else
                    'Infecção' if node in data_18s.columns[1:] else
                    'Fonte Alimentar' if node in data_12s.columns[1:] else
                    'Desconhecido'
                )
            })

    detailed_community_df = pd.DataFrame(detailed_community_data)

    # =======================================================
    # 14) Salvar relatórios em um arquivo Excel
    # =======================================================
    output_excel_path = "C:/Users/lucas/Analise_redes/community_detailed_report.xlsx"
    with pd.ExcelWriter(output_excel_path) as writer:
        community_metrics_df.to_excel(writer, sheet_name="Community Metrics", index=False)
        detailed_community_df.to_excel(writer, sheet_name="Detailed Communities", index=False)
        node_metrics_df.to_excel(writer, sheet_name="Node Metrics", index=False)

    print(f"Relatório detalhado salvo em '{output_excel_path}'.")

    # =======================================================
    # 15) Visualizar a rede destacando as comunidades
    # =======================================================
    def plot_network_with_communities(graph, communities, layout='spring'):
        """
        Plota o grafo destacando as comunidades.
        :param graph: Grafo NetworkX.
        :param communities: Lista de conjuntos (ou listas) com os nós de cada comunidade.
        :param layout: Pode ser 'spring', 'kamada_kawai', etc.
        """
        # Escolher o layout
        if layout == 'spring':
            pos = nx.spring_layout(graph, seed=42)
        elif layout == 'kamada_kawai':
            pos = nx.kamada_kawai_layout(graph)
        elif layout == 'spectral':
            pos = nx.spectral_layout(graph)
        else:
            pos = nx.spring_layout(graph, seed=42)
            print(f"Layout '{layout}' não reconhecido. Usando spring_layout por padrão.")

        plt.figure(figsize=(15, 10))

        import matplotlib.colors as mcolors

        # Aqui acessamos o colormap 'hsv' sem usar o parâmetro extra (que gerou o erro)
        cmap = plt.colormaps['hsv']  # Acesso direto ao colormap

        # Geramos N cores uniformemente distribuídas, onde N = número de comunidades
        n = len(communities)
        colors = [mcolors.rgb2hex(cmap(i / n)) for i in range(n)]

        # Desenhar nós de cada comunidade
        for i, community in enumerate(communities):
            community_nodes = list(community)
            nx.draw_networkx_nodes(
                graph,
                pos,
                nodelist=community_nodes,
                node_color=colors[i],
                label=f"Comunidade {i + 1}",
                node_size=300
            )

        # Desenhar as arestas e rótulos
        nx.draw_networkx_edges(graph, pos, edge_color='gray', alpha=0.5)
        nx.draw_networkx_labels(graph, pos, font_size=8)

        plt.title("Rede Combinada com Comunidades", fontsize=16)
        plt.legend(loc="best", fontsize=10)
        plt.axis('off')  # Remover eixos, opcional
        plt.show()

    # Chamar a função de plot com o layout desejado
    plot_network_with_communities(G_combined, communities, layout='kamada_kawai')

    # Se quiser comparar outro layout, basta chamar novamente:
    # plot_network_with_communities(G_combined, communities, layout='spring')

finally:
    # Restaura o stdout original e fecha o arquivo de log
    sys.stdout = original_stdout
    log_file.close()

# ========================= FIM DO CÓDIGO ========================= #
