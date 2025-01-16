# Teias Vivas

Script de análise de redes que combina dados de planilhas Excel, 
detecta comunidades e gera relatórios e visualizações com Python.

## Descrição Geral
Este repositório contém um exemplo de script em Python para análise de redes, 
que utiliza bibliotecas como **pandas**, **networkx** e **matplotlib** para:

1. Ler dados de **planilhas Excel** (planilhas “18S” e “12S”),  
2. Criar um **grafo combinado** (indivíduos conectados a infecções ou fontes alimentares),  
3. **Detectar comunidades** na rede (utilizando algoritmos de modularidade e propagação de rótulos),  
4. Calcular diferentes **métricas de centralidade** (grau, proximidade, intermediação, autovetor),  
5. Gerar **relatórios** em Excel e **visualizar** a rede com as comunidades coloridas.

## Estrutura do Projeto
├── analise_redes.py # Script principal com todo o fluxo de análise de redes

├── Redes_Mestrado.xlsx # Exemplo de arquivo Excel com planilhas '18S' e '12S'

├── logs.txt # Arquivo de log gerado, redirecionando prints do script

└── (demais arquivos/pastas) # Outras dependências

## Dependências

```bash
# Para instalar as bibliotecas necessárias, use:
pip install pandas networkx matplotlib openpyxl

# 1) Clone ou baixe este repositório:
git clone https://github.com/seu-usuario/teias-vivas.git
cd teias-vivas

# 2) Coloque o arquivo 'Redes_Mestrado.xlsx' (ou outro Excel no mesmo formato)
#    e verifique se o script 'analise_redes.py' aponta para o caminho correto.

# 3) Rode o script principal:
python analise_redes.py

# 4) A saída de logs será registrada em 'logs.txt',
#    e um arquivo Excel de relatórios será gerado em:
#    'community_detailed_report.xlsx'.

# 5) Durante a execução, uma janela pode aparecer exibindo o grafo,
#    destacando as comunidades (caso esteja em ambiente gráfico).
```
## Explicações Rápidas do Script

### Leitura das Planilhas
O script carrega as planilhas `18S` e `12S` do Excel, assegurando que cada coluna (exceto a primeira) seja binária (0 ou 1).

### Construção do Grafo
Cada indivíduo pode ter uma ou mais infecções (18S) e/ou fontes alimentares (12S). Essas relações são convertidas em arestas no grafo.

### Comunidades
Usa-se o algoritmo “Greedy Modularity” e o “Label Propagation” para agrupar nós, imprimindo quantas comunidades cada método detectou.

### Métricas
Calculam-se diversas centralidades para cada nó (grau, proximidade, intermediação e autovetor) e estatísticas para cada comunidade (tamanho, densidade, grau médio, etc.).

### Relatórios
É gerado um arquivo Excel contendo três planilhas:
- **Community Metrics**: estatísticas por comunidade;
- **Detailed Communities**: lista de nós (indivíduos/infecções/fontes) em cada comunidade;
- **Node Metrics**: métricas de cada nó.

### Plot
É exibido um gráfico com cores diferentes para cada comunidade, usando o layout escolhido (`spring` ou `kamada_kawai`).

### Observações
- Certifique-se de manter a mesma estrutura de colunas e planilhas (“18S” e “12S”) caso use dados próprios.
- Para layouts e algoritmos adicionais, consulte a [documentação do NetworkX](https://networkx.org/documentation/stable/) e explore parâmetros como `seed` e a função `nx.draw_*` para customizações.
