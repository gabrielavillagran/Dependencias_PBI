# Semantic Model Insights

Ferramenta de análise estática para modelos semânticos do Power BI. Processa arquivos TMDL (Tabular Model Definition Language) e estrutura de relatórios para gerar análises de dependências, complexidade e uso de medidas DAX.

---

## O que faz

Analisa projetos Power BI (`.pbip`) para responder perguntas comuns durante manutenção e refatoração:

- Quais medidas dependem da medida X?
- Quais medidas não estão sendo usadas em nenhum visual?
- Qual a complexidade relativa das medidas DAX?
- Quais páginas concentram maior número de medidas?

## Principais funcionalidades

### Análise de Complexidade

Calcula um score de 0-100 para cada medida DAX baseado em 5 dimensões:
- **D1 - Funções de iteração**: SUMX, RANKX, FILTER, etc.
- **D2 - Manipulação de contexto**: CALCULATE, ALL, KEEPFILTERS
- **D3 - Estrutura de código**: Linhas, uso de VAR, comentários
- **D4 - Número de dependentes**: Quantas outras medidas dependem dela
- **D5 - Anti-patterns**: Práticas que prejudicam performance

### Detecção de medidas órfãs

Identifica medidas que:
1. Não são referenciadas por nenhuma outra medida (análise de grafo DAX)
2. Não aparecem em nenhum visual do relatório (análise de estrutura)

Essas medidas são candidatas seguras para remoção.

### Grafo de dependências

Visualização interativa das relações entre medidas, colunas e tabelas. Permite explorar:
- **Dependências**: O que a medida usa (antecedentes)
- **Dependentes**: O que usa a medida (impacto de mudanças)

### Análise por página

Mostra distribuição de medidas por página do relatório, incluindo:
- Total de medidas únicas por página
- Complexidade média das medidas
- Medida mais complexa
- Medida mais reutilizada

## Tecnologias

- **Streamlit**: Interface web
- **NetworkX**: Construção e análise de grafos de dependência
- **Plotly**: Visualizações interativas
- **PyVis**: Renderização de grafos
- **openpyxl**: Geração de relatórios Excel

## Como usar

1. Compacte a pasta do seu projeto Power BI (`.pbip`) em um arquivo ZIP
2. Faça upload na interface do Streamlit
3. A ferramenta processa automaticamente:
   - Arquivos TMDL do modelo semântico
   - Estrutura de páginas e visuais do relatório

## Requisitos

Ver `requirements.txt` para dependências Python.

## Estrutura do projeto

```
├── app.py                          # Aplicação principal Streamlit
├── requirements.txt                # Dependências Python
└── README.md                       # Este arquivo
```

## Limitações conhecidas

- Requer projetos no formato `.pbip` (não funciona com `.pbix`)
- Análise de complexidade é baseada em heurísticas, não em telemetria real do engine
- Não detecta medidas usadas em RLS (Row Level Security) ou parâmetrosWhat-If
