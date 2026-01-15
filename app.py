import streamlit as st
import pandas as pd
import networkx as nx
from pyvis.network import Network
import tempfile
import os
import json
from streamlit.components.v1 import html
import io

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(layout="wide", page_title="Grafo de Dependências PBI")

# --- CSS PERSONALIZADO COM ANIMAÇÕES (MELHORIA 26) ---
st.markdown("""
    <style>
        /* Sidebar responsiva - Removido CSS fixo para evitar quebra de layout nativo */
        
        /* Animações suaves nos cards */
        .stMetric { 
            background-color: #f8f9fb; 
            padding: 10px; 
            border-radius: 10px; 
            border: 1px solid #e6e9ef;
            transition: all 0.3s ease;
        }
        .stMetric:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            border-color: #5E9AE9;
        }
        
        /* Animação de fade-in */
        .element-container {
            animation: fadeIn 0.5s ease-in;
        }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        /* Efeito hover nos botões */
        .stDownloadButton > button {
            transition: all 0.3s ease;
        }
        .stDownloadButton > button:hover {
            transform: scale(1.05);
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        }
        
        /* Zoom da página para 70% */
        .main .block-container {
            zoom: 0.7;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Grafo de Dependências — Power BI")

# --- FUNÇÕES AUXILIARES ---
def limpar_dax(texto):
    if pd.isnull(texto) or texto == "None":
        return ""
    return str(texto).replace("_x000D_", "").strip()

@st.cache_data(ttl=3600)
def processar_dependencias(df_bytes, tipos_selecionados):
    """Processa dependências com cache para melhor performance"""
    df = pd.read_csv(io.BytesIO(df_bytes), sep=None, engine='python', encoding='utf-8-sig')
    df.columns = df.columns.str.strip()
    return df

@st.cache_data
def construir_grafo(arestas_tuple):
    """Constrói grafo NetworkX com cache"""
    G = nx.DiGraph()
    G.add_edges_from(arestas_tuple)
    return G

def calcular_health_score(medidas_orfas, total_medidas, complexidade_media):
    """Calcula health score do modelo - Versão Simplificada e Intuitiva"""
    # Começa com 100
    score = 100
    
    # Deduz por órfãs SE FOR MUITAS (>40%)
    taxa_orfas = len(medidas_orfas) / max(total_medidas, 1)
    if taxa_orfas > 0.4:
        # Penaliza APENAS o excesso acima de 40%
        score -= (taxa_orfas - 0.4) * 50
    
    # Deduz por complexidade extrema
    if complexidade_media > 7:
        # Muito complexo = difícil manter
        score -= (complexidade_media - 7) * 5
    elif complexidade_media < 2:
        # Muito simples = subutilizado
        score -= (2 - complexidade_media) * 5
    
    return round(max(0, min(100, score)), 1)


def calcular_complexidade_dax(expressao):
    """Calcula score de complexidade DAX (função antiga - mantida para compatibilidade)"""
    import re
    if not expressao or expressao == "":
        return 0
    # Conta funções DAX (em maiúsculas seguidas de parênteses)
    funcoes = len(re.findall(r'\b[A-Z]{2,}\(', expressao))
    # Conta aninhamentos
    aninhamentos = expressao.count('(')
    # Complexidade = funções + aninhamentos / 10
    return min(10, (funcoes + aninhamentos / 10))

def calcular_complexity_score(expressao, nome_medida="", medidas_dependentes=0):
    """
    Calcula Complexity  Score (0-100) com 5 dimensões baseado em SQLBI + Microsoft Learn.
    
    Dimensões:
    - D1: Funções (SUMX, RANKX, FILTER, etc)
    - D2: CALCULATE e contexto
    - D3: Estrutura (linhas, VAR, comentários)
    - D4: Dependências
    - D5: Anti-patterns
    
    Returns:
        (score, classificacao, detalhes)
    """
    import re
    
    if not expressao or expressao == "":
        return 0, "🟢 Simples", []
    
    score = 0
    detalhes = []
    
    # === D1: FUNÇÕES (Peso Alto) ===
    funcoes_peso = {
        'SUMX': 8, 'AVERAGEX': 8, 'MINX': 8, 'MAXX': 8,
        'RANKX': 12,
        'FILTER': 10,
        'ADDCOLUMNS': 10,
        'SUMMARIZE': 12, 'SUMMARIZECOLUMNS': 12,
        'GENERATE': 15,
        'EARLIER': 20,
        'PATH': 8, 'CONTAINSROW': 8
    }
    
    for func, penalty in funcoes_peso.items():
        count = expressao.upper().count(func)
        if count > 0:
            score += count * penalty
            detalhes.append(f"D1: {func} ({count}x) = +{count * penalty}")
    
    # === D2: CALCULATE E CONTEXTO ===
    calculate_count = expressao.upper().count('CALCULATE')
    if calculate_count > 0:
        score += calculate_count * 5
        detalhes.append(f"D2: CALCULATE ({calculate_count}x) = +{calculate_count * 5}")
    
    # Múltiplos filtros em CALCULATE
    calc_pattern = r'CALCULATE\s*\([^,]+,([^)]+)\)'
    for match in re.findall(calc_pattern, expressao, re.IGNORECASE):
        filters = match.count(',') + 1
        if filters > 1:
            penalty = (filters - 1) * 3
            score += penalty
            detalhes.append(f"D2: CALCULATE c/ {filters} filtros = +{penalty}")
    
    # ALL, ALLEXCEPT, REMOVEFILTERS
    context_funcs = {'ALL': 6, 'ALLEXCEPT': 6, 'REMOVEFILTERS': 6, 'KEEPFILTERS': 3}
    for func, penalty in context_funcs.items():
        count = expressao.upper().count(func)
        if count > 0:
            score += count * penalty
            detalhes.append(f"D2: {func} ({count}x) = +{count * penalty}")
    
    # === D3: ESTRUTURA ===
    linhas = expressao.count('\n') + 1
    if linhas > 20:
        score += 10
        detalhes.append(f"D3: >20 linhas ({linhas}) = +10")
    elif linhas > 10:
        score += 5
        detalhes.append(f"D3: >10 linhas ({linhas}) = +5")
    
    # Bônus: VAR
    if 'VAR' in expressao.upper():
        var_count = expressao.upper().count('VAR')
        bonus = var_count * 5
        score -= bonus
        detalhes.append(f"D3: VAR ({var_count}x) = -{bonus} (bônus)")
    
    # Bônus: Comentários
    comentarios = expressao.count('--') + expressao.count('//')
    if comentarios > 0:
        bonus = min(comentarios * 2, 10)
        score -= bonus
        detalhes.append(f"D3: Comentários ({comentarios}) = -{bonus} (bônus)")
    
    # === D4: DEPENDÊNCIAS ===
    if medidas_dependentes > 0:
        penalty = medidas_dependentes * 2
        score += penalty
        detalhes.append(f"D4: {medidas_dependentes} dependentes = +{penalty}")
    
    # === D5: ANTI-PATTERNS ===
    if re.search(r'FILTER\s*\(\s*ALL\s*\(', expressao, re.IGNORECASE):
        score += 20
        detalhes.append("D5: FILTER(ALL(Tabela)) = +20")
    
    if 'DATE' in expressao.upper() and not any(x in expressao.upper() for x in ['SAMEPERIODLASTYEAR', 'DATESYTD', 'TOTALYTD', 'DATEADD']):
        score += 8
        detalhes.append("D5: Time intelligence manual = +8")
    
    # === CLASSIFICAÇÃO ===
    final_score = min(100, max(0, score))
    
    if final_score <= 20:
        classificacao = "🟢 Simples"
    elif final_score <= 40:
        classificacao = "🟡 Moderada"
    elif final_score <= 60:
        classificacao = "🟠 Complexa"
    elif final_score <= 80:
        classificacao = "🔴 Muito Complexa"
    else:
        classificacao = "⚫ Crítica"
    
    return final_score, classificacao, detalhes


def gerar_relatorio_texto(metricas, medidas_orfas, medidas_impacto, top_complexas=None):
    """Gera relatório em texto para download (MELHORIA 24)"""
    
    # Formatar lista de medidas complexas
    secao_complexidade = ""
    if top_complexas:
        lista_formatada = chr(10).join(f"  • {m['medida']} (Score: {m['score']} - {m['classificacao']})" for m in top_complexas[:10])
        secao_complexidade = f"""
🔥 TOP 10 MEDIDAS MAIS COMPLEXAS (CRÍTICAS)
────────────────────────────────────────────────────────────
{lista_formatada}
"""

    relatorio = f"""═══════════════════════════════════════════════════════════
            RELATÓRIO DE DEPENDÊNCIAS DAX - POWER BI
═══════════════════════════════════════════════════════════

📊 MÉTRICAS GERAIS
────────────────────────────────────────────────────────────
Objetos no Modelo: {metricas.get('objetos', 0)}
Nós no Grafo: {metricas.get('nos', 0)}
Relacionamentos: {metricas.get('relacionamentos', 0)}
Medidas Órfãs: {metricas.get('orfas', 0)}
Impacto Total: {metricas.get('impacto', 0)}
{secao_complexidade}
⚠️ MEDIDAS ÓRFÃS ({len(medidas_orfas)})
────────────────────────────────────────────────────────────
O que são Medidas Órfãs?
Medidas órfãs são medidas que NÃO são referenciadas por nenhuma outra 
medida no modelo. Isso pode significar:
  • Medidas finais usadas diretamente em visuais (normal)
  • Medidas obsoletas que podem ser removidas (limpeza recomendada)
  • Oportunidades de refatoração

Lista de Medidas Órfãs:
{chr(10).join(f"  • {m}" for m in sorted(list(medidas_orfas))) if medidas_orfas else "  Nenhuma medida órfã detectada!"}

📊 ANÁLISE DE IMPACTO
────────────────────────────────────────────────────────────
{chr(10).join(f"  • {m['medida']}: {m['impacto']} objetos dependentes" for m in medidas_impacto) if medidas_impacto else "  Nenhuma medida selecionada"}

═══════════════════════════════════════════════════════════
Relatório gerado automaticamente
═══════════════════════════════════════════════════════════
"""
    return relatorio

# --- 1. SESSÃO DE INSTRUÇÕES E DOWNLOAD ---
with st.expander("📖 Como gerar o arquivo de dependências?", expanded=False):
    st.markdown("""
    ### Passo a Passo:
    1. Abra o seu relatório no **Power BI Desktop**.
    2. Vá até a aba **Exibição** e selecione **Visualização de consulta DAX**.
    3. Copie o código abaixo, cole na janela de consulta e clique em **Executar**.
    """)

    dax_query = """EVALUATE
VAR Medidas = INFO.MEASURES()
VAR Dependencias = INFO.CALCDEPENDENCY()
RETURN
SELECTCOLUMNS(
    FILTER(Dependencias, [OBJECT_TYPE] = "MEASURE"),
    "Tipo Origem", [REFERENCED_OBJECT_TYPE],
    "Origem", [REFERENCED_OBJECT],
    "Expressão Origem", IF(
        [REFERENCED_OBJECT_TYPE] = "MEASURE",
        MAXX(
            FILTER(Medidas, [Name] = [REFERENCED_OBJECT]),
            [Expression]
        ),
        BLANK()
    ),
    "Tipo Destino", [OBJECT_TYPE],
    "Destino", [OBJECT],
    "Expressão Destino", 
        MAXX(
            FILTER(Medidas, [Name] = [OBJECT]),
            [Expression]
        )
)"""

    st.code(dax_query, language="sql")

    st.markdown("""
    5. Após executar a consulta, clique em **Copiar** com a opção **Células selecionadas** e, em seguida, cole os dados no arquivo CSV disponibilizado para download..
    """)

    # Imagem ilustrativa do passo 4
    st.image(
        "img/celulas_selecionadas.png",
        caption="Selecione todas as células do resultado da consulta DAX e copie",
        width=500
    )

    st.markdown("""
    6. Os dados devem estar da seguinte forma """)

    st.image(
        "img/modelo_dados_inseridos.png",
        caption="Selecione todas as células do resultado da consulta DAX e copie",
        width=500
    )


    # --- CORREÇÃO: GERAR CSV COM BOM PARA ACENTOS ---
    buffer_csv = io.BytesIO()
    # Adiciona o BOM do UTF-8 para o Excel reconhecer o til (~) e acentos
    buffer_csv.write('\ufeff'.encode('utf-8'))
    
    df_template = pd.DataFrame(columns=[
        "[Tipo Origem]", "[Origem]", "[Expressão Origem]", 
        "[Tipo Destino]", "[Destino]", "[Expressão Destino]"
    ])
    
    df_template.to_csv(buffer_csv, index=False, sep=";", encoding='utf-8', mode='ab')
    
    st.download_button(
        label="📥 Baixar Modelo CSV para Preencher",
        data=buffer_csv.getvalue(),
        file_name="modelo_dependencias.csv",
        mime="text/csv"
    )

# --- 2. CARREGAMENTO ---
uploaded_file = st.file_uploader("Envie o arquivo preenchido", type=["csv", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith('.xlsx'):
        df = pd.read_excel(uploaded_file)
    else:
        # --- CORREÇÃO: LER CSV COM UTF-8-SIG PARA ACENTOS ---
        df = pd.read_csv(uploaded_file, sep=None, engine='python', encoding='utf-8-sig')

    df.columns = df.columns.str.strip()
    
    col_origem, col_destino = "[Origem]", "[Destino]"
    col_tipo_origem, col_exp_origem, col_exp_destino = "[Tipo Origem]", "[Expressão Origem]", "[Expressão Destino]"

    if col_origem in df.columns and col_destino in df.columns:
        df[col_origem] = df[col_origem].astype(str).replace('nan', None)
        df[col_destino] = df[col_destino].astype(str).replace('nan', None)
        df = df.dropna(subset=[col_origem, col_destino])

        # --- 3. FILTROS ---
        st.sidebar.header("⚙️ Filtros do Grafo")
        tipos_disponiveis = sorted(df[col_tipo_origem].unique().astype(str).tolist())
        selecionar_todos = st.sidebar.checkbox("Selecionar todos os tipos", value=True)
        
        if selecionar_todos:
            tipos_selecionados = st.sidebar.multiselect("Filtrar Origens por Tipo:", options=tipos_disponiveis, default=tipos_disponiveis)
        else:
            tipos_selecionados = st.sidebar.multiselect("Filtrar Origens por Tipo:", options=tipos_disponiveis, default=[])

        df_filtrado = df[df[col_tipo_origem].isin(tipos_selecionados)]
        todas_destinos = sorted([str(m) for m in df_filtrado[col_destino].unique()])
        
        # 🔍 MELHORIA 1: BUSCA DE MEDIDAS
        st.sidebar.markdown("---")
        busca_medida = st.sidebar.text_input("🔍 Buscar Medida:", "", placeholder="Digite para filtrar...")
        
        # Filtrar medidas baseado na busca
        if busca_medida:
            medidas_filtradas = [m for m in todas_destinos if busca_medida.lower() in m.lower()]
            st.sidebar.caption(f"📊 {len(medidas_filtradas)} medidas encontradas")
        else:
            medidas_filtradas = todas_destinos
        
        medidas_selecionadas = st.sidebar.multiselect(
            "Selecione as Medidas Destino (Raízes):", 
            options=medidas_filtradas, 
            default=[]
        )
        
        # 🔄 DIREÇÃO DO GRAFO
        st.sidebar.markdown("---")
        st.sidebar.subheader("🔄 Direção do Grafo")
        direcao_grafo = st.sidebar.radio(
            "Escolha o que visualizar:",
            options=[
                "⬇️ Dependências (do que a medida depende)",
                "⬆️ Dependentes (quem depende da medida)"
            ],
            index=0,
            help="⬇️ Dependências: mostra as colunas, tabelas e medidas que a raiz usa\n"
                 "⬆️ Dependentes: mostra quais outras medidas dependem da raiz"
        )

        # Mapeamento de Info Limpo
        info_map = {}
        for _, row in df.iterrows():
            dest = str(row[col_destino])
            orig = str(row[col_origem])
            info_map[dest] = {"exp": limpar_dax(row[col_exp_destino]), "tipo": "MEASURE"}
            if orig not in info_map or not info_map[orig]["exp"]:
                info_map[orig] = {"exp": limpar_dax(row[col_exp_origem]), "tipo": str(row[col_tipo_origem])}

        if medidas_selecionadas:
            arestas, visitados, fila = [], set(), list(medidas_selecionadas)
            
            # Determinar modo de navegação baseado na escolha do usuário
            modo_dependencias = "⬇️" in direcao_grafo or "🔄" in direcao_grafo
            modo_dependentes = "⬆️" in direcao_grafo or "🔄" in direcao_grafo
            
            while fila:
                atual = fila.pop(0)
                if atual not in visitados:
                    visitados.add(atual)
                    
                    # MODO 1: Dependências (do que depende) - direção original
                    if modo_dependencias:
                        filhos = df_filtrado[df_filtrado[col_destino] == atual][col_origem].tolist()
                        for filho in filhos:
                            arestas.append((atual, filho))
                            if filho not in visitados: fila.append(filho)
                    
                    # MODO 2: Dependentes (quem depende) - direção reversa
                    if modo_dependentes:
                        pais = df_filtrado[df_filtrado[col_origem] == atual][col_destino].tolist()
                        for pai in pais:
                            arestas.append((pai, atual))
                            if pai not in visitados: fila.append(pai)
            
            G = nx.DiGraph()
            G.add_edges_from(arestas)

            # ⚠️ MELHORIA 2: DETECÇÃO DE MEDIDAS ÓRFÃS
            origens_unicas = set(df_filtrado[col_origem].unique())
            destinos_unicos = set(df_filtrado[col_destino].unique())
            medidas_orfas = destinos_unicos - origens_unicas
            
            # 📊 MELHORIA 5: ANÁLISE DE IMPACTO
            total_impacto = 0
            if medidas_selecionadas:
                for medida in medidas_selecionadas:
                    if medida in G:
                        descendentes = nx.descendants(G, medida)
                        total_impacto += len(descendentes)
            
            
            # --- 4. MÉTRICAS APRIMORADAS ---
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Objetos no Modelo", df[col_origem].nunique())
            c2.metric("Nós no Grafo", len(G.nodes()))
            c3.metric("Relacionamentos", len(arestas))
            c4.metric(
                "Medidas Órfãs", 
                len(medidas_orfas),
                help="🔍 MEDIDAS ÓRFÃS são medidas que não são referenciadas por nenhuma outra medida no modelo. Elas podem indicar: (1) medidas finais usadas diretamente em visuais, (2) medidas obsoletas que podem ser removidas, ou (3) oportunidades de refatoração. Um número alto sugere revisão de limpeza."
            )
            
            # === CÁLCULO DE COMPLEXIDADE DAX ===


            # Calcular score de complexidade apenas das medidas selecionadas
            medidas_selecionadas_complexas = []
            for nome_medida in medidas_selecionadas:
                info = info_map.get(nome_medida, {})
                if info.get("tipo") == "MEASURE":
                    exp = info.get("exp", "")
                    n_dependentes = df[col_destino].value_counts().to_dict().get(nome_medida, 0)
                    score, classificacao, _ = calcular_complexity_score(exp, nome_medida, n_dependentes)
                    medidas_selecionadas_complexas.append({'medida': nome_medida, 'score': score, 'classificacao': classificacao})

            scores_complexidade = [m['score'] for m in medidas_selecionadas_complexas]
            score_medio = round(sum(scores_complexidade) / len(scores_complexidade), 1) if scores_complexidade else 0

            # Classificação do modelo (das selecionadas)
            if score_medio <= 20:
                class_modelo = "🟢 Simples"
            elif score_medio <= 40:
                class_modelo = "🟡 Moderada"
            elif score_medio <= 60:
                class_modelo = "🟠 Complexa"
            elif score_medio <= 80:
                class_modelo = "🔴 Muito Complexa"
            else:
                class_modelo = "⚫ Crítica"
            
            c5.metric(
                "📊 Complexidade DAX",
                f"{score_medio}/100",
                delta=class_modelo,
                help="Score de 0 a 100 baseado em 5 dimensões (Funções, Contexto, Estrutura, Dependências, Anti-patterns).\n\n👇 Abra a seção 'ℹ️ Entenda o Cálculo' abaixo para ver a tabela de regras completa."
            )
            
            # ℹ️ TABELA DE REGRAS DE COMPLEXIDADE (Expansível)
            with st.expander("ℹ️ Entenda o Cálculo da Complexidade (Tabela de Regras)", expanded=False):
                st.markdown("""
                ### 📊 Como o Score é calculado?
                O score começa em **0** e acumula pontos de penalidade. Quanto menor, melhor.
                O KPI final é normalizado para **0-100**.
                
                #### 1️⃣ D1: Funções DAX (Peso Alto)
                | Função | Penalidade | Motivo |
                |---|---|---|
                | `SUMX`, `AVERAGEX`, `MINX`, `MAXX` | **+8** pts | Iterador (força Formula Engine) |
                | `RANKX` | **+12** pts | Custo computacional alto |
                | `FILTER` | **+10** pts | Iteração muitas vezes desnecessária |
                | `ADDCOLUMNS` | **+10** pts | Materialização temporária |
                | `SUMMARIZE`, `SUMMARIZECOLUMNS` | **+12** pts | Complexo para otimizar |
                | `GENERATE` | **+15** pts | Cross join custoso |
                | `EARLIER` | **+20** pts | Difícil leitura/manutenção |
                
                #### 2️⃣ D2: Contexto e CALCULATE
                | Regra | Penalidade |
                |---|---|
                | Cada `CALCULATE` | **+5** pts |
                | `CALCULATE` com múltiplos filtros | **+3** pts por filtro extra |
                | `ALL`, `ALLEXCEPT`, `REMOVEFILTERS` | **+6** pts |
                | `KEEPFILTERS` | **+3** pts |
                
                #### 3️⃣ D3: Estrutura do Código
                | Métrica | Pontos |
                |---|---|
                | Mais de 10 linhas | **+5** pts |
                | Mais de 20 linhas | **+10** pts |
                | Uso de Variáveis (`VAR`) | **-5** pts (BÔNUS 🟢) |
                | Comentários (`--` ou `//`) | **-2** pts (BÔNUS 🟢) |
                
                #### 4️⃣ D4: Dependências
                | Item | Penalidade |
                |---|---|
                | Por medida dependente | **+2** pts |
                
                #### 5️⃣ D5: Anti-patterns (Erros Comuns)
                | Anti-pattern | Penalidade |
                |---|---|
                | `FILTER(ALL(Tabela))` | **+20** pts (Muito ineficiente) |
                | Time Intelligence Manual | **+8** pts (Use funções nativas) |
                """)
        

            # --- 6. GERAÇÃO DO GRAFO COM ÍCONES (MELHORIA 27) ---
            # Ícones por tipo de objeto
            cores = {"MEASURE": "#88B995", "COLUMN": "#5E9AE9", "TABLE": "#F4A460", "UNKNOWN": "#CCCCCC"}
            icones = {"MEASURE": "📊", "COLUMN": "📋", "TABLE": "📁", "UNKNOWN": "❓"}
            
            net = Network(height="600px", width="100%", directed=True, bgcolor="#ffffff")
            for node in G.nodes():
                tipo = info_map.get(node, {}).get("tipo", "UNKNOWN")
                icone = icones.get(tipo, icones["UNKNOWN"])
                # Label com ícone
                label_com_icone = f"{icone} {node}"
                net.add_node(
                    node, 
                    label=label_com_icone, 
                    title=f"Clique para ver o DAX\nTipo: {tipo}", 
                    color=cores.get(tipo, cores["UNKNOWN"]), 
                    shape="box", 
                    margin=10, 
                    font={"face": "Segoe UI", "size": 14}
                )
            for u, v in G.edges():
                net.add_edge(u, v, color="#CCCCCC", width=1)

            net.set_options(json.dumps({
                "nodes": {"shadow": True},
                "layout": {"hierarchical": {"enabled": True, "direction": "UD", "sortMethod": "directed", "levelSeparation": 150, "nodeSpacing": 200}},
                "physics": {"enabled": False},
                "interaction": {"hover": True}
            }))

            path = os.path.join(tempfile.gettempdir(), "graph_pbi.html")
            net.save_graph(path)
            with open(path, 'r', encoding='utf-8') as f:
                html_content = f.read()

            info_json = json.dumps(info_map)
            custom_js = f"""
            <div id="dax-panel" style="position:fixed; top:20px; right:20px; width:450px; max-height:80vh; background:#ffffff; color:#31333f; border-radius:12px; padding:20px; overflow-y:auto; z-index:99999; display:none; box-shadow: 0 4px 16px rgba(0,0,0,0.15); font-family: 'Source Sans Pro', sans-serif; border: 1px solid #e6e9ef;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px; border-bottom: 1px solid #e6e9ef; padding-bottom:10px;">
                    <div>
                        <div id="panel-title" style="font-size:16px; font-weight:bold; color:#1f77b4; margin-bottom:2px;">Objeto</div>
                        <div id="panel-type" style="font-size:11px; color:#7d7d7d; text-transform:uppercase; font-weight: 600;">TIPO</div>
                    </div>
                    <button onclick="document.getElementById('dax-panel').style.display='none'" style="background:none; border:none; color:#999; cursor:pointer; font-size:24px; line-height:1; padding:0 5px;">&times;</button>
                </div>
                <pre id="panel-exp" style="white-space: pre-wrap; word-wrap: break-word; font-size: 14px; line-height: 1.6; color: #000000; margin: 0; background-color: #f0f2f6; padding: 16px; border-radius: 8px; font-family: 'Source Code Pro', monospace;"></pre>
            </div>
            <script>
                var infoData = {info_json};
                network.on("click", function (params) {{
                    if (params.nodes.length > 0) {{
                        var id = params.nodes[0];
                        var item = infoData[id];
                        document.getElementById('panel-title').innerText = id;
                        document.getElementById('panel-type').innerText = item.tipo;
                        document.getElementById('panel-exp').innerText = (item.exp && item.exp !== 'None') ? item.exp : "Sem expressão DAX disponível.";
                        document.getElementById('dax-panel').style.display = 'block';
                    }}
                }});
            </script>
            """
            full_html = html_content.replace("</body>", f"{custom_js}</body>")
            
            # 📸 MELHORIA 3 & 24: EXPORTAÇÃO HTML E RELATÓRIO
            # Gerar relatório de texto
            metricas_relatorio = {
                'objetos': df[col_origem].nunique(),
                'nos': len(G.nodes()),
                'relacionamentos': len(arestas),
                'orfas': len(medidas_orfas),
                'impacto': total_impacto
            }
            
            # Gerar lista de impacto
            medidas_impacto_lista = []
            for medida in medidas_selecionadas:
                if medida in G:
                    desc = nx.descendants(G, medida)
                    medidas_impacto_lista.append({'medida': medida, 'impacto': len(desc)})
            
            # Preparar top complexas se disponível
            top_complexas_export = sorted(todas_medidas_complexas, key=lambda x: x['score'], reverse=True) if 'todas_medidas_complexas' in locals() and todas_medidas_complexas else []

            relatorio_txt = gerar_relatorio_texto(metricas_relatorio, medidas_orfas, medidas_impacto_lista, top_complexas_export)
            
            # Botões de exportação acima do grafo
            col_export1, col_export2, col_spacer = st.columns([1, 1, 6])
            
            with col_export1:
                st.download_button(
                    label="📸 Exportar HTML",
                    data=full_html,
                    file_name="grafo_dependencias.html",
                    mime="text/html",
                    help="Baixe o grafo como arquivo HTML interativo",
                    use_container_width=True
                )
            
            with col_export2:
                st.download_button(
                    label="📄 Exportar Relatório",
                    data=relatorio_txt,
                    file_name="relatorio_dependencias.txt",
                    mime="text/plain",
                    help="Baixe relatório completo em formato texto",
                    use_container_width=True
                )
            
            # Legenda acima do grafo
            legenda_html = '<div style="display:flex;align-items:center;gap:20px;padding:12px 0;font-size:14px;"><span style="font-weight:600;margin-right:10px;">Legenda:</span>'
            for k, v in cores.items():
                icone = icones.get(k, "")
                legenda_html += f'<div style="display:inline-flex;align-items:center;gap:6px;"><div style="width:12px;height:12px;background:{v};border-radius:2px;"></div><span>{icone} {k}</span></div>'
            legenda_html += '</div>'
            st.markdown(legenda_html, unsafe_allow_html=True)
            
            # Grafo em largura total
            html(full_html, height=650)
            
            # 📊 ANÁLISE GLOBAL DO MODELO
            st.markdown("---")
            st.subheader("📊 Análise Global do Modelo")
            
            # Calcular dependentes globais (quantas vezes cada medida é usada)
            # Se A depende de B, então B aparece como destino de A.
            # Dependentes de X = Quantas vezes X aparece como destino.
            global_dependentes_count = df[col_destino].value_counts().to_dict()
            
            todas_medidas_complexas = []
            
            # Iterar sobre TODAS as medidas do modelo (info_map)
            for nome_medida, info in info_map.items():
                if info.get("tipo") == "MEASURE":
                    exp = info.get("exp", "")
                    # Pegar número de dependentes globalmente
                    n_dependentes = global_dependentes_count.get(nome_medida, 0)
                    
                    score, classificacao, _ = calcular_complexity_score(exp, nome_medida, n_dependentes)
                    todas_medidas_complexas.append({
                        'medida': nome_medida, 
                        'score': score, 
                        'classificacao': classificacao
                    })
            
            if todas_medidas_complexas:
                # Ordenar por score decrescente
                top_complexas_global = sorted(todas_medidas_complexas, key=lambda x: x['score'], reverse=True)
                
                # Criar DataFrame do ranking
                df_rank_global = pd.DataFrame(top_complexas_global)
                df_rank_global = df_rank_global[['medida', 'score', 'classificacao']]
                df_rank_global.columns = ['Medida', 'Score', 'Classificação']
                
                # Calcular dependentes por medida usando GRAFO GLOBAL (todas as relações do modelo, não filtradas)
                # Construir grafo global com todas as dependências do df_filtrado
                G_global = nx.DiGraph()
                G_global.add_edges_from([(row[col_destino], row[col_origem]) for _, row in df_filtrado.iterrows()])
                
                lista_dependentes = []
                for nome_medida in info_map:
                    if info_map[nome_medida].get("tipo") == "MEASURE":
                        # Contar todos os objetos que dependem desta medida (descendentes no grafo global)
                        if nome_medida in G_global:
                            descendentes = nx.descendants(G_global, nome_medida)
                            total_objetos = len(descendentes)
                            # Contar apenas medidas entre os descendentes
                            total_medidas = sum(1 for d in descendentes if info_map.get(d, {}).get("tipo") == "MEASURE")
                        else:
                            total_objetos = 0
                            total_medidas = 0
                        
                        lista_dependentes.append({
                            "Medida": nome_medida,
                            "Total de Objetos Dependentes": total_objetos,
                            "Medidas Dependentes": total_medidas
                        })

                df_dependentes = pd.DataFrame(lista_dependentes)
                df_dependentes = df_dependentes.sort_values(by=["Total de Objetos Dependentes", "Medidas Dependentes"], ascending=False)

                # Exibir as duas tabelas lado a lado
                col_ranking, col_dependentes = st.columns(2)
                
                with col_ranking:
                    st.markdown("##### 🏆 Ranking de Complexidade")
                    st.dataframe(
                        df_rank_global,
                        column_config={
                            "Medida": st.column_config.TextColumn("Medida", width="large"),
                            "Score": st.column_config.ProgressColumn(
                                "Score de Complexidade",
                                help="Score de 0 a 100",
                                format="%d",
                                min_value=0,
                                max_value=100,
                            ),
                            "Classificação": st.column_config.TextColumn("Classificação", width="medium"),
                        },
                        hide_index=True,
                        use_container_width=True,
                        height=400
                    )
                
                with col_dependentes:
                    st.markdown("##### 🔗 Medidas com Mais Dependentes")
                    st.dataframe(
                        df_dependentes,
                        column_config={
                            "Medida": st.column_config.TextColumn("Medida", width="large"),
                            "Total de Objetos Dependentes": st.column_config.NumberColumn("Total de Objetos Dependentes"),
                            "Medidas Dependentes": st.column_config.NumberColumn("Medidas Dependentes"),
                        },
                        hide_index=True,
                        use_container_width=True,
                        height=400
                    )
            
            # 📊 ANÁLISE DE IMPACTO DETALHADA
            if medidas_selecionadas:
                st.markdown("---")
                st.subheader("📊 Análise de Impacto por Medida")
                
                cols_impacto = st.columns(min(3, len(medidas_selecionadas)))
                for idx, medida in enumerate(medidas_selecionadas):
                    with cols_impacto[idx % 3]:
                        if medida in G:
                            descendentes = nx.descendants(G, medida)
                            with st.container():
                                st.metric(
                                    label=f"🎯 {medida}",
                                    value=f"{len(descendentes)} objetos",
                                    help=f"Alterar esta medida impactará {len(descendentes)} objetos dependentes"
                                )
                                if len(descendentes) > 0:
                                    with st.expander("Ver dependências"):
                                        st.write(sorted(list(descendentes)))


        else:
            st.info("Selecione pelo menos uma Medida Raiz na barra lateral.")
    else:
        st.error("Colunas [Origem] ou [Destino] não encontradas no arquivo.")
else:
    st.info("Aguardando upload do arquivo para gerar o grafo.")