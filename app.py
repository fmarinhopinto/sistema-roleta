import io
import os
import re
import unicodedata
from pathlib import Path
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Sistema Roleta - Controle Comercial",
    page_icon="🎯",
    layout="wide"
)

# --- DIRETÓRIOS LOCAIS ---
BASE_DIR = Path(__file__).resolve().parent
PASTA_ENTRADA = BASE_DIR / "arquivos_entrada"
PASTA_SAIDA = BASE_DIR / "planilhas_geradas"

PASTA_ENTRADA.mkdir(parents=True, exist_ok=True)
PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

st.title("🎯 Sistema Roleta - Controle Comercial")
st.caption(f"Diretório base: `{BASE_DIR}`")

# --- FUNÇÕES DE SUPORTE ---
def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    texto = re.sub(r'[\.\-\/\,\_]', ' ', texto)
    return " ".join(texto.split())

@st.cache_data(show_spinner=False)
def carregar_dados(origem):
    nome = origem.name if hasattr(origem, 'name') else str(origem)
    if nome.lower().endswith('.csv'):
        try:
            return pd.read_csv(origem, sep=';', encoding='latin1')
        except Exception:
            if hasattr(origem, 'seek'):
                origem.seek(0)
            return pd.read_csv(origem)
    else:
        return pd.read_excel(origem)

def identificar_coluna(colunas, possiveis_nomes):
    for nome in possiveis_nomes:
        for col in colunas:
            col_norm = normalizar_texto(col).lower()
            if nome.lower() in col_norm:
                return col
    return colunas[0] if len(colunas) > 0 else None

def buscar_arquivo_padrao(lista_arquivos, termos):
    for f in lista_arquivos:
        f_norm = normalizar_texto(f).lower()
        if any(t in f_norm for t in termos):
            return f
    return "Selecione..."

def limpar_valor(val):
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace('R$', '').replace('.', '').replace(',', '.').strip()
    return pd.to_numeric(s, errors='coerce') or 0.0

def estilizar_aba(ws, df_dados, col_id_nome, col_id_val, col_cli_nome, col_data_dt, titulo_col_id="PEDIDO", titulo_status=None):
    """
    Formata e preenche uma aba no padrão exato da Roleta:
    - Colunas A-G: Analítico
    - Colunas I-K: Resumo por Vendedor com fórmulas =SUM()
    """
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="000000")
    fill_cabecalho = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    borda_fina = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    fonte_alerta_vermelho = Font(name="Calibri", size=11, bold=True, color="9C0006")
    fill_alerta_vermelho = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # Cabeçalho Principal (A1:G1)
    cabecalhos_dados = ['VENDEDOR', titulo_col_id, 'CLIENTE', 'VALOR', 'STATUS', 'DATA', 'BRINDE ROLETA']
    for idx, h in enumerate(cabecalhos_dados, 1):
        cell = ws.cell(row=1, column=idx, value=h)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Cabeçalho Lateral (I1:K1)
    label_qtd = 'NUMERO DE PEDIDOS' if titulo_col_id == "PEDIDO" else 'NUMERO DE PROPOSTAS'
    cabecalhos_resumo = ['VENDEDOR', label_qtd, 'TOTAL']
    for idx, h in enumerate(cabecalhos_resumo, 9):
        cell = ws.cell(row=1, column=idx, value=h)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")

    grupos_vendedores = {}
    linha_atual = 2

    for _, r in df_dados.iterrows():
        vend = r['VENDEDOR_CLEAN']
        tem_duas_compras = bool(r.get('TEM_DUAS_COMPRAS', False))
        obs_brinde = str(r.get('BRINDE ROLETA', '')) if pd.notna(r.get('BRINDE ROLETA', '')) else ""

        ws.cell(row=linha_atual, column=1, value=vend)
        ws.cell(row=linha_atual, column=2, value=r[col_id_nome])
        ws.cell(row=linha_atual, column=3, value=str(r[col_cli_nome]))

        c_val = ws.cell(row=linha_atual, column=4, value=r['VALOR_NUM'])
        c_val.number_format = 'R$ #,##0.00'

        status_val = r['STATUS'] if 'STATUS' in r else (titulo_status or "")
        ws.cell(row=linha_atual, column=5, value=status_val)

        c_dt = ws.cell(row=linha_atual, column=6, value=r[col_data_dt].date() if pd.notna(r[col_data_dt]) else "")
        c_dt.number_format = 'DD/MM/YYYY'
        c_dt.alignment = Alignment(horizontal="center")

        ws.cell(row=linha_atual, column=7, value=obs_brinde)

        for c in range(1, 8):
            cell_item = ws.cell(row=linha_atual, column=c)
            cell_item.border = borda_fina
            if tem_duas_compras:
                cell_item.fill = fill_alerta_vermelho
                cell_item.font = fonte_alerta_vermelho
                if c == 7:
                    cell_item.alignment = Alignment(horizontal="center", vertical="center")

        if vend not in grupos_vendedores:
            grupos_vendedores[vend] = []
        grupos_vendedores[vend].append(linha_atual)

        linha_atual += 1

    # Tabela Resumo (Colunas I, J, K)
    linha_resumo = 2
    for vend, linhas in sorted(grupos_vendedores.items()):
        primeira = linhas[0]
        ultima = linhas[-1]
        qtd = len(linhas)

        ws.cell(row=linha_resumo, column=9, value=vend)
        ws.cell(row=linha_resumo, column=10, value=qtd).alignment = Alignment(horizontal="center")

        c_tot = ws.cell(row=linha_resumo, column=11, value=f"=SUM(D{primeira}:D{ultima})")
        c_tot.number_format = 'R$ #,##0.00'

        for c in range(9, 12):
            ws.cell(row=linha_resumo, column=c).border = borda_fina

        linha_resumo += 1

    # Auto-ajuste de colunas
    for col in ws.columns:
        letra = get_column_letter(col[0].column)
        tamanho_max = max(len(str(c.value or '')) for c in col)
        ws.column_dimensions[letra].width = max(tamanho_max + 3, 12)

def estilizar_aba_propostas(ws, df_resumo_adm):
    """
    Formata e preenche a Aba 2 no padrão exato da Roleta (Tabela Dinâmica ADM):
    - Coluna A: Vendedor
    - Coluna B: Qtd
    Cabeçalho com preenchimento amarelo (#FFFF00), negrito e bordas finas.
    """
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="000000")
    fill_cabecalho = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    borda_fina = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    fonte_dados = Font(name="Calibri", size=11, bold=False, color="000000")

    # Cabeçalho (A1:B1)
    c1 = ws.cell(row=1, column=1, value="Vendedor")
    c1.font = fonte_cabecalho
    c1.fill = fill_cabecalho
    c1.border = borda_fina
    c1.alignment = Alignment(horizontal="left", vertical="center")

    c2 = ws.cell(row=1, column=2, value="Qtd")
    c2.font = fonte_cabecalho
    c2.fill = fill_cabecalho
    c2.border = borda_fina
    c2.alignment = Alignment(horizontal="center", vertical="center")

    # Linhas de dados
    for idx, r in enumerate(df_resumo_adm.itertuples(), start=2):
        c_vend = ws.cell(row=idx, column=1, value=str(r.Vendedor))
        c_vend.font = fonte_dados
        c_vend.border = borda_fina
        c_vend.alignment = Alignment(horizontal="left", vertical="center")

        c_qtd = ws.cell(row=idx, column=2, value=int(r.Qtd))
        c_qtd.font = fonte_dados
        c_qtd.border = borda_fina
        c_qtd.alignment = Alignment(horizontal="center", vertical="center")

    # Auto-ajuste de largura das colunas
    max_len_v = max([len(str(r.Vendedor)) for r in df_resumo_adm.itertuples()] + [8])
    ws.column_dimensions['A'].width = max(max_len_v + 4, 22)
    ws.column_dimensions['B'].width = 12


# --- SELEÇÃO / UPLOAD DOS ARQUIVOS ---
st.subheader("1. Arquivos de Entrada (4 Bases)")

modo_leitura = st.radio(
    "Como deseja carregar os arquivos?",
    ["Fazer Upload pelo Navegador", f"Ler direto da pasta local ({PASTA_ENTRADA.name} ou raiz)"],
    horizontal=True
)

file_novos = None
file_reativados = None
file_fortpel = None
file_adm = None

if modo_leitura == "Fazer Upload pelo Navegador":
    c1, c2 = st.columns(2)
    c3, c4 = st.columns(2)
    with c1:
        file_novos = st.file_uploader("📥 1. BI Clientes Novos", type=["xlsx", "xls", "csv"], key="up_novos")
    with c2:
        file_reativados = st.file_uploader("📥 2. BI Clientes Reativados", type=["xlsx", "xls", "csv"], key="up_reat")
    with c3:
        file_fortpel = st.file_uploader("📥 3. Acompanhamento Fortpel (Pedidos)", type=["xlsx", "xls", "csv"], key="up_fort")
    with c4:
        file_adm = st.file_uploader("📥 4. Planilha ADM (Orçamentos Efetivados)", type=["xlsx", "xls", "csv"], key="up_adm")

else:
    # Busca tanto na pasta 'arquivos_entrada' quanto na pasta raiz do sistema
    arquivos_locais = sorted(list(set(
        [f.name for f in PASTA_ENTRADA.iterdir() if f.is_file() and not f.name.startswith('~$') and f.suffix.lower() in ['.xlsx', '.xls', '.csv']] +
        [f.name for f in BASE_DIR.iterdir() if f.is_file() and not f.name.startswith('~$') and f.suffix.lower() in ['.xlsx', '.xls', '.csv']]
    )))

    def resolver_caminho(nome):
        if (PASTA_ENTRADA / nome).exists():
            return PASTA_ENTRADA / nome
        return BASE_DIR / nome

    if not arquivos_locais:
        st.info(f"Nenhum arquivo encontrado em `{PASTA_ENTRADA}` ou `{BASE_DIR}`. Coloque as 4 planilhas lá ou use a opção de Upload.")
    else:
        padrao_novos = buscar_arquivo_padrao(arquivos_locais, ["novo"])
        padrao_reat = buscar_arquivo_padrao(arquivos_locais, ["reativ"])
        padrao_fort = buscar_arquivo_padrao(arquivos_locais, ["acomp", "fortpel"])
        padrao_adm = buscar_arquivo_padrao(arquivos_locais, ["adm", "orcamento", "proposta"])

        c1, c2 = st.columns(2)
        c3, c4 = st.columns(2)
        opcoes = ["Selecione..."] + arquivos_locais

        with c1:
            idx_n = opcoes.index(padrao_novos) if padrao_novos in opcoes else 0
            nome_n = st.selectbox("1. BI Clientes Novos:", opcoes, index=idx_n, key="loc_novos")
            if nome_n != "Selecione...":
                file_novos = resolver_caminho(nome_n)

        with c2:
            idx_r = opcoes.index(padrao_reat) if padrao_reat in opcoes else 0
            nome_r = st.selectbox("2. BI Clientes Reativados:", opcoes, index=idx_r, key="loc_reat")
            if nome_r != "Selecione...":
                file_reativados = resolver_caminho(nome_r)

        with c3:
            idx_f = opcoes.index(padrao_fort) if padrao_fort in opcoes else 0
            nome_f = st.selectbox("3. Acompanhamento Fortpel:", opcoes, index=idx_f, key="loc_fort")
            if nome_f != "Selecione...":
                file_fortpel = resolver_caminho(nome_f)

        with c4:
            idx_a = opcoes.index(padrao_adm) if padrao_adm in opcoes else 0
            nome_a = st.selectbox("4. Planilha ADM (Orçamentos Efetivados):", opcoes, index=idx_a, key="loc_adm")
            if nome_a != "Selecione...":
                file_adm = resolver_caminho(nome_a)


# --- MAPEAMENTO E PROCESSAMENTO ---
if file_novos and file_reativados and file_fortpel and file_adm:
    try:
        with st.spinner("Lendo os 4 arquivos Excel/CSV..."):
            df_novos_raw = carregar_dados(file_novos)
            df_reativados_raw = carregar_dados(file_reativados)
            df_fortpel_raw = carregar_dados(file_fortpel)
            df_adm_raw = carregar_dados(file_adm)
    except Exception as e:
        st.error(f"Erro ao ler arquivos: {e}")
        st.stop()

    st.success("As 4 bases foram carregadas com sucesso!")

    st.subheader("2. Mapeamento das Colunas Chave")
    tab_bi, tab_fortpel, tab_adm = st.tabs(["Bases do BI", "Acompanhamento Fortpel", "Planilha ADM (Orçamentos)"])

    with tab_bi:
        c_b1, c_b2 = st.columns(2)
        with c_b1:
            col_novos_cli = st.selectbox(
                "Coluna Cliente (BI Novos)",
                df_novos_raw.columns,
                index=df_novos_raw.columns.get_loc(identificar_coluna(df_novos_raw.columns, ["cliente", "razao", "nome"]))
            )
            col_novos_val = identificar_coluna(df_novos_raw.columns, ["venda", "valor", "total", "faturado"])
        with c_b2:
            col_reat_cli = st.selectbox(
                "Coluna Cliente (BI Reativados)",
                df_reativados_raw.columns,
                index=df_reativados_raw.columns.get_loc(identificar_coluna(df_reativados_raw.columns, ["cliente", "razao", "nome"]))
            )
            col_reat_val = identificar_coluna(df_reativados_raw.columns, ["venda", "valor", "total", "faturado"])

    with tab_fortpel:
        c_f1, c_f2, c_f3, c_f4, c_f5 = st.columns(5)
        with c_f1:
            col_f_cli = st.selectbox("Cliente", df_fortpel_raw.columns, index=df_fortpel_raw.columns.get_loc(identificar_coluna(df_fortpel_raw.columns, ["cliente", "razao", "nome"])))
        with c_f2:
            col_f_ped = st.selectbox("Pedido", df_fortpel_raw.columns, index=df_fortpel_raw.columns.get_loc(identificar_coluna(df_fortpel_raw.columns, ["pedido", "numero", "ordem"])))
        with c_f3:
            col_f_vend = st.selectbox("Vendedor", df_fortpel_raw.columns, index=df_fortpel_raw.columns.get_loc(identificar_coluna(df_fortpel_raw.columns, ["operador", "vendedor", "representante", "usuario", "gestor"])))
        with c_f4:
            col_f_val = st.selectbox("Valor", df_fortpel_raw.columns, index=df_fortpel_raw.columns.get_loc(identificar_coluna(df_fortpel_raw.columns, ["valor", "total", "faturado", "preco", "somadevl"])))
        with c_f5:
            col_f_data = st.selectbox("Data", df_fortpel_raw.columns, index=df_fortpel_raw.columns.get_loc(identificar_coluna(df_fortpel_raw.columns, ["dt_inser", "dt_saida", "data", "emissao", "faturamento", "dt"])))

        # Opção inteligente para valor da Aba 1 (caso Fortpel tenha apenas soma total)
        tem_valor_bi = (col_novos_val is not None) and (col_reat_val is not None)
        origem_valor_aba1 = st.radio(
            "Origem do Valor da Aba 1 (Novos/Reativados):",
            [
                "Usar coluna 'Venda' do BI (Novos e Reativados) [Recomendado se o Acompanhamento tiver total geral]",
                "Usar coluna selecionada do Acompanhamento Fortpel"
            ] if tem_valor_bi else ["Usar coluna selecionada do Acompanhamento Fortpel"],
            index=0 if tem_valor_bi else 0
        )

    with tab_adm:
        st.markdown("📊 **Tabela Dinâmica ADM (Aba 2: Propostas Enviadas)**")
        st.caption("A aba de propostas é gerada pela contagem de registros do campo Status por Vendedor.")
        c_a1, c_a2 = st.columns(2)
        with c_a1:
            idx_v = df_adm_raw.columns.get_loc(identificar_coluna(df_adm_raw.columns, ["vendedor", "operador", "usuario", "representante"]))
            col_a_vend = st.selectbox("Coluna Vendedor (ADM):", df_adm_raw.columns, index=idx_v)
        with c_a2:
            idx_s = df_adm_raw.columns.get_loc(identificar_coluna(df_adm_raw.columns, ["status", "situacao", "fase"]))
            col_a_status = st.selectbox("Coluna Status / Contagem (ADM):", df_adm_raw.columns, index=idx_s)

        todos_status_adm = sorted([str(s) for s in df_adm_raw[col_a_status].dropna().unique()])
        status_adm_selecionados = st.multiselect(
            "Filtrar por Status específicos (opcional - por padrão conta todos):",
            options=todos_status_adm,
            default=todos_status_adm,
            help="Selecione quais status entram na contagem. Por padrão, todos os status presentes na planilha são contabilizados."
        )

    # Conversão de Datas (Apenas Fortpel requer data)
    df_fortpel = df_fortpel_raw.copy()
    df_fortpel['__DATA_DT'] = pd.to_datetime(df_fortpel[col_f_data], dayfirst=True, errors='coerce')

    datas_validas = df_fortpel['__DATA_DT'].dropna()
    min_date = datas_validas.min() if not datas_validas.empty else pd.Timestamp.now()
    max_date = datas_validas.max() if not datas_validas.empty else pd.Timestamp.now()

    # --- FILTRO POR PERÍODO ---
    st.subheader("3. Período e Filtro de Datas")

    c_filt1, c_filt2 = st.columns([1, 1])
    with c_filt1:
        modo_filtro_roleta = st.radio(
            "Filtro de Data para a Roleta (Aba 1 - Novos e Reativados):",
            [
                "Opção B: Filtrar pelo intervalo de datas selecionado (Inclusivo, recomendado)",
                "Opção A: Considerar tudo que vier no BI (Sem restrição de data na Roleta)"
            ],
            index=0,
            help="A Opção B restringe os pedidos rigorosamente entre as datas selecionadas, evitando trazer pedidos inseridos após a data limite ou pedidos antigos."
        )

        criterio_roleta = st.radio(
            "Critério de Contagem da Roleta (Aba 1):",
            [
                "1 Registro por Cliente Único (Padrão BI - 1 linha por cliente, soma exata)",
                "Todos os Pedidos Individuais (Pode gerar múltiplas linhas se o cliente fez mais de um pedido)"
            ],
            index=0,
            help="No relatório do BI, cada cliente conta como 1 registro. Marque '1 Registro por Cliente Único' para bater 100% com o relatório do BI."
        )

    with c_filt2:
        data_sel = st.date_input(
            "Selecione o intervalo de datas (para Aba 2 e filtro da Roleta):",
            value=(min_date.date(), max_date.date())
        )

    if isinstance(data_sel, tuple) and len(data_sel) == 2:
        dt_inicio, dt_fim = data_sel

        if st.button("🚀 Processar e Gerar Planilha Completa", type="primary"):
            with st.spinner("Processando cruzamentos das 2 abas..."):
                # ==========================================
                # ABA 1: REATIVADOS - NOVOS
                # ==========================================
                df_novos_clean = df_novos_raw.dropna(subset=[col_novos_cli]).copy()
                df_novos_clean['__CLI_NORM'] = df_novos_clean[col_novos_cli].apply(normalizar_texto)
                clientes_novos = set(df_novos_clean['__CLI_NORM'])

                df_reat_clean = df_reativados_raw.dropna(subset=[col_reat_cli]).copy()
                df_reat_clean['__CLI_NORM'] = df_reat_clean[col_reat_cli].apply(normalizar_texto)
                clientes_reativados = set(df_reat_clean['__CLI_NORM'])

                df_fortpel['__CLI_NORM'] = df_fortpel[col_f_cli].apply(normalizar_texto)

                def atribuir_status(nome_norm):
                    if nome_norm in clientes_novos:
                        return "NOVOS"
                    elif nome_norm in clientes_reativados:
                        return "REATIVADO"
                    return None

                df_fortpel['STATUS'] = df_fortpel['__CLI_NORM'].apply(atribuir_status)

                status_validos = ['NOVOS', 'NOVO', 'REATIVADO', 'REATIVADOS']
                mask_status = df_fortpel['STATUS'].astype(str).str.upper().isin(status_validos)

                if "Sem restrição" in modo_filtro_roleta:
                    # Sem restrição de data
                    df_aba1 = df_fortpel[mask_status].copy()
                else:
                    # FILTRAR POR PERÍODO INCLUSIVO
                    mask_data_f = (df_fortpel['__DATA_DT'].dt.date >= dt_inicio) & (df_fortpel['__DATA_DT'].dt.date <= dt_fim)
                    df_aba1 = df_fortpel[mask_data_f & mask_status].copy()

                # Desduplicação por Cliente Único (Padrão BI) se selecionado
                if "Cliente Único" in criterio_roleta:
                    df_aba1 = df_aba1.sort_values(by=['__DATA_DT', col_f_ped], ascending=[True, True])
                    df_aba1 = df_aba1.drop_duplicates(subset=['__CLI_NORM'], keep='first')

                # Mapa de valores do BI se selecionado
                if "BI" in origem_valor_aba1:
                    mapa_valores_bi = {}
                    if col_novos_val:
                        for _, row in df_novos_clean.iterrows():
                            mapa_valores_bi[row['__CLI_NORM']] = limpar_valor(row[col_novos_val])
                    if col_reat_val:
                        for _, row in df_reat_clean.iterrows():
                            mapa_valores_bi[row['__CLI_NORM']] = limpar_valor(row[col_reat_val])
                    df_aba1['VALOR_NUM'] = df_aba1['__CLI_NORM'].map(mapa_valores_bi).fillna(0.0)
                else:
                    df_aba1['VALOR_NUM'] = df_aba1[col_f_val].apply(limpar_valor)

                # ========================================================
                # REGRA: VERIFICAR OUTRO PEDIDO NO MESMO MÊS DO PERÍODO
                # ========================================================
                meses_periodo = set(pd.date_range(dt_inicio, dt_fim).to_period('M'))
                mask_meses_f = df_fortpel['__DATA_DT'].dt.to_period('M').isin(meses_periodo)
                pedidos_no_mes = df_fortpel[mask_meses_f].groupby('__CLI_NORM')[col_f_ped].nunique().to_dict()

                def verificar_duas_compras(cli_norm):
                    qtd = pedidos_no_mes.get(cli_norm, 1)
                    return qtd > 1

                def gerar_rotulo_brinde(cli_norm):
                    qtd = pedidos_no_mes.get(cli_norm, 1)
                    if qtd == 2:
                        return "DUAS COMPRAS"
                    elif qtd > 2:
                        return f"DUAS COMPRAS ({qtd} PEDIDOS)"
                    return ""

                df_aba1['TEM_DUAS_COMPRAS'] = df_aba1['__CLI_NORM'].apply(verificar_duas_compras)
                df_aba1['BRINDE ROLETA'] = df_aba1['__CLI_NORM'].apply(gerar_rotulo_brinde)

                # ==========================================
                # ABA 2: PROPOSTAS ENVIADAS (TABELA DINÂMICA ADM)
                # ==========================================
                df_adm_filtrado = df_adm_raw.dropna(subset=[col_a_vend]).copy()
                if status_adm_selecionados:
                    df_adm_filtrado = df_adm_filtrado[df_adm_filtrado[col_a_status].astype(str).isin(status_adm_selecionados)]

                df_aba2 = (
                    df_adm_filtrado.groupby(col_a_vend)[col_a_status]
                    .count()
                    .reset_index()
                )
                df_aba2.columns = ['Vendedor', 'Qtd']
                df_aba2 = df_aba2.sort_values(by=['Qtd', 'Vendedor'], ascending=[False, True])

                # Dataframes finais oficiais de cada aba:
                df_roleta_novos_reativados = df_aba1
                df_propostas = df_aba2

                # ==========================================
                # GERAÇÃO DO WORKBOOK COM AS DUAS ABAS
                # ==========================================
                wb = openpyxl.Workbook()

                # Nome das abas dinâmicas mantendo as datas selecionadas
                dia_i = dt_inicio.strftime('%d')
                dia_f = dt_fim.strftime('%d')
                nome_aba1 = f"{dia_i}-{dia_f} Reativados - Novos"[:31]
                nome_aba2 = f"{dia_i}-{dia_f} Propostas Enviadas"[:31]

                # Preencher Aba 1 (df_roleta_novos_reativados)
                ws1 = wb.active
                ws1.title = nome_aba1

                if not df_roleta_novos_reativados.empty:
                    df_roleta_novos_reativados['VENDEDOR_CLEAN'] = df_roleta_novos_reativados[col_f_vend].astype(str).str.strip().str.upper()
                    df_roleta_novos_reativados = df_roleta_novos_reativados.sort_values(by=['VENDEDOR_CLEAN', '__DATA_DT', col_f_cli])
                    estilizar_aba(ws1, df_roleta_novos_reativados, col_f_ped, col_f_val, col_f_cli, '__DATA_DT', titulo_col_id="PEDIDO")
                else:
                    ws1.cell(row=1, column=1, value="Nenhum pedido encontrado para Novos/Reativados no período.")

                # Preencher Aba 2 (df_propostas - Tabela Dinâmica Vendedor x Qtd)
                ws2 = wb.create_sheet(title=nome_aba2)

                if not df_propostas.empty:
                    estilizar_aba_propostas(ws2, df_propostas)
                else:
                    ws2.cell(row=1, column=1, value="Nenhum registro encontrado na planilha ADM.")

                # Salvar na pasta planilhas_geradas
                nome_arquivo_saida = f"Controle_roleta_{dt_inicio.strftime('%d%m')}_a_{dt_fim.strftime('%d%m')}.xlsx"
                caminho_salvar_local = PASTA_SAIDA / nome_arquivo_saida
                wb.save(caminho_salvar_local)

                # Buffer para Download
                buffer_download = io.BytesIO()
                wb.save(buffer_download)
                buffer_download.seek(0)

                st.success("✅ Planilha gerada com as duas abas com sucesso!")
                st.info(f"📁 Arquivo salvo em: `{caminho_salvar_local}`")

                t1, t2 = st.tabs([f"Aba 1: {nome_aba1}", f"Aba 2: {nome_aba2}"])
                with t1:
                    if not df_aba1.empty:
                        qtd_duas_compras = int(df_aba1['TEM_DUAS_COMPRAS'].sum()) if 'TEM_DUAS_COMPRAS' in df_aba1 else 0
                        if qtd_duas_compras > 0:
                            st.warning(f"⚠️ {qtd_duas_compras} cliente(s) com mais de uma compra no mesmo mês identificado(s) (destacados com 'DUAS COMPRAS' e linha em vermelho na planilha)!")
                        st.dataframe(df_aba1[['VENDEDOR_CLEAN', col_f_ped, col_f_cli, 'VALOR_NUM', 'STATUS', '__DATA_DT', 'BRINDE ROLETA']], use_container_width=True)
                    else:
                        st.write("Sem registros no período para a Aba 1.")
                with t2:
                    if not df_aba2.empty:
                        st.dataframe(df_aba2, use_container_width=True)
                        st.metric("Total de Propostas Efetivadas", int(df_aba2['Qtd'].sum()))
                    else:
                        st.write("Sem registros na planilha ADM.")

                st.download_button(
                    label="📥 Baixar Planilha Consolidada (2 Abas) (.xlsx)",
                    data=buffer_download,
                    file_name=nome_arquivo_saida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
