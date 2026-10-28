import io
import os
import re
import unicodedata
import datetime
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
        nome_norm = normalizar_texto(nome).lower()
        for col in colunas:
            col_norm = normalizar_texto(col).lower()
            if (nome_norm in col_norm) or (nome.lower() in str(col).lower()):
                return col
    return colunas[0] if len(colunas) > 0 else None

def buscar_arquivo_padrao(lista_arquivos, termos):
    for f in lista_arquivos:
        f_norm = normalizar_texto(f).lower()
        if any(t in f_norm for t in termos):
            return f
    return "Selecione..."

def limpar_valor(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace('R$', '').strip()
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
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

    cabecalhos_dados = ['VENDEDOR', titulo_col_id, 'CLIENTE', 'VALOR', 'STATUS', 'DATA', 'BRINDE ROLETA']
    for idx, h in enumerate(cabecalhos_dados, 1):
        cell = ws.cell(row=1, column=idx, value=h)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")

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

    for col in ws.columns:
        letra = get_column_letter(col[0].column)
        tamanho_max = max(len(str(c.value or '')) for c in col)
        ws.column_dimensions[letra].width = max(tamanho_max + 3, 12)

def estilizar_aba_propostas(ws, df_resumo_adm):
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="000000")
    fill_cabecalho = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    borda_fina = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    fonte_dados = Font(name="Calibri", size=11, bold=False, color="000000")

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

    for idx, r in enumerate(df_resumo_adm.itertuples(), start=2):
        c_vend = ws.cell(row=idx, column=1, value=str(r.Vendedor))
        c_vend.font = fonte_dados
        c_vend.border = borda_fina
        c_vend.alignment = Alignment(horizontal="left", vertical="center")

        c_qtd = ws.cell(row=idx, column=2, value=int(r.Qtd))
        c_qtd.font = fonte_dados
        c_qtd.border = borda_fina
        c_qtd.alignment = Alignment(horizontal="center", vertical="center")

    max_len_v = max([len(str(r.Vendedor)) for r in df_resumo_adm.itertuples()] + [8])
    ws.column_dimensions['A'].width = max(max_len_v + 4, 22)
    ws.column_dimensions['B'].width = 12

def estilizar_aba_produtos(ws, df_top_notas, df_lideres_cat):
    fonte_titulo = Font(name="Calibri", size=13, bold=True, color="1F4E79")
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="000000")
    fill_cab_azul = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_cab_amarelo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    borda_fina = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    fonte_dados = Font(name="Calibri", size=11, bold=False, color="000000")
    fonte_destaque = Font(name="Calibri", size=11, bold=True, color="000000")

    ws.cell(row=1, column=1, value="🏆 TOP MAIORES VENDAS EM NOTA ÚNICA").font = fonte_titulo

    headers_1 = ["POSIÇÃO", "VENDEDOR", "NOTA", "VALOR TOTAL"]
    for idx, h in enumerate(headers_1, 1):
        c = ws.cell(row=2, column=idx, value=h)
        c.font = fonte_cabecalho
        c.fill = fill_cab_azul
        c.border = borda_fina
        c.alignment = Alignment(horizontal="center", vertical="center")

    linha_atual = 3
    if df_top_notas is not None and not df_top_notas.empty:
        for pos, r in enumerate(df_top_notas.head(10).itertuples(), start=1):
            c_pos = ws.cell(row=linha_atual, column=1, value=f"{pos}º")
            c_pos.alignment = Alignment(horizontal="center")
            c_pos.font = fonte_destaque if pos == 1 else fonte_dados
            c_pos.border = borda_fina

            c_vend = ws.cell(row=linha_atual, column=2, value=str(r.Vendedor))
            c_vend.font = fonte_destaque if pos == 1 else fonte_dados
            c_vend.border = borda_fina

            nota_val = int(r.Nota) if pd.notna(r.Nota) and isinstance(r.Nota, (int, float)) and r.Nota == int(r.Nota) else str(r.Nota)
            c_nota = ws.cell(row=linha_atual, column=3, value=nota_val)
            c_nota.alignment = Alignment(horizontal="center")
            c_nota.font = fonte_dados
            c_nota.border = borda_fina

            c_val = ws.cell(row=linha_atual, column=4, value=float(r.Valor_Total))
            c_val.number_format = 'R$ #,##0.00'
            c_val.font = fonte_destaque if pos == 1 else fonte_dados
            c_val.border = borda_fina

            linha_atual += 1
    else:
        ws.cell(row=3, column=1, value="Sem dados de nota fiscal disponíveis.")
        linha_atual = 4

    linha_atual += 2
    ws.cell(row=linha_atual, column=1, value="🏷️ MAIOR VENDA POR CATEGORIA DE PRODUTO").font = fonte_titulo
    linha_atual += 1

    headers_2 = ["CATEGORIA", "VENDEDOR LÍDER", "TOTAL VENDIDO"]
    for idx, h in enumerate(headers_2, 1):
        c = ws.cell(row=linha_atual, column=idx, value=h)
        c.font = fonte_cabecalho
        c.fill = fill_cab_amarelo
        c.border = borda_fina
        c.alignment = Alignment(horizontal="center", vertical="center")

    linha_atual += 1
    if df_lideres_cat is not None and not df_lideres_cat.empty:
        for r in df_lideres_cat.itertuples():
            c_cat = ws.cell(row=linha_atual, column=1, value=str(r.Categoria))
            c_cat.font = fonte_dados
            c_cat.border = borda_fina

            c_vend = ws.cell(row=linha_atual, column=2, value=str(r.Vendedor))
            c_vend.font = fonte_destaque
            c_vend.border = borda_fina

            c_val = ws.cell(row=linha_atual, column=3, value=float(r.Valor_Total))
            c_val.number_format = 'R$ #,##0.00'
            c_val.font = fonte_dados
            c_val.border = borda_fina

            linha_atual += 1
    else:
        ws.cell(row=linha_atual, column=1, value="Sem dados de categoria disponíveis.")

    for col in ws.columns:
        letra = get_column_letter(col[0].column)
        tamanho_max = max(len(str(c.value or '')) for c in col)
        ws.column_dimensions[letra].width = max(tamanho_max + 3, 14)


CATEGORIAS_ROLETA = [
    "ALIMENTOS E BEBIDAS",
    "EMBALAGENS E DESCARTAVEIS",
    "EPI'S",
    "HIGIENE E LIMPEZA",
    "INFORMÁTICA",
    "MRO",
    "OFFICE"
]

def extrair_ranking_categoria(df_cat_raw, categoria_alvo, df_fortpel=None, df_adm=None):
    if df_cat_raw is None or df_cat_raw.empty:
        return pd.DataFrame(), 0.0, pd.DataFrame()

    mapa_cli_vend = {}
    if df_fortpel is not None and not df_fortpel.empty:
        col_c = None
        for cand in ['nome_cliente', 'cliente', 'razao', 'nome']:
            for c in df_fortpel.columns:
                if cand in str(c).lower():
                    col_c = c
                    break
            if col_c:
                break

        col_v = None
        for cand in ['carteira', 'operador', 'vendedor', 'gestor']:
            for c in df_fortpel.columns:
                if cand in str(c).lower():
                    col_v = c
                    break
            if col_v:
                break

        if col_c and col_v:
            df_f_clean = df_fortpel.dropna(subset=[col_c, col_v]).copy()
            df_f_clean['__C_NORM'] = df_f_clean[col_c].apply(normalizar_texto)
            grp = df_f_clean.groupby('__C_NORM')[col_v].agg(lambda x: x.mode()[0] if not x.mode().empty else x.iloc[0])
            for cli, v in grp.items():
                mapa_cli_vend[cli] = str(v).strip().upper()

    if df_adm is not None and not df_adm.empty:
        col_ac = None
        col_av = None
        for cand in ['nome do cliente', 'cliente', 'nome']:
            for c in df_adm.columns:
                if cand in str(c).lower():
                    col_ac = c
                    break
            if col_ac:
                break
        for cand in ['vendedor do cliente', 'vendedor', 'operador']:
            for c in df_adm.columns:
                if cand in str(c).lower():
                    col_av = c
                    break
            if col_av:
                break
        if col_ac and col_av:
            for _, r in df_adm.dropna(subset=[col_ac, col_av]).iterrows():
                cn = normalizar_texto(r[col_ac])
                if cn and cn not in mapa_cli_vend:
                    mapa_cli_vend[cn] = str(r[col_av]).strip().upper()

    df_work = df_cat_raw.copy()
    if str(df_work.iloc[0, 0]).strip().lower() in ['cliente', 'codigo', 'cod']:
        df_work = df_work.iloc[1:].copy()

    df_work = df_work[~df_work.iloc[:, 0].astype(str).str.upper().str.contains('TOTAL|FILTROS APLICADOS', na=False)].copy()

    cols_map = {}
    for col in df_work.columns:
        c_norm = normalizar_texto(col).replace("'", "").replace(" ", "")
        for cat in CATEGORIAS_ROLETA:
            cat_norm = normalizar_texto(cat).replace("'", "").replace(" ", "")
            if cat_norm == c_norm or cat_norm in c_norm or c_norm in cat_norm:
                cols_map[cat] = col

    lideres_todas = []

    if len(cols_map) >= 2:
        col_id_primeira = df_work.columns[0]
        df_work['__ID_NORM'] = df_work[col_id_primeira].apply(normalizar_texto)

        def obter_vendedor(id_norm):
            if id_norm in mapa_cli_vend:
                return mapa_cli_vend[id_norm]
            return id_norm

        df_work['__VENDEDOR'] = df_work['__ID_NORM'].apply(obter_vendedor)

        for cat in CATEGORIAS_ROLETA:
            c_nome = cols_map.get(cat)
            if c_nome and c_nome in df_work.columns:
                df_work[f'__VAL_{cat}'] = df_work[c_nome].apply(limpar_valor)
                res_cat = (
                    df_work.groupby('__VENDEDOR', as_index=False)[f'__VAL_{cat}']
                    .sum()
                    .rename(columns={'__VENDEDOR': 'Vendedor', f'__VAL_{cat}': 'Valor_Total'})
                    .sort_values(by='Valor_Total', ascending=False)
                )
                res_cat = res_cat[res_cat['Valor_Total'] > 0]
                if not res_cat.empty:
                    top_v = res_cat.iloc[0]
                    lideres_todas.append({
                        'Categoria': cat,
                        'Vendedor': top_v['Vendedor'],
                        'Valor_Total': top_v['Valor_Total'],
                        'Total_Categoria': res_cat['Valor_Total'].sum()
                    })

        col_alvo = cols_map.get(categoria_alvo)
        if col_alvo and col_alvo in df_work.columns:
            df_work['__VAL_ALVO'] = df_work[col_alvo].apply(limpar_valor)
            df_rank_alvo = (
                df_work.groupby('__VENDEDOR', as_index=False)
                .agg(
                    Total_Vendido=('__VAL_ALVO', 'sum'),
                    Qtd_Itens=('__VAL_ALVO', lambda s: (s > 0).sum())
                )
                .rename(columns={'__VENDEDOR': 'Vendedor'})
                .sort_values(by='Total_Vendido', ascending=False)
            )
            df_rank_alvo = df_rank_alvo[df_rank_alvo['Total_Vendido'] > 0]
            total_alvo = df_rank_alvo['Total_Vendido'].sum()
            return df_rank_alvo, total_alvo, pd.DataFrame(lideres_todas)

    else:
        col_c = identificar_coluna(df_work.columns, ["categoria", "grupo", "linha", "subgrupo", "descr. produto", "descricao"])
        col_v = identificar_coluna(df_work.columns, ["vendedor", "operador", "gestor", "representante", "carteira"])
        col_val = identificar_coluna(df_work.columns, ["vendas total", "venda total", "venda", "valor total", "valor", "total", "preco"])

        if col_c and col_v and col_val:
            df_work['__VALOR_NUM'] = df_work[col_val].apply(limpar_valor)
            df_work['__VEND_LIMPO'] = df_work[col_v].astype(str).str.strip().str.upper()

            for cat in CATEGORIAS_ROLETA:
                cat_norm = normalizar_texto(cat)
                m_cat = df_work[col_c].apply(lambda x: cat_norm in normalizar_texto(x) or normalizar_texto(x) in cat_norm)
                df_sub = df_work[m_cat]
                if not df_sub.empty:
                    grp = (
                        df_sub.groupby('__VEND_LIMPO', as_index=False)['__VALOR_NUM']
                        .sum()
                        .rename(columns={'__VEND_LIMPO': 'Vendedor', '__VALOR_NUM': 'Valor_Total'})
                        .sort_values(by='Valor_Total', ascending=False)
                    )
                    grp = grp[grp['Valor_Total'] > 0]
                    if not grp.empty:
                        lideres_todas.append({
                            'Categoria': cat,
                            'Vendedor': grp.iloc[0]['Vendedor'],
                            'Valor_Total': grp.iloc[0]['Valor_Total'],
                            'Total_Categoria': grp['Valor_Total'].sum()
                        })

            cat_alvo_norm = normalizar_texto(categoria_alvo)
            m_alvo = df_work[col_c].apply(lambda x: cat_alvo_norm in normalizar_texto(x) or normalizar_texto(x) in cat_alvo_norm)
            df_sub_alvo = df_work[m_alvo]
            if not df_sub_alvo.empty:
                df_rank_alvo = (
                    df_sub_alvo.groupby('__VEND_LIMPO', as_index=False)
                    .agg(
                        Total_Vendido=('__VALOR_NUM', 'sum'),
                        Qtd_Itens=('__VALOR_NUM', 'count')
                    )
                    .rename(columns={'__VEND_LIMPO': 'Vendedor'})
                    .sort_values(by='Total_Vendido', ascending=False)
                )
                df_rank_alvo = df_rank_alvo[df_rank_alvo['Total_Vendido'] > 0]
                total_alvo = df_rank_alvo['Total_Vendido'].sum()
                return df_rank_alvo, total_alvo, pd.DataFrame(lideres_todas)

    return pd.DataFrame(), 0.0, pd.DataFrame(lideres_todas)

def extrair_maior_venda_nf_faturada(df_nf_raw, dt_inicio=None, dt_fim=None):
    if df_nf_raw is None or df_nf_raw.empty:
        return pd.DataFrame()

    col_n = identificar_coluna(df_nf_raw.columns, ["n_nota_obs", "nota fiscal", "nota", "nf", "numero da nota", "nr nota", "num_nota", "documento", "pedido"])
    col_v = identificar_coluna(df_nf_raw.columns, ["vendedor", "operador", "gestor", "carteira", "representante", "usuario"])
    col_val = identificar_coluna(df_nf_raw.columns, ["vendas total", "venda total", "venda", "valor total", "valor", "total", "vl_total", "faturado", "preco venda"])
    col_dt = identificar_coluna(df_nf_raw.columns, ["dt_nota_hist", "dt_emissao", "data emissao", "data faturamento", "dt_saida", "data", "dt", "emissao"])

    if not col_n or not col_v or not col_val:
        return pd.DataFrame()

    df_work = df_nf_raw.dropna(subset=[col_n, col_v]).copy()
    df_work['__VAL'] = df_work[col_val].apply(limpar_valor)
    df_work['__VEND'] = df_work[col_v].astype(str).str.strip().str.upper()

    if col_dt and col_dt in df_work.columns and dt_inicio and dt_fim:
        df_work['__DT'] = pd.to_datetime(df_work[col_dt], errors='coerce')
        mask_dt = (df_work['__DT'].dt.date >= dt_inicio) & (df_work['__DT'].dt.date <= dt_fim)
        if mask_dt.sum() > 0:
            df_work = df_work[mask_dt].copy()

    df_work['__NUM_NOTA'] = pd.to_numeric(df_work[col_n], errors='coerce')
    df_valid = df_work[(df_work['__NUM_NOTA'].notna()) & (df_work['__NUM_NOTA'] > 0)].copy()
    if df_valid.empty:
        df_valid = df_work.copy()

    df_grp = (
        df_valid.groupby([col_n, '__VEND'], as_index=False)['__VAL']
        .sum()
        .rename(columns={col_n: 'Nota', '__VEND': 'Vendedor', '__VAL': 'Valor_Total'})
        .sort_values(by='Valor_Total', ascending=False)
    )
    df_grp = df_grp[df_grp['Valor_Total'] > 0]
    return df_grp

def extrair_maior_venda_nota(df_fortpel, col_nota, col_vend, col_val, dt_inicio=None, dt_fim=None):
    if df_fortpel is None or df_fortpel.empty or not col_nota or not col_vend or not col_val:
        return pd.DataFrame()

    df_f = df_fortpel.copy()
    df_f['__VALOR_NUM'] = df_f[col_val].apply(limpar_valor)
    df_f['__VEND_LIMPO'] = df_f[col_vend].astype(str).str.strip().str.upper()

    if '__DATA_DT' in df_f.columns and pd.notna(dt_inicio) and pd.notna(dt_fim):
        mask_dt = (df_f['__DATA_DT'].dt.date >= dt_inicio) & (df_f['__DATA_DT'].dt.date <= dt_fim)
        df_sub = df_f[mask_dt].copy()
        if df_sub.empty:
            df_sub = df_f.copy()
    else:
        df_sub = df_f.copy()

    df_f_num = pd.to_numeric(df_sub[col_nota], errors='coerce')
    df_sub_valid = df_sub[(df_f_num.notna()) & (df_f_num > 0)].copy()
    if df_sub_valid.empty:
        df_sub_valid = df_sub.copy()

    df_grp = (
        df_sub_valid.groupby([col_nota, '__VEND_LIMPO'], as_index=False)['__VALOR_NUM']
        .sum()
        .rename(columns={col_nota: 'Nota', '__VEND_LIMPO': 'Vendedor', '__VALOR_NUM': 'Valor_Total'})
        .sort_values(by='Valor_Total', ascending=False)
    )
    df_grp = df_grp[df_grp['Valor_Total'] > 0]
    return df_grp

def obter_id_arquivo(f):
    if f is None:
        return "None"
    if hasattr(f, 'name') and hasattr(f, 'size'):
        return f"{f.name}_{f.size}"
    return str(f)

def carregar_planilha_ajustada(arquivo):
    """
    Lê uma planilha Excel previamente gerada e ajustada manualmente pelo usuário.
    Retorna dicionário contendo df_aba1 (Novos/Reativados) e df_aba2 (Propostas).
    """
    xl = pd.ExcelFile(arquivo)
    res = {}

    s1 = None
    for s in xl.sheet_names:
        if any(term in s.lower() for term in ['reativad', 'novo', 'aba 1']):
            s1 = s
            break
    if not s1 and len(xl.sheet_names) > 0:
        s1 = xl.sheet_names[0]

    if s1:
        df1 = pd.read_excel(arquivo, sheet_name=s1)
        df1 = df1.iloc[:, :7].dropna(how='all').copy()
        df1.columns = [str(c).strip().upper() for c in df1.columns]

        col_v = identificar_coluna(df1.columns, ['vendedor', 'operador'])
        col_c = identificar_coluna(df1.columns, ['cliente', 'nome'])
        col_vl = identificar_coluna(df1.columns, ['valor', 'total'])
        col_st = identificar_coluna(df1.columns, ['status'])
        col_p = identificar_coluna(df1.columns, ['pedido', 'nota'])

        df1_clean = pd.DataFrame()
        if col_v: df1_clean['Vendedor'] = df1[col_v].astype(str).str.strip().str.upper()
        if col_c: df1_clean['Cliente'] = df1[col_c].astype(str).str.strip().str.upper()
        if col_vl: df1_clean['VALOR_NUM'] = df1[col_vl].apply(limpar_valor)
        if col_st: df1_clean['Status'] = df1[col_st].astype(str).str.strip().str.upper()
        if col_p: df1_clean['Pedido'] = df1[col_p]
        res['df_aba1'] = df1_clean
        res['nome_aba1'] = s1

    s2 = None
    for s in xl.sheet_names:
        if any(term in s.lower() for term in ['proposta', 'adm', 'aba 2']):
            s2 = s
            break
    if s2:
        df2 = pd.read_excel(arquivo, sheet_name=s2).dropna(how='all')
        col_v = identificar_coluna(df2.columns, ['vendedor', 'operador'])
        col_q = identificar_coluna(df2.columns, ['qtd', 'quantidade', 'total'])
        if col_v and col_q:
            df2_clean = df2[[col_v, col_q]].copy()
            df2_clean.columns = ['Vendedor', 'Qtd']
            df2_clean['Vendedor'] = df2_clean['Vendedor'].astype(str).str.strip().str.upper()
            df2_clean['Qtd'] = pd.to_numeric(df2_clean['Qtd'], errors='coerce').fillna(0).astype(int)
            res['df_aba2'] = df2_clean
            res['nome_aba2'] = s2

    return res

def padronizar_df_aba1(df):
    if df is None or df.empty:
        return pd.DataFrame()
    df_c = df.copy()
    col_v = identificar_coluna(df_c.columns, ["vendedor_clean", "vendedor", "operador"])
    col_cli = identificar_coluna(df_c.columns, ["cliente_clean", "cliente", "nome"])
    col_val = identificar_coluna(df_c.columns, ["valor_num", "valor", "total"])
    col_ped = identificar_coluna(df_c.columns, ["pedido", "nota"])
    col_st = identificar_coluna(df_c.columns, ["status", "status_clean"])

    res = pd.DataFrame()
    if col_v: res['Vendedor'] = df_c[col_v].astype(str).str.strip().str.upper()
    if col_cli: res['Cliente'] = df_c[col_cli].astype(str).str.strip().str.upper()
    if col_val: res['VALOR_NUM'] = df_c[col_val].apply(limpar_valor)
    if col_ped: res['Pedido'] = df_c[col_ped]
    if col_st: res['Status'] = df_c[col_st].astype(str).str.strip().str.upper()
    return res

def padronizar_df_aba2(df):
    if df is None or df.empty:
        return pd.DataFrame()
    df_c = df.copy()
    col_v = identificar_coluna(df_c.columns, ["vendedor", "operador"])
    col_q = identificar_coluna(df_c.columns, ["qtd", "quantidade", "total"])
    res = pd.DataFrame()
    if col_v and col_q:
        res['Vendedor'] = df_c[col_v].astype(str).str.strip().str.upper()
        res['Qtd'] = pd.to_numeric(df_c[col_q], errors='coerce').fillna(0).astype(int)
    return res

def gerar_excel_resultados_finais(rank1, rank2, rank3, rank4, cat_nome):
    wb = openpyxl.Workbook()
    fonte_titulo = Font(name="Calibri", size=13, bold=True, color="1F4E79")
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="000000")
    fill_cab_azul = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_cab_amarelo = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    borda_fina = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    fonte_dados = Font(name="Calibri", size=11, bold=False, color="000000")
    fonte_destaque = Font(name="Calibri", size=11, bold=True, color="000000")

    # ABA 1: PÓDIO DOS CAMPEÕES
    ws0 = wb.active
    ws0.title = "🏆 Campeões Roleta"
    ws0.cell(row=1, column=1, value="🏆 RESULTADOS FINAIS DA ROLETA COMERCIAL").font = fonte_titulo

    headers_podio = ["RANKING", "CRITÉRIO DE PREMIAÇÃO", "VENDEDOR CAMPEÃO", "RESULTADO / MÉTRICA"]
    for idx, h in enumerate(headers_podio, 1):
        c = ws0.cell(row=3, column=idx, value=h)
        c.font = fonte_cabecalho
        c.fill = fill_cab_amarelo
        c.border = borda_fina
        c.alignment = Alignment(horizontal="center", vertical="center")

    podio_rows = []
    # 1. Novos/Reativados
    if rank1 is not None and not rank1.empty:
        top1 = rank1.iloc[0]
        podio_rows.append(("1º Ranking", "Maior Nº de Clientes Novos/Reativados", str(top1['Vendedor']), f"{int(top1['Qtd_Clientes'])} clientes (R$ {top1['Valor_Total']:,.2f})"))
    else:
        podio_rows.append(("1º Ranking", "Maior Nº de Clientes Novos/Reativados", "-", "Sem dados"))

    # 2. Propostas Enviadas
    if rank2 is not None and not rank2.empty:
        top2 = rank2.iloc[0]
        podio_rows.append(("2º Ranking", "Maior Nº de Propostas Enviadas", str(top2['Vendedor']), f"{int(top2['Total_Propostas'])} propostas"))
    else:
        podio_rows.append(("2º Ranking", "Maior Nº de Propostas Enviadas", "-", "Sem dados"))

    # 3. Maior Pedido Faturado
    if rank3 is not None and not rank3.empty:
        top3 = rank3.iloc[0]
        nota_val = int(top3['Nota']) if isinstance(top3['Nota'], (int, float)) and top3['Nota'] == int(top3['Nota']) else str(top3['Nota'])
        podio_rows.append(("3º Ranking", "Maior Pedido Faturado (Nota Única)", str(top3['Vendedor']), f"Nota {nota_val} (R$ {top3['Valor_Total']:,.2f})"))
    else:
        podio_rows.append(("3º Ranking", "Maior Pedido Faturado (Nota Única)", "-", "Sem dados"))

    # 4. Maior Venda Categoria
    if rank4 is not None and not rank4.empty:
        top4 = rank4.iloc[0]
        podio_rows.append(("4º Ranking", f"Maior Venda Categoria: {cat_nome}", str(top4['Vendedor']), f"R$ {top4['Total_Vendido']:,.2f}"))
    else:
        podio_rows.append(("4º Ranking", f"Maior Venda Categoria: {cat_nome}", "-", "Sem dados"))

    for pos, (r_rank, r_crit, r_vend, r_met) in enumerate(podio_rows, start=4):
        c_r = ws0.cell(row=pos, column=1, value=r_rank)
        c_c = ws0.cell(row=pos, column=2, value=r_crit)
        c_v = ws0.cell(row=pos, column=3, value=r_vend)
        c_m = ws0.cell(row=pos, column=4, value=r_met)
        for c in (c_r, c_c, c_v, c_m):
            c.font = fonte_destaque
            c.border = borda_fina
        c_r.alignment = Alignment(horizontal="center")

    for col in ws0.columns:
        letra = get_column_letter(col[0].column)
        ws0.column_dimensions[letra].width = max(max(len(str(c.value or '')) for c in col) + 4, 18)

    # ABA 2: RANKING 1 (NOVOS E REATIVADOS)
    ws1 = wb.create_sheet("1. Novos e Reativados")
    ws1.append(["POSIÇÃO", "VENDEDOR", "QTD CLIENTES", "VALOR TOTAL PEDIDOS (DESEMPATE)"])
    for c in ws1[1]:
        c.font = fonte_cabecalho
        c.fill = fill_cab_azul
        c.border = borda_fina
    if rank1 is not None and not rank1.empty:
        for pos, r in enumerate(rank1.itertuples(), 1):
            c_pos = ws1.cell(row=pos+1, column=1, value=f"{pos}º")
            c_vend = ws1.cell(row=pos+1, column=2, value=str(r.Vendedor))
            c_qtd = ws1.cell(row=pos+1, column=3, value=int(r.Qtd_Clientes))
            c_val = ws1.cell(row=pos+1, column=4, value=float(r.Valor_Total))
            c_val.number_format = 'R$ #,##0.00'
            for c in (c_pos, c_vend, c_qtd, c_val):
                c.font = fonte_destaque if pos == 1 else fonte_dados
                c.border = borda_fina
            c_pos.alignment = Alignment(horizontal="center")
            c_qtd.alignment = Alignment(horizontal="center")

    for col in ws1.columns:
        letra = get_column_letter(col[0].column)
        ws1.column_dimensions[letra].width = max(max(len(str(c.value or '')) for c in col) + 4, 16)

    # ABA 3: RANKING 2 (PROPOSTAS)
    ws2 = wb.create_sheet("2. Propostas Enviadas")
    ws2.append(["POSIÇÃO", "VENDEDOR", "TOTAL PROPOSTAS ENVIADAS"])
    for c in ws2[1]:
        c.font = fonte_cabecalho
        c.fill = fill_cab_azul
        c.border = borda_fina
    if rank2 is not None and not rank2.empty:
        for pos, r in enumerate(rank2.itertuples(), 1):
            c_pos = ws2.cell(row=pos+1, column=1, value=f"{pos}º")
            c_vend = ws2.cell(row=pos+1, column=2, value=str(r.Vendedor))
            c_qtd = ws2.cell(row=pos+1, column=3, value=int(r.Total_Propostas))
            for c in (c_pos, c_vend, c_qtd):
                c.font = fonte_destaque if pos == 1 else fonte_dados
                c.border = borda_fina
            c_pos.alignment = Alignment(horizontal="center")
            c_qtd.alignment = Alignment(horizontal="center")

    for col in ws2.columns:
        letra = get_column_letter(col[0].column)
        ws2.column_dimensions[letra].width = max(max(len(str(c.value or '')) for c in col) + 4, 16)

    # ABA 4: RANKING 3 (MAIOR PEDIDO)
    ws3 = wb.create_sheet("3. Maior Pedido Faturado")
    ws3.append(["POSIÇÃO", "VENDEDOR", "NOTA / PEDIDO", "VALOR TOTAL FATURADO"])
    for c in ws3[1]:
        c.font = fonte_cabecalho
        c.fill = fill_cab_azul
        c.border = borda_fina
    if rank3 is not None and not rank3.empty:
        for pos, r in enumerate(rank3.head(20).itertuples(), 1):
            c_pos = ws3.cell(row=pos+1, column=1, value=f"{pos}º")
            c_vend = ws3.cell(row=pos+1, column=2, value=str(r.Vendedor))
            nota_val = int(r.Nota) if pd.notna(r.Nota) and isinstance(r.Nota, (int, float)) and r.Nota == int(r.Nota) else str(r.Nota)
            c_nota = ws3.cell(row=pos+1, column=3, value=nota_val)
            c_val = ws3.cell(row=pos+1, column=4, value=float(r.Valor_Total))
            c_val.number_format = 'R$ #,##0.00'
            for c in (c_pos, c_vend, c_nota, c_val):
                c.font = fonte_destaque if pos == 1 else fonte_dados
                c.border = borda_fina
            c_pos.alignment = Alignment(horizontal="center")
            c_nota.alignment = Alignment(horizontal="center")

    for col in ws3.columns:
        letra = get_column_letter(col[0].column)
        ws3.column_dimensions[letra].width = max(max(len(str(c.value or '')) for c in col) + 4, 16)

    # ABA 5: RANKING 4 (CATEGORIA)
    ws4 = wb.create_sheet("4. Vendas por Categoria")
    ws4.append(["POSIÇÃO", "VENDEDOR", f"TOTAL VENDIDO ({cat_nome})"])
    for c in ws4[1]:
        c.font = fonte_cabecalho
        c.fill = fill_cab_azul
        c.border = borda_fina
    if rank4 is not None and not rank4.empty:
        for pos, r in enumerate(rank4.itertuples(), 1):
            c_pos = ws4.cell(row=pos+1, column=1, value=f"{pos}º")
            c_vend = ws4.cell(row=pos+1, column=2, value=str(r.Vendedor))
            c_val = ws4.cell(row=pos+1, column=3, value=float(r.Total_Vendido))
            c_val.number_format = 'R$ #,##0.00'
            for c in (c_pos, c_vend, c_val):
                c.font = fonte_destaque if pos == 1 else fonte_dados
                c.border = borda_fina
            c_pos.alignment = Alignment(horizontal="center")

    for col in ws4.columns:
        letra = get_column_letter(col[0].column)
        ws4.column_dimensions[letra].width = max(max(len(str(c.value or '')) for c in col) + 4, 18)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ========================================================
# MODAL DIALOGS: UPLOAD AJUSTADA E RESULTADOS FINAIS
# ========================================================
@st.dialog("📤 UPLOAD DE PLANILHA AJUSTADA", width="large")
def modal_upload_ajustada():
    st.markdown("### 📤 Upload de Planilha Ajustada (Pós-Edição Manual)")
    st.caption("Utilize este campo quando tiver realizado ajustes no Excel baixado (ex: caso um cliente tenha mais de 1 pedido ou ajustes de brindes). As planilhas ajustadas substituirão as bases brutas na apuração dos **Resultados Finais**.")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        st.markdown("#### 👤 Planilha Ajustada — FRANCINE")
        file_fra = st.file_uploader("Selecione o arquivo Excel ajustado de Francine:", type=["xlsx", "xls"], key="modal_up_fra")
        if file_fra is not None:
            dados_fra = carregar_planilha_ajustada(file_fra)
            if 'df_aba1' in dados_fra:
                st.session_state['df_aba1_francine_ajustada'] = dados_fra['df_aba1']
                st.session_state['df_aba2_francine_ajustada'] = dados_fra.get('df_aba2', pd.DataFrame())
                st.session_state['nome_arquivo_ajustada_fra'] = file_fra.name
                st.success(f"✅ Francine: {len(dados_fra['df_aba1'])} clientes novos/reativados carregados!")

        if st.session_state.get('df_aba1_francine_ajustada') is not None:
            st.info(f"Arquivo ativo: `{st.session_state.get('nome_arquivo_ajustada_fra', 'Ajustada')}`")
            if st.button("🗑️ Remover Planilha Ajustada de Francine", key="del_aj_fra"):
                st.session_state.pop('df_aba1_francine_ajustada', None)
                st.session_state.pop('df_aba2_francine_ajustada', None)
                st.session_state.pop('nome_arquivo_ajustada_fra', None)
                st.rerun()

    with col_u2:
        st.markdown("#### 👤 Planilha Ajustada — FILIPE")
        file_fil = st.file_uploader("Selecione o arquivo Excel ajustado de Filipe:", type=["xlsx", "xls"], key="modal_up_fil")
        if file_fil is not None:
            dados_fil = carregar_planilha_ajustada(file_fil)
            if 'df_aba1' in dados_fil:
                st.session_state['df_aba1_filipe_ajustada'] = dados_fil['df_aba1']
                st.session_state['df_aba2_filipe_ajustada'] = dados_fil.get('df_aba2', pd.DataFrame())
                st.session_state['nome_arquivo_ajustada_fil'] = file_fil.name
                st.success(f"✅ Filipe: {len(dados_fil['df_aba1'])} clientes novos/reativados carregados!")

        if st.session_state.get('df_aba1_filipe_ajustada') is not None:
            st.info(f"Arquivo ativo: `{st.session_state.get('nome_arquivo_ajustada_fil', 'Ajustada')}`")
            if st.button("🗑️ Remover Planilha Ajustada de Filipe", key="del_aj_fil"):
                st.session_state.pop('df_aba1_filipe_ajustada', None)
                st.session_state.pop('df_aba2_filipe_ajustada', None)
                st.session_state.pop('nome_arquivo_ajustada_fil', None)
                st.rerun()

    st.markdown("---")
    if st.button("Concluir e Salvar", type="primary", use_container_width=True):
        st.rerun()


@st.dialog("🏆 RESULTADOS FINAIS DA ROLETA COMERCIAL", width="large")
def modal_resultados_finais():
    st.markdown("## 🏆 Resultados Finais Consolidados (Roleta Comercial)")
    st.caption("Cruzamento e consolidação dos resultados de **FRANCINE** e **FILIPE** para eleição dos campeões.")

    # 1. Recupera dados de FRANCINE (prioriza versão ajustada se enviada)
    df1_fra = st.session_state.get('df_aba1_francine_ajustada')
    fra_usando_ajustada = df1_fra is not None
    if df1_fra is None:
        df1_fra = st.session_state.get('df_aba1_francine')

    df2_fra = st.session_state.get('df_aba2_francine_ajustada')
    if df2_fra is None:
        df2_fra = st.session_state.get('df_aba2_francine')

    df_top_fra = st.session_state.get('df_top_notas_francine', pd.DataFrame())
    df_cat_fra = st.session_state.get('df_categorias_raw_francine')
    df_rank_cat_fra = st.session_state.get('df_rank_cat_francine', pd.DataFrame())

    # 2. Recupera dados de FILIPE (prioriza versão ajustada se enviada)
    df1_fil = st.session_state.get('df_aba1_filipe_ajustada')
    fil_usando_ajustada = df1_fil is not None
    if df1_fil is None:
        df1_fil = st.session_state.get('df_aba1_filipe')

    df2_fil = st.session_state.get('df_aba2_filipe_ajustada')
    if df2_fil is None:
        df2_fil = st.session_state.get('df_aba2_filipe')

    df_top_fil = st.session_state.get('df_top_notas_filipe', pd.DataFrame())
    df_cat_fil = st.session_state.get('df_categorias_raw_filipe')
    df_rank_cat_fil = st.session_state.get('df_rank_cat_filipe', pd.DataFrame())

    tem_dados_fra = (df1_fra is not None and not df1_fra.empty) or (df2_fra is not None and not df2_fra.empty)
    tem_dados_fil = (df1_fil is not None and not df1_fil.empty) or (df2_fil is not None and not df2_fil.empty)

    if not (tem_dados_fra or tem_dados_fil):
        st.warning("⚠️ Atenção: Nenhuma planilha foi gerada ou enviada ainda. Por favor, processe as planilhas de Francine e/ou Filipe nos painéis abaixo, ou faça o upload das planilhas ajustadas no botão **'UPLOAD DE PLANILHA AJUSTADA'**.")
        return

    # Banner de fontes de dados utilizadas
    info_fra = "📝 Ajustada" if fra_usando_ajustada else ("✅ Gerada" if tem_dados_fra else "⚪ Sem dados")
    info_fil = "📝 Ajustada" if fil_usando_ajustada else ("✅ Gerada" if tem_dados_fil else "⚪ Sem dados")
    st.info(f"📊 **Bases utilizadas para apuração:** Francine (**{info_fra}**) | Filipe (**{info_fil}**)")

    # ========================================================
    # CÁLCULO DOS 4 RANKINGS OFICIAIS
    # ========================================================

    # --- 1. MAIOR NÚMERO DE CLIENTES NOVOS / REATIVADOS ---
    # Critério: soma clientes por vendedor. Desempate: valor dos pedidos
    p_df1_fra = padronizar_df_aba1(df1_fra)
    p_df1_fil = padronizar_df_aba1(df1_fil)
    df1_comb = pd.concat([p_df1_fra, p_df1_fil], ignore_index=True)

    rank1 = pd.DataFrame()
    if not df1_comb.empty and 'Vendedor' in df1_comb.columns:
        rank1 = (
            df1_comb.groupby('Vendedor', as_index=False)
            .agg(
                Qtd_Clientes=('Cliente', 'count') if 'Cliente' in df1_comb.columns else ('Vendedor', 'count'),
                Valor_Total=('VALOR_NUM', 'sum') if 'VALOR_NUM' in df1_comb.columns else ('Vendedor', lambda x: 0.0)
            )
            .sort_values(by=['Qtd_Clientes', 'Valor_Total'], ascending=[False, False])
            .reset_index(drop=True)
        )

    # --- 2. MAIOR NÚMERO DE PROPOSTAS ENVIADAS ---
    p_df2_fra = padronizar_df_aba2(df2_fra)
    p_df2_fil = padronizar_df_aba2(df2_fil)
    df2_comb = pd.concat([p_df2_fra, p_df2_fil], ignore_index=True)

    rank2 = pd.DataFrame()
    if not df2_comb.empty and 'Vendedor' in df2_comb.columns:
        rank2 = (
            df2_comb.groupby('Vendedor', as_index=False)['Qtd']
            .sum()
            .rename(columns={'Qtd': 'Total_Propostas'})
            .sort_values(by='Total_Propostas', ascending=False)
            .reset_index(drop=True)
        )

    # --- 3. MAIOR PEDIDO FATURADO ---
    df_notas_comb = pd.concat([df_top_fra, df_top_fil], ignore_index=True)
    rank3 = pd.DataFrame()
    if not df_notas_comb.empty and 'Valor_Total' in df_notas_comb.columns:
        rank3 = (
            df_notas_comb.sort_values(by='Valor_Total', ascending=False)
            .drop_duplicates(subset=['Nota', 'Vendedor'] if 'Nota' in df_notas_comb.columns and 'Vendedor' in df_notas_comb.columns else None)
            .reset_index(drop=True)
        )
    elif not df1_comb.empty and 'VALOR_NUM' in df1_comb.columns:
        # Fallback caso notas não tenham sido carregadas separadamente
        rank3 = (
            df1_comb.sort_values(by='VALOR_NUM', ascending=False)
            .rename(columns={'Pedido': 'Nota', 'VALOR_NUM': 'Valor_Total'})
            .reset_index(drop=True)
        )

    # --- 4. MAIOR VENDA DA CATEGORIA DETERMINADA PREVIAMENTE ---
    cat_sugerida = st.session_state.get('categoria_solicitada_francine') or st.session_state.get('categoria_solicitada_filipe') or "HIGIENE E LIMPEZA"
    idx_cat_final = CATEGORIAS_ROLETA.index(cat_sugerida) if cat_sugerida in CATEGORIAS_ROLETA else 0

    col_sel_cat, _ = st.columns([1, 1])
    with col_sel_cat:
        cat_escolhida = st.selectbox(
            "🏷️ Categoria Avaliada para o 4º Ranking:",
            CATEGORIAS_ROLETA,
            index=idx_cat_final,
            key="modal_sel_cat_final"
        )

    # Se tiver bases brutas de categorias, reextrai com precisão para a categoria selecionada
    partes_cat = []
    if df_cat_fra is not None and not df_cat_fra.empty:
        df_r_fra, _, _ = extrair_ranking_categoria(df_cat_fra, cat_escolhida, st.session_state.get('df_fortpel_francine'), st.session_state.get('df_adm_raw_francine'))
        if not df_r_fra.empty:
            partes_cat.append(df_r_fra)
    elif not df_rank_cat_fra.empty and st.session_state.get('categoria_solicitada_francine') == cat_escolhida:
        partes_cat.append(df_rank_cat_fra)

    if df_cat_fil is not None and not df_cat_fil.empty:
        df_r_fil, _, _ = extrair_ranking_categoria(df_cat_fil, cat_escolhida, st.session_state.get('df_fortpel_filipe'), st.session_state.get('df_adm_raw_filipe'))
        if not df_r_fil.empty:
            partes_cat.append(df_r_fil)
    elif not df_rank_cat_fil.empty and st.session_state.get('categoria_solicitada_filipe') == cat_escolhida:
        partes_cat.append(df_rank_cat_fil)

    rank4 = pd.DataFrame()
    if partes_cat:
        df_cat_comb = pd.concat(partes_cat, ignore_index=True)
        col_v_cat = identificar_coluna(df_cat_comb.columns, ["vendedor"])
        col_val_cat = identificar_coluna(df_cat_comb.columns, ["total_vendido", "valor_total", "valor"])
        if col_v_cat and col_val_cat:
            df_cat_comb['Vendedor'] = df_cat_comb[col_v_cat].astype(str).str.strip().str.upper()
            df_cat_comb['Total_Vendido'] = df_cat_comb[col_val_cat].apply(limpar_valor)
            rank4 = (
                df_cat_comb.groupby('Vendedor', as_index=False)['Total_Vendido']
                .sum()
                .sort_values(by='Total_Vendido', ascending=False)
                .reset_index(drop=True)
            )

    # ========================================================
    # PAINEL COM OS 4 CAMPEÕES EM DESTAQUE (CARDS NO TOPO)
    # ========================================================
    st.markdown("### 🥇 Vendedores Campeões da Roleta")
    c_m1, c_m2 = st.columns(2)
    c_m3, c_m4 = st.columns(2)

    with c_m1:
        st.markdown("**1 - Maior Nº de Clientes Novos / Reativados**")
        if not rank1.empty:
            top1 = rank1.iloc[0]
            st.metric(
                label=f"🥇 {top1['Vendedor']}",
                value=f"{int(top1['Qtd_Clientes'])} Clientes",
                delta=f"R$ {top1['Valor_Total']:,.2f} (Critério Desempate)".replace(',', 'X').replace('.', ',').replace('X', '.')
            )
        else:
            st.info("Sem dados de novos/reativados.")

    with c_m2:
        st.markdown("**2 - Maior Nº de Propostas Enviadas**")
        if not rank2.empty:
            top2 = rank2.iloc[0]
            st.metric(
                label=f"🥇 {top2['Vendedor']}",
                value=f"{int(top2['Total_Propostas'])} Propostas"
            )
        else:
            st.info("Sem dados de propostas ADM.")

    with c_m3:
        st.markdown("**3 - Maior Pedido Faturado (Nota Única)**")
        if not rank3.empty:
            top3 = rank3.iloc[0]
            nota_val = int(top3['Nota']) if isinstance(top3['Nota'], (int, float)) and top3['Nota'] == int(top3['Nota']) else str(top3['Nota'])
            val_fmt = f"R$ {top3['Valor_Total']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            st.metric(
                label=f"🥇 {top3['Vendedor']}",
                value=val_fmt,
                delta=f"Nota/Pedido: {nota_val}"
            )
        else:
            st.info("Sem dados de notas/pedidos faturados.")

    with c_m4:
        st.markdown(f"**4 - Maior Venda Categoria: `{cat_escolhida}`**")
        if not rank4.empty:
            top4 = rank4.iloc[0]
            val_cat_fmt = f"R$ {top4['Total_Vendido']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            st.metric(
                label=f"🥇 {top4['Vendedor']}",
                value=val_cat_fmt,
                delta=f"Categoria: {cat_escolhida}"
            )
        else:
            st.info("Sem dados da categoria selecionada.")

    st.markdown("---")

    # ========================================================
    # TABELAS DETALHADAS DE CADA RANKING
    # ========================================================
    st.markdown("### 📋 Classificação Completa por Critério")
    tabs_finais = st.tabs([
        "👥 1. Novos / Reativados",
        "📑 2. Propostas Enviadas",
        "📦 3. Maior Pedido Faturado",
        f"🏷️ 4. Vendas {cat_escolhida}"
    ])

    with tabs_finais[0]:
        if not rank1.empty:
            st.caption("Classificado pelo maior número de clientes. Em caso de empate, desempatado pelo valor total dos pedidos.")
            df_disp1 = rank1.copy()
            df_disp1.insert(0, 'Posição', [f"🥇 1º" if i == 0 else (f"🥈 2º" if i == 1 else (f"🥉 3º" if i == 2 else f"{i+1}º")) for i in range(len(df_disp1))])
            df_disp1['Valor_Total'] = df_disp1['Valor_Total'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            df_disp1 = df_disp1.rename(columns={'Qtd_Clientes': 'Qtd Clientes Novos/Reat.', 'Valor_Total': 'Valor Total dos Pedidos'})
            st.dataframe(df_disp1, use_container_width=True)
        else:
            st.write("Sem registros para este critério.")

    with tabs_finais[1]:
        if not rank2.empty:
            df_disp2 = rank2.copy()
            df_disp2.insert(0, 'Posição', [f"🥇 1º" if i == 0 else (f"🥈 2º" if i == 1 else (f"🥉 3º" if i == 2 else f"{i+1}º")) for i in range(len(df_disp2))])
            df_disp2 = df_disp2.rename(columns={'Total_Propostas': 'Total de Propostas Enviadas'})
            st.dataframe(df_disp2, use_container_width=True)
        else:
            st.write("Sem registros para este critério.")

    with tabs_finais[2]:
        if not rank3.empty:
            df_disp3 = rank3.head(20).copy()
            df_disp3.insert(0, 'Posição', [f"🥇 1º" if i == 0 else (f"🥈 2º" if i == 1 else (f"🥉 3º" if i == 2 else f"{i+1}º")) for i in range(len(df_disp3))])
            df_disp3['Valor_Total'] = df_disp3['Valor_Total'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            df_disp3 = df_disp3.rename(columns={'Nota': 'Nota Fiscal / Pedido', 'Valor_Total': 'Valor Total Faturado'})
            st.dataframe(df_disp3, use_container_width=True)
        else:
            st.write("Sem registros para este critério.")

    with tabs_finais[3]:
        if not rank4.empty:
            df_disp4 = rank4.copy()
            df_disp4.insert(0, 'Posição', [f"🥇 1º" if i == 0 else (f"🥈 2º" if i == 1 else (f"🥉 3º" if i == 2 else f"{i+1}º")) for i in range(len(df_disp4))])
            df_disp4['Total_Vendido'] = df_disp4['Total_Vendido'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            df_disp4 = df_disp4.rename(columns={'Total_Vendido': f'Total Vendido ({cat_escolhida})'})
            st.dataframe(df_disp4, use_container_width=True)
        else:
            st.write(f"Sem registros para a categoria {cat_escolhida}.")

    st.markdown("---")

    # ========================================================
    # EXPORTAÇÃO COMPLETA DOS RESULTADOS EM EXCEL
    # ========================================================
    bytes_excel_final = gerar_excel_resultados_finais(rank1, rank2, rank3, rank4, cat_escolhida)
    nome_arq_final = f"Resultados_Finais_Roleta_Consolidado_{datetime.date.today().strftime('%d%m%Y')}.xlsx"

    col_exp1, col_exp2 = st.columns([2, 1])
    with col_exp1:
        st.download_button(
            label="📥 Exportar Relatório Consolidado dos Resultados Finais (.xlsx)",
            data=bytes_excel_final,
            file_name=nome_arq_final,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            key="dl_btn_resultados_finais"
        )
    with col_exp2:
        if st.button("Fechar Tela de Resultados", use_container_width=True):
            st.rerun()


# ========================================================
# CABEÇALHO E SELETOR DO OPERADOR (FRANCINE / FILIPE)
# ========================================================
st.title("🎯 Sistema Roleta - Controle Comercial")

status_fra_txt = "📝 Ajustada" if st.session_state.get('df_aba1_francine_ajustada') is not None else ("✅ Gerado" if st.session_state.get('processado_francine') else "⚪ Pendente")
status_fil_txt = "📝 Ajustada" if st.session_state.get('df_aba1_filipe_ajustada') is not None else ("✅ Gerado" if st.session_state.get('processado_filipe') else "⚪ Pendente")

col_top_sel, col_top_status = st.columns([2, 1])
with col_top_sel:
    operador_ativo = st.radio(
        "👤 Selecione o Operador / Perfil:",
        ["FRANCINE", "FILIPE"],
        horizontal=True,
        key="sel_operador_global",
        help="Cada operador possui sua própria área de trabalho e gera sua planilha independente."
    )
with col_top_status:
    st.markdown("**Status dos Relatórios:**")
    st.caption(f"FRANCINE: **{status_fra_txt}** | FILIPE: **{status_fil_txt}**")

# ========================================================
# BOTÕES DE AÇÕES GLOBAIS (PLANILHA AJUSTADA E RESULTADOS FINAIS)
# ========================================================
col_btn_acoes1, col_btn_acoes2 = st.columns(2)
with col_btn_acoes1:
    if st.button("📤 UPLOAD DE PLANILHA AJUSTADA", use_container_width=True, help="Clique para fazer o upload da planilha ajustada no Excel"):
        modal_upload_ajustada()

with col_btn_acoes2:
    if st.button("🏆 GERAR RESULTADOS FINAIS", type="primary", use_container_width=True, help="Cruza as planilhas de Francine e Filipe e abre a tela com os 4 rankings consolidados"):
        modal_resultados_finais()

# Alerta caso alguma planilha ajustada esteja ativa
tem_ajustada_fra = st.session_state.get('df_aba1_francine_ajustada') is not None
tem_ajustada_fil = st.session_state.get('df_aba1_filipe_ajustada') is not None
if tem_ajustada_fra or tem_ajustada_fil:
    nomes_aj = []
    if tem_ajustada_fra: nomes_aj.append("FRANCINE")
    if tem_ajustada_fil: nomes_aj.append("FILIPE")
    st.success(f"📝 **Planilha Ajustada ativa para:** {', '.join(nomes_aj)}. O cálculo dos Resultados Finais priorizará os dados ajustados automaticamente!")

st.markdown("---")


# ========================================================
# FUNÇÃO PARA RENDERIZAR O PAINEL DE UM OPERADOR ESPECÍFICO
# ========================================================
def renderizar_painel_operador(nome_op):
    prefixo = nome_op.lower()

    st.subheader(f"📋 Painel de Controle Comercial — {nome_op}")
    st.caption(f"Arquivos e parâmetros específicos para a planilha de **{nome_op}**.")

    # --- SEÇÃO 1: UPLOAD DOS 6 CAMPOS ---
    st.markdown("#### 1. Arquivos de Entrada (Bases BI, Fortpel, ADM, Categorias e NFs)")

    modo_leitura = st.radio(
        f"Como deseja carregar os arquivos para {nome_op}?",
        ["Fazer Upload pelo Navegador", f"Ler direto da pasta local ({PASTA_ENTRADA.name} ou raiz)"],
        horizontal=True,
        key=f"modo_leitura_{prefixo}"
    )

    file_novos = None
    file_reativados = None
    file_fortpel = None
    file_adm = None
    file_categorias = None
    file_nf = None
    categoria_solicitada = "HIGIENE E LIMPEZA"

    if modo_leitura == "Fazer Upload pelo Navegador":
        c1, c2 = st.columns(2)
        with c1:
            file_novos = st.file_uploader(f"📥 1. BI Clientes Novos ({nome_op})", type=["xlsx", "xls", "csv"], key=f"up_novos_{prefixo}")
        with c2:
            file_reativados = st.file_uploader(f"📥 2. BI Clientes Reativados ({nome_op})", type=["xlsx", "xls", "csv"], key=f"up_reat_{prefixo}")

        c3, c4 = st.columns(2)
        with c3:
            file_fortpel = st.file_uploader(f"📥 3. Acompanhamento Fortpel ({nome_op})", type=["xlsx", "xls", "csv"], key=f"up_fort_{prefixo}")
        with c4:
            file_adm = st.file_uploader(f"📥 4. Planilha ADM ({nome_op})", type=["xlsx", "xls", "csv"], key=f"up_adm_{prefixo}")

        c5, c6 = st.columns(2)
        with c5:
            file_categorias = st.file_uploader(f"📥 5. Planilha de Categorias do BI ({nome_op})", type=["xlsx", "xls", "csv"], key=f"up_cat_{prefixo}")
            categoria_solicitada = st.selectbox(
                "🏷️ Categoria Solicitada (para o Campo 5):",
                CATEGORIAS_ROLETA,
                index=3,
                help="Selecione a categoria para identificar o vendedor com a maior venda.",
                key=f"sel_cat_req_{prefixo}"
            )
        with c6:
            file_nf = st.file_uploader(f"📥 6. Planilha de NFs Faturadas ({nome_op})", type=["xlsx", "xls", "csv"], key=f"up_nf_{prefixo}")

    else:
        arquivos_locais = sorted(list(set(
            [f.name for f in PASTA_ENTRADA.iterdir() if f.is_file() and not f.name.startswith('~$') and f.suffix.lower() in ['.xlsx', '.xls', '.csv']] +
            [f.name for f in BASE_DIR.iterdir() if f.is_file() and not f.name.startswith('~$') and f.suffix.lower() in ['.xlsx', '.xls', '.csv']]
        )))

        def resolver_caminho(nome):
            if (PASTA_ENTRADA / nome).exists():
                return PASTA_ENTRADA / nome
            return BASE_DIR / nome

        if not arquivos_locais:
            st.info(f"Nenhum arquivo encontrado em `{PASTA_ENTRADA}` ou `{BASE_DIR}`. Coloque as planilhas lá ou use a opção de Upload.")
        else:
            padrao_novos = buscar_arquivo_padrao(arquivos_locais, ["novo"])
            padrao_reat = buscar_arquivo_padrao(arquivos_locais, ["reativ"])
            padrao_fort = buscar_arquivo_padrao(arquivos_locais, ["acomp", "fortpel"])
            padrao_adm = buscar_arquivo_padrao(arquivos_locais, ["adm", "orcamento", "proposta"])
            padrao_cat = buscar_arquivo_padrao(arquivos_locais, ["categoria", "cat", "data (7", "data (8"])
            padrao_nf = buscar_arquivo_padrao(arquivos_locais, ["prod", "nota", "fatur", "nf"])

            opcoes = ["Selecione..."] + arquivos_locais

            c1, c2 = st.columns(2)
            with c1:
                idx_n = opcoes.index(padrao_novos) if padrao_novos in opcoes else 0
                nome_n = st.selectbox(f"1. BI Clientes Novos ({nome_op}):", opcoes, index=idx_n, key=f"loc_novos_{prefixo}")
                if nome_n != "Selecione...":
                    file_novos = resolver_caminho(nome_n)

            with c2:
                idx_r = opcoes.index(padrao_reat) if padrao_reat in opcoes else 0
                nome_r = st.selectbox(f"2. BI Clientes Reativados ({nome_op}):", opcoes, index=idx_r, key=f"loc_reat_{prefixo}")
                if nome_r != "Selecione...":
                    file_reativados = resolver_caminho(nome_r)

            c3, c4 = st.columns(2)
            with c3:
                idx_f = opcoes.index(padrao_fort) if padrao_fort in opcoes else 0
                nome_f = st.selectbox(f"3. Acompanhamento Fortpel ({nome_op}):", opcoes, index=idx_f, key=f"loc_fort_{prefixo}")
                if nome_f != "Selecione...":
                    file_fortpel = resolver_caminho(nome_f)

            with c4:
                idx_a = opcoes.index(padrao_adm) if padrao_adm in opcoes else 0
                nome_a = st.selectbox(f"4. Planilha ADM ({nome_op}):", opcoes, index=idx_a, key=f"loc_adm_{prefixo}")
                if nome_a != "Selecione...":
                    file_adm = resolver_caminho(nome_a)

            c5, c6 = st.columns(2)
            with c5:
                idx_p = opcoes.index(padrao_cat) if padrao_cat in opcoes else 0
                nome_p = st.selectbox(f"5. Planilha de Categorias ({nome_op}):", opcoes, index=idx_p, key=f"loc_cat_{prefixo}")
                if nome_p != "Selecione...":
                    file_categorias = resolver_caminho(nome_p)
                categoria_solicitada = st.selectbox(
                    "🏷️ Categoria Solicitada (para o Campo 5):",
                    CATEGORIAS_ROLETA,
                    index=3,
                    help="Selecione a categoria para identificar o vendedor com a maior venda.",
                    key=f"sel_cat_req_loc_{prefixo}"
                )

            with c6:
                idx_nf = opcoes.index(padrao_nf) if padrao_nf in opcoes else 0
                nome_nf = st.selectbox(f"6. Planilha de NFs Faturadas ({nome_op}):", opcoes, index=idx_nf, key=f"loc_nf_{prefixo}")
                if nome_nf != "Selecione...":
                    file_nf = resolver_caminho(nome_nf)

    # --- SEÇÃO 2: CONFIGURAÇÕES DE PERÍODO E PARÂMETROS ---
    st.markdown("#### 2. Configuração de Período e Parâmetros")

    hoje = datetime.date.today()
    inicio_padrao = hoje.replace(day=1)
    fim_padrao = hoje

    c_p1, c_p2 = st.columns(2)
    with c_p1:
        data_sel = st.date_input(
            f"📅 Intervalo de Datas ({nome_op}):",
            value=(inicio_padrao, fim_padrao),
            format="DD/MM/YYYY",
            key=f"data_sel_{prefixo}"
        )

        modo_filtro_roleta = st.radio(
            f"Filtro de Data para a Roleta (Aba 1 - {nome_op}):",
            [
                "Opção B: Filtrar pelo intervalo de datas selecionado (Inclusivo, recomendado)",
                "Opção A: Considerar tudo que vier no BI (Sem restrição de data na Roleta)"
            ],
            index=0,
            key=f"modo_filtro_{prefixo}"
        )

    with c_p2:
        criterio_roleta = st.radio(
            f"Critério de Contagem da Roleta (Aba 1 - {nome_op}):",
            [
                "1 Registro por Cliente Único (Padrão BI - 1 linha por cliente, soma exata)",
                "Todos os Pedidos Individuais (Pode gerar múltiplas linhas se o cliente fez mais de um pedido)"
            ],
            index=0,
            key=f"criterio_roleta_{prefixo}"
        )

        origem_valor_aba1 = st.radio(
            f"Origem do Valor da Aba 1 (Novos/Reativados - {nome_op}):",
            [
                "Usar coluna 'Venda' do BI (Novos e Reativados) [Recomendado se o Acompanhamento tiver total geral]",
                "Usar coluna de Valor do Acompanhamento Fortpel"
            ],
            index=0,
            key=f"origem_valor_{prefixo}"
        )

    # --- SEÇÃO 3: BOTÃO DE PROCESSAMENTO ---
    st.markdown("---")
    btn_col1, btn_col2, btn_col3 = st.columns([1, 2, 1])
    with btn_col2:
        botao_gerar = st.button(
            f"🚀 Gerar Planilha de {nome_op}",
            type="primary",
            use_container_width=True,
            key=f"btn_gerar_{prefixo}"
        )

    fingerprint_atual = (
        obter_id_arquivo(file_novos),
        obter_id_arquivo(file_reativados),
        obter_id_arquivo(file_fortpel),
        obter_id_arquivo(file_adm),
        obter_id_arquivo(file_categorias),
        obter_id_arquivo(file_nf),
        str(data_sel),
        modo_filtro_roleta,
        criterio_roleta,
        origem_valor_aba1
    )

    if st.session_state.get(f'fingerprint_{prefixo}') != fingerprint_atual:
        st.session_state[f'processado_{prefixo}'] = False

    # --- SEÇÃO 4: PROCESSAMENTO SOB DEMANDA ---
    if botao_gerar:
        if not (file_novos and file_reativados and file_fortpel and file_adm):
            st.error(f"⚠️ Atenção: Por favor, selecione ao menos as 4 bases principais (Novos, Reativados, Fortpel e ADM) de {nome_op} antes de gerar a planilha.")
        elif not (isinstance(data_sel, tuple) and len(data_sel) == 2):
            st.warning("⚠️ Atenção: Por favor, selecione a data de início e a data de fim no calendário do período.")
        else:
            dt_inicio, dt_fim = data_sel
            with st.spinner(f"⏳ Processando bases de {nome_op} e gerando tabelas da Roleta..."):
                try:
                    df_novos_raw = carregar_dados(file_novos)
                    df_reativados_raw = carregar_dados(file_reativados)
                    df_fortpel_raw = carregar_dados(file_fortpel)
                    df_adm_raw = carregar_dados(file_adm)
                    df_categorias_raw = carregar_dados(file_categorias) if file_categorias else None
                    df_nf_raw = carregar_dados(file_nf) if file_nf else None

                    col_novos_cli = identificar_coluna(df_novos_raw.columns, ["cliente", "razao", "nome"])
                    col_novos_val = identificar_coluna(df_novos_raw.columns, ["venda", "valor", "total", "faturado"])

                    col_reat_cli = identificar_coluna(df_reativados_raw.columns, ["cliente", "razao", "nome"])
                    col_reat_val = identificar_coluna(df_reativados_raw.columns, ["venda", "valor", "total", "faturado"])

                    col_f_cli = identificar_coluna(df_fortpel_raw.columns, ["cliente", "razao", "nome"])
                    col_f_ped = identificar_coluna(df_fortpel_raw.columns, ["pedido", "numero", "ordem"])
                    col_f_nota = identificar_coluna(df_fortpel_raw.columns, ["nota", "nf", "n_nota", "n_nota_obs", "pedido"])
                    col_f_vend = identificar_coluna(df_fortpel_raw.columns, ["carteira", "operador", "vendedor", "representante", "gestor"])
                    col_f_val = identificar_coluna(df_fortpel_raw.columns, ["valor", "total", "faturado", "preco", "somadevl"])
                    col_f_data = identificar_coluna(df_fortpel_raw.columns, ["dt_inser", "dt inser", "data inser", "dt_saida", "data", "emissao", "faturamento", "dt"])

                    col_a_vend = identificar_coluna(df_adm_raw.columns, ["vendedor", "operador", "usuario", "representante"])
                    col_a_status = identificar_coluna(df_adm_raw.columns, ["status", "situacao", "fase"])

                    df_fortpel = df_fortpel_raw.copy()
                    df_fortpel['__DATA_DT'] = pd.to_datetime(df_fortpel[col_f_data], dayfirst=True, errors='coerce')

                    # ABA 1: REATIVADOS - NOVOS
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
                        mask_data_f = (df_fortpel['__DATA_DT'].dt.date <= dt_fim) if pd.notna(dt_fim) else True
                        df_aba1 = df_fortpel[mask_data_f & mask_status].copy()
                    else:
                        mask_data_f = (df_fortpel['__DATA_DT'].dt.date >= dt_inicio) & (df_fortpel['__DATA_DT'].dt.date <= dt_fim)
                        df_aba1 = df_fortpel[mask_data_f & mask_status].copy()

                        if df_aba1.empty and mask_status.sum() > 0:
                            mask_data_f_flex = (df_fortpel['__DATA_DT'].dt.date <= dt_fim) if pd.notna(dt_fim) else True
                            df_aba1 = df_fortpel[mask_data_f_flex & mask_status].copy()

                    if "Cliente Único" in criterio_roleta:
                        df_aba1 = df_aba1.sort_values(by=['__DATA_DT', col_f_ped], ascending=[True, True])
                        df_aba1 = df_aba1.drop_duplicates(subset=['__CLI_NORM'], keep='first')

                    if "BI" in origem_valor_aba1 and col_novos_val and col_reat_val:
                        mapa_valores_bi = {}
                        for _, row in df_novos_clean.iterrows():
                            mapa_valores_bi[row['__CLI_NORM']] = limpar_valor(row[col_novos_val])
                        for _, row in df_reat_clean.iterrows():
                            mapa_valores_bi[row['__CLI_NORM']] = limpar_valor(row[col_reat_val])
                        df_aba1['VALOR_NUM'] = df_aba1['__CLI_NORM'].map(mapa_valores_bi).fillna(0.0)
                    else:
                        df_aba1['VALOR_NUM'] = df_aba1[col_f_val].apply(limpar_valor)

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

                    # ABA 2: PROPOSTAS ENVIADAS (TABELA DINÂMICA ADM)
                    df_adm_filtrado = df_adm_raw.dropna(subset=[col_a_vend]).copy()
                    df_aba2 = (
                        df_adm_filtrado.groupby(col_a_vend)[col_a_status]
                        .count()
                        .reset_index()
                    )
                    df_aba2.columns = ['Vendedor', 'Qtd']
                    df_aba2 = df_aba2.sort_values(by=['Qtd', 'Vendedor'], ascending=[False, True])

                    # 🏆 MAIOR VENDA EM NOTA ÚNICA
                    if df_nf_raw is not None and not df_nf_raw.empty:
                        df_top_notas = extrair_maior_venda_nf_faturada(df_nf_raw, dt_inicio=dt_inicio, dt_fim=dt_fim)
                    else:
                        df_top_notas = extrair_maior_venda_nota(
                            df_fortpel=df_fortpel,
                            col_nota=col_f_nota,
                            col_vend=col_f_vend,
                            col_val=col_f_val,
                            dt_inicio=dt_inicio,
                            dt_fim=dt_fim
                        )

                    # 🏷️ CATEGORIAS DO BI
                    df_rank_cat = pd.DataFrame()
                    total_geral_cat = 0.0
                    df_lideres_cat = pd.DataFrame()

                    if df_categorias_raw is not None:
                        df_rank_cat, total_geral_cat, df_lideres_cat = extrair_ranking_categoria(
                            df_cat_raw=df_categorias_raw,
                            categoria_alvo=categoria_solicitada,
                            df_fortpel=df_fortpel,
                            df_adm=df_adm_raw
                        )

                    # GERAÇÃO DO WORKBOOK EXCEL (3 ABAS)
                    wb = openpyxl.Workbook()

                    dia_i = dt_inicio.strftime('%d')
                    dia_f = dt_fim.strftime('%d')
                    nome_aba1 = f"{dia_i}-{dia_f} Reativados - Novos"[:31]
                    nome_aba2 = f"{dia_i}-{dia_f} Propostas Enviadas"[:31]

                    ws1 = wb.active
                    ws1.title = nome_aba1
                    if not df_aba1.empty:
                        df_aba1['VENDEDOR_CLEAN'] = df_aba1[col_f_vend].astype(str).str.strip().str.upper()
                        df_aba1 = df_aba1.sort_values(by=['VENDEDOR_CLEAN', '__DATA_DT', col_f_cli])
                        estilizar_aba(ws1, df_aba1, col_f_ped, col_f_val, col_f_cli, '__DATA_DT', titulo_col_id="PEDIDO")
                    else:
                        ws1.cell(row=1, column=1, value="Nenhum pedido encontrado para Novos/Reativados no período.")

                    ws2 = wb.create_sheet(title=nome_aba2)
                    if not df_aba2.empty:
                        estilizar_aba_propostas(ws2, df_aba2)
                    else:
                        ws2.cell(row=1, column=1, value="Nenhum registro encontrado na planilha ADM.")

                    ws3 = wb.create_sheet(title="Destaques Roleta")
                    estilizar_aba_produtos(ws3, df_top_notas, df_lideres_cat)

                    nome_arquivo_saida = f"Controle_roleta_{nome_op}_{dt_inicio.strftime('%d%m')}_a_{dt_fim.strftime('%d%m')}.xlsx"
                    caminho_salvar_local = PASTA_SAIDA / nome_arquivo_saida
                    wb.save(caminho_salvar_local)

                    buffer_download = io.BytesIO()
                    wb.save(buffer_download)
                    buffer_download.seek(0)

                    # Salva no Session State sob o prefixo deste operador
                    st.session_state[f'processado_{prefixo}'] = True
                    st.session_state[f'fingerprint_{prefixo}'] = fingerprint_atual
                    st.session_state[f'df_aba1_{prefixo}'] = df_aba1
                    st.session_state[f'df_aba2_{prefixo}'] = df_aba2
                    st.session_state[f'df_top_notas_{prefixo}'] = df_top_notas
                    st.session_state[f'df_rank_cat_{prefixo}'] = df_rank_cat
                    st.session_state[f'total_geral_cat_{prefixo}'] = total_geral_cat
                    st.session_state[f'df_lideres_cat_{prefixo}'] = df_lideres_cat
                    st.session_state[f'df_categorias_raw_{prefixo}'] = df_categorias_raw
                    st.session_state[f'df_fortpel_{prefixo}'] = df_fortpel
                    st.session_state[f'df_adm_raw_{prefixo}'] = df_adm_raw
                    st.session_state[f'categoria_solicitada_{prefixo}'] = categoria_solicitada
                    st.session_state[f'col_f_ped_{prefixo}'] = col_f_ped
                    st.session_state[f'col_f_cli_{prefixo}'] = col_f_cli
                    st.session_state[f'col_f_vend_{prefixo}'] = col_f_vend
                    st.session_state[f'nome_aba1_{prefixo}'] = nome_aba1
                    st.session_state[f'nome_aba2_{prefixo}'] = nome_aba2
                    st.session_state[f'caminho_salvar_local_{prefixo}'] = caminho_salvar_local
                    st.session_state[f'buffer_bytes_{prefixo}'] = buffer_download.getvalue()
                    st.session_state[f'nome_arquivo_saida_{prefixo}'] = nome_arquivo_saida
                    st.session_state[f'mapeamento_colunas_{prefixo}'] = {
                        "BI Novos": f"Cliente: `{col_novos_cli}` | Valor: `{col_novos_val}`",
                        "BI Reativados": f"Cliente: `{col_reat_cli}` | Valor: `{col_reat_val}`",
                        "Acompanhamento Fortpel": f"Cliente: `{col_f_cli}` | Pedido: `{col_f_ped}` | Vendedor: `{col_f_vend}` | Valor: `{col_f_val}` | Data: `{col_f_data}`",
                        "Planilha ADM": f"Vendedor: `{col_a_vend}` | Status: `{col_a_status}`",
                        "NFs Faturadas (Campo 6)": "Ativo" if df_nf_raw is not None else "Não enviado (usando Fortpel)",
                        "Categorias (Campo 5)": "Ativo" if df_categorias_raw is not None else "Não enviado"
                    }

                except Exception as e:
                    st.error(f"❌ Ocorreu um erro ao processar as planilhas de {nome_op}: {e}")
                    st.session_state[f'processado_{prefixo}'] = False

    # --- SEÇÃO 5: EXIBIÇÃO DOS RESULTADOS DESTE OPERADOR ---
    if st.session_state.get(f'processado_{prefixo}', False):
        df_aba1_s = st.session_state[f'df_aba1_{prefixo}']
        df_aba2_s = st.session_state[f'df_aba2_{prefixo}']
        nome_aba1_s = st.session_state[f'nome_aba1_{prefixo}']
        nome_aba2_s = st.session_state[f'nome_aba2_{prefixo}']
        caminho_local_s = st.session_state[f'caminho_salvar_local_{prefixo}']
        buffer_bytes_s = st.session_state[f'buffer_bytes_{prefixo}']
        nome_saida_s = st.session_state[f'nome_arquivo_saida_{prefixo}']

        st.success(f"✅ Tabelas de {nome_op} processadas e consolidadas com sucesso!")
        st.info(f"📁 Arquivo Excel salvo em: `{caminho_local_s}`")

        abas_titulos = [
            f"Aba 1: {nome_aba1_s}",
            f"Aba 2: {nome_aba2_s}",
            "Aba 3: 🏆 Destaques Roleta (Nota Única & Categorias)"
        ]
        abas_res = st.tabs(abas_titulos)

        with abas_res[0]:
            if not df_aba1_s.empty:
                qtd_duas_compras = int(df_aba1_s['TEM_DUAS_COMPRAS'].sum()) if 'TEM_DUAS_COMPRAS' in df_aba1_s else 0
                if qtd_duas_compras > 0:
                    st.warning(f"⚠️ {qtd_duas_compras} cliente(s) com mais de uma compra no mesmo mês identificado(s) (destacados com 'DUAS COMPRAS' e linha em vermelho na planilha)!")
                st.dataframe(df_aba1_s[['VENDEDOR_CLEAN', st.session_state[f'col_f_ped_{prefixo}'], st.session_state[f'col_f_cli_{prefixo}'], 'VALOR_NUM', 'STATUS', '__DATA_DT', 'BRINDE ROLETA']], use_container_width=True)
            else:
                st.write(f"Sem registros no período para a Aba 1 ({nome_op}).")

        with abas_res[1]:
            if not df_aba2_s.empty:
                st.dataframe(df_aba2_s, use_container_width=True)
                st.metric("Total de Propostas Efetivadas", int(df_aba2_s['Qtd'].sum()))
            else:
                st.write(f"Sem registros na planilha ADM ({nome_op}).")

        with abas_res[2]:
            df_top_notas_s = st.session_state.get(f'df_top_notas_{prefixo}', pd.DataFrame())
            df_cat_raw_s = st.session_state.get(f'df_categorias_raw_{prefixo}')
            df_fort_s = st.session_state.get(f'df_fortpel_{prefixo}')
            df_adm_s = st.session_state.get(f'df_adm_raw_{prefixo}')
            cat_padrao = st.session_state.get(f'categoria_solicitada_{prefixo}', 'HIGIENE E LIMPEZA')

            st.markdown(f"### 🏆 Maior Venda em Nota Única — {nome_op}")
            st.caption("Calculado a partir da análise das NFs faturadas no período (Campo 6 ou base Fortpel).")

            if not df_top_notas_s.empty:
                top1_nota = df_top_notas_s.iloc[0]
                nota_formatada = int(top1_nota['Nota']) if isinstance(top1_nota['Nota'], (int, float)) and top1_nota['Nota'] == int(top1_nota['Nota']) else str(top1_nota['Nota'])
                val_formatado = f"R$ {top1_nota['Valor_Total']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

                c_m1, c_m2, c_m3 = st.columns(3)
                with c_m1:
                    st.metric("Vendedor Campeão (Nota Única)", str(top1_nota['Vendedor']))
                with c_m2:
                    st.metric("Nota Fiscal", nota_formatada)
                with c_m3:
                    st.metric("Valor Total da Nota", val_formatado)

                with st.expander("📋 Ver Ranking Completo das Maiores Vendas em Nota Única"):
                    df_notas_disp = df_top_notas_s.copy()
                    df_notas_disp['Valor_Total'] = df_notas_disp['Valor_Total'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                    st.dataframe(df_notas_disp, use_container_width=True)
            else:
                st.info(f"Informações de notas fiscais não disponíveis para o período ({nome_op}).")

            st.markdown("---")

            st.markdown(f"### 🏷️ Consulta de Maior Venda por Categoria Solicitada — {nome_op}")
            st.caption("Selecione a categoria desejada para visualizar o vendedor que fez a maior venda:")

            idx_cat_sel = CATEGORIAS_ROLETA.index(cat_padrao) if cat_padrao in CATEGORIAS_ROLETA else 0
            cat_escolhida = st.selectbox(
                f"Selecione a Categoria ({nome_op}):",
                CATEGORIAS_ROLETA,
                index=idx_cat_sel,
                key=f"sel_cat_destaque_{prefixo}"
            )

            if df_cat_raw_s is not None:
                df_rank_escolhido, total_escolhido, df_lideres_todas = extrair_ranking_categoria(
                    df_cat_raw=df_cat_raw_s,
                    categoria_alvo=cat_escolhida,
                    df_fortpel=df_fort_s,
                    df_adm=df_adm_s
                )

                if not df_rank_escolhido.empty:
                    lider_cat = df_rank_escolhido.iloc[0]
                    val_lider_fmt = f"R$ {lider_cat['Total_Vendido']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    total_geral_fmt = f"R$ {total_escolhido:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

                    col_c1, col_c2, col_c3 = st.columns(3)
                    with col_c1:
                        st.metric("Vendedor Líder da Categoria", str(lider_cat['Vendedor']))
                    with col_c2:
                        st.metric(f"Vendido por {lider_cat['Vendedor']}", val_lider_fmt)
                    with col_c3:
                        st.metric(f"Total Geral ({cat_escolhida})", total_geral_fmt)

                    st.markdown(f"**Ranking de Vendedores na categoria `{cat_escolhida}`:**")
                    df_cat_disp = df_rank_escolhido.copy()
                    df_cat_disp['Total_Vendido'] = df_cat_disp['Total_Vendido'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                    st.dataframe(df_cat_disp, use_container_width=True)

                if not df_lideres_todas.empty:
                    with st.expander("📊 Ver Tabela Geral de Líderes das 7 Categorias"):
                        df_lid_disp = df_lideres_todas.copy()
                        df_lid_disp['Valor_Total'] = df_lid_disp['Valor_Total'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                        df_lid_disp['Total_Categoria'] = df_lid_disp['Total_Categoria'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                        st.dataframe(df_lid_disp, use_container_width=True)
            else:
                st.info(f"ℹ️ Para consultar os vendedores líderes por categoria, faça o upload da Planilha de Categorias do BI no Campo 5 ({nome_op}).")

        with st.expander(f"🔍 Mapeamento das Colunas Identificadas ({nome_op})"):
            for base, det in st.session_state.get(f'mapeamento_colunas_{prefixo}', {}).items():
                st.markdown(f"- **{base}**: {det}")

        st.download_button(
            label=f"📥 Baixar Planilha Consolidada de {nome_op} (.xlsx)",
            data=buffer_bytes_s,
            file_name=nome_saida_s,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_btn_{prefixo}"
        )

    else:
        st.info(f"👆 Selecione ou envie as planilhas nos campos acima para **{nome_op}**, configure o período comercial e clique no botão **'🚀 Gerar Planilha de {nome_op}'** para iniciar o processamento.")


# ========================================================
# RENDERIZAÇÃO DO PAINEL ATIVO
# ========================================================
renderizar_painel_operador(operador_ativo)
