import io
import os
import re
import unicodedata
import datetime
import time
import zoneinfo
import pickle
from pathlib import Path
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- CONFIGURAÇÃO DE FUSO HORÁRIO OFICIAL DO BRASIL (BRASÍLIA / UTC-3) ---
try:
    os.environ['TZ'] = 'America/Sao_Paulo'
    if hasattr(time, 'tzset'):
        time.tzset()
except Exception:
    pass

try:
    FUSO_BRASIL = zoneinfo.ZoneInfo("America/Sao_Paulo")
except Exception:
    FUSO_BRASIL = datetime.timezone(datetime.timedelta(hours=-3))

def obter_agora_brasil():
    """Retorna datetime atual sempre no fuso oficial do Brasil (America/Sao_Paulo - Brasília)."""
    try:
        return datetime.datetime.now(FUSO_BRASIL)
    except Exception:
        return datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-3)))

def obter_hoje_brasil():
    """Retorna a data atual (date) no fuso oficial do Brasil."""
    return obter_agora_brasil().date()

def obter_timestamp_formatado_brasil():
    """Retorna timestamp formatado 'DD/MM HH:MM' no horário oficial do Brasil."""
    return obter_agora_brasil().strftime("%d/%m %H:%M")

def converter_timestamp_utc_para_brasil(ts_str):
    """
    Converte um timestamp string gravado anteriormente em UTC no servidor na nuvem
    (ex: '28/10 17:50') para o horário oficial de Brasília (ex: '28/10 14:50').
    """
    if not isinstance(ts_str, str) or not ts_str.strip():
        return ts_str
    try:
        ts_limpo = ts_str.strip()
        ano_atual = obter_agora_brasil().year
        dt = datetime.datetime.strptime(f"{ano_atual}/{ts_limpo}", "%Y/%d/%m %H:%M")
        dt_br = dt - datetime.timedelta(hours=3)
        return dt_br.strftime("%d/%m %H:%M")
    except Exception:
        return ts_str

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Sistema Roleta - Controle Comercial",
    page_icon="🎯",
    layout="wide"
)

# --- DIRETÓRIOS LOCAIS E COMPARTILHADOS ---
BASE_DIR = Path(__file__).resolve().parent
PASTA_ENTRADA = BASE_DIR / "arquivos_entrada"
PASTA_SAIDA = BASE_DIR / "planilhas_geradas"
PASTA_STORAGE = BASE_DIR / "dados_compartilhados"

PASTA_ENTRADA.mkdir(parents=True, exist_ok=True)
PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
PASTA_STORAGE.mkdir(parents=True, exist_ok=True)

# --- FUNÇÕES DE PERSISTÊNCIA COMPARTILHADA (DISCO DO SERVIDOR) ---
def salvar_dados_operador(nome_op, dados_dict):
    """
    Persiste os dados de um operador em disco para sincronização
    entre navegadores e computadores diferentes (Filipe <-> Francine).
    """
    if isinstance(dados_dict, dict):
        dados_dict['tz_brasil'] = True
    prefixo = nome_op.lower()
    arquivo_pkl = PASTA_STORAGE / f"{prefixo}_dados.pkl"
    try:
        with open(arquivo_pkl, "wb") as f:
            pickle.dump(dados_dict, f)
    except Exception as e:
        st.error(f"Erro ao sincronizar dados de {nome_op}: {e}")

def carregar_dados_operador(nome_op):
    """
    Carrega os dados persistidos de um operador diretamente do disco do servidor.
    Se o timestamp foi salvo anteriormente em UTC pelo servidor na nuvem,
    migra automaticamente para o Horário de Brasília (UTC-3).
    """
    prefixo = nome_op.lower()
    arquivo_pkl = PASTA_STORAGE / f"{prefixo}_dados.pkl"
    if arquivo_pkl.exists():
        try:
            with open(arquivo_pkl, "rb") as f:
                dados = pickle.load(f)

            if isinstance(dados, dict) and 'timestamp' in dados:
                if not dados.get('tz_brasil'):
                    dados['timestamp'] = converter_timestamp_utc_para_brasil(dados['timestamp'])
                    dados['tz_brasil'] = True
                    try:
                        with open(arquivo_pkl, "wb") as f_save:
                            pickle.dump(dados, f_save)
                    except Exception:
                        pass
            return dados
        except Exception as e:
            print(f"Erro ao ler dados de {nome_op}: {e}")
    return None

def limpar_dados_operador(nome_op):
    """
    Remove os dados persistidos de um operador para iniciar uma nova rodada limpa.
    """
    prefixo = nome_op.lower()
    arquivo_pkl = PASTA_STORAGE / f"{prefixo}_dados.pkl"
    if arquivo_pkl.exists():
        try:
            arquivo_pkl.unlink()
        except Exception:
            pass

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

def obter_id_arquivo(f):
    if f is None:
        return ""
    if hasattr(f, 'name') and hasattr(f, 'size'):
        return f"{f.name}_{f.size}"
    if isinstance(f, (str, Path)):
        return str(f)
    return str(f)

def estilizar_aba(ws, df_dados, col_id_nome, col_id_val, col_cli_nome, col_data_dt, titulo_col_id="PEDIDO", titulo_status=None):
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

def carregar_planilha_ajustada(arquivo):
    xl = pd.ExcelFile(arquivo)
    res = {
        'df_aba1': pd.DataFrame(),
        'df_aba2': pd.DataFrame(),
        'df_top_notas': pd.DataFrame(),
        'df_cats': pd.DataFrame(),
        'df_lideres_cat': pd.DataFrame()
    }

    # 1. ABA 1: NOVOS E REATIVADOS
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

    # 2. ABA 2: PROPOSTAS ENVIADAS
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

    # 3. ABA 3: DESTAQUES ROLETA (TOP NOTAS E MAIORES VENDAS POR CATEGORIA)
    s3 = None
    for s in xl.sheet_names:
        if any(term in s.lower() for term in ['destaque', 'roleta', 'aba 3', 'produtos', 'categoria']):
            s3 = s
            break
    if s3:
        try:
            df3 = pd.read_excel(arquivo, sheet_name=s3, header=None)
            notas = []
            cats = []
            lendo_notas = False
            lendo_cat = False

            for idx, row in df3.iterrows():
                vals = [str(x).strip() for x in row.dropna().tolist()]
                if not vals:
                    continue
                linha_str = " ".join(vals).upper()

                if "TOP MAIORES VENDAS" in linha_str or "NOTA UNICA" in linha_str or "NOTA ÚNICA" in linha_str:
                    lendo_notas = True
                    lendo_cat = False
                    continue

                if "CATEGORIA DE PRODUTO" in linha_str or "MAIOR VENDA POR CATEGORIA" in linha_str:
                    lendo_cat = True
                    lendo_notas = False
                    continue

                if lendo_notas:
                    if "POSI" in linha_str and "VENDEDOR" in linha_str:
                        continue
                    if len(row) >= 4 and pd.notna(row[1]) and pd.notna(row[3]):
                        try:
                            val_num = limpar_valor(row[3])
                            if val_num > 0:
                                notas.append({
                                    'Vendedor': str(row[1]).strip().upper(),
                                    'Nota': row[2],
                                    'Valor_Total': val_num
                                })
                        except Exception:
                            pass

                if lendo_cat:
                    if "CATEGORIA" in linha_str and "VENDEDOR" in linha_str:
                        continue
                    if len(row) >= 3 and pd.notna(row[0]) and pd.notna(row[1]) and pd.notna(row[2]):
                        try:
                            val_num = limpar_valor(row[2])
                            if val_num > 0:
                                cats.append({
                                    'Categoria': str(row[0]).strip().upper(),
                                    'Vendedor': str(row[1]).strip().upper(),
                                    'Total_Vendido': val_num,
                                    'Valor_Total': val_num
                                })
                        except Exception:
                            pass

            res['df_top_notas'] = pd.DataFrame(notas)
            res['df_cats'] = pd.DataFrame(cats)
            res['df_lideres_cat'] = pd.DataFrame(cats)
            res['nome_aba3'] = s3
        except Exception as e:
            print(f"Erro ao ler Aba 3 da planilha ajustada: {e}")

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
    if rank1 is not None and not rank1.empty:
        top1 = rank1.iloc[0]
        podio_rows.append(("1º Ranking", "Maior Nº de Clientes Novos/Reativados", str(top1['Vendedor']), f"{int(top1['Qtd_Clientes'])} clientes (R$ {top1['Valor_Total']:,.2f})"))
    else:
        podio_rows.append(("1º Ranking", "Maior Nº de Clientes Novos/Reativados", "-", "Sem dados"))

    if rank2 is not None and not rank2.empty:
        top2 = rank2.iloc[0]
        podio_rows.append(("2º Ranking", "Maior Nº de Propostas Enviadas", str(top2['Vendedor']), f"{int(top2['Total_Propostas'])} propostas"))
    else:
        podio_rows.append(("2º Ranking", "Maior Nº de Propostas Enviadas", "-", "Sem dados"))

    if rank3 is not None and not rank3.empty:
        top3 = rank3.iloc[0]
        nota_val = int(top3['Nota']) if isinstance(top3['Nota'], (int, float)) and top3['Nota'] == int(top3['Nota']) else str(top3['Nota'])
        podio_rows.append(("3º Ranking", "Maior Pedido Faturado (Nota Única)", str(top3['Vendedor']), f"Nota {nota_val} (R$ {top3['Valor_Total']:,.2f})"))
    else:
        podio_rows.append(("3º Ranking", "Maior Pedido Faturado (Nota Única)", "-", "Sem dados"))

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
    st.caption("Utilize este campo quando tiver realizado ajustes manuais no Excel (ex: caso um cliente tenha mais de 1 pedido ou ajustes de brindes). As planilhas ajustadas são salvas no servidor e sincronizadas para ambos os operadores.")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        st.markdown("#### 👤 Planilha Ajustada — FRANCINE")
        file_fra = st.file_uploader("Selecione o arquivo Excel ajustado de Francine:", type=["xlsx", "xls"], key="modal_up_fra")
        if file_fra is not None:
            dados_fra = carregar_planilha_ajustada(file_fra)
            if 'df_aba1' in dados_fra:
                # Salva o arquivo em disco
                caminho_fra_disk = PASTA_STORAGE / "francine_ajustada.xlsx"
                with open(caminho_fra_disk, "wb") as f_out:
                    f_out.write(file_fra.getvalue())

                df_top_fra = dados_fra.get('df_top_notas', pd.DataFrame())
                df_cats_fra = dados_fra.get('df_cats', pd.DataFrame())

                # Atualiza objeto de persistência
                existente_fra = carregar_dados_operador("FRANCINE") or {}
                existente_fra.update({
                    'timestamp': obter_timestamp_formatado_brasil(),
                    'tz_brasil': True,
                    'processado': True,
                    'usando_ajustada': True,
                    'df_aba1': dados_fra['df_aba1'],
                    'df_aba2': dados_fra.get('df_aba2', pd.DataFrame()),
                    'df_top_notas': df_top_fra,
                    'df_cats': df_cats_fra,
                    'df_lideres_cat': df_cats_fra,
                    'nome_arquivo_ajustada': file_fra.name
                })
                salvar_dados_operador("FRANCINE", existente_fra)
                st.session_state['df_aba1_francine_ajustada'] = dados_fra['df_aba1']
                st.session_state['df_aba2_francine_ajustada'] = dados_fra.get('df_aba2', pd.DataFrame())
                st.session_state['df_top_francine_ajustada'] = df_top_fra
                st.session_state['df_cats_francine_ajustada'] = df_cats_fra
                st.success(f"✅ Francine: {len(dados_fra['df_aba1'])} clientes carregados e sincronizados!")

        fra_disco = carregar_dados_operador("FRANCINE")
        if fra_disco and fra_disco.get('usando_ajustada'):
            st.info(f"Arquivo ativo: `{fra_disco.get('nome_arquivo_ajustada', 'Planilha Ajustada')}` ({fra_disco.get('timestamp')})")
            if st.button("🗑️ Remover Planilha Ajustada de Francine", key="del_aj_fra"):
                limpar_dados_operador("FRANCINE")
                st.session_state.pop('df_aba1_francine_ajustada', None)
                st.session_state.pop('df_aba2_francine_ajustada', None)
                st.session_state.pop('df_top_francine_ajustada', None)
                st.session_state.pop('df_cats_francine_ajustada', None)
                st.rerun()

    with col_u2:
        st.markdown("#### 👤 Planilha Ajustada — FILIPE")
        file_fil = st.file_uploader("Selecione o arquivo Excel ajustado de Filipe:", type=["xlsx", "xls"], key="modal_up_fil")
        if file_fil is not None:
            dados_fil = carregar_planilha_ajustada(file_fil)
            if 'df_aba1' in dados_fil:
                caminho_fil_disk = PASTA_STORAGE / "filipe_ajustada.xlsx"
                with open(caminho_fil_disk, "wb") as f_out:
                    f_out.write(file_fil.getvalue())

                df_top_fil = dados_fil.get('df_top_notas', pd.DataFrame())
                df_cats_fil = dados_fil.get('df_cats', pd.DataFrame())

                existente_fil = carregar_dados_operador("FILIPE") or {}
                existente_fil.update({
                    'timestamp': obter_timestamp_formatado_brasil(),
                    'tz_brasil': True,
                    'processado': True,
                    'usando_ajustada': True,
                    'df_aba1': dados_fil['df_aba1'],
                    'df_aba2': dados_fil.get('df_aba2', pd.DataFrame()),
                    'df_top_notas': df_top_fil,
                    'df_cats': df_cats_fil,
                    'df_lideres_cat': df_cats_fil,
                    'nome_arquivo_ajustada': file_fil.name
                })
                salvar_dados_operador("FILIPE", existente_fil)
                st.session_state['df_aba1_filipe_ajustada'] = dados_fil['df_aba1']
                st.session_state['df_aba2_filipe_ajustada'] = dados_fil.get('df_aba2', pd.DataFrame())
                st.session_state['df_top_filipe_ajustada'] = df_top_fil
                st.session_state['df_cats_filipe_ajustada'] = df_cats_fil
                st.success(f"✅ Filipe: {len(dados_fil['df_aba1'])} clientes carregados e sincronizados!")

        fil_disco = carregar_dados_operador("FILIPE")
        if fil_disco and fil_disco.get('usando_ajustada'):
            st.info(f"Arquivo ativo: `{fil_disco.get('nome_arquivo_ajustada', 'Planilha Ajustada')}` ({fil_disco.get('timestamp')})")
            if st.button("🗑️ Remover Planilha Ajustada de Filipe", key="del_aj_fil"):
                limpar_dados_operador("FILIPE")
                st.session_state.pop('df_aba1_filipe_ajustada', None)
                st.session_state.pop('df_aba2_filipe_ajustada', None)
                st.session_state.pop('df_top_filipe_ajustada', None)
                st.session_state.pop('df_cats_filipe_ajustada', None)
                st.rerun()

    st.markdown("---")
    if st.button("Concluir e Salvar", type="primary", use_container_width=True):
        st.rerun()


@st.dialog("🏆 RESULTADOS FINAIS DA ROLETA COMERCIAL", width="large")
def modal_resultados_finais():
    st.markdown("## 🏆 Resultados Finais Consolidados (Roleta Comercial)")
    st.caption("Cruzamento e consolidação sincronizada das planilhas de **FRANCINE** e **FILIPE**.")

    # 1. Recupera dados de FRANCINE diretamente do disco sincronizado
    disco_fra = carregar_dados_operador("FRANCINE") or {}
    df1_fra = disco_fra.get('df_aba1')
    df2_fra = disco_fra.get('df_aba2')
    df_top_fra = disco_fra.get('df_top_notas', pd.DataFrame())
    df_cat_fra = disco_fra.get('df_categorias_raw')
    df_rank_cat_fra = disco_fra.get('df_rank_cat', pd.DataFrame())
    df_cats_fra = disco_fra.get('df_lideres_cat')
    if df_cats_fra is None or (isinstance(df_cats_fra, pd.DataFrame) and df_cats_fra.empty):
        df_cats_fra = disco_fra.get('df_cats', pd.DataFrame())
    fra_usando_ajustada = bool(disco_fra.get('usando_ajustada'))

    # Fallback para session_state local se não estiver no disco
    if df1_fra is None:
        df1_fra = st.session_state.get('df_aba1_francine_ajustada') or st.session_state.get('df_aba1_francine')
    if df2_fra is None:
        df2_fra = st.session_state.get('df_aba2_francine_ajustada') or st.session_state.get('df_aba2_francine')
    if (df_top_fra is None or df_top_fra.empty) and 'df_top_francine_ajustada' in st.session_state:
        df_top_fra = st.session_state.get('df_top_francine_ajustada', pd.DataFrame())
    if (df_cats_fra is None or df_cats_fra.empty) and 'df_cats_francine_ajustada' in st.session_state:
        df_cats_fra = st.session_state.get('df_cats_francine_ajustada', pd.DataFrame())

    # Auto-recuperação da Aba 3 de Francine se notas ou categorias estiverem vazias
    caminho_fra_aj = PASTA_STORAGE / "francine_ajustada.xlsx"
    if (df_top_fra is None or df_top_fra.empty or df_cats_fra is None or df_cats_fra.empty) and caminho_fra_aj.exists():
        try:
            d_extra_fra = carregar_planilha_ajustada(caminho_fra_aj)
            if (df_top_fra is None or df_top_fra.empty) and not d_extra_fra.get('df_top_notas', pd.DataFrame()).empty:
                df_top_fra = d_extra_fra['df_top_notas']
            if (df_cats_fra is None or df_cats_fra.empty) and not d_extra_fra.get('df_cats', pd.DataFrame()).empty:
                df_cats_fra = d_extra_fra['df_cats']
        except Exception:
            pass
    elif (df_top_fra is None or df_top_fra.empty or df_cats_fra is None or df_cats_fra.empty) and disco_fra.get('caminho_salvar_local'):
        p_fra_local = Path(disco_fra.get('caminho_salvar_local'))
        if p_fra_local.exists():
            try:
                d_extra_fra = carregar_planilha_ajustada(p_fra_local)
                if (df_top_fra is None or df_top_fra.empty) and not d_extra_fra.get('df_top_notas', pd.DataFrame()).empty:
                    df_top_fra = d_extra_fra['df_top_notas']
                if (df_cats_fra is None or df_cats_fra.empty) and not d_extra_fra.get('df_cats', pd.DataFrame()).empty:
                    df_cats_fra = d_extra_fra['df_cats']
            except Exception:
                pass

    # 2. Recupera dados de FILIPE diretamente do disco sincronizado
    disco_fil = carregar_dados_operador("FILIPE") or {}
    df1_fil = disco_fil.get('df_aba1')
    df2_fil = disco_fil.get('df_aba2')
    df_top_fil = disco_fil.get('df_top_notas', pd.DataFrame())
    df_cat_fil = disco_fil.get('df_categorias_raw')
    df_rank_cat_fil = disco_fil.get('df_rank_cat', pd.DataFrame())
    df_cats_fil = disco_fil.get('df_lideres_cat')
    if df_cats_fil is None or (isinstance(df_cats_fil, pd.DataFrame) and df_cats_fil.empty):
        df_cats_fil = disco_fil.get('df_cats', pd.DataFrame())
    fil_usando_ajustada = bool(disco_fil.get('usando_ajustada'))

    if df1_fil is None:
        df1_fil = st.session_state.get('df_aba1_filipe_ajustada') or st.session_state.get('df_aba1_filipe')
    if df2_fil is None:
        df2_fil = st.session_state.get('df_aba2_filipe_ajustada') or st.session_state.get('df_aba2_filipe')
    if (df_top_fil is None or df_top_fil.empty) and 'df_top_filipe_ajustada' in st.session_state:
        df_top_fil = st.session_state.get('df_top_filipe_ajustada', pd.DataFrame())
    if (df_cats_fil is None or df_cats_fil.empty) and 'df_cats_filipe_ajustada' in st.session_state:
        df_cats_fil = st.session_state.get('df_cats_filipe_ajustada', pd.DataFrame())

    # Auto-recuperação da Aba 3 de Filipe se notas ou categorias estiverem vazias
    caminho_fil_aj = PASTA_STORAGE / "filipe_ajustada.xlsx"
    if (df_top_fil is None or df_top_fil.empty or df_cats_fil is None or df_cats_fil.empty) and caminho_fil_aj.exists():
        try:
            d_extra_fil = carregar_planilha_ajustada(caminho_fil_aj)
            if (df_top_fil is None or df_top_fil.empty) and not d_extra_fil.get('df_top_notas', pd.DataFrame()).empty:
                df_top_fil = d_extra_fil['df_top_notas']
            if (df_cats_fil is None or df_cats_fil.empty) and not d_extra_fil.get('df_cats', pd.DataFrame()).empty:
                df_cats_fil = d_extra_fil['df_cats']
        except Exception:
            pass
    elif (df_top_fil is None or df_top_fil.empty or df_cats_fil is None or df_cats_fil.empty) and disco_fil.get('caminho_salvar_local'):
        p_fil_local = Path(disco_fil.get('caminho_salvar_local'))
        if p_fil_local.exists():
            try:
                d_extra_fil = carregar_planilha_ajustada(p_fil_local)
                if (df_top_fil is None or df_top_fil.empty) and not d_extra_fil.get('df_top_notas', pd.DataFrame()).empty:
                    df_top_fil = d_extra_fil['df_top_notas']
                if (df_cats_fil is None or df_cats_fil.empty) and not d_extra_fil.get('df_cats', pd.DataFrame()).empty:
                    df_cats_fil = d_extra_fil['df_cats']
            except Exception:
                pass

    tem_dados_fra = (df1_fra is not None and not df1_fra.empty) or (df2_fra is not None and not df2_fra.empty) or (df_top_fra is not None and not df_top_fra.empty)
    tem_dados_fil = (df1_fil is not None and not df1_fil.empty) or (df2_fil is not None and not df2_fil.empty) or (df_top_fil is not None and not df_top_fil.empty)

    if not (tem_dados_fra or tem_dados_fil):
        st.warning("⚠️ Atenção: Nenhuma planilha foi gerada ou enviada ainda. Por favor, processe as planilhas nos painéis abaixo ou faça o upload das planilhas ajustadas no botão **'UPLOAD DE PLANILHA AJUSTADA'**.")
        return

    ts_fra_txt = disco_fra.get('timestamp', '')
    ts_fil_txt = disco_fil.get('timestamp', '')
    info_fra = f"📝 Ajustada ({ts_fra_txt})" if fra_usando_ajustada else (f"✅ Gerada ({ts_fra_txt})" if tem_dados_fra else "⚪ Sem dados")
    info_fil = f"📝 Ajustada ({ts_fil_txt})" if fil_usando_ajustada else (f"✅ Gerada ({ts_fil_txt})" if tem_dados_fil else "⚪ Sem dados")
    st.info(f"📊 **Bases sincronizadas:** Francine (**{info_fra}**) | Filipe (**{info_fil}**)")

    # 1. MAIOR NÚMERO DE CLIENTES NOVOS / REATIVADOS
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

    # 2. MAIOR NÚMERO DE PROPOSTAS ENVIADAS
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

    # 3. MAIOR PEDIDO FATURADO (CRUZAMENTO DE NOTAS DE AMBOS OS OPERADORES)
    list_top = []
    if df_top_fra is not None and not df_top_fra.empty:
        t_fra = df_top_fra.copy()
        col_v = identificar_coluna(t_fra.columns, ['vendedor', 'operador'])
        col_n = identificar_coluna(t_fra.columns, ['nota', 'pedido', 'documento'])
        col_val = identificar_coluna(t_fra.columns, ['valor_total', 'valor', 'total'])
        if col_v and col_val:
            t_fra_c = pd.DataFrame()
            t_fra_c['Vendedor'] = t_fra[col_v].astype(str).str.strip().str.upper()
            t_fra_c['Nota'] = t_fra[col_n] if col_n else '-'
            t_fra_c['Valor_Total'] = t_fra[col_val].apply(limpar_valor)
            t_fra_c['Origem'] = 'FRANCINE'
            list_top.append(t_fra_c)

    if df_top_fil is not None and not df_top_fil.empty:
        t_fil = df_top_fil.copy()
        col_v = identificar_coluna(t_fil.columns, ['vendedor', 'operador'])
        col_n = identificar_coluna(t_fil.columns, ['nota', 'pedido', 'documento'])
        col_val = identificar_coluna(t_fil.columns, ['valor_total', 'valor', 'total'])
        if col_v and col_val:
            t_fil_c = pd.DataFrame()
            t_fil_c['Vendedor'] = t_fil[col_v].astype(str).str.strip().str.upper()
            t_fil_c['Nota'] = t_fil[col_n] if col_n else '-'
            t_fil_c['Valor_Total'] = t_fil[col_val].apply(limpar_valor)
            t_fil_c['Origem'] = 'FILIPE'
            list_top.append(t_fil_c)

    rank3 = pd.DataFrame()
    if list_top:
        df_notas_comb = pd.concat(list_top, ignore_index=True)
        df_notas_comb = df_notas_comb[df_notas_comb['Valor_Total'] > 0]
        rank3 = (
            df_notas_comb.sort_values(by='Valor_Total', ascending=False)
            .drop_duplicates(subset=['Nota', 'Vendedor'] if 'Nota' in df_notas_comb.columns else ['Vendedor'])
            .reset_index(drop=True)
        )
    elif not df1_comb.empty and 'VALOR_NUM' in df1_comb.columns:
        rank3 = (
            df1_comb.sort_values(by='VALOR_NUM', ascending=False)
            .rename(columns={'Pedido': 'Nota', 'VALOR_NUM': 'Valor_Total'})
            .reset_index(drop=True)
        )

    # 4. MAIOR VENDA DA CATEGORIA DETERMINADA PREVIAMENTE
    cat_sugerida = disco_fra.get('categoria_solicitada') or disco_fil.get('categoria_solicitada') or st.session_state.get('categoria_solicitada_francine') or "HIGIENE E LIMPEZA"
    idx_cat_final = CATEGORIAS_ROLETA.index(cat_sugerida) if cat_sugerida in CATEGORIAS_ROLETA else 0

    col_sel_cat, _ = st.columns([1, 1])
    with col_sel_cat:
        cat_escolhida = st.selectbox(
            "🏷️ Categoria Avaliada para o 4º Ranking:",
            CATEGORIAS_ROLETA,
            index=idx_cat_final,
            key="modal_sel_cat_final"
        )

    cat_escolhida_norm = normalizar_texto(cat_escolhida).replace("'", "").replace(" ", "")

    partes_cat = []

    # Processa Francine para a categoria escolhida
    df_r_fra = pd.DataFrame()
    if df_cat_fra is not None and not df_cat_fra.empty:
        df_r_fra, _, _ = extrair_ranking_categoria(df_cat_fra, cat_escolhida, disco_fra.get('df_fortpel'), disco_fra.get('df_adm_raw'))

    if df_r_fra is not None and not df_r_fra.empty:
        col_v = identificar_coluna(df_r_fra.columns, ['vendedor', 'operador'])
        col_val = identificar_coluna(df_r_fra.columns, ['total_vendido', 'valor_total', 'valor'])
        if col_v and col_val:
            df_r_fra_c = pd.DataFrame()
            df_r_fra_c['Vendedor'] = df_r_fra[col_v].astype(str).str.strip().str.upper()
            df_r_fra_c['Total_Vendido'] = df_r_fra[col_val].apply(limpar_valor)
            df_r_fra_c['Origem'] = 'FRANCINE'
            partes_cat.append(df_r_fra_c)
    elif df_cats_fra is not None and not df_cats_fra.empty:
        col_cat = identificar_coluna(df_cats_fra.columns, ['categoria', 'cat'])
        col_v = identificar_coluna(df_cats_fra.columns, ['vendedor', 'operador'])
        col_val = identificar_coluna(df_cats_fra.columns, ['total_vendido', 'valor_total', 'valor', 'total'])
        if col_cat and col_v and col_val:
            m = df_cats_fra[col_cat].apply(lambda x: cat_escolhida_norm in normalizar_texto(x).replace("'", "").replace(" ", "") or normalizar_texto(x).replace("'", "").replace(" ", "") in cat_escolhida_norm)
            sub_c = df_cats_fra[m].copy()
            if not sub_c.empty:
                df_sub_c = pd.DataFrame()
                df_sub_c['Vendedor'] = sub_c[col_v].astype(str).str.strip().str.upper()
                df_sub_c['Total_Vendido'] = sub_c[col_val].apply(limpar_valor)
                df_sub_c['Origem'] = 'FRANCINE'
                partes_cat.append(df_sub_c)
    elif not df_rank_cat_fra.empty:
        cat_salva_fra = normalizar_texto(disco_fra.get('categoria_solicitada', '')).replace("'", "").replace(" ", "")
        if cat_salva_fra == cat_escolhida_norm:
            col_v = identificar_coluna(df_rank_cat_fra.columns, ['vendedor', 'operador'])
            col_val = identificar_coluna(df_rank_cat_fra.columns, ['total_vendido', 'valor_total', 'valor'])
            if col_v and col_val:
                df_rf_c = pd.DataFrame()
                df_rf_c['Vendedor'] = df_rank_cat_fra[col_v].astype(str).str.strip().str.upper()
                df_rf_c['Total_Vendido'] = df_rank_cat_fra[col_val].apply(limpar_valor)
                df_rf_c['Origem'] = 'FRANCINE'
                partes_cat.append(df_rf_c)

    # Processa Filipe para a categoria escolhida
    df_r_fil = pd.DataFrame()
    if df_cat_fil is not None and not df_cat_fil.empty:
        df_r_fil, _, _ = extrair_ranking_categoria(df_cat_fil, cat_escolhida, disco_fil.get('df_fortpel'), disco_fil.get('df_adm_raw'))

    if df_r_fil is not None and not df_r_fil.empty:
        col_v = identificar_coluna(df_r_fil.columns, ['vendedor', 'operador'])
        col_val = identificar_coluna(df_r_fil.columns, ['total_vendido', 'valor_total', 'valor'])
        if col_v and col_val:
            df_r_fil_c = pd.DataFrame()
            df_r_fil_c['Vendedor'] = df_r_fil[col_v].astype(str).str.strip().str.upper()
            df_r_fil_c['Total_Vendido'] = df_r_fil[col_val].apply(limpar_valor)
            df_r_fil_c['Origem'] = 'FILIPE'
            partes_cat.append(df_r_fil_c)
    elif df_cats_fil is not None and not df_cats_fil.empty:
        col_cat = identificar_coluna(df_cats_fil.columns, ['categoria', 'cat'])
        col_v = identificar_coluna(df_cats_fil.columns, ['vendedor', 'operador'])
        col_val = identificar_coluna(df_cats_fil.columns, ['total_vendido', 'valor_total', 'valor', 'total'])
        if col_cat and col_v and col_val:
            m = df_cats_fil[col_cat].apply(lambda x: cat_escolhida_norm in normalizar_texto(x).replace("'", "").replace(" ", "") or normalizar_texto(x).replace("'", "").replace(" ", "") in cat_escolhida_norm)
            sub_c = df_cats_fil[m].copy()
            if not sub_c.empty:
                df_sub_c = pd.DataFrame()
                df_sub_c['Vendedor'] = sub_c[col_v].astype(str).str.strip().str.upper()
                df_sub_c['Total_Vendido'] = sub_c[col_val].apply(limpar_valor)
                df_sub_c['Origem'] = 'FILIPE'
                partes_cat.append(df_sub_c)
    elif not df_rank_cat_fil.empty:
        cat_salva_fil = normalizar_texto(disco_fil.get('categoria_solicitada', '')).replace("'", "").replace(" ", "")
        if cat_salva_fil == cat_escolhida_norm:
            col_v = identificar_coluna(df_rank_cat_fil.columns, ['vendedor', 'operador'])
            col_val = identificar_coluna(df_rank_cat_fil.columns, ['total_vendido', 'valor_total', 'valor'])
            if col_v and col_val:
                df_rfl_c = pd.DataFrame()
                df_rfl_c['Vendedor'] = df_rank_cat_fil[col_v].astype(str).str.strip().str.upper()
                df_rfl_c['Total_Vendido'] = df_rank_cat_fil[col_val].apply(limpar_valor)
                df_rfl_c['Origem'] = 'FILIPE'
                partes_cat.append(df_rfl_c)

    rank4 = pd.DataFrame()
    if partes_cat:
        df_cat_comb = pd.concat(partes_cat, ignore_index=True)
        df_cat_comb = df_cat_comb[df_cat_comb['Total_Vendido'] > 0]
        rank4 = (
            df_cat_comb.groupby('Vendedor', as_index=False)['Total_Vendido']
            .sum()
            .sort_values(by='Total_Vendido', ascending=False)
            .reset_index(drop=True)
        )

    # PAINEL COM OS 4 CAMPEÕES EM DESTAQUE (CARDS NO TOPO)
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
            origem_top3 = f" ({top3['Origem']})" if 'Origem' in top3 and pd.notna(top3['Origem']) else ""
            st.metric(
                label=f"🥇 {top3['Vendedor']}",
                value=val_fmt,
                delta=f"Nota/Pedido: {nota_val}{origem_top3}"
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

    # TABELAS DETALHADAS DE CADA RANKING
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

    bytes_excel_final = gerar_excel_resultados_finais(rank1, rank2, rank3, rank4, cat_escolhida)
    nome_arq_final = f"Resultados_Finais_Roleta_Consolidado_{obter_hoje_brasil().strftime('%d%m%Y')}.xlsx"

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
# CABEÇALHO, SELETOR E SINCRONIZAÇÃO
# ========================================================
st.title("🎯 Sistema Roleta - Controle Comercial")

# Obtém status lendo diretamente do disco do servidor (compartilhado entre máquinas)
dados_fra_disk = carregar_dados_operador("FRANCINE")
dados_fil_disk = carregar_dados_operador("FILIPE")

if dados_fra_disk:
    ts_f = dados_fra_disk.get('timestamp', '')
    status_fra_txt = f"📝 Ajustada ({ts_f})" if dados_fra_disk.get('usando_ajustada') else f"✅ Gerado ({ts_f})"
else:
    status_fra_txt = "⚪ Pendente"

if dados_fil_disk:
    ts_l = dados_fil_disk.get('timestamp', '')
    status_fil_txt = f"📝 Ajustada ({ts_l})" if dados_fil_disk.get('usando_ajustada') else f"✅ Gerado ({ts_l})"
else:
    status_fil_txt = "⚪ Pendente"

col_top_sel, col_top_status, col_top_sync = st.columns([2, 2, 1])
with col_top_sel:
    operador_ativo = st.radio(
        "👤 Selecione o Operador / Perfil:",
        ["FRANCINE", "FILIPE"],
        horizontal=True,
        key="sel_operador_global",
        help="Cada operador possui sua própria área de trabalho e seus dados ficam salvos e sincronizados no servidor."
    )
with col_top_status:
    st.markdown("**Status dos Relatórios (Sincronizado):**")
    st.caption(f"FRANCINE: **{status_fra_txt}** | FILIPE: **{status_fil_txt}**")
with col_top_sync:
    st.write("")
    if st.button("🔄 Sincronizar", help="Atualiza e sincroniza as planilhas salvas por ambos os computadores", use_container_width=True):
        st.rerun()

# ========================================================
# BOTÕES DE AÇÕES GLOBAIS
# ========================================================
col_btn_acoes1, col_btn_acoes2 = st.columns(2)
with col_btn_acoes1:
    if st.button("📤 UPLOAD DE PLANILHA AJUSTADA", use_container_width=True, help="Clique para fazer o upload da planilha ajustada no Excel"):
        modal_upload_ajustada()

with col_btn_acoes2:
    if st.button("🏆 GERAR RESULTADOS FINAIS", type="primary", use_container_width=True, help="Cruza as planilhas de Francine e Filipe e abre a tela com os 4 rankings consolidados"):
        modal_resultados_finais()

tem_ajustada_fra = dados_fra_disk and dados_fra_disk.get('usando_ajustada')
tem_ajustada_fil = dados_fil_disk and dados_fil_disk.get('usando_ajustada')
if tem_ajustada_fra or tem_ajustada_fil:
    nomes_aj = []
    if tem_ajustada_fra: nomes_aj.append("FRANCINE")
    if tem_ajustada_fil: nomes_aj.append("FILIPE")
    st.success(f"📝 **Planilha Ajustada sincronizada para:** {', '.join(nomes_aj)}. A apuração dos Resultados Finais utilizará os dados ajustados automaticamente!")

st.markdown("---")


# ========================================================
# PAINEL INDIVIDUAL DO OPERADOR (COM PERSISTÊNCIA COMPARTILHADA)
# ========================================================
def renderizar_painel_operador(nome_op):
    prefixo = nome_op.lower()

    col_tit1, col_tit2 = st.columns([3, 1])
    with col_tit1:
        st.subheader(f"📋 Painel de Controle Comercial — {nome_op}")
        st.caption(f"Arquivos e parâmetros específicos para a planilha de **{nome_op}**.")
    with col_tit2:
        # Botão para limpar dados persistidos deste operador se necessário
        if carregar_dados_operador(nome_op):
            if st.button(f"🗑️ Limpar dados de {nome_op}", key=f"btn_limpar_{prefixo}", help=f"Apaga a planilha salva de {nome_op} no servidor para iniciar do zero"):
                limpar_dados_operador(nome_op)
                st.session_state[f'processado_{prefixo}'] = False
                st.rerun()

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

    hoje = obter_hoje_brasil()
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
            with st.spinner(f"⏳ Processando bases de {nome_op} e salvando de forma sincronizada no servidor..."):
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

                    # ========================================================
                    # PERSISTÊNCIA COMPARTILHADA EM DISCO
                    # ========================================================
                    dados_operador_disco = {
                        'timestamp': obter_timestamp_formatado_brasil(),
                        'tz_brasil': True,
                        'processado': True,
                        'usando_ajustada': False,
                        'df_aba1': df_aba1,
                        'df_aba2': df_aba2,
                        'df_top_notas': df_top_notas,
                        'df_rank_cat': df_rank_cat,
                        'total_geral_cat': total_geral_cat,
                        'df_lideres_cat': df_lideres_cat,
                        'df_categorias_raw': df_categorias_raw,
                        'categoria_solicitada': categoria_solicitada,
                        'df_fortpel': df_fortpel,
                        'df_adm_raw': df_adm_raw,
                        'col_f_ped': col_f_ped,
                        'col_f_cli': col_f_cli,
                        'col_f_vend': col_f_vend,
                        'nome_aba1': nome_aba1,
                        'nome_aba2': nome_aba2,
                        'caminho_salvar_local': str(caminho_salvar_local),
                        'buffer_bytes': buffer_download.getvalue(),
                        'nome_arquivo_saida': nome_arquivo_saida,
                        'mapeamento_colunas': {
                            "BI Novos": f"Cliente: `{col_novos_cli}` | Valor: `{col_novos_val}`",
                            "BI Reativados": f"Cliente: `{col_reat_cli}` | Valor: `{col_reat_val}`",
                            "Acompanhamento Fortpel": f"Cliente: `{col_f_cli}` | Pedido: `{col_f_ped}` | Vendedor: `{col_f_vend}` | Valor: `{col_f_val}` | Data: `{col_f_data}`",
                            "Planilha ADM": f"Vendedor: `{col_a_vend}` | Status: `{col_a_status}`",
                            "NFs Faturadas (Campo 6)": "Ativo" if df_nf_raw is not None else "Não enviado (usando Fortpel)",
                            "Categorias (Campo 5)": "Ativo" if df_categorias_raw is not None else "Não enviado"
                        }
                    }
                    salvar_dados_operador(nome_op, dados_operador_disco)

                    st.session_state[f'processado_{prefixo}'] = True
                    st.session_state[f'fingerprint_{prefixo}'] = fingerprint_atual
                    st.session_state[f'dados_disco_{prefixo}'] = dados_operador_disco
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Ocorreu um erro ao processar as planilhas de {nome_op}: {e}")
                    st.session_state[f'processado_{prefixo}'] = False

    # --- SEÇÃO 5: EXIBIÇÃO DOS RESULTADOS (SINCRONIZADO DO DISCO) ---
    dados_salvos_op = carregar_dados_operador(nome_op)
    if dados_salvos_op and dados_salvos_op.get('processado', False):
        df_aba1_s = dados_salvos_op['df_aba1']
        df_aba2_s = dados_salvos_op['df_aba2']
        nome_aba1_s = dados_salvos_op.get('nome_aba1', 'Reativados - Novos')
        nome_aba2_s = dados_salvos_op.get('nome_aba2', 'Propostas Enviadas')
        caminho_local_s = dados_salvos_op.get('caminho_salvar_local', '')
        buffer_bytes_s = dados_salvos_op.get('buffer_bytes')
        nome_saida_s = dados_salvos_op.get('nome_arquivo_saida', f"Controle_{nome_op}.xlsx")
        col_ped_s = dados_salvos_op.get('col_f_ped', 'Pedido')
        col_cli_s = dados_salvos_op.get('col_f_cli', 'Cliente')
        ts_op = dados_salvos_op.get('timestamp', '')

        st.success(f"✅ Planilha de {nome_op} sincronizada no servidor! (Última geração: {ts_op})")
        if caminho_local_s:
            st.info(f"📁 Arquivo Excel salvo em: `{caminho_local_s}`")

        abas_titulos = [
            f"Aba 1: {nome_aba1_s}",
            f"Aba 2: {nome_aba2_s}",
            "Aba 3: 🏆 Destaques Roleta (Nota Única & Categorias)"
        ]
        abas_res = st.tabs(abas_titulos)

        with abas_res[0]:
            if df_aba1_s is not None and not df_aba1_s.empty:
                qtd_duas_compras = int(df_aba1_s['TEM_DUAS_COMPRAS'].sum()) if 'TEM_DUAS_COMPRAS' in df_aba1_s else 0
                if qtd_duas_compras > 0:
                    st.warning(f"⚠️ {qtd_duas_compras} cliente(s) com mais de uma compra no mesmo mês identificado(s) (destacados com 'DUAS COMPRAS' e linha em vermelho na planilha)!")

                df_aba1_disp = pd.DataFrame()

                # Vendedor
                col_v = identificar_coluna(df_aba1_s.columns, ['vendedor_clean', 'vendedor', 'operador'])
                if col_v:
                    df_aba1_disp['Vendedor'] = df_aba1_s[col_v].astype(str).str.strip().str.upper()

                # Pedido
                col_p = col_ped_s if (col_ped_s and col_ped_s in df_aba1_s.columns) else identificar_coluna(df_aba1_s.columns, ['pedido', 'nota', 'documento'])
                if col_p and col_p in df_aba1_s.columns:
                    df_aba1_disp['Pedido'] = df_aba1_s[col_p].astype(str)

                # Cliente
                col_c = col_cli_s if (col_cli_s and col_cli_s in df_aba1_s.columns) else identificar_coluna(df_aba1_s.columns, ['cliente', 'razao', 'nome'])
                if col_c and col_c in df_aba1_s.columns:
                    df_aba1_disp['Cliente'] = df_aba1_s[col_c].astype(str).str.strip().str.upper()

                # Valor
                col_vl = identificar_coluna(df_aba1_s.columns, ['valor_num', 'valor', 'total'])
                if col_vl:
                    df_aba1_disp['Valor'] = df_aba1_s[col_vl].apply(limpar_valor).apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

                # Status
                col_st = identificar_coluna(df_aba1_s.columns, ['status', 'status_clean'])
                if col_st:
                    df_aba1_disp['Status'] = df_aba1_s[col_st].astype(str).str.strip().str.upper()

                # Data
                if '__DATA_DT' in df_aba1_s.columns and pd.api.types.is_datetime64_any_dtype(df_aba1_s['__DATA_DT']):
                    df_aba1_disp['Data'] = df_aba1_s['__DATA_DT'].dt.strftime('%d/%m/%Y').fillna('')
                else:
                    col_dt = identificar_coluna(df_aba1_s.columns, ['data', 'dt'])
                    if col_dt and col_dt in df_aba1_s.columns:
                        df_aba1_disp['Data'] = df_aba1_s[col_dt].astype(str)

                # Brinde Roleta
                if 'BRINDE ROLETA' in df_aba1_s.columns:
                    df_aba1_disp['Brinde Roleta'] = df_aba1_s['BRINDE ROLETA'].fillna('')
                elif 'TEM_DUAS_COMPRAS' in df_aba1_s.columns:
                    df_aba1_disp['Brinde Roleta'] = df_aba1_s['TEM_DUAS_COMPRAS'].apply(lambda x: "DUAS COMPRAS" if x else "")

                if df_aba1_disp.empty:
                    cols_unicas = list(dict.fromkeys(df_aba1_s.columns))
                    df_aba1_disp = df_aba1_s[cols_unicas].copy()

                # Garante que não haja colunas duplicadas
                df_aba1_disp = df_aba1_disp.loc[:, ~df_aba1_disp.columns.duplicated()]
                st.dataframe(df_aba1_disp, use_container_width=True)
            else:
                st.write(f"Sem registros no período para a Aba 1 ({nome_op}).")

        with abas_res[1]:
            if df_aba2_s is not None and not df_aba2_s.empty:
                df_aba2_disp = df_aba2_s.loc[:, ~df_aba2_s.columns.duplicated()].copy()
                st.dataframe(df_aba2_disp, use_container_width=True)
                st.metric("Total de Propostas Efetivadas", int(df_aba2_disp['Qtd'].sum()) if 'Qtd' in df_aba2_disp else 0)
            else:
                st.write(f"Sem registros na planilha ADM ({nome_op}).")

        with abas_res[2]:
            df_top_notas_s = dados_salvos_op.get('df_top_notas', pd.DataFrame())
            df_cat_raw_s = dados_salvos_op.get('df_categorias_raw')
            df_fort_s = dados_salvos_op.get('df_fortpel')
            df_adm_s = dados_salvos_op.get('df_adm_raw')
            cat_padrao = dados_salvos_op.get('categoria_solicitada', 'HIGIENE E LIMPEZA')

            st.markdown(f"### 🏆 Maior Venda em Nota Única — {nome_op}")
            st.caption("Calculado a partir da análise das NFs faturadas no período (Campo 6 ou base Fortpel).")

            if df_top_notas_s is not None and not df_top_notas_s.empty:
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
                    df_notas_disp = df_top_notas_s.loc[:, ~df_top_notas_s.columns.duplicated()].copy()
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
                    df_cat_disp = df_rank_escolhido.loc[:, ~df_rank_escolhido.columns.duplicated()].copy()
                    df_cat_disp['Total_Vendido'] = df_cat_disp['Total_Vendido'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                    st.dataframe(df_cat_disp, use_container_width=True)

                if not df_lideres_todas.empty:
                    with st.expander("📊 Ver Tabela Geral de Líderes das 7 Categorias"):
                        df_lid_disp = df_lideres_todas.loc[:, ~df_lideres_todas.columns.duplicated()].copy()
                        df_lid_disp['Valor_Total'] = df_lid_disp['Valor_Total'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                        df_lid_disp['Total_Categoria'] = df_lid_disp['Total_Categoria'].apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                        st.dataframe(df_lid_disp, use_container_width=True)
            else:
                st.info(f"ℹ️ Para consultar os vendedores líderes por categoria, faça o upload da Planilha de Categorias do BI no Campo 5 ({nome_op}).")

        with st.expander(f"🔍 Mapeamento das Colunas Identificadas ({nome_op})"):
            for base, det in dados_salvos_op.get('mapeamento_colunas', {}).items():
                st.markdown(f"- **{base}**: {det}")

        if buffer_bytes_s:
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
