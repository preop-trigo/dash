from pathlib import Path
from io import BytesIO
from urllib.parse import urlparse, parse_qsl, urlencode, urlunparse
import unicodedata
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from streamlit_autorefresh import st_autorefresh

# =========================================================
# CONFIGURACAO
# =========================================================
st.set_page_config(page_title="Dashboard Pre-Op 2026", page_icon="📊", layout="wide")

# Atualiza a pagina automaticamente a cada 1 hora.
INTERVALO_ATUALIZACAO_SEG = 60 * 60
st_autorefresh(interval=INTERVALO_ATUALIZACAO_SEG * 1000, key="atualizacao_automatica")

SHAREPOINT_URL = "https://trigogrupo-my.sharepoint.com/:x:/g/personal/isabel_mendes_grupotrigo_com_br/IQBQBfuaDt2BRZ3DnDhUAB64AXlXAjiIByO_jURmo2U7syo?e=k04bNY"
ARQUIVO_PADRAO = "PROJETOS PREOP 2026 - DASH.xlsx"
ABA_METAS = "Inaugurações 2026"
ABA_GERAL = "Geral"
LOGO_ARQUIVO = "trigo_logo.webp"
META_PERCENTUAL = 0.80
ANO = 2026

TIPOS_PROJETO = [
    "INAUGURACAO",
    "REFORMA",
    "REPASSE",
    "MUDANCA DE PONTO",
    "VIRADA DIGITAL",
    "MIGRACAO",
]
TIPO_LABELS = {
    "INAUGURACAO": "Inauguração",
    "REFORMA": "Reforma",
    "REPASSE": "Repasse",
    "MUDANCA DE PONTO": "Mudança de Ponto",
    "VIRADA DIGITAL": "Virada Digital",
    "MIGRACAO": "Migração",
}

ORANGE = "#F58220"
DARK = "#222222"
GOLD = "#D5A61C"
TEAL = "#187A7B"
PINK = "#D6527C"
GRAY = "#8B8B8B"
RED = "#C93C3C"
GREEN = "#2E8B57"
BLUE = "#356A9A"
PURPLE = "#7655A6"

STATUS_MAP = {
    "PLANEJAMENTO": "EM PROJETO",
    "AGUARDANDO INICIO": "EM PROJETO",
    "EXECUCAO": "EM OBRA",
    "FINALIZACAO": "EM FINALIZAÇÃO",
    "PAUSADO": "PAUSADAS",
}
STATUS_ORDER = ["EM PROJETO", "EM OBRA", "EM FINALIZAÇÃO", "PAUSADAS"]
STATUS_COLORS = {
    "EM PROJETO": GOLD,
    "EM OBRA": TEAL,
    "EM FINALIZAÇÃO": PINK,
    "PAUSADAS": GRAY,
}
RISCO_ORDER = ["BAIXO RISCO", "MEDIO RISCO", "ALTO RISCO"]
RISCO_COLORS = {"BAIXO RISCO": GREEN, "MEDIO RISCO": GOLD, "ALTO RISCO": RED}
TIPO_COLORS = {
    "Inauguração": ORANGE,
    "Reforma": BLUE,
    "Repasse": TEAL,
    "Mudança de Ponto": PURPLE,
    "Virada Digital": GOLD,
    "Migração": PINK,
}

# ---- Board Geral (replica a aba Board_geral da planilha) ----
BOARD_TIPOS = ["INAUGURACAO", "REFORMA", "REPASSE", "MUDANCA DE PONTO", "MIGRACAO", "VIRADA DIGITAL"]
BOARD_MARCAS = {
    "SPOLETO": "Spoleto",
    "CHINA IN BOX": "China in Box",
    "GENDAI": "Gendai",
    "KONI": "Koni",
    "ASA ACAI": "Asa Açaí",
    "PIZZARIA SPOLETO": "Pizzaria Spoleto",
    "CASA DO PAO DE QUEIJO": "Casa do Pão de Queijo",
    "2V": "2V",
}
BOARD_FASES = ["PLANEJAMENTO", "EXECUCAO", "FINALIZACAO", "PAUSADO", "AGUARDANDO INICIO"]
FASE_LABELS = {
    "PLANEJAMENTO": "Planejamento",
    "EXECUCAO": "Execução",
    "FINALIZACAO": "Finalização",
    "PAUSADO": "Pausado",
    "AGUARDANDO INICIO": "Aguardando início",
}
FASE_COLORS = {
    "Planejamento": GOLD,
    "Execução": TEAL,
    "Finalização": PINK,
    "Pausado": GRAY,
    "Aguardando início": BLUE,
}


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return " ".join(texto.split())


def soma_numerica(valores):
    serie = pd.to_numeric(pd.Series(valores), errors="coerce")
    return float(serie.fillna(0).sum())


def pct(num, den):
    return num / den if den else 0


def url_download_sharepoint(url: str) -> str:
    """Adiciona download=1 ao link compartilhado do SharePoint."""
    p = urlparse(url)
    query = dict(parse_qsl(p.query, keep_blank_values=True))
    query["download"] = "1"
    return urlunparse((p.scheme, p.netloc, p.path, p.params, urlencode(query), p.fragment))


@st.cache_data(ttl=INTERVALO_ATUALIZACAO_SEG, show_spinner=False)
def baixar_sharepoint(url: str):
    """Tenta baixar a planilha diretamente do link compartilhado.

    Funciona quando o link permite download direto para o usuário/ambiente.
    Se o SharePoint exigir uma sessão corporativa que o Python não possua,
    a aplicação usa o arquivo local como contingência.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
    }
    r = requests.get(url_download_sharepoint(url), headers=headers, timeout=30, allow_redirects=True)
    r.raise_for_status()
    content_type = (r.headers.get("content-type") or "").lower()
    conteudo = r.content
    # XLSX e um ZIP e normalmente inicia com PK.
    if not conteudo.startswith(b"PK"):
        if "text/html" in content_type or b"<html" in conteudo[:500].lower():
            raise PermissionError(
                "O SharePoint solicitou autenticação. Sincronize a planilha pelo OneDrive "
                "ou use o arquivo local até que o acesso direto esteja liberado."
            )
        raise ValueError("O conteúdo recebido do SharePoint não parece ser um arquivo XLSX.")
    return conteudo


@st.cache_data(ttl=INTERVALO_ATUALIZACAO_SEG, show_spinner=False)
def carregar_excel(fonte: str, arquivo_local: str):
    """Retorna bytes do XLSX + descricao da fonte efetivamente usada."""
    if fonte == "SharePoint":
        try:
            conteudo = baixar_sharepoint(SHAREPOINT_URL)
            return conteudo, "SharePoint (online)", None
        except Exception as e:
            caminho = Path(arquivo_local)
            if caminho.exists():
                return caminho.read_bytes(), "Arquivo local (contingência)", str(e)
            raise

    caminho = Path(arquivo_local)
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    return caminho.read_bytes(), "Arquivo local", None


@st.cache_data(ttl=INTERVALO_ATUALIZACAO_SEG, show_spinner=False)
def carregar_dados(fonte: str, arquivo_local: str):
    conteudo, fonte_usada, aviso_fonte = carregar_excel(fonte, arquivo_local)

    # ---------- METAS X REALIZADO ----------
    df_meta_raw = pd.read_excel(BytesIO(conteudo), sheet_name=ABA_METAS, header=None, engine="openpyxl")
    registros = []
    marca_atual = None
    total_oficial_meta = None
    total_oficial_realizado = None
    lendo_total = False

    for _, row in df_meta_raw.iterrows():
        col_a = row.iloc[0] if len(row) > 0 else None
        col_b = normalizar_texto(row.iloc[1]) if len(row) > 1 else ""
        nome_a = normalizar_texto(col_a)

        # A linha TOTAL da planilha e a fonte oficial para os KPIs executivos.
        # Coluna Q (indice 16) corresponde ao campo "Total" da aba Inaugurações 2026.
        if nome_a == "TOTAL":
            lendo_total = True

        if lendo_total:
            if col_b == "ORCADO":
                valor_total = pd.to_numeric(row.iloc[16] if len(row) > 16 else None, errors="coerce")
                if pd.isna(valor_total):
                    meses = list(row.iloc[2:8]) + list(row.iloc[9:15])
                    valor_total = soma_numerica(meses)
                total_oficial_meta = int(valor_total)
            elif col_b == "REALIZADO":
                valor_total = pd.to_numeric(row.iloc[16] if len(row) > 16 else None, errors="coerce")
                if pd.isna(valor_total):
                    meses = list(row.iloc[2:8]) + list(row.iloc[9:15])
                    valor_total = soma_numerica(meses)
                total_oficial_realizado = int(valor_total)
                break
            continue

        if col_b == "ORCADO":
            if pd.notna(col_a):
                marca_atual = str(col_a).strip()
            if marca_atual:
                meses = list(row.iloc[2:8]) + list(row.iloc[9:15])
                registros.append({"Marca": marca_atual, "Meta": int(soma_numerica(meses)), "Realizado": 0})
        elif col_b == "REALIZADO" and marca_atual:
            meses = list(row.iloc[2:8]) + list(row.iloc[9:15])
            if registros and registros[-1]["Marca"] == marca_atual:
                registros[-1]["Realizado"] = int(soma_numerica(meses))

    df_metas = pd.DataFrame(registros)
    if df_metas.empty:
        raise ValueError("Não foi possível localizar Orçado/Realizado na aba Inaugurações 2026.")
    df_metas["Atingimento"] = df_metas.apply(lambda x: pct(x["Realizado"], x["Meta"]), axis=1)
    df_metas["Faltam"] = (df_metas["Meta"] - df_metas["Realizado"]).clip(lower=0)

    # ---------- BASE GERAL ----------
    df = pd.read_excel(BytesIO(conteudo), sheet_name=ABA_GERAL, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    obrigatorias = ["Status", "Data Prevista", "Data Realizada", "Marca", "Restaurante", "Tipo", "Etapa", "Passagem de bastão", "Data Contrato"]
    faltantes = [c for c in obrigatorias if c not in df.columns]
    if faltantes:
        raise ValueError(f"Colunas não encontradas na aba Geral: {', '.join(faltantes)}")

    # Copia da base Geral sem transformacoes, usada pela pagina Board Geral.
    geral_raw = df.copy()

    for c in ["Data Prevista", "Data Realizada", "Passagem de bastão", "Data Contrato"]:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    for c in ["Status", "Tipo", "Etapa", "Risco", "Marca", "Responsável"]:
        if c in df.columns:
            df[f"{c}_norm"] = df[c].map(normalizar_texto)

    # Mantem apenas os tipos solicitados, sem misturar REFORMA / REPASSE.
    df["Tipo Painel"] = df["Tipo_norm"].map(TIPO_LABELS)
    base_tipos = df[df["Tipo_norm"].isin(TIPOS_PROJETO)].copy()

    # Carteira ativa de todos os tipos. Usa Ano Previsto quando disponível;
    # caso contrario, usa o ano da Data Prevista.
    if "Ano Previsto" in base_tipos.columns:
        ano_prev = pd.to_numeric(base_tipos["Ano Previsto"], errors="coerce")
        filtro_ano = ano_prev.eq(ANO) | (ano_prev.isna() & base_tipos["Data Prevista"].dt.year.eq(ANO))
    else:
        filtro_ano = base_tipos["Data Prevista"].dt.year.eq(ANO)

    ativos = base_tipos[(base_tipos["Status_norm"] == "PREVISTO") & filtro_ano].copy()
    ativos["Grupo Status"] = ativos["Etapa_norm"].map(STATUS_MAP)

    # Inclui os previstos mesmo se a etapa ainda nao estiver classificada;
    # para graficos de status, somente os reconhecidos entram na agregacao.
    hoje = pd.Timestamp.today().normalize()
    ativos["Prazo Planejado (dias)"] = (ativos["Data Contrato"] - ativos["Passagem de bastão"]).dt.days
    ativos.loc[ativos["Prazo Planejado (dias)"] < 0, "Prazo Planejado (dias)"] = pd.NA
    ativos["Dias Consumidos"] = (hoje - ativos["Data Contrato"]).dt.days
    ativos.loc[ativos["Dias Consumidos"] < 0, "Dias Consumidos"] = pd.NA
    ativos["Dias até Prevista"] = (ativos["Data Prevista"] - hoje).dt.days
    ativos["% Prazo Consumido"] = ativos["Dias Consumidos"] / ativos["Prazo Planejado (dias)"].replace(0, pd.NA)

    # Projetos realizados em 2026, para prazo real.
    realizados_tipos = base_tipos[
        (base_tipos["Status_norm"] == "REALIZADO")
        & (base_tipos["Data Realizada"].dt.year == ANO)
    ].copy()
    realizados_tipos["Prazo Planejado (dias)"] = (realizados_tipos["Data Contrato"] - realizados_tipos["Passagem de bastão"]).dt.days
    realizados_tipos["Tempo Real (dias)"] = (realizados_tipos["Data Realizada"] - realizados_tipos["Passagem de bastão"]).dt.days
    realizados_tipos["Desvio do prazo (dias)"] = (realizados_tipos["Data Realizada"] - realizados_tipos["Data Contrato"]).dt.days
    realizados_tipos["No Prazo"] = realizados_tipos["Desvio do prazo (dias)"].le(0)
    for c in ["Prazo Planejado (dias)", "Tempo Real (dias)"]:
        realizados_tipos.loc[realizados_tipos[c] < 0, c] = pd.NA

    # Inauguracoes realizadas - mantidas para comparar com meta.
    inauguracoes_realizadas = realizados_tipos[realizados_tipos["Tipo_norm"] == "INAUGURACAO"].copy()

    totais_oficiais = {
        "Meta": total_oficial_meta if total_oficial_meta is not None else int(df_metas["Meta"].sum()),
        "Realizado": total_oficial_realizado if total_oficial_realizado is not None else int(df_metas["Realizado"].sum()),
    }

    return df_metas, ativos, realizados_tipos, inauguracoes_realizadas, totais_oficiais, geral_raw, fonte_usada, aviso_fonte


# =========================================================
# BOARD GERAL - mesmos calculos da aba Board_geral, sobre a aba Geral
# =========================================================
def calcular_board(geral: pd.DataFrame, mes_ref: int):
    """Reproduz a aba Board_geral da planilha a partir da aba Geral.

    Regras (identicas as formulas COUNTIFS da planilha):
    - Carteira = Status "PREVISTO".
    - Marca/Tipo/Responsavel/Etapa comparados sem acento e sem diferenca de caixa.
    - Em andamento = Total - Pausado - Aguardando inicio.
    - Realizado 2026/2025 = Status "REALIZADO" com Ano Previsto = 2026/2025.
    - Mesmo periodo 2025 = Realizado 2025 com Mes Previsto <= mes atual.
    """
    g = geral.copy()
    g["_status"] = g["Status"].map(normalizar_texto)
    g["_tipo"] = g["Tipo"].map(normalizar_texto)
    g["_etapa"] = g["Etapa"].map(normalizar_texto)
    marca = g["Marca"].map(normalizar_texto)
    g["_marca"] = marca.where(marca.isin(BOARD_MARCAS), "OUTRAS")

    if "Responsável" in g.columns:
        resp_orig = g["Responsável"].map(lambda v: "" if pd.isna(v) else str(v).strip())
    else:
        resp_orig = pd.Series("", index=g.index)
    g["_resp"] = resp_orig.map(normalizar_texto).replace({"": "SEM RESPONSAVEL", "?": "SEM RESPONSAVEL"})
    rotulos_resp = {"SEM RESPONSAVEL": "Sem responsável"}
    for norm, orig in zip(g["_resp"], resp_orig):
        if norm not in rotulos_resp and orig:
            rotulos_resp[norm] = orig

    g["_ano"] = pd.to_numeric(g["Ano Previsto"], errors="coerce") if "Ano Previsto" in g.columns else float("nan")
    mes_col = "Mês Previsto" if "Mês Previsto" in g.columns else None
    g["_mes"] = pd.to_numeric(g[mes_col], errors="coerce") if mes_col else float("nan")

    prev = g[(g["_status"] == "PREVISTO") & (g["_tipo"].isin(BOARD_TIPOS))].copy()
    real = g[(g["_status"] == "REALIZADO") & (g["_tipo"].isin(BOARD_TIPOS))].copy()
    real26 = real[real["_ano"] == 2026]
    real25 = real[real["_ano"] == 2025]
    real25_per = real25[real25["_mes"] <= mes_ref]

    # Responsaveis: os que possuem carteira ativa, do maior para o menor.
    resp_ordem = prev["_resp"].value_counts().index.tolist()

    # ---------- KPIs ----------
    kpis = {"TOTAL": len(prev)}
    for t in BOARD_TIPOS:
        kpis[t] = int((prev["_tipo"] == t).sum())

    # ---------- Tabela 1: Responsavel x Tipo x Marca ----------
    cols_marca = list(BOARD_MARCAS.items())
    if (prev["_marca"] == "OUTRAS").any():
        cols_marca.append(("OUTRAS", "Outras marcas"))

    linhas = []
    for r in resp_ordem:
        for t in BOARD_TIPOS:
            sub = prev[(prev["_resp"] == r) & (prev["_tipo"] == t)]
            if sub.empty:
                continue
            linha = {"Responsável": rotulos_resp.get(r, r), "Tipo": TIPO_LABELS[t]}
            for chave, rotulo in cols_marca:
                linha[rotulo] = int((sub["_marca"] == chave).sum())
            total = len(sub)
            pausado = int((sub["_etapa"] == "PAUSADO").sum())
            aguard = int((sub["_etapa"] == "AGUARDANDO INICIO").sum())
            linha["Total"] = total
            linha["Pausado"] = pausado
            linha["Aguardando início"] = aguard
            linha["Em andamento"] = total - pausado - aguard
            linhas.append(linha)
    tab_marca = pd.DataFrame(linhas)
    if not tab_marca.empty:
        tot = {"Responsável": "TOTAL", "Tipo": ""}
        for c in tab_marca.columns[2:]:
            tot[c] = int(tab_marca[c].sum())
        tab_marca = pd.concat([tab_marca, pd.DataFrame([tot])], ignore_index=True)

    # ---------- Tabela 2: Responsavel x Fase x Tipo ----------
    linhas = []
    for r in resp_ordem:
        for fase in BOARD_FASES + ["REALIZADO 2026"]:
            linha = {"Responsável": rotulos_resp.get(r, r),
                     "Fase": FASE_LABELS.get(fase, "Realizado 2026")}
            for t in BOARD_TIPOS:
                if fase == "REALIZADO 2026":
                    n = int(((real26["_resp"] == r) & (real26["_tipo"] == t)).sum())
                else:
                    n = int(((prev["_resp"] == r) & (prev["_tipo"] == t) & (prev["_etapa"] == fase)).sum())
                linha[TIPO_LABELS[t]] = n
            linhas.append(linha)
    tab_fase = pd.DataFrame(linhas)
    cols_tipo = [TIPO_LABELS[t] for t in BOARD_TIPOS]
    tab_fase["Total"] = tab_fase[cols_tipo].sum(axis=1)

    # ---------- Resumo por tipo ----------
    def por_tipo(df_, filtro=None):
        base = df_ if filtro is None else df_[filtro(df_)]
        return [int((base["_tipo"] == t).sum()) for t in BOARD_TIPOS]

    em_andamento = por_tipo(prev, lambda d: d["_etapa"].isin(["PLANEJAMENTO", "EXECUCAO", "FINALIZACAO"]))
    pausados = por_tipo(prev, lambda d: d["_etapa"].isin(["PAUSADO", "AGUARDANDO INICIO"]))
    r26 = por_tipo(real26)
    r25 = por_tipo(real25)
    r25p = por_tipo(real25_per)
    total_tipo = [a + b + c for a, b, c in zip(em_andamento, pausados, r26)]
    delta = [a - b for a, b in zip(r26, r25p)]

    resumo = pd.DataFrame(
        [total_tipo, em_andamento, pausados, r26, r25, r25p, delta],
        index=["Total (carteira + realizado 2026)", "Total em andamento", "Total pausados / não iniciados",
               "Total realizado 2026", "Realizado 2025", f"Realizado mesmo período/2025 (até mês {mes_ref})",
               "Δ mesmo período (2026 − 2025)"],
        columns=cols_tipo,
    )
    resumo["Total"] = resumo.sum(axis=1)
    resumo = resumo.reset_index().rename(columns={"index": "Indicador"})

    return {
        "kpis": kpis,
        "tab_marca": tab_marca,
        "tab_fase": tab_fase,
        "resumo": resumo,
        "cols_tipo": cols_tipo,
        "resp_ordem": [rotulos_resp.get(r, r) for r in resp_ordem],
    }


def card_html(titulo, valor, subtitulo, cor):
    return f"""
    <div style="background:white;border-radius:16px;padding:18px 20px;border-top:6px solid {cor};
                box-shadow:0 2px 10px rgba(0,0,0,0.07);min-height:122px;">
        <div style="font-size:13px;color:#666;font-weight:700;letter-spacing:.4px;">{titulo}</div>
        <div style="font-size:36px;color:{DARK};font-weight:800;line-height:1.15;margin-top:5px;">{valor}</div>
        <div style="font-size:12px;color:#777;margin-top:5px;">{subtitulo}</div>
    </div>
    """


def grafico_vazio(texto):
    fig = go.Figure()
    fig.add_annotation(text=texto, x=.5, y=.5, showarrow=False, font=dict(size=16, color="#777"))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    fig.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10), paper_bgcolor="white", plot_bgcolor="white")
    return fig


def destacar_total(linha):
    """Realca linhas de total nas tabelas do Board Geral."""
    if str(linha.iloc[0]).upper().startswith("TOTAL"):
        return ["font-weight:700;background-color:#FFF0E1"] * len(linha)
    return [""] * len(linha)


# =========================================================
# ESTILO
# =========================================================
st.markdown(f"""
<style>
.stApp {{ background: linear-gradient(180deg,#FFF8F1 0,#F7F7F7 260px); }}
.block-container {{ padding-top: 1rem; padding-bottom: 2.5rem; max-width: 1500px; }}
h1,h2,h3 {{ color: {DARK}; letter-spacing:-0.02em; }}
[data-testid="stSidebar"] {{ background:#FFFFFF; border-right:1px solid #ECECEC; }}
[data-testid="stTabs"] button {{ font-weight:700; }}
div[data-testid="stDataFrame"] {{ border-radius:14px; overflow:hidden; box-shadow:0 2px 12px rgba(0,0,0,.05); }}
.section-note {{ color:#737373;font-size:12px;margin-top:-8px;margin-bottom:10px; }}
.source-pill {{display:inline-block;background:#fff;border:1px solid #e8e8e8;border-radius:20px;padding:6px 12px;color:#555;font-size:12px;}}
</style>
""", unsafe_allow_html=True)

logo_path = Path(__file__).parent / LOGO_ARQUIVO
head_logo, head_title = st.columns([1.15, 4], vertical_alignment="center")
with head_logo:
    if logo_path.exists():
        # Usa exatamente o arquivo enviado pelo usuário; nenhuma logo é recriada por HTML/CSS.
        st.image(str(logo_path), use_container_width=True)
    else:
        st.warning(f"Logo não encontrada: {LOGO_ARQUIVO}")
with head_title:
    st.markdown(f"""
    <div style="background:{ORANGE};padding:18px 24px;border-radius:18px;margin-bottom:8px;">
      <div style="font-size:13px;color:white;font-weight:700;letter-spacing:1.2px;">PRÉ-OP | GESTÃO DE PROJETOS</div>
      <div style="font-size:32px;color:white;font-weight:800;">Dashboard de Projetos 2026</div>
      <div style="font-size:14px;color:white;opacity:.92;">Metas, carteira, risco, prazos e Board Geral por responsável</div>
    </div>
    """, unsafe_allow_html=True)

caminho_default = Path(__file__).parent / ARQUIVO_PADRAO
with st.sidebar:
    st.header("Fonte de dados")
    fonte_escolhida = st.radio("Origem", ["SharePoint", "Arquivo local"], index=0)
    st.caption("SharePoint é a fonte padrão. O painel tenta atualizar a base online a cada 1 hora.")
    caminho_digitado = st.text_input("Arquivo local / contingência", value=str(caminho_default))
    if st.button("🔄 Atualizar dados agora", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    df_metas, ativos, realizados_tipos, inauguracoes_realizadas, totais_oficiais, geral_raw, fonte_usada, aviso_fonte = carregar_dados(
        fonte_escolhida, caminho_digitado
    )
except Exception as e:
    st.error(f"Erro ao carregar a planilha: {e}")
    st.stop()

status_col1, status_col2 = st.columns([3, 2])
with status_col1:
    st.markdown(f"<span class='source-pill'>Fonte carregada: <b>{fonte_usada}</b></span>", unsafe_allow_html=True)
with status_col2:
    st.caption(f"Última leitura do painel: {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}")
if aviso_fonte:
    st.warning(
        "O acesso direto ao SharePoint não foi autenticado pelo Python. O painel está usando o arquivo local como contingência. "
        f"Detalhe: {aviso_fonte}"
    )

# =========================================================
# FILTROS
# =========================================================
marcas_disponiveis = sorted(set(df_metas["Marca"].dropna()) | set(ativos["Marca"].dropna()) | set(realizados_tipos["Marca"].dropna()))
tipos_disponiveis = [TIPO_LABELS[t] for t in TIPOS_PROJETO]

with st.sidebar:
    st.divider()
    st.header("Filtros")
    st.caption("Os filtros valem para todas as páginas, exceto o Board Geral, que usa a base completa da aba Geral.")
    marcas_filtro = st.multiselect("Marca", marcas_disponiveis, default=marcas_disponiveis)
    tipos_filtro = st.multiselect("Tipo de projeto", tipos_disponiveis, default=tipos_disponiveis)
    responsaveis = sorted(ativos["Responsável"].dropna().astype(str).unique()) if "Responsável" in ativos else []
    responsaveis_filtro = st.multiselect("Responsável", responsaveis, default=responsaveis)

metas = df_metas[df_metas["Marca"].isin(marcas_filtro)].copy()
ativos_f = ativos[(ativos["Marca"].isin(marcas_filtro)) & (ativos["Tipo Painel"].isin(tipos_filtro))].copy()
real_f = realizados_tipos[(realizados_tipos["Marca"].isin(marcas_filtro)) & (realizados_tipos["Tipo Painel"].isin(tipos_filtro))].copy()
real_inaug = inauguracoes_realizadas[inauguracoes_realizadas["Marca"].isin(marcas_filtro)].copy()
if responsaveis:
    ativos_f = ativos_f[ativos_f["Responsável"].astype(str).isin(responsaveis_filtro)].copy()

# =========================================================
# ABAS
# =========================================================
tab_exec, tab_oper, tab_prazos, tab_board, tab_detalhe = st.tabs([
    "Visão Executiva", "Visão Operacional", "Gestão de Prazos", "Board Geral", "Detalhamento"
])

# =========================================================
# VISAO EXECUTIVA - FOCO EM GRAFICOS
# =========================================================
with tab_exec:
    # Orçado total = soma das linhas por marca. Meta = 80% do orçado.
    orcado_total = int(metas["Meta"].sum())
    meta_total = int(round(orcado_total * META_PERCENTUAL))
    realizado_total = int(metas["Realizado"].sum())
    atingimento = pct(realizado_total, meta_total)
    faltam = max(meta_total - realizado_total, 0)
    carteira_inaug = len(ativos_f[ativos_f["Tipo_norm"] == "INAUGURACAO"])
    cobertura = pct(realizado_total + carteira_inaug, meta_total)

    c1, c2, c3, c4, c5 = st.columns(5)
    for col, t, v, s, cor in [
        (c1, "ORÇADO TOTAL", orcado_total, "lojas previstas no orçamento", BLUE),
        (c2, "META 2026", meta_total, "80% do total orçado", ORANGE),
        (c3, "REALIZADO", realizado_total, "inaugurações concluídas", TEAL),
        (c4, "ATINGIMENTO", f"{atingimento:.1%}", "realizado ÷ meta", GOLD),
        (c5, "FALTAM", faltam, "para atingir a meta", PINK),
    ]:
        with col:
            st.markdown(card_html(t, v, s, cor), unsafe_allow_html=True)

    st.write("")
    left, right = st.columns([0.9, 1.4])
    with left:
        st.subheader("Atingimento da meta")
        gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=atingimento * 100,
            number={"suffix": "%", "valueformat": ".1f"},
            delta={"reference": 100, "relative": False, "suffix": " p.p."},
            gauge={
                "axis": {"range": [0, max(110, atingimento * 100 + 10)]},
                "bar": {"color": ORANGE},
                "steps": [
                    {"range": [0, 70], "color": "#F3F3F3"},
                    {"range": [70, 100], "color": "#FFF0E1"},
                ],
                "threshold": {"line": {"color": GREEN, "width": 4}, "thickness": .8, "value": 100},
            },
            title={"text": f"{realizado_total} realizadas de {meta_total}"},
        ))
        gauge.update_layout(height=360, margin=dict(l=30, r=30, t=60, b=20), paper_bgcolor="white")
        st.plotly_chart(gauge, use_container_width=True)
        st.caption(f"Cobertura com a carteira ativa de inaugurações: {cobertura:.1%}")

    with right:
        st.subheader("Orçado x realizado por marca")
        plot_df = metas.rename(columns={"Meta": "Orçado"}).melt(id_vars="Marca", value_vars=["Orçado", "Realizado"], var_name="Indicador", value_name="Quantidade")
        fig = px.bar(
            plot_df, x="Marca", y="Quantidade", color="Indicador", barmode="group", text_auto=True,
            color_discrete_map={"Orçado": "#C9C9C9", "Realizado": ORANGE}
        )
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), plot_bgcolor="white", paper_bgcolor="white",
                          legend_title_text="", xaxis_title="", yaxis_title="Lojas")
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Carteira ativa por tipo de projeto")
        tipo_counts = ativos_f["Tipo Painel"].value_counts().reindex(tipos_disponiveis, fill_value=0).reset_index()
        tipo_counts.columns = ["Tipo", "Quantidade"]
        tipo_counts = tipo_counts[tipo_counts["Quantidade"] > 0]
        if not tipo_counts.empty:
            fig = px.pie(tipo_counts, names="Tipo", values="Quantidade", hole=.55, color="Tipo", color_discrete_map=TIPO_COLORS)
            fig.update_traces(textposition="inside", textinfo="percent+label")
            fig.update_layout(height=390, margin=dict(l=10, r=10, t=10, b=10), legend_title_text="")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio("Sem projetos ativos nos filtros selecionados"), use_container_width=True)

    with c2:
        st.subheader("Risco da carteira")
        risco_counts = ativos_f.get("Risco_norm", pd.Series(index=ativos_f.index, dtype=str)).value_counts().reindex(RISCO_ORDER, fill_value=0)
        rd = risco_counts.rename_axis("Risco").reset_index(name="Quantidade")
        fig = px.bar(rd, x="Risco", y="Quantidade", text_auto=True, color="Risco", color_discrete_map=RISCO_COLORS)
        fig.update_layout(height=390, showlegend=False, plot_bgcolor="white", paper_bgcolor="white", xaxis_title="", yaxis_title="Projetos")
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Cobertura da meta por marca")
    pipe_inaug = ativos_f[ativos_f["Tipo_norm"] == "INAUGURACAO"].groupby("Marca").size().rename("Pipeline ativo")
    cobertura_df = metas[["Marca", "Meta", "Realizado"]].merge(pipe_inaug, on="Marca", how="left")
    cobertura_df["Pipeline ativo"] = cobertura_df["Pipeline ativo"].fillna(0).astype(int)
    cobertura_df["Cobertura projetada"] = cobertura_df["Realizado"] + cobertura_df["Pipeline ativo"]
    cob_melt = cobertura_df.melt(
        id_vars="Marca", value_vars=["Meta", "Realizado", "Cobertura projetada"], var_name="Indicador", value_name="Quantidade"
    )
    fig = px.bar(
        cob_melt, x="Marca", y="Quantidade", color="Indicador", barmode="group", text_auto=True,
        color_discrete_map={"Meta": "#C9C9C9", "Realizado": ORANGE, "Cobertura projetada": TEAL}
    )
    fig.update_layout(height=410, plot_bgcolor="white", paper_bgcolor="white", legend_title_text="", xaxis_title="", yaxis_title="Lojas")
    st.plotly_chart(fig, use_container_width=True)

# =========================================================
# VISAO OPERACIONAL - TODOS OS TIPOS DE PROJETO
# =========================================================
with tab_oper:
    st.subheader("Carteira operacional 2026")
    st.markdown("<div class='section-note'>Inclui Inauguração, Reforma, Repasse, Mudança de Ponto, Virada Digital e Migração.</div>", unsafe_allow_html=True)

    tipo_counts = ativos_f["Tipo Painel"].value_counts().reindex(tipos_disponiveis, fill_value=0)
    cols = st.columns(3)
    for i, tipo in enumerate(tipos_disponiveis):
        with cols[i % 3]:
            st.markdown(card_html(tipo.upper(), int(tipo_counts[tipo]), "projetos previstos em 2026", TIPO_COLORS[tipo]), unsafe_allow_html=True)

    st.write("")
    st.subheader("Etapa por tipo de projeto")
    status_base = ativos_f[ativos_f["Grupo Status"].notna()].copy()
    if not status_base.empty:
        status_tipo = status_base.groupby(["Tipo Painel", "Grupo Status"]).size().reset_index(name="Quantidade")
        fig = px.bar(
            status_tipo, x="Tipo Painel", y="Quantidade", color="Grupo Status", barmode="stack", text_auto=True,
            category_orders={"Grupo Status": STATUS_ORDER, "Tipo Painel": tipos_disponiveis},
            color_discrete_map=STATUS_COLORS,
        )
        fig.update_layout(height=430, plot_bgcolor="white", paper_bgcolor="white", legend_title_text="", xaxis_title="", yaxis_title="Projetos")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.plotly_chart(grafico_vazio("Sem etapas classificadas nos filtros selecionados"), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Risco por tipo")
        risco_tipo = ativos_f[ativos_f.get("Risco_norm", "").isin(RISCO_ORDER)].groupby(["Tipo Painel", "Risco_norm"]).size().reset_index(name="Quantidade")
        if not risco_tipo.empty:
            fig = px.bar(
                risco_tipo, x="Tipo Painel", y="Quantidade", color="Risco_norm", barmode="stack", text_auto=True,
                color_discrete_map=RISCO_COLORS, category_orders={"Tipo Painel": tipos_disponiveis, "Risco_norm": RISCO_ORDER}
            )
            fig.update_layout(height=390, plot_bgcolor="white", paper_bgcolor="white", legend_title_text="", xaxis_title="", yaxis_title="Projetos")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio("Sem informação de risco"), use_container_width=True)

    with c2:
        st.subheader("Projetos por responsável")
        if "Responsável" in ativos_f.columns and not ativos_f.empty:
            resp = ativos_f.groupby("Responsável").size().reset_index(name="Projetos").sort_values("Projetos")
            fig = px.bar(resp, x="Projetos", y="Responsável", orientation="h", text_auto=True)
            fig.update_traces(marker_color=ORANGE, textposition="outside")
            fig.update_layout(height=max(390, 36 * len(resp) + 80), plot_bgcolor="white", paper_bgcolor="white", xaxis_title="Projetos", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio("Sem informação de responsável"), use_container_width=True)

# =========================================================
# GESTAO DE PRAZOS
# =========================================================
with tab_prazos:
    st.subheader("Gestão de Prazos")
    st.markdown(
        "<div class='section-note'><b>Base do cálculo:</b> Prazo Planejado = Data Contrato − Passagem de bastão | "
        "Tempo Real = Data Realizada − Passagem de bastão | Desvio = Data Realizada − Data Contrato | "
        "Idade ativa = Hoje − Data Contrato.</div>",
        unsafe_allow_html=True,
    )

    prazo_planejado = ativos_f["Prazo Planejado (dias)"].dropna()
    prazo_medio = prazo_planejado.mean() if not prazo_planejado.empty else 0
    dias_consumidos = ativos_f["Dias Consumidos"].dropna()
    idade_media = dias_consumidos.mean() if not dias_consumidos.empty else 0
    realizados_validos = real_f[real_f["Desvio do prazo (dias)"].notna()].copy()
    aderencia = pct(int(realizados_validos["No Prazo"].sum()), len(realizados_validos))
    atrasados = realizados_validos[realizados_validos["Desvio do prazo (dias)"] > 0]
    atraso_medio = atrasados["Desvio do prazo (dias)"].mean() if not atrasados.empty else 0
    tempo_real = real_f["Tempo Real (dias)"].dropna()
    tempo_real_medio = tempo_real.mean() if not tempo_real.empty else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    prazo_cards = [
        (c1, "PRAZO PLANEJADO MÉDIO", f"{prazo_medio:.0f} dias", "Passagem de bastão → Data Contrato", BLUE),
        (c2, "TEMPO REAL MÉDIO", f"{tempo_real_medio:.0f} dias", "Passagem de bastão → Data Realizada", ORANGE),
        (c3, "IDADE MÉDIA ATIVA", f"{idade_media:.0f} dias", "Data Contrato → hoje", TEAL),
        (c4, "ADERÊNCIA AO PRAZO", f"{aderencia:.1%}", "Data Realizada ≤ Data Contrato", GREEN if aderencia >= .8 else GOLD),
        (c5, "ATRASO MÉDIO", f"{atraso_medio:.0f} dias", "somente projetos atrasados", RED if atraso_medio > 0 else GREEN),
    ]
    for col, t, v, s, cor in prazo_cards:
        with col:
            st.markdown(card_html(t, v, s, cor), unsafe_allow_html=True)

    st.write("")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Prazo planejado médio por tipo")
        base = ativos_f[ativos_f["Prazo Planejado (dias)"].notna()].copy()
        if not base.empty:
            pm = base.groupby("Tipo Painel")["Prazo Planejado (dias)"].mean().round().reset_index().sort_values("Prazo Planejado (dias)")
            fig = px.bar(pm, x="Prazo Planejado (dias)", y="Tipo Painel", orientation="h", text="Prazo Planejado (dias)", color="Tipo Painel", color_discrete_map=TIPO_COLORS)
            fig.update_traces(texttemplate="%{text:.0f} dias", textposition="outside")
            fig.update_layout(height=400, showlegend=False, plot_bgcolor="white", paper_bgcolor="white", xaxis_title="Dias", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio("Sem Passagem de bastão + Data Contrato"), use_container_width=True)

    with c2:
        st.subheader("Tempo real médio por tipo")
        base = real_f[real_f["Tempo Real (dias)"].notna()].copy()
        if not base.empty:
            tr = base.groupby("Tipo Painel")["Tempo Real (dias)"].mean().round().reset_index().sort_values("Tempo Real (dias)")
            fig = px.bar(tr, x="Tempo Real (dias)", y="Tipo Painel", orientation="h", text="Tempo Real (dias)", color="Tipo Painel", color_discrete_map=TIPO_COLORS)
            fig.update_traces(texttemplate="%{text:.0f} dias", textposition="outside")
            fig.update_layout(height=400, showlegend=False, plot_bgcolor="white", paper_bgcolor="white", xaxis_title="Dias", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio("Sem Passagem de bastão + Data Realizada"), use_container_width=True)

    st.subheader("Projetos ativos próximos ou acima da Data Prevista")
    prazo_status = ativos_f[ativos_f["Data Prevista"].notna()].copy()
    prazo_status["Situação prazo"] = pd.cut(
        prazo_status["Dias até Prevista"],
        bins=[-10**9, -1, 15, 30, 10**9],
        labels=["ATRASADO", "ATÉ 15 DIAS", "16 A 30 DIAS", "ACIMA DE 30 DIAS"],
    )
    ordem_prazo = ["ATRASADO", "ATÉ 15 DIAS", "16 A 30 DIAS", "ACIMA DE 30 DIAS"]
    pc = prazo_status["Situação prazo"].value_counts().reindex(ordem_prazo, fill_value=0).reset_index()
    pc.columns = ["Situação", "Projetos"]
    fig = px.bar(pc, x="Situação", y="Projetos", text_auto=True, color="Situação",
                 color_discrete_map={"ATRASADO": RED, "ATÉ 15 DIAS": ORANGE, "16 A 30 DIAS": GOLD, "ACIMA DE 30 DIAS": GREEN})
    fig.update_layout(height=360, showlegend=False, plot_bgcolor="white", paper_bgcolor="white", xaxis_title="", yaxis_title="Projetos")
    st.plotly_chart(fig, use_container_width=True)

# =========================================================
# BOARD GERAL - replica da aba Board_geral, calculada sobre a aba Geral
# =========================================================
with tab_board:
    mes_ref = datetime.now().month
    board = calcular_board(geral_raw, mes_ref)
    kpis = board["kpis"]

    st.subheader("Board Geral")
    st.markdown(
        "<div class='section-note'>Mesma lógica da aba Board_geral da planilha, calculada sobre a aba Geral: "
        "carteira = Status <b>Previsto</b>; Em andamento = Total − Pausado − Aguardando início; "
        f"Realizado 2025/2026 pelo Ano Previsto; mesmo período = Mês Previsto ≤ {mes_ref} (mês atual). "
        "Esta página não usa os filtros da barra lateral.</div>",
        unsafe_allow_html=True,
    )

    # ---- Cards: total e por tipo ----
    cards = st.columns(7)
    with cards[0]:
        st.markdown(card_html("EM ANDAMENTO", kpis["TOTAL"], "total de projetos previstos", DARK), unsafe_allow_html=True)
    for col, t in zip(cards[1:], BOARD_TIPOS):
        rotulo = TIPO_LABELS[t]
        with col:
            st.markdown(card_html(rotulo.upper(), kpis[t], "previstos", TIPO_COLORS[rotulo]), unsafe_allow_html=True)

    # ---- Tabela 1: responsavel x tipo x marca ----
    st.write("")
    st.subheader("Carteira por responsável, tipo e marca")
    st.markdown("<div class='section-note'>Linhas sem projetos foram ocultadas.</div>", unsafe_allow_html=True)
    if board["tab_marca"].empty:
        st.info("Sem projetos previstos na aba Geral.")
    else:
        st.dataframe(
            board["tab_marca"].style.apply(destacar_total, axis=1),
            use_container_width=True, hide_index=True,
            height=min(38 * (len(board["tab_marca"]) + 1) + 3, 760),
        )

    # ---- Graficos ----
    tab_fase = board["tab_fase"]
    cols_tipo = board["cols_tipo"]
    g1, g2 = st.columns(2)
    with g1:
        st.subheader("Fases da carteira por responsável")
        fase_plot = tab_fase[tab_fase["Fase"] != "Realizado 2026"].groupby(["Responsável", "Fase"], sort=False)["Total"].sum().reset_index()
        fase_plot = fase_plot[fase_plot["Total"] > 0]
        if not fase_plot.empty:
            fig = px.bar(
                fase_plot, x="Responsável", y="Total", color="Fase", barmode="stack", text_auto=True,
                category_orders={"Responsável": board["resp_ordem"], "Fase": [FASE_LABELS[f] for f in BOARD_FASES]},
                color_discrete_map=FASE_COLORS,
            )
            fig.update_layout(height=400, plot_bgcolor="white", paper_bgcolor="white", legend_title_text="", xaxis_title="", yaxis_title="Projetos")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio("Sem projetos previstos"), use_container_width=True)

    with g2:
        st.subheader("Realizado 2026 x mesmo período de 2025")
        resumo = board["resumo"]
        lin_26 = resumo[resumo["Indicador"] == "Total realizado 2026"].iloc[0]
        lin_25 = resumo[resumo["Indicador"].str.startswith("Realizado mesmo período")].iloc[0]
        comp = pd.DataFrame({
            "Tipo": cols_tipo * 2,
            "Período": ["2026"] * len(cols_tipo) + ["Mesmo período 2025"] * len(cols_tipo),
            "Quantidade": [int(lin_26[c]) for c in cols_tipo] + [int(lin_25[c]) for c in cols_tipo],
        })
        fig = px.bar(
            comp, x="Tipo", y="Quantidade", color="Período", barmode="group", text_auto=True,
            color_discrete_map={"2026": ORANGE, "Mesmo período 2025": "#C9C9C9"},
            category_orders={"Tipo": cols_tipo},
        )
        fig.update_layout(height=400, plot_bgcolor="white", paper_bgcolor="white", legend_title_text="", xaxis_title="", yaxis_title="Projetos")
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    # ---- Tabela 2: responsavel x fase x tipo ----
    st.subheader("Tipo e fase por responsável")
    st.dataframe(
        tab_fase, use_container_width=True, hide_index=True,
        height=min(38 * (len(tab_fase) + 1) + 3, 760),
    )

    # ---- Resumo total ----
    st.subheader("Totais e comparativo com 2025")
    st.dataframe(
        board["resumo"], use_container_width=True, hide_index=True,
    )

# =========================================================
# DETALHAMENTO
# =========================================================
with tab_detalhe:
    st.subheader("Projetos ativos")
    colunas = [c for c in [
        "Tipo Painel", "Marca", "Restaurante", "Passagem de bastão", "Data Contrato", "Data Prevista", "Prazo Planejado (dias)",
        "Dias Consumidos", "% Prazo Consumido", "Dias até Prevista", "Responsável", "Etapa", "Grupo Status",
        "Risco", "Comentários"
    ] if c in ativos_f.columns]
    detalhe = ativos_f[colunas].copy()
    for c in ["Passagem de bastão", "Data Contrato", "Data Prevista"]:
        if c in detalhe:
            detalhe[c] = pd.to_datetime(detalhe[c], errors="coerce").dt.strftime("%d/%m/%Y")
    if "% Prazo Consumido" in detalhe:
        detalhe["% Prazo Consumido"] = detalhe["% Prazo Consumido"].map(lambda x: f"{x:.0%}" if pd.notna(x) else "")
    st.dataframe(detalhe.sort_values(["Tipo Painel", "Marca", "Restaurante"]), use_container_width=True, hide_index=True)

    st.subheader("Projetos realizados em 2026 - prazos")
    cols_real = [c for c in [
        "Tipo Painel", "Marca", "Restaurante", "Passagem de bastão", "Data Contrato", "Data Prevista", "Data Realizada",
        "Prazo Planejado (dias)", "Tempo Real (dias)", "Desvio do prazo (dias)", "Responsável"
    ] if c in real_f.columns]
    real_ex = real_f[cols_real].copy()
    for c in ["Passagem de bastão", "Data Contrato", "Data Prevista", "Data Realizada"]:
        if c in real_ex:
            real_ex[c] = real_ex[c].dt.strftime("%d/%m/%Y")
    if "Desvio do prazo (dias)" in real_ex:
        real_ex = real_ex.sort_values("Desvio do prazo (dias)", ascending=False)
    st.dataframe(real_ex, use_container_width=True, hide_index=True)

st.caption(
    "Fonte padrão: SharePoint. Tipos operacionais: Inauguração, Reforma, Repasse, Mudança de Ponto, Virada Digital e Migração. "
    "Prazo planejado = Data Contrato − Passagem de bastão. Tempo real = Data Realizada − Passagem de bastão. "
    "Desvio = Data Realizada − Data Contrato. Idade ativa = Hoje − Data Contrato. Atualização automática a cada 1 hora."
)
