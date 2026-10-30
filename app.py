import html
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

# ---- Resumo de aberturas (Acompanhamento de Aberturas) ----
RESUMO_ORDEM = [
    "SPOLETO", "KONI", "CHINA IN BOX", "GENDAI",
    "PIZZARIA SPOLETO", "ASA ACAI", "2V", "CASA DO PAO DE QUEIJO",
]
RESUMO_ROTULOS = {
    "SPOLETO": "Spoleto",
    "KONI": "Koni",
    "CHINA IN BOX": "China in Box",
    "GENDAI": "Gendai",
    "PIZZARIA SPOLETO": "Pizzaria Spoleto",
    "ASA ACAI": "Asa Açaí",
    "2V": "2V",
    "CASA DO PAO DE QUEIJO": "Casa do Pão de Queijo",
}
# Nome usado na aba "Inaugurações 2026" -> nome usado na aba Geral
RESUMO_ALIAS = {"CPQ": "CASA DO PAO DE QUEIJO"}


# Etapas (coluna N da aba Geral) para o grafico "Etapa por tipo de projeto"
ETAPA_LABELS = {
    "PLANEJAMENTO": "Planejamento",
    "EXECUCAO": "Execução",
    "FINALIZACAO": "Finalização",
    "PAUSADO": "Pausado",
    "AGUARDANDO INICIO": "Aguardando início",
}
ETAPA_ORDER = ["Planejamento", "Execução", "Finalização", "Pausado", "Aguardando início", "Sem etapa"]
ETAPA_COLORS = {
    "Planejamento": GOLD,
    "Execução": TEAL,
    "Finalização": PINK,
    "Pausado": GRAY,
    "Aguardando início": BLUE,
    "Sem etapa": "#C9C9C9",
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

    # Copia da base Geral sem transformacoes, usada pela pagina Resumo.
    geral_raw = df.copy()

    # Remove espacos sobrando no nome da marca (ex.: "CHINA IN BOX ").
    df["Marca"] = df["Marca"].where(df["Marca"].isna(), df["Marca"].astype(str).str.strip())

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

    # Todos os projetos previstos da aba Geral (qualquer ano e qualquer Tipo da coluna L).
    previstos_todos = df[df["Status_norm"] == "PREVISTO"].copy()
    previstos_todos["Tipo Painel"] = [
        TIPO_LABELS.get(n, str(o).strip().title() if pd.notna(o) and str(o).strip() else "Sem tipo")
        for n, o in zip(previstos_todos["Tipo_norm"], previstos_todos["Tipo"])
    ]
    previstos_todos["Etapa Painel"] = previstos_todos["Etapa_norm"].map(ETAPA_LABELS).fillna("Sem etapa")
    if "Responsável" in previstos_todos.columns:
        previstos_todos["Responsável"] = previstos_todos["Responsável"].map(
            lambda v: "Sem responsável" if pd.isna(v) or str(v).strip() in ("", "?") else str(v).strip()
        )
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

    return df_metas, ativos, realizados_tipos, inauguracoes_realizadas, totais_oficiais, geral_raw, previstos_todos, fonte_usada, aviso_fonte


# =========================================================
# RESUMO - Acompanhamento de Aberturas (inaugurações por marca)
# =========================================================
def calcular_resumo(df_metas: pd.DataFrame, geral: pd.DataFrame, ano: int = ANO):
    """Tabela "Acompanhamento de Aberturas", calculada a partir da planilha.

    - Meta 2026: Orçado da aba "Inaugurações 2026".
    - Inaugurado: aba Geral, Tipo Inauguração, Status Realizado, Ano Previsto = 2026.
    - Pausados / Não iniciados: Status Previsto com Etapa Pausado ou Aguardando início.
    - Em andamento para 2026: Previsto, fora de pausados, com ano de referência <= 2026.
    - Em andamento para 2027: Previsto, fora de pausados, com ano de referência > 2026.
      Ano de referência = Ano Previsto; se vazio, ano da Data Prevista; se também vazio,
      projetos em Execução/Finalização contam como 2026 e os demais como 2027.
    - Projeção 2026 = Inaugurado + Em andamento para 2026; seta = Projeção − Meta.
    """
    g = geral.copy()
    g["_status"] = g["Status"].map(normalizar_texto)
    g["_tipo"] = g["Tipo"].map(normalizar_texto)
    g["_etapa"] = g["Etapa"].map(normalizar_texto)
    g["_marca"] = g["Marca"].map(normalizar_texto).replace(RESUMO_ALIAS)

    ano_prev = pd.to_numeric(g["Ano Previsto"], errors="coerce")
    data_prev = pd.to_datetime(g["Data Prevista"], errors="coerce")
    ano_ref = ano_prev.fillna(data_prev.dt.year)
    sem_ano = ano_ref.isna()
    em_obra = g["_etapa"].isin(["EXECUCAO", "FINALIZACAO"])
    ano_ref = ano_ref.mask(sem_ano & em_obra, ano).mask(sem_ano & ~em_obra, ano + 1)

    inaug = g[g["_tipo"] == "INAUGURACAO"]
    realizadas = inaug[(inaug["_status"] == "REALIZADO") & (ano_prev.loc[inaug.index] == ano)]
    previstas = inaug[inaug["_status"] == "PREVISTO"]
    pausa = previstas["_etapa"].isin(["PAUSADO", "AGUARDANDO INICIO"])
    ref_prev = ano_ref.loc[previstas.index]
    em_2026 = previstas[~pausa & (ref_prev <= ano)]
    em_2027 = previstas[~pausa & (ref_prev > ano)]
    pausadas = previstas[pausa]

    metas = df_metas.copy()
    metas["_marca"] = metas["Marca"].map(normalizar_texto).replace(RESUMO_ALIAS)
    meta_por_marca = metas.groupby("_marca")["Meta"].sum()

    chaves = list(RESUMO_ORDEM)
    extras = set(realizadas["_marca"]) | set(previstas["_marca"]) | set(meta_por_marca.index)
    chaves += sorted(m for m in extras if m and m not in chaves)

    linhas = []
    for k in chaves:
        meta = int(meta_por_marca.get(k, 0))
        ina = int((realizadas["_marca"] == k).sum())
        a26 = int((em_2026["_marca"] == k).sum())
        pau = int((pausadas["_marca"] == k).sum())
        a27 = int((em_2027["_marca"] == k).sum())
        if k not in RESUMO_ORDEM and not any([meta, ina, a26, pau, a27]):
            continue
        proj = ina + a26
        linhas.append({
            "Marca": RESUMO_ROTULOS.get(k, k.title()),
            "Meta 2026": meta, "Inaugurado": ina, "Em andamento 2026": a26,
            "Pausados / Não iniciados": pau, "Em andamento 2027": a27,
            "Projeção 2026": proj, "Δ": proj - meta,
        })
    tabela = pd.DataFrame(linhas)
    total = {"Marca": "TOTAL 2026"}
    for c in tabela.columns[1:]:
        total[c] = int(tabela[c].sum())
    return tabela, total


def html_resumo(tabela: pd.DataFrame, total: dict) -> str:
    """Monta a tabela "Acompanhamento de Aberturas" em HTML."""
    colunas = [
        ("Meta 2026", "🎯", "Meta 2026", "#0B1B4D", "#0B1B4D"),
        ("Inaugurado", "🚀", "Inaugurado", "#1FA64A", "#1B8A3E"),
        ("Em andamento 2026", "⏳", "Em Andamento<br>para 2026", "#F5A800", "#E69A00"),
        ("Pausados / Não iniciados", "⏸", "Pausados / Não<br>iniciados", "#E8481F", "#D8321A"),
        ("Em andamento 2027", "📈", "Em Andamento<br>para 2027", "#2D5BFF", "#2D5BFF"),
        ("Projeção 2026", "📊", "Projeção 2026", "#3A3F5C", "#2B2F45"),
    ]

    def seta(delta):
        if delta > 0:
            return f"<span style='color:#1B8A3E;font-weight:800;margin-left:8px;'>▲ +{delta}</span>"
        if delta < 0:
            return f"<span style='color:#D8321A;font-weight:800;margin-left:8px;'>▼ {delta}</span>"
        return "<span style='color:#F5A800;font-weight:800;margin-left:8px;'>—</span>"

    th = "<th></th>" + "".join(
        f"<th style='padding:0 5px;'><div style='font-size:20px;text-align:center;'>{ico}</div>"
        f"<div style='background:{cor};color:white;border-radius:10px 10px 0 0;padding:10px 6px;"
        f"font-size:13px;font-weight:800;text-align:center;line-height:1.2;'>{rot}</div></th>"
        for _, ico, rot, cor, _ in colunas
    )

    def linha(reg, eh_total=False):
        rotulo_bg = "background:#0B1B4D;color:white;" if eh_total else "background:white;color:#222;"
        celulas = (
            f"<td style='padding:12px 14px;font-weight:800;font-size:14px;border-radius:12px 0 0 12px;"
            f"{rotulo_bg}white-space:nowrap;'>{reg['Marca']}</td>"
        )
        for nome, _, _, _, cor_txt in colunas:
            valor = reg[nome]
            extra = seta(reg["Δ"]) if nome == "Projeção 2026" else ""
            cor = "#222" if nome in ("Meta 2026", "Projeção 2026") else cor_txt
            celulas += (
                f"<td style='background:white;text-align:center;padding:12px 6px;font-size:17px;"
                f"font-weight:800;color:{cor};'>{valor}{extra}</td>"
            )
        return f"<tr style='box-shadow:0 1px 6px rgba(0,0,0,.08);'>{celulas}</tr>"

    corpo = "".join(linha(r) for r in tabela.to_dict("records")) + linha(total, eh_total=True)
    return (
        "<div style='overflow-x:auto;'>"
        "<table style='width:100%;border-collapse:separate;border-spacing:0 6px;min-width:860px;'>"
        f"<thead><tr>{th}</tr></thead><tbody>{corpo}</tbody></table></div>"
    )


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
      <div style="font-size:14px;color:white;opacity:.92;">Metas, carteira, risco, prazos e acompanhamento de aberturas</div>
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
    df_metas, ativos, realizados_tipos, inauguracoes_realizadas, totais_oficiais, geral_raw, previstos_todos, fonte_usada, aviso_fonte = carregar_dados(
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
marcas_disponiveis = sorted(set(df_metas["Marca"].dropna()) | set(ativos["Marca"].dropna()) | set(realizados_tipos["Marca"].dropna()) | set(previstos_todos["Marca"].dropna()))
tipos_disponiveis = [TIPO_LABELS[t] for t in TIPOS_PROJETO]

with st.sidebar:
    st.divider()
    st.header("Filtros")
    st.caption("Os filtros valem para todas as páginas, exceto a página Resumo, que usa a base completa.")
    marcas_filtro = st.multiselect("Marca", marcas_disponiveis, default=marcas_disponiveis)
    tipos_filtro = st.multiselect("Tipo de projeto", tipos_disponiveis, default=tipos_disponiveis)
    responsaveis = sorted(set(ativos["Responsável"].dropna().astype(str)) | set(previstos_todos["Responsável"].dropna().astype(str))) if "Responsável" in ativos else []
    responsaveis_filtro = st.multiselect("Responsável", responsaveis, default=responsaveis)

metas = df_metas[df_metas["Marca"].isin(marcas_filtro)].copy()
ativos_f = ativos[(ativos["Marca"].isin(marcas_filtro)) & (ativos["Tipo Painel"].isin(tipos_filtro))].copy()
real_f = realizados_tipos[(realizados_tipos["Marca"].isin(marcas_filtro)) & (realizados_tipos["Tipo Painel"].isin(tipos_filtro))].copy()
real_inaug = inauguracoes_realizadas[inauguracoes_realizadas["Marca"].isin(marcas_filtro)].copy()
if responsaveis:
    ativos_f = ativos_f[ativos_f["Responsável"].astype(str).isin(responsaveis_filtro)].copy()

# Todos os projetos previstos (aba Geral, qualquer ano e tipo) com os filtros da barra lateral.
previstos_f = previstos_todos[previstos_todos["Marca"].isin(marcas_filtro)].copy()
if len(tipos_filtro) < len(tipos_disponiveis):
    previstos_f = previstos_f[previstos_f["Tipo Painel"].isin(tipos_filtro)].copy()
if responsaveis and set(responsaveis_filtro) != set(responsaveis):
    previstos_f = previstos_f[previstos_f["Responsável"].astype(str).isin(responsaveis_filtro)].copy()

# =========================================================
# ABAS
# =========================================================
tab_exec, tab_oper, tab_prazos, tab_resumo, tab_agenda, tab_detalhe = st.tabs([
    "Visão Executiva", "Visão Operacional", "Gestão de Prazos", "Resumo", "Agenda", "Detalhamento"
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
    st.markdown("<div class='section-note'>Todos os projetos previstos da aba Geral, coloridos pela Etapa (coluna N).</div>", unsafe_allow_html=True)
    tipos_ordem = tipos_disponiveis + sorted(t for t in previstos_f["Tipo Painel"].unique() if t not in tipos_disponiveis)
    if not previstos_f.empty:
        status_tipo = previstos_f.groupby(["Tipo Painel", "Etapa Painel"]).size().reset_index(name="Quantidade")
        fig = px.bar(
            status_tipo, x="Tipo Painel", y="Quantidade", color="Etapa Painel", barmode="stack", text_auto=True,
            category_orders={"Etapa Painel": ETAPA_ORDER, "Tipo Painel": tipos_ordem},
            color_discrete_map=ETAPA_COLORS,
        )
        fig.update_layout(height=430, plot_bgcolor="white", paper_bgcolor="white", legend_title_text="", xaxis_title="", yaxis_title="Projetos")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.plotly_chart(grafico_vazio("Sem projetos previstos nos filtros selecionados"), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Risco por tipo")
        st.markdown("<div class='section-note'>Todos os projetos previstos da aba Geral.</div>", unsafe_allow_html=True)
        risco_tipo = previstos_f[previstos_f["Risco_norm"].isin(RISCO_ORDER)].groupby(["Tipo Painel", "Risco_norm"]).size().reset_index(name="Quantidade")
        if not risco_tipo.empty:
            fig = px.bar(
                risco_tipo, x="Tipo Painel", y="Quantidade", color="Risco_norm", barmode="stack", text_auto=True,
                color_discrete_map=RISCO_COLORS, category_orders={"Tipo Painel": tipos_ordem, "Risco_norm": RISCO_ORDER}
            )
            fig.update_layout(height=390, plot_bgcolor="white", paper_bgcolor="white", legend_title_text="", xaxis_title="", yaxis_title="Projetos")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio("Sem informação de risco"), use_container_width=True)

    with c2:
        st.subheader("Projetos por responsável")
        st.markdown("<div class='section-note'>Projetos previstos da aba Geral, por todos os tipos da coluna L.</div>", unsafe_allow_html=True)
        if not previstos_f.empty:
            resp = previstos_f.groupby(["Responsável", "Tipo Painel"]).size().reset_index(name="Projetos")
            ordem_resp = resp.groupby("Responsável")["Projetos"].sum().sort_values().index.tolist()
            fig = px.bar(
                resp, x="Projetos", y="Responsável", color="Tipo Painel", orientation="h", barmode="stack", text_auto=True,
                category_orders={"Responsável": ordem_resp, "Tipo Painel": tipos_ordem},
                color_discrete_map=TIPO_COLORS,
            )
            fig.update_layout(height=max(390, 36 * len(ordem_resp) + 100), plot_bgcolor="white", paper_bgcolor="white",
                              legend_title_text="", xaxis_title="Projetos", yaxis_title="")
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
    # ---------- Inaugurações: idade dos ativos e tempo médio por marca ----------
    st.divider()
    st.subheader("Inaugurações: idade dos projetos ativos")
    st.markdown(
        "<div class='section-note'>Todos os projetos <b>previstos</b> do tipo Inauguração (qualquer ano). "
        "Idade = hoje − Passagem de bastão. O losango marca o prazo planejado (Data Contrato − Passagem de bastão): "
        "barra verde = dentro do prazo, vermelha = acima do prazo, cinza = sem Data Contrato válida.</div>",
        unsafe_allow_html=True,
    )
    hoje_ref = pd.Timestamp.today().normalize()
    idade_base = previstos_f[previstos_f["Tipo_norm"] == "INAUGURACAO"].copy()
    idade_base["Idade (dias)"] = (hoje_ref - idade_base["Passagem de bastão"]).dt.days
    total_inaug_prev = len(idade_base)
    idade_base = idade_base[idade_base["Idade (dias)"].notna() & (idade_base["Idade (dias)"] >= 0)].copy()
    # Prazo planejado = Data Contrato - Passagem de bastão (ignora datas ausentes ou invertidas).
    idade_base["Prazo planejado (dias)"] = (idade_base["Data Contrato"] - idade_base["Passagem de bastão"]).dt.days
    idade_base.loc[~(idade_base["Prazo planejado (dias)"] > 0), "Prazo planejado (dias)"] = float("nan")
    idade_base["Situação"] = "Sem prazo planejado"
    com_prazo = idade_base["Prazo planejado (dias)"].notna()
    idade_base.loc[com_prazo & (idade_base["Idade (dias)"] <= idade_base["Prazo planejado (dias)"]), "Situação"] = "Dentro do prazo"
    idade_base.loc[com_prazo & (idade_base["Idade (dias)"] > idade_base["Prazo planejado (dias)"]), "Situação"] = "Acima do prazo"
    if not idade_base.empty:
        idade_base = idade_base.sort_values("Idade (dias)")
        ordem_y = idade_base["Restaurante"].tolist()
        cores_situacao = {"Dentro do prazo": GREEN, "Acima do prazo": RED, "Sem prazo planejado": "#C9C9C9"}
        fig = go.Figure()
        for situacao, cor in cores_situacao.items():
            bloco = idade_base[idade_base["Situação"] == situacao]
            if bloco.empty:
                continue
            fig.add_trace(go.Bar(
                x=bloco["Idade (dias)"], y=bloco["Restaurante"], orientation="h", name=situacao, marker_color=cor,
                text=[f"{int(i)} d" for i in bloco["Idade (dias)"]], textposition="outside",
                customdata=bloco[["Marca", "Etapa Painel", "Responsável", "Prazo planejado (dias)"]].values,
                hovertemplate="<b>%{y}</b><br>Marca: %{customdata[0]}<br>Etapa: %{customdata[1]}<br>"
                              "Responsável: %{customdata[2]}<br>Idade: %{x} dias<br>"
                              "Prazo planejado: %{customdata[3]} dias<extra></extra>",
            ))
        marcador = idade_base[com_prazo.loc[idade_base.index]]
        fig.add_trace(go.Scatter(
            x=marcador["Prazo planejado (dias)"], y=marcador["Restaurante"], mode="markers", name="Prazo planejado",
            marker=dict(symbol="diamond", size=11, color=DARK, line=dict(color="white", width=1)),
            hovertemplate="<b>%{y}</b><br>Prazo planejado: %{x} dias<extra></extra>",
        ))
        fig.update_layout(
            height=max(420, 26 * len(idade_base) + 140), plot_bgcolor="white", paper_bgcolor="white",
            legend_title_text="", xaxis_title="Dias desde a passagem de bastão", yaxis_title="",
            yaxis=dict(categoryorder="array", categoryarray=ordem_y), barmode="overlay",
        )
        st.plotly_chart(fig, use_container_width=True)
        n_acima = int((idade_base["Situação"] == "Acima do prazo").sum())
        n_dentro = int((idade_base["Situação"] == "Dentro do prazo").sum())
        n_sem = int((idade_base["Situação"] == "Sem prazo planejado").sum())
        st.caption(
            f"{len(idade_base)} de {total_inaug_prev} inaugurações previstas têm Passagem de bastão preenchida · "
            f"idade média: {idade_base['Idade (dias)'].mean():.0f} dias · mais antiga: {int(idade_base['Idade (dias)'].max())} dias · "
            f"acima do prazo planejado: {n_acima} · dentro do prazo: {n_dentro} · sem Data Contrato válida: {n_sem}."
        )
    else:
        st.plotly_chart(grafico_vazio("Sem inaugurações previstas com Passagem de bastão"), use_container_width=True)

    st.subheader("Inaugurações: tempo médio por marca")
    st.markdown(
        "<div class='section-note'>Inaugurações realizadas em 2026 (Data Realizada), independente do ano da passagem de bastão. "
        "Tempo = Data Realizada − Passagem de bastão.</div>",
        unsafe_allow_html=True,
    )
    tempo_base = real_f[real_f["Tipo_norm"] == "INAUGURACAO"].copy()
    tempo_base["Tempo Real (dias)"] = pd.to_numeric(tempo_base["Tempo Real (dias)"], errors="coerce")
    total_inaug_real = len(tempo_base)
    tempo_base = tempo_base[tempo_base["Tempo Real (dias)"].notna()]
    if not tempo_base.empty:
        tm = tempo_base.groupby("Marca")["Tempo Real (dias)"].agg(["mean", "count"]).reset_index()
        tm["Tempo médio (dias)"] = tm["mean"].round()
        tm["Rótulo"] = tm.apply(lambda r: f"{r['Tempo médio (dias)']:.0f} dias ({int(r['count'])})", axis=1)
        tm = tm.sort_values("Tempo médio (dias)")
        fig = px.bar(tm, x="Marca", y="Tempo médio (dias)", text="Rótulo", category_orders={"Marca": tm["Marca"].tolist()})
        fig.update_traces(marker_color=ORANGE, textposition="outside")
        fig.update_layout(height=400, plot_bgcolor="white", paper_bgcolor="white", xaxis_title="", yaxis_title="Dias")
        st.plotly_chart(fig, use_container_width=True)
        geral_media = tempo_base["Tempo Real (dias)"].mean()
        st.caption(
            f"{len(tempo_base)} de {total_inaug_real} inaugurações realizadas em 2026 têm Passagem de bastão preenchida · "
            f"média geral: {geral_media:.0f} dias · entre parênteses, a quantidade de projetos de cada marca."
        )
    else:
        st.plotly_chart(grafico_vazio("Sem inaugurações realizadas em 2026 com Passagem de bastão"), use_container_width=True)

# =========================================================
# RESUMO - Acompanhamento de Aberturas
# =========================================================
with tab_resumo:
    st.subheader("Acompanhamento de Aberturas")
    tabela_resumo, total_resumo = calcular_resumo(df_metas, geral_raw)
    st.markdown(html_resumo(tabela_resumo, total_resumo), unsafe_allow_html=True)
    st.markdown(
        "<div class='section-note' style='margin-top:10px;'>"
        "<b>Como é calculado</b> (inaugurações, aba Geral): <b>Meta 2026</b> = orçado da aba Inaugurações 2026 · "
        "<b>Inaugurado</b> = realizadas com Ano Previsto 2026 · "
        "<b>Pausados / Não iniciados</b> = previstas com etapa Pausado ou Aguardando início · "
        "<b>Em andamento 2026 / 2027</b> = demais previstas, pelo ano previsto (sem ano: em Execução/Finalização contam como 2026, "
        "as outras como 2027) · <b>Projeção 2026</b> = Inaugurado + Em andamento 2026 · seta = Projeção − Meta. "
        "Esta página não usa os filtros da barra lateral."
        "</div>",
        unsafe_allow_html=True,
    )

# =========================================================
# AGENDA - proximas inauguracoes (mes vigente + 2 meses)
# =========================================================
MESES_PT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]


def limpar_farol(valor):
    """Devolve o farol da coluna V; ignora vazios e erros do Excel (#REF!)."""
    if pd.isna(valor):
        return ""
    txt = str(valor).strip()
    return "" if txt.startswith("#") or txt.lower() == "nan" else txt


def html_agenda(sub: pd.DataFrame, hoje: pd.Timestamp) -> str:
    """Tabela: Data prevista - Unidade - Farol - Observação."""
    cab = "".join(
        f"<th style='text-align:{al};padding:8px 10px;font-size:12px;color:#555;border-bottom:2px solid {ORANGE};"
        f"white-space:nowrap;'>{t}</th>"
        for t, al in [("Data prevista", "left"), ("Unidade", "left"), ("Farol", "center"), ("Observação", "left")]
    )
    linhas = ""
    for _, r in sub.iterrows():
        atrasada = r["Data Prevista"] < hoje
        cor_data = f"color:{RED};font-weight:800;" if atrasada else f"color:{DARK};font-weight:700;"
        farol = html.escape(limpar_farol(r.get("FAROL RISCOS")))
        obs = html.escape("" if pd.isna(r.get("Comentários")) else str(r.get("Comentários")).strip())
        linhas += (
            "<tr style='border-bottom:1px solid #F0F0F0;vertical-align:top;'>"
            f"<td style='padding:9px 10px;white-space:nowrap;{cor_data}'>{r['Data Prevista'].strftime('%d/%m/%Y')}</td>"
            f"<td style='padding:9px 10px;font-weight:700;color:{DARK};min-width:220px;'>{html.escape(str(r['Restaurante']))}</td>"
            f"<td style='padding:9px 10px;text-align:center;font-size:18px;'>{farol}</td>"
            f"<td style='padding:9px 10px;color:#444;font-size:13px;line-height:1.35;'>{obs}</td>"
            "</tr>"
        )
    return (
        "<div style='background:white;border-radius:14px;padding:6px 10px;box-shadow:0 2px 10px rgba(0,0,0,.06);margin-bottom:14px;overflow-x:auto;'>"
        f"<table style='width:100%;border-collapse:collapse;'><thead><tr>{cab}</tr></thead><tbody>{linhas}</tbody></table></div>"
    )


with tab_agenda:
    st.subheader("Agenda de próximas inaugurações")
    st.markdown(
        "<div class='section-note'>Projetos <b>previstos</b> da aba Geral, pela Data Prevista, no mês atual e nos 2 meses seguintes, "
        "separados por tipo. Datas em vermelho já passaram e o projeto ainda não foi realizado. Segue os filtros da barra lateral.</div>",
        unsafe_allow_html=True,
    )
    hoje_ag = pd.Timestamp.today().normalize()
    meses_ag = [hoje_ag.to_period("M") + i for i in range(3)]
    ag = previstos_f.copy()
    ag["Data Prevista"] = pd.to_datetime(ag["Data Prevista"], errors="coerce")
    ag = ag[ag["Data Prevista"].notna()].copy()
    ag["_periodo"] = ag["Data Prevista"].dt.to_period("M")
    tipos_ag = tipos_disponiveis + sorted(t for t in ag["Tipo Painel"].unique() if t not in tipos_disponiveis)

    abas_mes = st.tabs([
        f"{MESES_PT[p.month - 1]}/{p.year} ({int((ag['_periodo'] == p).sum())})" for p in meses_ag
    ])
    for periodo, aba in zip(meses_ag, abas_mes):
        with aba:
            bloco_mes = ag[ag["_periodo"] == periodo]
            if bloco_mes.empty:
                st.info("Nenhum projeto previsto para este mês nos filtros selecionados.")
                continue
            for tipo in tipos_ag:
                sub = bloco_mes[bloco_mes["Tipo Painel"] == tipo].sort_values(["Data Prevista", "Restaurante"])
                if sub.empty:
                    continue
                st.markdown(
                    f"<div style='margin:6px 0 6px 2px;font-size:17px;font-weight:800;color:{TIPO_COLORS.get(tipo, DARK)};'>"
                    f"{html.escape(tipo)} <span style='font-size:13px;color:#777;font-weight:600;'>· {len(sub)}</span></div>",
                    unsafe_allow_html=True,
                )
                st.markdown(html_agenda(sub, hoje_ag), unsafe_allow_html=True)

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
