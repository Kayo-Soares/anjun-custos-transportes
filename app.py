"""
Dashboard de Custos de Transporte - Streamlit
================================================
Reproduz os indicadores do painel Power BI (Custo Total, Custo por
Fornecedor, Custo por Pacote e Custo por KM) de forma interativa.

Como rodar:
    1. Coloque este arquivo na mesma pasta dos .xlsx mensais
       (ou use o botão de upload dentro do app)
    2. No terminal:  streamlit run streamlit_app.py
    3. Abre automaticamente no navegador em http://localhost:8501
"""

import glob
import os

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------
# 1. CONFIGURAÇÃO
# ----------------------------------------------------------------------
st.set_page_config(page_title="Custos de Transporte | 运输成本", layout="wide", page_icon="🚛")

# ----------------------------------------------------------------------
# Paleta e estilo — tema "operação logística": petróleo + âmbar
# ----------------------------------------------------------------------
COR_PRIMARIA = "#0F3D3E"      # petróleo escuro (base)
COR_PRIMARIA_CLARA = "#1F6F5C"  # petróleo médio (barras, linha)
COR_ACENTO = "#E8871E"        # âmbar (destaque / alerta)
COR_FUNDO_CARD = "#FFFFFF"
COR_TEXTO_SECUNDARIO = "#6B7280"
PALETA_ROTAS = ["#0F3D3E", "#1F6F5C", "#3B9C8B", "#7FC4B4", "#E8871E", "#F2B36B", "#C9CBA3", "#A0A8B0"]

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600&display=swap');

    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
    h1, h2, h3 {{ font-family: 'Space Grotesk', sans-serif; }}

    .kpi-card {{
        background: {COR_FUNDO_CARD};
        border-radius: 14px;
        padding: 18px 20px;
        border-left: 5px solid {COR_PRIMARIA_CLARA};
        box-shadow: 0 1px 3px rgba(15, 61, 62, 0.08), 0 1px 2px rgba(15, 61, 62, 0.06);
        margin-bottom: 8px;
    }}
    .kpi-card.acento {{ border-left-color: {COR_ACENTO}; }}
    .kpi-label {{
        font-size: 0.78rem; font-weight: 600; letter-spacing: .03em;
        color: {COR_TEXTO_SECUNDARIO}; text-transform: uppercase; margin-bottom: 4px;
    }}
    .kpi-value {{
        font-family: 'Space Grotesk', sans-serif; font-size: 1.65rem; font-weight: 700;
        color: {COR_PRIMARIA}; line-height: 1.1;
    }}
    .section-title {{
        font-family: 'Space Grotesk', sans-serif; font-weight: 700; font-size: 1.15rem;
        color: {COR_PRIMARIA}; margin-top: 6px; margin-bottom: 2px;
    }}
</style>
""", unsafe_allow_html=True)

INPUT_DIR = os.path.dirname(os.path.abspath(__file__))
SHEET_NAME = "Custo Secundaria"

COLUMN_MAP = {
    "序号\nNo.": "no",
    "两星期      Quinzena": "quinzena",
    "日期\nData": "data",
    "任务名称\nMotivo": "motivo",
    "部门/大区\nSetor": "setor",
    "车型\nModelo": "modelo",
    "车牌\nPlaca": "placa",
    "自有/三方\nProprio/Terceiro": "proprio_terceiro",
    "司机\nMotorista": "motorista",
    "承运商\nFornecedor": "fornecedor",
    "线路名称\nRota": "rota",
    "线路类型\nTipo transferencia": "tipo_transferencia",
    "方向\nDirecao": "direcao",
    "理论装载量\nCarga teoria": "carga_teoria",
    "实际装载量\nCarga real": "carga_real",       # usada como "Pacotes" no painel
    "卷                              Volumes": "volumes",
    "运输里程\nKM": "km",
    "发车时间\nSaida": "saida",
    "到车时间\nChegada": "chegada",
    "运行时间\nTempo de viagem": "tempo_viagem",
    "折旧成本\nDespesa de Depreciação": "despesa_depreciacao",
    "司机成本                            Custo Motorista": "custo_motorista",
    "运费\nFrete": "frete",
    "燃油费\nCombustivel": "combustivel",
    "维修费\nManutencao": "manutencao",
    "路桥费\nPedagio": "pedagio",
    "保险费\nSeguro": "seguro",
    "其他成本\nOutros custos": "outros_custos",
    "成本费用\nCusto total": "custo_total",
    "每包价值                    Valor por Pacote": "valor_por_pacote_orig",
    "每公里价值                Valor por KM ": "valor_por_km_orig",
    "备注\nObs": "obs",
}


# ----------------------------------------------------------------------
# 2. CARGA E LIMPEZA (com cache — só reprocessa se os arquivos mudarem)
# ----------------------------------------------------------------------
@st.cache_data(show_spinner="Lendo e consolidando planilhas de custo...")
def carregar_base(arquivos: list) -> pd.DataFrame:
    partes = []
    ignorados = []
    for arq in arquivos:
        nome_arquivo = arq.name if hasattr(arq, "name") else os.path.basename(arq)
        try:
            df = pd.read_excel(arq, sheet_name=SHEET_NAME)
        except ValueError:
            # arquivo não tem a aba "Custo Secundaria" (ex: planilha de programação) — ignora
            ignorados.append(nome_arquivo)
            continue
        df = df.rename(columns=COLUMN_MAP)
        df["arquivo_origem"] = nome_arquivo
        partes.append(df)

    if ignorados:
        st.sidebar.caption(f"⚠️ Ignorado(s) (sem aba '{SHEET_NAME}'): {', '.join(ignorados)}")

    if not partes:
        return pd.DataFrame()

    base = pd.concat(partes, ignore_index=True)
    base["data"] = pd.to_datetime(base["data"], errors="coerce")
    base["ano_mes"] = base["data"].dt.to_period("M").astype(str)

    # Remove espaços extras no início/fim dos campos de texto — evita que
    # "MAX LOG" e "MAX LOG " (com espaço) virem duas categorias diferentes
    col_texto = ["fornecedor", "rota", "motorista", "setor", "modelo", "placa",
                 "proprio_terceiro", "tipo_transferencia", "direcao", "motivo"]
    for c in col_texto:
        if c in base.columns:
            base[c] = base[c].astype(str).str.strip().replace({"nan": pd.NA})

    num_cols = ["carga_teoria", "carga_real", "volumes", "km", "custo_total",
                "frete", "combustivel", "manutencao", "pedagio", "seguro", "outros_custos"]
    for c in num_cols:
        base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0)

    base = base[(base["custo_total"] > 0) & (base["rota"].notna())]
    return base


def localizar_arquivos_na_pasta() -> list[str]:
    arquivos = sorted(glob.glob(os.path.join(INPUT_DIR, "*.xlsx")))
    arquivos = [a for a in arquivos if not os.path.basename(a).startswith("~$")]
    return arquivos


# ----------------------------------------------------------------------
# 2b. PROGRAMAÇÃO DE TRANSFERÊNCIA (planilha de estrutura diferente:
#     blocos por estação, sem cabeçalho fixo — precisa de parser próprio)
# ----------------------------------------------------------------------
DIAS_SEMANA = ["seg", "ter", "qua", "qui", "sex", "sab", "dom"]
DIAS_LABEL = {"seg": "Seg", "ter": "Ter", "qua": "Qua", "qui": "Qui",
              "sex": "Sex", "sab": "Sáb", "dom": "Dom"}


@st.cache_data(show_spinner="Lendo programação de transferência...")
def carregar_programacao(arquivo) -> pd.DataFrame:
    raw = pd.read_excel(arquivo, sheet_name=0, header=None)

    registros = []
    estacao_atual = None
    for i in range(len(raw)):
        row = raw.iloc[i]
        # Cabeçalho de bloco: coluna 0 preenchida e colunas 1-12 todas vazias
        if pd.notna(row[0]) and row[1:13].isna().all() and str(row[0]).strip().upper() != "ROTA":
            estacao_atual = str(row[0]).strip()
            continue
        # Linha de dados: coluna 0 é um número de rota, dentro de um bloco já identificado
        if estacao_atual and pd.notna(row[0]):
            try:
                rota_num = int(row[0])
            except (ValueError, TypeError):
                continue
            reg = {
                "estacao": estacao_atual,
                "rota": rota_num,
                "tipo_operacao": row[1],
                "estacao_partida": row[2],
                "iata_percurso": row[3],
                "horario_saida": row[4],
                "obs": row[12] if len(row) > 12 else None,
            }
            for idx, dia in enumerate(DIAS_SEMANA):
                reg[dia] = pd.notna(row[5 + idx])
            registros.append(reg)

    return pd.DataFrame(registros)


def localizar_arquivo_programacao_na_pasta():
    """Procura um .xlsx cujo nome sugira ser a planilha de programação."""
    candidatos = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))
    for c in candidatos:
        nome = os.path.basename(c).upper()
        if "PROGRAMA" in nome or "TRANSFER" in nome:
            return c
    return None


# ----------------------------------------------------------------------
# 3. FONTE DOS DADOS: upload manual OU pasta local
# ----------------------------------------------------------------------
st.sidebar.header("📂 Dados | 数据")
uploads = st.sidebar.file_uploader(
    "Envie os arquivos .xlsx (opcional) | 上传 .xlsx 文件（可选）", type="xlsx", accept_multiple_files=True
)

if uploads:
    base = carregar_base(uploads)
    fonte = f"{len(uploads)} arquivo(s) enviado(s) manualmente | 手动上传 {len(uploads)} 个文件"
else:
    arquivos_locais = localizar_arquivos_na_pasta()
    if not arquivos_locais:
        st.warning(
            "Nenhum .xlsx encontrado na pasta do app. "
            "Envie os arquivos pelo campo ao lado ou coloque-os na mesma pasta deste script.\n\n"
            "未在文件夹中找到 .xlsx 文件。请通过左侧上传，或将文件放在与本脚本相同的文件夹中。"
        )
        st.stop()
    base = carregar_base(arquivos_locais)
    fonte = f"{len(arquivos_locais)} arquivo(s) da pasta local | 本地文件夹中 {len(arquivos_locais)} 个文件"

st.sidebar.caption(f"Fonte | 数据来源: {fonte}")

# ----------------------------------------------------------------------
# 3b. FONTE DA PROGRAMAÇÃO DE TRANSFERÊNCIA (opcional, planilha separada)
# ----------------------------------------------------------------------
st.sidebar.header("📅 Programação | 时间表")
prog_upload = st.sidebar.file_uploader(
    "Envie a planilha de programação (opcional) | 上传时间表（可选）", type="xlsx", key="prog_upload"
)

if prog_upload is not None:
    prog_df = carregar_programacao(prog_upload)
else:
    prog_caminho = localizar_arquivo_programacao_na_pasta()
    prog_df = carregar_programacao(prog_caminho) if prog_caminho else pd.DataFrame()

# ----------------------------------------------------------------------
# 4. FILTROS
# ----------------------------------------------------------------------
st.sidebar.header("🔎 Filtros | 筛选条件")

meses = sorted(base["ano_mes"].dropna().unique())
meses_sel = st.sidebar.multiselect("Mês | 月份", meses, default=meses)

rotas = sorted(base["rota"].dropna().unique())
rotas_sel = st.sidebar.multiselect("Rota | 线路", rotas, default=rotas)

fornecedores = sorted(base["fornecedor"].dropna().unique())
fornecedores_sel = st.sidebar.multiselect("Fornecedor | 承运商", fornecedores, default=fornecedores)

df = base[
    base["ano_mes"].isin(meses_sel)
    & base["rota"].isin(rotas_sel)
    & base["fornecedor"].isin(fornecedores_sel)
]

if df.empty:
    st.warning("Nenhum registro para os filtros selecionados. | 所选筛选条件下没有记录。")
    st.stop()


# ----------------------------------------------------------------------
# ABAS PRINCIPAIS
# ----------------------------------------------------------------------
tab_custos, tab_programacao = st.tabs(["💰 Custos | 成本", "📅 Programação | 时间表"])

with tab_custos:
    # ----------------------------------------------------------------------
    # 5. INDICADORES (KPIs)
    # ----------------------------------------------------------------------
    st.title("🚛 Relatório de Transporte Secundária | 二线运输报告")
    st.caption(f"Período | 期间: {df['data'].min().strftime('%d/%m/%Y')} a {df['data'].max().strftime('%d/%m/%Y')}")

    custo_total = df["custo_total"].sum()
    total_km = df["km"].sum()
    total_pacotes = df["carga_real"].sum()
    custo_medio_pacote = custo_total / total_pacotes if total_pacotes else 0
    custo_medio_km = custo_total / total_km if total_km else 0
    total_fretes = len(df)


    def kpi_card(col, label, valor, acento=False):
        classe = "kpi-card acento" if acento else "kpi-card"
        col.markdown(
            f'<div class="{classe}"><div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{valor}</div></div>',
            unsafe_allow_html=True,
        )


    c1, c2, c3, c4, c5, c6 = st.columns(6)
    kpi_card(c1, "Custo Total | 成本总额", f"R$ {custo_total:,.2f}")
    kpi_card(c2, "Custo por Pacote | 每包价值", f"R$ {custo_medio_pacote:,.2f}")
    kpi_card(c3, "Custo por KM | 每公里价值", f"R$ {custo_medio_km:,.2f}")
    kpi_card(c4, "Nº de Fretes | 运单数量", f"{total_fretes:,}")
    kpi_card(c5, "Total de KM Rodados | 总行驶里程", f"{total_km:,.0f} km", acento=True)
    kpi_card(c6, "Total de Pacotes | 总包裹数", f"{total_pacotes:,.0f}", acento=True)

    st.divider()

    # ----------------------------------------------------------------------
    # 6. CUSTO POR FORNECEDOR + CUSTO POR PACOTE / KM (donuts)
    # ----------------------------------------------------------------------
    col1, col2, col3 = st.columns([1.2, 1, 1])

    with col1:
        st.markdown('<div class="section-title">Custo por Fornecedor | 承运商成本</div>', unsafe_allow_html=True)
        tab_fornecedor = (df.groupby("fornecedor")["custo_total"].sum()
                             .sort_values(ascending=False).reset_index())
        tab_fornecedor["percentual"] = tab_fornecedor["custo_total"] / tab_fornecedor["custo_total"].sum() * 100
        tab_fornecedor["rotulo"] = tab_fornecedor.apply(
            lambda r: f"R$ {r['custo_total']:,.0f}  ({r['percentual']:.1f}%)", axis=1)

        fig = px.bar(tab_fornecedor, x="custo_total", y="fornecedor", orientation="h",
                     text="rotulo", color="custo_total",
                     color_continuous_scale=[COR_PRIMARIA_CLARA, COR_ACENTO])
        fig.update_traces(textposition="outside", cliponaxis=False,
                           hovertemplate="<b>%{y}</b><br>Custo | 成本: R$ %{x:,.2f}<extra></extra>")
        fig.update_layout(
            yaxis=dict(autorange="reversed", title=""),
            xaxis=dict(title="Custo Total (R$) | 成本总额", showgrid=True, gridcolor="#EEF0EF"),
            height=460, margin=dict(l=0, r=60, t=10, b=0),
            coloraxis_showscale=False, plot_bgcolor="white", paper_bgcolor="white",
            font=dict(family="Inter, sans-serif", size=12, color=COR_PRIMARIA),
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown('<div class="section-title">Custo por Pacote | 每包价值</div>', unsafe_allow_html=True)
        tab_pacote = (df.groupby("rota")
                        .agg(custo_total=("custo_total", "sum"), pacotes=("carga_real", "sum"))
                        .reset_index())
        tab_pacote["custo_por_pacote"] = tab_pacote["custo_total"] / tab_pacote["pacotes"].replace(0, pd.NA)
        tab_pacote = tab_pacote.sort_values("custo_total", ascending=False)
        top_pacote = tab_pacote.head(8).copy()
        if len(tab_pacote) > 8:
            outros = pd.DataFrame([{
                "rota": "Outras rotas | 其他线路",
                "custo_total": tab_pacote["custo_total"].iloc[8:].sum(),
                "pacotes": tab_pacote["pacotes"].iloc[8:].sum(),
                "custo_por_pacote": tab_pacote["custo_total"].iloc[8:].sum() / max(tab_pacote["pacotes"].iloc[8:].sum(), 1),
            }])
            top_pacote = pd.concat([top_pacote, outros], ignore_index=True)

        fig = px.pie(top_pacote, values="custo_total", names="rota", hole=0.55,
                     color_discrete_sequence=PALETA_ROTAS)
        fig.update_traces(
            textinfo="percent", textposition="inside", textfont_size=11,
            customdata=top_pacote[["custo_por_pacote"]],
            hovertemplate="<b>%{label}</b><br>Custo | 成本: R$ %{value:,.2f}<br>Custo/pacote | 每包成本: R$ %{customdata[0]:,.2f}<extra></extra>",
            marker=dict(line=dict(color="white", width=2)),
        )
        fig.update_layout(
            height=460, margin=dict(l=0, r=0, t=10, b=60),
            legend=dict(orientation="h", yanchor="top", y=-0.05, font=dict(size=9)),
            annotations=[dict(text=f"R$ {custo_medio_pacote:,.2f}<br><span style='font-size:10px'>média geral | 平均值</span>",
                               x=0.5, y=0.5, font_size=15, showarrow=False, font_color=COR_PRIMARIA)],
            paper_bgcolor="white",
        )
        st.plotly_chart(fig, use_container_width=True)
        with st.expander("Ver tabela | 查看表格"):
            st.dataframe(tab_pacote, use_container_width=True)

    with col3:
        st.markdown('<div class="section-title">Custo por KM | 每公里价值</div>', unsafe_allow_html=True)
        tab_km = (df.groupby("rota")
                    .agg(custo_total=("custo_total", "sum"), km=("km", "sum"))
                    .reset_index())
        tab_km["custo_por_km"] = tab_km["custo_total"] / tab_km["km"].replace(0, pd.NA)
        tab_km = tab_km.sort_values("custo_total", ascending=False)
        top_km = tab_km.head(8).copy()
        if len(tab_km) > 8:
            outros = pd.DataFrame([{
                "rota": "Outras rotas | 其他线路",
                "custo_total": tab_km["custo_total"].iloc[8:].sum(),
                "km": tab_km["km"].iloc[8:].sum(),
                "custo_por_km": tab_km["custo_total"].iloc[8:].sum() / max(tab_km["km"].iloc[8:].sum(), 1),
            }])
            top_km = pd.concat([top_km, outros], ignore_index=True)

        fig = px.pie(top_km, values="custo_total", names="rota", hole=0.55,
                     color_discrete_sequence=PALETA_ROTAS)
        fig.update_traces(
            textinfo="percent", textposition="inside", textfont_size=11,
            customdata=top_km[["custo_por_km"]],
            hovertemplate="<b>%{label}</b><br>Custo | 成本: R$ %{value:,.2f}<br>Custo/km | 每公里成本: R$ %{customdata[0]:,.2f}<extra></extra>",
            marker=dict(line=dict(color="white", width=2)),
        )
        fig.update_layout(
            height=460, margin=dict(l=0, r=0, t=10, b=60),
            legend=dict(orientation="h", yanchor="top", y=-0.05, font=dict(size=9)),
            annotations=[dict(text=f"R$ {custo_medio_km:,.2f}<br><span style='font-size:10px'>média geral | 平均值</span>",
                               x=0.5, y=0.5, font_size=15, showarrow=False, font_color=COR_PRIMARIA)],
            paper_bgcolor="white",
        )
        st.plotly_chart(fig, use_container_width=True)
        with st.expander("Ver tabela | 查看表格"):
            st.dataframe(tab_km, use_container_width=True)

    st.divider()

    # ----------------------------------------------------------------------
    # 7. CUSTO MÊS A MÊS
    # ----------------------------------------------------------------------
    st.markdown('<div class="section-title">Custo Total — Mês a Mês | 月度成本趋势</div>', unsafe_allow_html=True)

    tab_mensal = (df.groupby("ano_mes")
                    .agg(custo_total=("custo_total", "sum"), fretes=("custo_total", "count"),
                         pacotes=("carga_real", "sum"), km=("km", "sum"))
                    .reset_index().sort_values("ano_mes"))
    tab_mensal["variacao_pct"] = tab_mensal["custo_total"].pct_change() * 100
    tab_mensal["custo_por_pacote"] = tab_mensal["custo_total"] / tab_mensal["pacotes"].replace(0, pd.NA)
    tab_mensal["rotulo_variacao"] = tab_mensal["variacao_pct"].map(
        lambda v: "" if pd.isna(v) else f"{'+' if v > 0 else ''}{v:.1f}%")

    fig = go.Figure()

    # Linha 1 — Custo Total (eixo esquerdo, principal)
    fig.add_trace(go.Scatter(
        x=tab_mensal["ano_mes"], y=tab_mensal["custo_total"], name="Custo Total | 成本总额 (R$)",
        mode="lines+markers", yaxis="y1",
        line=dict(color=COR_PRIMARIA_CLARA, width=3, shape="linear"),
        marker=dict(size=8, color=COR_PRIMARIA_CLARA, line=dict(color="white", width=1.5)),
        fill="tozeroy", fillcolor="rgba(31, 111, 92, 0.08)",
        hovertemplate="Custo Total | 成本总额: R$ %{y:,.2f}<extra></extra>",
    ))

    # Linha 2 — Custo por Pacote (eixo direito #1)
    fig.add_trace(go.Scatter(
        x=tab_mensal["ano_mes"], y=tab_mensal["custo_por_pacote"], name="Custo por Pacote | 每包价值 (R$)",
        mode="lines+markers", yaxis="y2",
        line=dict(color=COR_ACENTO, width=2.5, dash="dot"),
        marker=dict(size=7, color=COR_ACENTO),
        hovertemplate="Custo por Pacote | 每包价值: R$ %{y:,.2f}<extra></extra>",
    ))

    # Linha 3 — KM Rodado (eixo direito #2)
    fig.add_trace(go.Scatter(
        x=tab_mensal["ano_mes"], y=tab_mensal["km"], name="KM Rodado | 行驶里程",
        mode="lines+markers", yaxis="y3",
        line=dict(color=COR_TEXTO_SECUNDARIO, width=2, dash="dash"),
        marker=dict(size=7, color=COR_TEXTO_SECUNDARIO),
        hovertemplate="KM Rodado | 行驶里程: %{y:,.0f} km<extra></extra>",
    ))

    # rótulos de variação percentual do Custo Total, coloridos por sinal
    for _, row in tab_mensal.iterrows():
        if row["rotulo_variacao"]:
            cor = "#C0392B" if row["variacao_pct"] > 0 else "#1F8F4F"
            fig.add_annotation(x=row["ano_mes"], y=row["custo_total"], yref="y1",
                                text=row["rotulo_variacao"], showarrow=False, yshift=16,
                                font=dict(size=10, color=cor, family="Inter"))

    fig.update_layout(
        xaxis=dict(title="", showgrid=False, domain=[0, 0.86]),
        yaxis=dict(title="Custo Total | 成本总额 (R$)", showgrid=True, gridcolor="#EEF0EF",
                   range=[0, tab_mensal["custo_total"].max() * 1.3],
                   title_font=dict(color=COR_PRIMARIA_CLARA), tickfont=dict(color=COR_PRIMARIA_CLARA)),
        yaxis2=dict(title="Custo/Pacote | 每包价值 (R$)", overlaying="y", side="right", showgrid=False,
                    range=[0, tab_mensal["custo_por_pacote"].max() * 1.4],
                    title_font=dict(color=COR_ACENTO), tickfont=dict(color=COR_ACENTO)),
        yaxis3=dict(title="KM | 行驶里程", overlaying="y", side="right", anchor="free", position=0.98,
                    showgrid=False, range=[0, tab_mensal["km"].max() * 1.4],
                    title_font=dict(color=COR_TEXTO_SECUNDARIO), tickfont=dict(color=COR_TEXTO_SECUNDARIO)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, font=dict(size=11)),
        height=440, margin=dict(l=0, r=10, t=40, b=0),
        plot_bgcolor="white", paper_bgcolor="white",
        hovermode="x unified",
        font=dict(family="Inter, sans-serif", size=12, color=COR_PRIMARIA),
    )
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("Ver tabela mês a mês | 查看月度表格"):
        st.dataframe(
            tab_mensal.style.format({"custo_total": "R$ {:,.2f}", "variacao_pct": "{:+.1f}%"}),
            use_container_width=True,
        )

    # ----------------------------------------------------------------------
    # 8. BASE CONSOLIDADA (para conferência)
    # ----------------------------------------------------------------------
    with st.expander("📋 Ver base consolidada completa | 查看完整汇总数据"):
        st.dataframe(df, use_container_width=True)
        st.download_button(
            "Baixar base consolidada (CSV) | 下载汇总数据 (CSV)",
            df.to_csv(index=False).encode("utf-8-sig"),
            file_name="base_consolidada_transporte.csv",
            mime="text/csv",
        )

# ----------------------------------------------------------------------
# ABA: PROGRAMAÇÃO DE TRANSFERÊNCIA
# ----------------------------------------------------------------------
with tab_programacao:
    st.title("📅 Programação de Transferência | 转运时间表")

    if prog_df.empty:
        st.info(
            "Nenhuma planilha de programação carregada ainda.\n\n"
            "Envie o arquivo pela barra lateral (seção **📅 Programação**), ou coloque-o "
            "na mesma pasta do app com 'Programação' ou 'Transferência' no nome do arquivo.\n\n"
            "尚未加载时间表文件。请通过侧边栏上传，或将文件放在应用所在文件夹中，"
            "文件名包含 'Programação' 或 'Transferência'。"
        )
    else:
        estacoes = sorted(prog_df["estacao"].unique())
        estacao_sel = st.selectbox("Estação | 站点", estacoes)

        dados_estacao = prog_df[prog_df["estacao"] == estacao_sel].sort_values("rota")

        # rótulo de cada rota, combinando número + tipo de operação + percurso
        dados_estacao = dados_estacao.copy()
        dados_estacao["rota_label"] = dados_estacao.apply(
            lambda r: f"Rota {r['rota']} · {r['tipo_operacao']} · {r['iata_percurso']}", axis=1
        )

        # ---- Heatmap semanal (dias ativos por rota) ----
        matriz = dados_estacao.set_index("rota_label")[DIAS_SEMANA].astype(int)
        matriz.columns = [DIAS_LABEL[d] for d in DIAS_SEMANA]

        hover_horario = dados_estacao.set_index("rota_label")["horario_saida"].astype(str)

        fig = go.Figure(data=go.Heatmap(
            z=matriz.values,
            x=matriz.columns,
            y=matriz.index,
            colorscale=[[0, "#F2F3F1"], [1, COR_PRIMARIA_CLARA]],
            showscale=False,
            xgap=4, ygap=4,
            hovertemplate="<b>%{y}</b><br>Dia: %{x}<br>Ativo: %{z}<extra></extra>",
        ))
        fig.update_layout(
            title=dict(text=f"Grade Semanal — {estacao_sel} | 每周计划", font=dict(size=14, color=COR_PRIMARIA)),
            height=120 + 40 * len(matriz),
            margin=dict(l=0, r=0, t=40, b=0),
            plot_bgcolor="white", paper_bgcolor="white",
            font=dict(family="Inter, sans-serif", size=12, color=COR_PRIMARIA),
            yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig, use_container_width=True)

        # ---- Tabela detalhada (horário, percurso, observações) ----
        st.markdown('<div class="section-title">Detalhamento das Rotas | 线路详情</div>', unsafe_allow_html=True)
        tabela = dados_estacao[[
            "rota", "tipo_operacao", "estacao_partida", "iata_percurso",
            "horario_saida", "seg", "ter", "qua", "qui", "sex", "sab", "dom", "obs",
        ]].rename(columns={
            "rota": "Rota", "tipo_operacao": "Tipo Operação", "estacao_partida": "Estação Partida",
            "iata_percurso": "IATA/Percurso", "horario_saida": "Horário Saída", "obs": "Obs",
            "seg": "Seg", "ter": "Ter", "qua": "Qua", "qui": "Qui", "sex": "Sex", "sab": "Sáb", "dom": "Dom",
        })
        st.dataframe(tabela, use_container_width=True, hide_index=True)

        with st.expander("📋 Ver programação completa (todas as estações) | 查看完整时间表"):
            st.dataframe(prog_df, use_container_width=True, hide_index=True)
            st.download_button(
                "Baixar programação (CSV) | 下载时间表 (CSV)",
                prog_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="programacao_transferencia.csv",
                mime="text/csv",
            )