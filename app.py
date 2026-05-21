import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl, io, os, warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="Tax Impact Dashboard", page_icon="📊", layout="wide",
                   initial_sidebar_state="collapsed")

st.markdown("""
<style>
.main .block-container{max-width:1400px;padding:1.5rem 2rem;}
section[data-testid="stSidebar"]{display:none;}
.header-box{background:linear-gradient(135deg,#0A2463,#1E5FA8);padding:22px 32px;
    border-radius:12px;margin-bottom:24px;}
.header-title{color:white;font-size:26px;font-weight:900;letter-spacing:2px;}
.header-sub{color:#D0E8FF;font-size:13px;margin-top:4px;}
.zona-title{font-size:15px;font-weight:900;color:white;background:#0A2463;
    padding:10px 20px;border-radius:8px;margin:20px 0 14px;letter-spacing:1px;
    display:inline-block;}
.kpi-box{background:#ffffff;border-radius:10px;padding:14px 18px;
    box-shadow:0 2px 8px rgba(0,0,0,.08);text-align:center;
    max-width:300px;width:300px;margin-right:10px;}
.kpi-lbl{font-size:10px;font-weight:700;color:#888;text-transform:uppercase;
    letter-spacing:.5px;margin-bottom:6px;}
.kpi-val{font-size:22px;font-weight:900;color:#0A2463;}
</style>""", unsafe_allow_html=True)

# ── DATOS ──────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def cargar():
    import requests
    RUTA_LOCAL = r"C:\Users\Arturo Aguilar\Documents\CLOSING\CASCO\Resumen_Inversiones_Casco.xlsx"
    if os.path.exists(RUTA_LOCAL):
        wb = openpyxl.load_workbook(RUTA_LOCAL, data_only=True)
    else:
        FILE_ID = st.secrets["FILE_ID"]
        r = requests.get(f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx")
        wb = openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)
    ws = wb["DATA"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            rows.append({"id": row[0], "zona": row[1], "proyecto": row[2],
                         "metrica": row[3], "actual": row[4], "sin_tx": row[5]})
    return pd.DataFrame(rows)

df_all = cargar()

def fmt_v(v, pct=False):
    if v is None or (isinstance(v, float) and pd.isna(v)): return "—"
    return f"{v*100:.2f}%" if pct else f"${v:,.0f}"

def es_pct_metrica(m):
    return m in ["CASH ON CASH", "IRR"]

# ── HEADER ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-box">
  <div class="header-title">PORTFOLIO — TAX IMPACT DASHBOARD</div>
  <div class="header-sub">Valor Actual vs Valor Sin Impuestos · CASH ON CASH · IRR · FCF FROM FINANCING</div>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN RESUMEN FCF FROM FINANCING
# ══════════════════════════════════════════════════════════════════════════════
def resumen_fcf(df):
    def get_fcf(zona):
        sub = df[(df["zona"] == zona) & (df["metrica"] == "FCF FROM FINANCING")]
        return sub["actual"].sum(), sub["sin_tx"].sum()

    a_ca, s_ca = get_fcf("CASCO ANTIGUO")
    a_sa, s_sa = get_fcf("SANTA ANA")
    a_tot = a_ca + a_sa
    s_tot = s_ca + s_sa

    st.markdown('<div class="zona-title">RESUMEN PORTFOLIO — FCF FROM FINANCING</div>', unsafe_allow_html=True)
    st.markdown("")

    filas = [
        [("FCF Sin Impuesto (Ley Casco) — Casco Antiguo", f"${s_ca:,.0f}"),
         ("FCF Sin Impuesto (Ley Casco) — Santa Ana",     f"${s_sa:,.0f}"),
         ("FCF Sin Impuesto (Ley Casco) — Total",         f"${s_tot:,.0f}")],
        [("FCF Con Impuesto — Casco Antiguo", f"${a_ca:,.0f}"),
         ("FCF Con Impuesto — Santa Ana",     f"${a_sa:,.0f}"),
         ("FCF Con Impuesto — Total",         f"${a_tot:,.0f}")],
        [("Variación FCF — Casco Antiguo", f"${s_ca-a_ca:+,.0f}"),
         ("Variación FCF — Santa Ana",     f"${s_sa-a_sa:+,.0f}"),
         ("Variación FCF — Total",         f"${s_tot-a_tot:+,.0f}")],
    ]
    for fila in filas:
        html = '<div style="display:flex;gap:10px;justify-content:center;margin-bottom:10px;">'
        for lbl, val in fila:
            html += f"""<div class="kpi-box">
                <div class="kpi-lbl">{lbl}</div>
                <div class="kpi-val">{val}</div>
            </div>"""
        html += '</div>'
        st.markdown(html, unsafe_allow_html=True)

    st.markdown("<hr style='border:none;border-top:2px solid #e0e0e0;margin:24px 0;'>",
                unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN: render_zona
# ══════════════════════════════════════════════════════════════════════════════
def render_zona(zona, metrica_grafico, kpi_orden, default_proyecto):
    st.markdown(f'<div class="zona-title">{zona}</div>', unsafe_allow_html=True)

    df_z = df_all[df_all["zona"] == zona]
    proyectos = sorted(df_z["proyecto"].unique().tolist())

    # Dropdown ancho reducido (~4cm) con opción "Todos"
    opciones = ["Todos"] + proyectos
    default_idx = opciones.index(default_proyecto) if default_proyecto in opciones else 0
    col_sel, _ = st.columns([1, 4])
    with col_sel:
        proyecto_sel = st.selectbox("Proyecto:", opciones, index=default_idx,
                                    key=f"sel_{zona}")

    df_proy = df_z.copy() if proyecto_sel == "Todos" else df_z[df_z["proyecto"] == proyecto_sel].copy()

    # ── 3 ETIQUETAS KPI ────────────────────────────────────────────────────────
    st.markdown("")
    cards_html = '<div style="display:flex;gap:10px;flex-wrap:wrap;justify-content:center;">'
    for m in kpi_orden:
        row_m = df_proy[df_proy["metrica"] == m]
        if row_m.empty:
            diff_val = "—"
        else:
            diff = row_m["sin_tx"].sum() - row_m["actual"].sum()
            diff_val = fmt_v(diff, pct=es_pct_metrica(m))
        cards_html += f"""<div class="kpi-box">
            <div class="kpi-lbl">Variación {m}</div>
            <div class="kpi-val">{diff_val}</div>
        </div>"""
    cards_html += '</div>'
    st.markdown(cards_html, unsafe_allow_html=True)

    st.markdown("<div style='margin-top:100px;'></div>", unsafe_allow_html=True)

    # ── GRÁFICO: filtrado por selección ───────────────────────────────────────
    df_graf = df_proy[df_proy["metrica"] == metrica_grafico].copy()
    es_pct  = es_pct_metrica(metrica_grafico)

    fig = go.Figure()
    bar_width = 0.25 if proyecto_sel != "Todos" else 0.35
    fig.add_trace(go.Bar(
        name="Con Impuesto", x=df_graf["proyecto"], y=df_graf["actual"],
        marker_color="#0A2463", width=bar_width,
        text=[f"{v*100:.1f}%" if es_pct else f"${v:,.0f}" for v in df_graf["actual"]],
        textposition="outside", textfont=dict(size=15)
    ))
    fig.add_trace(go.Bar(
        name="Sin Impuesto (Ley Casco)", x=df_graf["proyecto"], y=df_graf["sin_tx"],
        marker_color="#4C9BE8", width=bar_width,
        text=[f"{v*100:.1f}%" if es_pct else f"${v:,.0f}" for v in df_graf["sin_tx"]],
        textposition="outside", textfont=dict(size=15)
    ))
    fig.update_layout(
        title=dict(text=f"{metrica_grafico} — {zona}", font=dict(size=13, color="#0A2463")),
        barmode="group", bargap=0.4, bargroupgap=0.08, plot_bgcolor="white",
        width=900, height=350,
        margin=dict(t=40, b=70, l=40, r=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis=dict(tickformat=".0%" if es_pct else "$,.0f", gridcolor="#F0F0F0"),
        xaxis=dict(tickangle=-30, tickfont=dict(size=12))
    )

    # Título y gráfico en la misma columna centrada
    nombre_filtro = proyecto_sel if proyecto_sel != "Todos" else "Todos los proyectos"
    _, col_c, _ = st.columns([0.3, 3, 0.3])
    with col_c:
        st.markdown(f"""
        <div style="background:#EFF3FA;border:1px solid #C0D0F0;border-radius:8px;
             padding:8px 16px;margin-bottom:8px;text-align:center;">
            <span style="font-size:19px;font-weight:700;color:#0A2463;">
                {nombre_filtro}
            </span>
        </div>""", unsafe_allow_html=True)
        fig.update_layout(width=None, height=350)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("<div style='margin-top:30px;'></div>", unsafe_allow_html=True)

    # ── TABLA: TODOS los valores del proyecto seleccionado ────────────────────
    filas = []
    for _, r in df_proy.iterrows():
        m   = r["metrica"]
        pct = es_pct_metrica(m)
        a   = r["actual"]
        s   = r["sin_tx"]
        d   = s - a
        filas.append({
            "Descripción / Métrica": m,
            "Con Impuesto":              fmt_v(a, pct),
            "Sin Impuesto (Ley Casco)":  fmt_v(s, pct),
            "Diferencia":            (f"{d*100:+.2f}%" if pct else f"${d:+,.0f}"),
        })

    _, col_t, _ = st.columns([1, 2, 1])
    with col_t:
        st.markdown(f"**{proyecto_sel} — Detalle completo**")
        st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

resumen_fcf(df_all)

# ══════════════════════════════════════════════════════════════════════════════
# CASCO ANTIGUO
# ══════════════════════════════════════════════════════════════════════════════
render_zona(
    zona             = "CASCO ANTIGUO",
    metrica_grafico  = "CASH ON CASH",
    kpi_orden        = ["CASH ON CASH", "FCF FROM FINANCING", "IRR"],
    default_proyecto = "MANSION BALUARTE"
)

st.markdown("<hr style='border:none;border-top:2px solid #e0e0e0;margin:32px 0;'>",
            unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SANTA ANA
# ══════════════════════════════════════════════════════════════════════════════
render_zona(
    zona             = "SANTA ANA",
    metrica_grafico  = "IRR",
    kpi_orden        = ["IRR", "FCF FROM FINANCING", "CASH ON CASH"],
    default_proyecto = "CENTRA LINK"
)
