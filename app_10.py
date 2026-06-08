import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import base64
from pathlib import Path

# ── Paleta Constructora Londres ────────────────────────────────────────────────
AZUL     = "#1A1846"
NARANJO  = "#E18426"
HUESO    = "#F6F1E9"
GRIS     = "#ECECEC"
AZUL_MED = "#3D3B65"
TERRACOTA= "#B04D2F"
ARENA    = "#D9B98A"

st.set_page_config(
    page_title="Evaluación Subcontratos — Constructora Londres",
    page_icon="🏗️",
    layout="wide"
)

# ── CSS marca Londres ──────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@300;400;500;600;700&display=swap');

  html, body, [class*="css"] {{
    font-family: 'Hanken Grotesk', sans-serif;
    background-color: {HUESO};
    color: {AZUL};
  }}

  /* Header principal */
  .header-bar {{
    background-color: {AZUL};
    padding: 18px 32px;
    border-radius: 10px;
    display: flex;
    align-items: center;
    gap: 20px;
    margin-bottom: 24px;
  }}
  .header-bar h1 {{
    color: white;
    font-size: 1.6rem;
    font-weight: 600;
    margin: 0;
  }}
  .header-bar span {{
    color: {NARANJO};
    font-size: 0.9rem;
    font-weight: 500;
  }}

  /* KPI cards */
  .kpi-card {{
    background: white;
    border-radius: 10px;
    padding: 18px 20px;
    border-left: 4px solid {NARANJO};
    box-shadow: 0 2px 8px rgba(26,24,70,0.07);
  }}
  .kpi-label {{
    font-size: 0.78rem;
    font-weight: 500;
    color: {AZUL_MED};
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }}
  .kpi-value {{
    font-size: 2rem;
    font-weight: 700;
    color: {AZUL};
    line-height: 1.1;
  }}
  .kpi-sub {{
    font-size: 0.82rem;
    color: {NARANJO};
    font-weight: 600;
  }}

  /* Divisor naranja */
  .divider-naranja {{
    height: 3px;
    background: linear-gradient(90deg, {NARANJO}, transparent);
    border: none;
    margin: 28px 0 20px 0;
    border-radius: 2px;
  }}

  /* Títulos de sección */
  .seccion-titulo {{
    font-size: 1.1rem;
    font-weight: 700;
    color: {AZUL};
    margin-bottom: 4px;
  }}
  .seccion-sub {{
    font-size: 0.8rem;
    color: {AZUL_MED};
    margin-bottom: 16px;
  }}

  /* Sidebar */
  section[data-testid="stSidebar"] {{
    background-color: {AZUL} !important;
  }}
  section[data-testid="stSidebar"] * {{
    color: white !important;
  }}
  section[data-testid="stSidebar"] .stMultiSelect span {{
    background-color: {NARANJO} !important;
  }}

  /* Botones y tags */
  .stMultiSelect [data-baseweb="tag"] {{
    background-color: {NARANJO} !important;
  }}
</style>
""", unsafe_allow_html=True)


# ── Helpers ────────────────────────────────────────────────────────────────────
COLOR_ESTADO = {
    "APROBADO":  "#2ECC71",
    "MEJORAR":   NARANJO,
    "REPROBADO": TERRACOTA,
}

COLOR_AREA = {
    "TERRENO": AZUL,
    "RRHH":    AZUL_MED,
    "SSOMA":   NARANJO,
    "CALIDAD": ARENA,
}

def plotly_layout(fig):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Hanken Grotesk, sans-serif", color=AZUL),
        margin=dict(t=15, b=15, l=10, r=10),
    )
    fig.update_xaxes(gridcolor=GRIS, zeroline=False)
    fig.update_yaxes(gridcolor=GRIS, zeroline=False)
    return fig

def top5_bar(df_area, area, color):
    top = (
        df_area.groupby("SUBCONTRATO")[area]
        .mean().dropna().reset_index()
        .rename(columns={area: "Nota"})
        .sort_values("Nota", ascending=False).head(5)
        .sort_values("Nota", ascending=True)
    )
    top["Nota"] = top["Nota"].round(2)
    fig = px.bar(top, x="Nota", y="SUBCONTRATO", orientation="h",
                 text="Nota", range_x=[1, 7],
                 color_discrete_sequence=[color])
    fig.update_traces(textposition="outside", marker_line_width=0)
    fig = plotly_layout(fig)
    fig.update_layout(height=240, showlegend=False)
    return fig

@st.cache_data
def cargar_datos(file_bytes):
    wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb["Consolidado"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[5]
    data   = [r for r in rows[6:] if r[0] is not None and r[19] in ("APROBADO", "MEJORAR", "REPROBADO")]
    df = pd.DataFrame(data, columns=header)
    df = df.rename(columns={
        "N° EVA":             "N_EVA",
        "CÓDIGO OBRA":        "CODIGO_OBRA",
        "FECHA EVALUACIÓN":   "FECHA",
        "NOMBRE SUBCONTRATO": "SUBCONTRATO",
        "NOTA FINAL (1 a 7)": "NOTA_FINAL",
        "ESTADO":             "ESTADO",
        "TERRENO":            "TERRENO",
        "RRHH":               "RRHH",
        "SSOMA":              "SSOMA",
        "CALIDAD":            "CALIDAD",
    })
    for col in ["NOTA_FINAL", "TERRENO", "RRHH", "SSOMA", "CALIDAD"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["N_EVA"]       = pd.to_numeric(df["N_EVA"], errors="coerce")
    df["OBRA"]        = df["OBRA"].str.strip().str.title()
    df["SUBCONTRATO"] = df["SUBCONTRATO"].str.strip().str.title()
    df["ESTADO"]      = df["ESTADO"].str.strip().str.upper()
    if "ACTIVIDAD" in df.columns:
        df["ACTIVIDAD"] = df["ACTIVIDAD"].str.strip().str.title()
    return df


# ── Header con logo ────────────────────────────────────────────────────────────
# Intenta logo.png primero, luego logo.jpg
logo_path = Path("logo.png") if Path("logo.png").exists() else Path("logo.jpg")
if logo_path.exists():
    ext = "png" if logo_path.suffix == ".png" else "jpeg"
    with open(logo_path, "rb") as f:
        logo_b64 = base64.b64encode(f.read()).decode()
    logo_html = f'<img src="data:image/{ext};base64,{logo_b64}" style="height:70px;">'
else:
    logo_html = '<span style="color:#E18426; font-size:2rem;">🏗️</span>'

st.markdown(f"""
<div style="
  background-color: white;
  padding: 24px 32px;
  border-radius: 12px;
  display: flex;
  align-items: center;
  gap: 28px;
  border-bottom: 4px solid {NARANJO};
  margin-bottom: 28px;
  box-shadow: 0 2px 12px rgba(26,24,70,0.08);
">
  {logo_html}
  <div style="border-left: 2px solid {GRIS}; padding-left: 24px;">
    <div style="font-size: 2rem; font-weight: 800; color: {AZUL}; line-height: 1.1;">
      Evaluación de Subcontratos
    </div>
    <div style="font-size: 0.95rem; color: {NARANJO}; font-weight: 500; margin-top: 4px;">
      Sistema de Gestión de Proveedores
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Upload ─────────────────────────────────────────────────────────────────────
excel_path = Path("consolidado.xlsx")
if not excel_path.exists():
    st.error("No se encontró el archivo consolidado.xlsx")
    st.stop()

df = cargar_datos(excel_path.read_bytes())

# ── Filtros sidebar ────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Filtros")
    obras_disp = sorted(df["OBRA"].dropna().unique())
    obras_sel  = st.multiselect("Obra", obras_disp, default=obras_disp)
    evas_disp  = sorted(df["N_EVA"].dropna().unique())
    eva_sel    = st.multiselect("N° Evaluación", evas_disp, default=evas_disp)
    estados_sel = st.multiselect("Estado", ["APROBADO", "MEJORAR", "REPROBADO"],
                                  default=["APROBADO", "MEJORAR", "REPROBADO"])

df_f = df[
    df["OBRA"].isin(obras_sel) &
    df["N_EVA"].isin(eva_sel)  &
    df["ESTADO"].isin(estados_sel)
]

if df_f.empty:
    st.warning("Sin datos para los filtros seleccionados.")
    st.stop()

# ── KPIs ───────────────────────────────────────────────────────────────────────
total      = len(df_f)
aprobados  = (df_f["ESTADO"] == "APROBADO").sum()
mejorar    = (df_f["ESTADO"] == "MEJORAR").sum()
reprobados = (df_f["ESTADO"] == "REPROBADO").sum()
nota_prom  = df_f["NOTA_FINAL"].mean()

k1, k2, k3, k4, k5 = st.columns(5)
kpis = [
    (k1, "Total evaluaciones", total, ""),
    (k2, "✅ Aprobado",  aprobados,  f"{aprobados/total*100:.0f}%"),
    (k3, "⚠️ Mejorar",   mejorar,    f"{mejorar/total*100:.0f}%"),
    (k4, "❌ Reprobado", reprobados, f"{reprobados/total*100:.0f}%"),
    (k5, "Nota promedio", f"{nota_prom:.2f}", "sobre 7.0"),
]
for col, label, val, sub in kpis:
    col.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{val}</div>
      <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ── Fila 1 ─────────────────────────────────────────────────────────────────────
c1, c2 = st.columns(2)

with c1:
    st.markdown('<div class="seccion-titulo">Distribución de estados por obra</div>', unsafe_allow_html=True)
    estado_obra = df_f.groupby(["OBRA","ESTADO"]).size().reset_index(name="Cantidad")
    fig1 = px.bar(estado_obra, x="OBRA", y="Cantidad", color="ESTADO",
                  color_discrete_map=COLOR_ESTADO, barmode="stack", text_auto=True)
    fig1.update_layout(xaxis_tickangle=-30, legend_title="Estado")
    st.plotly_chart(plotly_layout(fig1), use_container_width=True)

with c2:
    st.markdown('<div class="seccion-titulo">Evolución de nota promedio</div>', unsafe_allow_html=True)
    evo = df_f.groupby("N_EVA")["NOTA_FINAL"].mean().reset_index()
    evo.columns = ["Evaluación", "Nota promedio"]
    fig2 = px.line(evo, x="Evaluación", y="Nota promedio", markers=True,
                   range_y=[1,7], color_discrete_sequence=[NARANJO])
    fig2.add_hline(y=5.5, line_dash="dash", line_color="#2ECC71", annotation_text="Aprobado ≥ 5.5")
    fig2.add_hline(y=4.0, line_dash="dash", line_color=NARANJO,  annotation_text="Mejorar ≥ 4.0")
    st.plotly_chart(plotly_layout(fig2), use_container_width=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ── Fila 2 ─────────────────────────────────────────────────────────────────────
c3, c4 = st.columns(2)

with c3:
    st.markdown('<div class="seccion-titulo">Nota promedio por área evaluada</div>', unsafe_allow_html=True)
    areas = ["TERRENO", "RRHH", "SSOMA", "CALIDAD"]
    notas_area = df_f[areas].mean().reset_index()
    notas_area.columns = ["Área", "Nota"]
    fig3 = px.bar(notas_area, x="Área", y="Nota", color="Área",
                  color_discrete_map=COLOR_AREA,
                  text=notas_area["Nota"].round(2), range_y=[1,7])
    fig3.update_layout(showlegend=False)
    st.plotly_chart(plotly_layout(fig3), use_container_width=True)

with c4:
    st.markdown('<div class="seccion-titulo">Subcontratos con más REPROBADO</div>', unsafe_allow_html=True)
    rep = (df_f[df_f["ESTADO"]=="REPROBADO"].groupby("SUBCONTRATO").size()
           .reset_index(name="Reprobados").sort_values("Reprobados", ascending=True).tail(10))
    if rep.empty:
        st.success("¡Sin subcontratos reprobados en la selección!")
    else:
        fig4 = px.bar(rep, x="Reprobados", y="SUBCONTRATO", orientation="h",
                      color_discrete_sequence=[TERRACOTA], text_auto=True)
        st.plotly_chart(plotly_layout(fig4), use_container_width=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ── Top 5 por área ─────────────────────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">🏆 Top 5 subcontratos por área evaluada</div>', unsafe_allow_html=True)
st.markdown('<div class="seccion-sub">Nota promedio más alta en cada dimensión</div>', unsafe_allow_html=True)

a1, a2, a3, a4 = st.columns(4)
for col, area, label in [
    (a1, "TERRENO", "🔵 Terreno (PT)"),
    (a2, "RRHH",    "🟣 RRHH"),
    (a3, "SSOMA",   "🟠 SSOMA"),
    (a4, "CALIDAD", "🟡 Calidad"),
]:
    with col:
        st.markdown(f"**{label}**")
        st.plotly_chart(top5_bar(df_f, area, COLOR_AREA[area]), use_container_width=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ── Top 5 por actividad ────────────────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">🔧 Top 5 subcontratos por actividad</div>', unsafe_allow_html=True)
st.markdown('<div class="seccion-sub">Los 5 mejores dentro de cada tipo de actividad</div>', unsafe_allow_html=True)

if "ACTIVIDAD" in df_f.columns:
    actividades = sorted(df_f["ACTIVIDAD"].dropna().unique())
    act_sel = st.selectbox("Selecciona una actividad", actividades)
    df_act = df_f[df_f["ACTIVIDAD"] == act_sel]
    top_act = (df_act.groupby("SUBCONTRATO")["NOTA_FINAL"].mean().dropna()
               .reset_index().rename(columns={"NOTA_FINAL":"Nota promedio"})
               .sort_values("Nota promedio", ascending=False).head(5)
               .sort_values("Nota promedio", ascending=True))
    top_act["Nota promedio"] = top_act["Nota promedio"].round(2)
    fig_act = px.bar(top_act, x="Nota promedio", y="SUBCONTRATO", orientation="h",
                     text="Nota promedio", range_x=[1,7],
                     color_discrete_sequence=[AZUL])
    fig_act.update_traces(textposition="outside")
    fig_act.update_layout(height=300)
    st.plotly_chart(plotly_layout(fig_act), use_container_width=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ── Ranking general ────────────────────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">Ranking general de subcontratos</div>', unsafe_allow_html=True)

ranking = (df_f.groupby("SUBCONTRATO")
           .agg(Evaluaciones=("NOTA_FINAL","count"),
                Nota_promedio=("NOTA_FINAL","mean"),
                Ultimo_estado=("ESTADO","last"))
           .reset_index().sort_values("Nota_promedio", ascending=False))
ranking["Nota_promedio"] = ranking["Nota_promedio"].round(2)

def color_estado(val):
    c = {"APROBADO":"#d4edda","MEJORAR":"#fff3cd","REPROBADO":"#f8d7da"}.get(val,"")
    return f"background-color: {c}"

try:
    styled = ranking.style.map(color_estado, subset=["Ultimo_estado"])
except AttributeError:
    styled = ranking.style.applymap(color_estado, subset=["Ultimo_estado"])

st.dataframe(styled, use_container_width=True, height=400)

with st.expander("Ver detalle completo"):
    cols_show = ["N_EVA","CODIGO_OBRA","OBRA","SUBCONTRATO","ACTIVIDAD",
                 "TERRENO","RRHH","SSOMA","CALIDAD","NOTA_FINAL","ESTADO"]
    cols_show = [c for c in cols_show if c in df_f.columns]
    st.dataframe(df_f[cols_show].sort_values(["OBRA","N_EVA"]), use_container_width=True)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="text-align:center; padding:20px; color:{AZUL_MED}; font-size:0.78rem; margin-top:30px;">
  Constructora Londres · Sistema de Evaluación de Subcontratos · {pd.Timestamp.now().year}
</div>
""", unsafe_allow_html=True)
