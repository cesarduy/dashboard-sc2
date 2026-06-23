import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import base64
from pathlib import Path
from PIL import Image

# ── Paleta Constructora Londres ────────────────────────────────────────────────
AZUL     = "#1A1846"
NARANJO  = "#E18426"
HUESO    = "#F6F1E9"
GRIS     = "#ECECEC"
AZUL_MED = "#3D3B65"
TERRACOTA= "#B04D2F"
ARENA    = "#D9B98A"

# ── Criterios de detalle por área ──────────────────────────────────────────────
CRITERIOS = {
    "TERRENO": [
        "Cumple los plazos internos de ejecución de la obra.",
        "Realiza trabajos en forma limpia, sin dañar trabajos ya realizados.",
        "Asiste a las reuniones de planificación.",
        "Cuenta con Supervisores competentes.",
        "Puntualidad.",
    ],
    "RRHH": [
        "Entrega certificado de la Inspección del Trabajo (Formulario n°30 y 30-1).",
        "Mantiene pactos de horas extras actualizado.",
        "Entrega certificado de Deuda Tributaria con sus respaldos, emitido por Tesorería General Republica.",
        "Entrega Copia de la entrega del Reglamento  Interno a los Trabajadores.",
        "Mantiene libro de Asistencia al dia.",
        "Entrega certificado y mantiene planilla de Leyes Sociales al día.",
        "Realiza y entrega Listado de Trabajadores Vigentes en obra mensual.",
        "Entrega Formulario  Pago de IVA (n°29).",
        "Presenta y entrega respaldo de Libro de Remuneraciones.",
        "Presentación de la Directiva de Funcionamiento con respaldo de jornada excepcional aprobada.",
        "Tiene curso SO10 con certificado vigente y en regla ",
    ],
    "SSOMA": [
        "Cumple con Implementación documental DS 44",
        "Participa en charlas y capacitaciones de obra.",
        "Manejo de accidentes con ingresos a mutualidad en obra",
        "Cumple con todos los protocolos Minsal que le apliquen",
        "Cumple con plan personalizado de actividades (foco en obra).",
        "Acata indicaciones de normas de seguridad  del depto SSOMA en Terreno.",
        "Colabora y participa en actividades extra de gestión en SSOMA.",
        "Utiliza de EPP y soluciones de seguridad en Terreno.",
    ],
    "CALIDAD": [
        "Cumple con las exigencias de calidad.",
        "Entrega protocolos",
        "Resuelve No Conformidades",
        "Demuestra conocimiento técnico.",
    ],
}

ETIQUETAS_CORTAS = {
    "Cumple los plazos internos de ejecución de la obra.": "Plazos",
    "Realiza trabajos en forma limpia, sin dañar trabajos ya realizados.": "Limpieza",
    "Asiste a las reuniones de planificación.": "Reuniones",
    "Cuenta con Supervisores competentes.": "Supervisores",
    "Puntualidad.": "Puntualidad",
    "Entrega certificado de la Inspección del Trabajo (Formulario n°30 y 30-1).": "F30/30-1",
    "Mantiene pactos de horas extras actualizado.": "Horas extras",
    "Entrega certificado de Deuda Tributaria con sus respaldos, emitido por Tesorería General Republica.": "Deuda tributaria",
    "Entrega Copia de la entrega del Reglamento  Interno a los Trabajadores.": "Reglamento interno",
    "Mantiene libro de Asistencia al dia.": "Libro asistencia",
    "Entrega certificado y mantiene planilla de Leyes Sociales al día.": "Leyes sociales",
    "Realiza y entrega Listado de Trabajadores Vigentes en obra mensual.": "Listado trab.",
    "Entrega Formulario  Pago de IVA (n°29).": "IVA F29",
    "Presenta y entrega respaldo de Libro de Remuneraciones.": "Libro remuner.",
    "Presentación de la Directiva de Funcionamiento con respaldo de jornada excepcional aprobada.": "Directiva jornada",
    "Tiene curso SO10 con certificado vigente y en regla ": "Curso SO10",
    "Cumple con Implementación documental DS 44": "DS44",
    "Participa en charlas y capacitaciones de obra.": "Charlas/Cap.",
    "Manejo de accidentes con ingresos a mutualidad en obra": "Accidentes",
    "Cumple con todos los protocolos Minsal que le apliquen": "Minsal",
    "Cumple con plan personalizado de actividades (foco en obra).": "Plan personal.",
    "Acata indicaciones de normas de seguridad  del depto SSOMA en Terreno.": "Normas SSOMA",
    "Colabora y participa en actividades extra de gestión en SSOMA.": "Gestión SSOMA",
    "Utiliza de EPP y soluciones de seguridad en Terreno.": "EPP",
    "Cumple con las exigencias de calidad.": "Exigencias cal.",
    "Entrega protocolos": "Protocolos",
    "Resuelve No Conformidades": "No conformid.",
    "Demuestra conocimiento técnico.": "Conocim. técn.",
}

TODOS_CRITERIOS = [c for lista in CRITERIOS.values() for c in lista]

# ── Favicon ────────────────────────────────────────────────────────────────────
_favicon_path = next((p for p in [Path("logo_icono.jpeg"), Path("logo_icono.png"), Path("logo.jpeg"), Path("logo.jpg")] if p.exists()), None)
_page_icon = Image.open(_favicon_path) if _favicon_path else "🏗️"

st.set_page_config(
    page_title="Evaluación Subcontratos — Constructora Londres",
    page_icon=_page_icon,
    layout="wide"
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@300;400;500;600;700&display=swap');
  html, body, [class*="css"] {{ font-family: 'Hanken Grotesk', sans-serif; background-color: {HUESO}; color: #0D0C25; }}
  h1, h2, h3, h4, h5, h6, p, span, div, label {{ color: #0D0C25; }}
  .stMarkdown, .stText {{ color: #0D0C25 !important; }}
  .kpi-card {{ background: white; border-radius: 10px; padding: 18px 20px; border-left: 4px solid {NARANJO}; box-shadow: 0 2px 8px rgba(26,24,70,0.07); }}
  .kpi-label {{ font-size: 0.78rem; font-weight: 500; color: {AZUL_MED}; text-transform: uppercase; letter-spacing: 0.05em; }}
  .kpi-value {{ font-size: 2rem; font-weight: 700; color: {AZUL}; line-height: 1.1; }}
  .kpi-sub {{ font-size: 0.82rem; color: {NARANJO}; font-weight: 600; }}
  .divider-naranja {{ height: 3px; background: linear-gradient(90deg, {NARANJO}, transparent); border: none; margin: 28px 0 20px 0; border-radius: 2px; }}
  .seccion-titulo {{ font-size: 1.1rem; font-weight: 700; color: {AZUL}; margin-bottom: 4px; }}
  .seccion-sub {{ font-size: 0.8rem; color: {AZUL_MED}; margin-bottom: 16px; }}
  section[data-testid="stSidebar"] {{ background-color: {AZUL} !important; }}
  section[data-testid="stSidebar"] * {{ color: white !important; }}
  section[data-testid="stSidebar"] .stMultiSelect span {{ background-color: {NARANJO} !important; }}
  section[data-testid="stSidebar"] .stMultiSelect label,
  section[data-testid="stSidebar"] .stMultiSelect label *,
  section[data-testid="stSidebar"] [data-testid="stWidgetLabel"],
  section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] * {{ background: transparent !important; background-color: transparent !important; background-image: none !important; box-shadow: none !important; }}
  .stMultiSelect [data-baseweb="tag"] {{ background-color: {NARANJO} !important; }}
  .stSelectbox > div > div, div[data-baseweb="select"] > div {{ background-color: #ffffff !important; border-color: #cccccc !important; }}
  div[data-baseweb="select"] span {{ color: {AZUL} !important; }}
  div[data-baseweb="popover"] ul {{ background-color: #ffffff !important; }}
  div[data-baseweb="popover"] li {{ color: {AZUL} !important; }}
  div[data-baseweb="popover"] li:hover {{ background-color: {GRIS} !important; }}
  .alerta-card {{ background: #fff8f0; border: 1px solid {NARANJO}; border-left: 4px solid {TERRACOTA}; border-radius: 8px; padding: 12px 16px; margin-bottom: 8px; }}
</style>
""", unsafe_allow_html=True)

# ── Helpers ────────────────────────────────────────────────────────────────────
COLOR_ESTADO = {"APROBADO": "#2ECC71", "MEJORAR": NARANJO, "REPROBADO": TERRACOTA}
COLOR_AREA   = {"TERRENO": AZUL, "RRHH": AZUL_MED, "SSOMA": NARANJO, "CALIDAD": ARENA}

def plotly_layout(fig):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Hanken Grotesk, sans-serif", color="#1A1846"),
        margin=dict(t=15, b=15, l=10, r=10),
    )
    fig.update_xaxes(gridcolor="#CCCCCC", zeroline=False, tickfont=dict(color="#1A1846", size=12))
    fig.update_yaxes(gridcolor="#CCCCCC", zeroline=False, tickfont=dict(color="#1A1846", size=12))
    fig.update_layout(legend=dict(font=dict(color="#1A1846", size=12)))
    return fig

def top5_bar(df_area, area, color):
    top = (df_area.groupby("SUBCONTRATO")[area].mean().dropna().reset_index()
           .rename(columns={area: "Nota"}).sort_values("Nota", ascending=False).head(5)
           .sort_values("Nota", ascending=True))
    top["Nota"] = top["Nota"].round(2)
    fig = px.bar(top, x="Nota", y="SUBCONTRATO", orientation="h", text="Nota",
                 range_x=[1, 7], color_discrete_sequence=[color])
    fig.update_traces(textposition="outside", marker_line_width=0)
    return plotly_layout(fig.update_layout(height=240, showlegend=False))

@st.cache_data
def cargar_datos(file_bytes):
    wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb["Consolidado"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[5]
    data = [r for r in rows[6:] if r[0] is not None and r[19] in ("APROBADO","MEJORAR","REPROBADO","NO RECOMENDADO","NO AUTORIZADO")]
    df = pd.DataFrame(data, columns=header)
    df = df.rename(columns={
        "N° EVA": "N_EVA", "CÓDIGO OBRA": "CODIGO_OBRA", "FECHA EVALUACIÓN": "FECHA",
        "NOMBRE SUBCONTRATO": "SUBCONTRATO", "NOTA FINAL (1 a 7)": "NOTA_FINAL",
        "ESTADO": "ESTADO", "TERRENO": "TERRENO", "RRHH": "RRHH", "SSOMA": "SSOMA", "CALIDAD": "CALIDAD",
    })
    num_cols = ["NOTA_FINAL","TERRENO","RRHH","SSOMA","CALIDAD"] + TODOS_CRITERIOS
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df["N_EVA"]       = pd.to_numeric(df["N_EVA"], errors="coerce")
    df["OBRA"]        = df["OBRA"].str.strip().str.title()
    df["SUBCONTRATO"] = df["SUBCONTRATO"].str.strip().str.title()
    df["ESTADO"]      = df["ESTADO"].str.strip().str.upper()
    if "ACTIVIDAD" in df.columns:
        df["ACTIVIDAD"] = df["ACTIVIDAD"].str.strip().str.title()
    return df

# ── Logo ───────────────────────────────────────────────────────────────────────
logo_path = next((p for p in [Path("logo.png"), Path("logo.jpg"), Path("logo.jpeg")] if p.exists()), None)
if logo_path:
    ext = "png" if logo_path.suffix == ".png" else "jpeg"
    with open(logo_path, "rb") as f:
        logo_b64 = base64.b64encode(f.read()).decode()
    logo_html = f'<img src="data:image/{ext};base64,{logo_b64}" style="height:70px;">'
else:
    logo_html = f'<span style="color:{NARANJO}; font-size:2rem;">🏗️</span>'

_header_placeholder = st.empty()

# ── Datos ──────────────────────────────────────────────────────────────────────
excel_path = Path("consolidado.xlsx")
if not excel_path.exists():
    st.error("No se encontró el archivo consolidado.xlsx")
    st.stop()

df = cargar_datos(excel_path.read_bytes())

# ── Fotos de obras ─────────────────────────────────────────────────────────────
FOTOS_OBRAS = {
    "Guillermo Feliú":  "foto_guillermo_feliu.jpeg",
    "José Venturelli":  "foto_jose_venturelli.jpg",
    "Pedro Luna Ii":    "foto_pedro_luna_ii.JPG",
    "Alfredo Helsby":   "foto_alfredo_helsby.JPG",
    "Armando Uribe":    "foto_armando_uribe.jpg",
    "Enrique Maturana": "foto_enrique_maturana.JPG",
    "Enriqueta Petit":  "foto_enriqueta_petit.JPG",
}

@st.cache_data
def obra_foto_b64(nombre_obra: str):
    archivo = FOTOS_OBRAS.get(nombre_obra)
    if archivo:
        p = Path(archivo)
        if p.exists():
            ext = p.suffix.lower().lstrip(".")
            if ext == "jpg": ext = "jpeg"
            with open(p, "rb") as f:
                return f"data:image/{ext};base64,{base64.b64encode(f.read()).decode()}"
    return None

# ── Sidebar ────────────────────────────────────────────────────────────────────
ORDEN_OBRAS = [
    "Delia Del Carril","Colina Ii","Armando Uribe","Pedro Luna I","Alfredo Helsby",
    "José Venturelli","Enrique Maturana","Enriqueta Petit","Pedro Luna Ii","Guillermo Feliú","Adriana Olguín",
]

if "obras_activas" not in st.session_state:
    st.session_state["obras_activas"] = set()

with st.sidebar:
    st.markdown("### Proyectos")
    obras_en_df = set(df["OBRA"].dropna().unique())
    obras_activas = st.session_state["obras_activas"]
    todas_activo  = len(obras_activas) == 0
    textos_activos = []
    if todas_activo: textos_activos.append("🏗️  Todos los proyectos")
    for obra in ORDEN_OBRAS:
        if obra in obras_activas:
            label = obra.replace(" Ii"," II") + ("" if obra in obras_en_df else " 🔜")
            textos_activos.append(label)
    textos_js = str(textos_activos).replace("'",'"')
    st.markdown(f"""
    <style>
      section[data-testid="stSidebar"] button {{ color:white!important; background-color:rgba(255,255,255,0.07)!important; border:3px solid transparent!important; border-radius:8px!important; font-weight:500!important; }}
      section[data-testid="stSidebar"] button:hover {{ background-color:rgba(255,255,255,0.14)!important; }}
      section[data-testid="stSidebar"] button p {{ color:white!important; }}
    </style>""", unsafe_allow_html=True)
    components.html(f"""<script>(function(){{var textos={textos_js};function marcar(){{var doc=window.parent.document;var sb=doc.querySelector('section[data-testid="stSidebar"]');if(!sb)return;sb.querySelectorAll('button').forEach(function(btn){{var p=btn.querySelector('p');if(!p)return;var txt=p.innerText.trim();if(textos.indexOf(txt)!==-1){{btn.style.borderColor='{NARANJO}';btn.style.backgroundColor='rgba(225,132,38,0.22)';btn.style.boxShadow='0 0 0 1px {NARANJO}'}}else{{btn.style.borderColor='transparent';btn.style.backgroundColor='rgba(255,255,255,0.07)';btn.style.boxShadow='none'}}}})}}setTimeout(marcar,100);setTimeout(marcar,400);setTimeout(marcar,900)}})();</script>""", height=0)
    if st.button("🏗️  Todos los proyectos", key="btn_todas", use_container_width=True):
        st.session_state["obras_activas"] = set(); st.rerun()
    st.markdown('<hr style="border-color:rgba(255,255,255,0.15);margin:6px 0 10px 0;">', unsafe_allow_html=True)
    for obra in ORDEN_OBRAS:
        es_futuro = obra not in obras_en_df
        label_obra = obra.replace(" Ii"," II")
        if st.button(f"{label_obra}{' 🔜' if es_futuro else ''}", key=f"btn_{obra}", use_container_width=True):
            nuevas = set(obras_activas)
            if obra in nuevas: nuevas.discard(obra)
            else: nuevas.add(obra)
            st.session_state["obras_activas"] = nuevas; st.rerun()
    obras_activas = st.session_state["obras_activas"]
    if len(obras_activas) == 0:
        obras_sel = [o for o in ORDEN_OBRAS if o in obras_en_df]
    else:
        obras_sel = [o for o in ORDEN_OBRAS if o in obras_activas and o in obras_en_df]

# ── Fondo dinámico ─────────────────────────────────────────────────────────────
obras_activas_count = len(st.session_state.get("obras_activas", set()))
_foto_fondo = obra_foto_b64(obras_sel[0]) if obras_activas_count == 1 and len(obras_sel) == 1 else None
if _foto_fondo:
    st.markdown(f"""<style>.stApp{{background-image:url("{_foto_fondo}");background-size:cover;background-attachment:fixed;background-position:center;}}.stApp::before{{content:"";position:fixed;inset:0;background:rgba(255,255,255,0.90);z-index:0;pointer-events:none;}}.stApp>*{{position:relative;z-index:1;}}</style>""", unsafe_allow_html=True)
else:
    st.markdown("""<style>.stApp{background-image:none!important;background-color:#ffffff!important;}.stApp::before{display:none;}</style>""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
if obras_activas_count == 1 and len(obras_sel) == 1:
    _sub = f'<div style="font-size:1.6rem;font-weight:800;color:{NARANJO};line-height:1.2;margin-top:4px;letter-spacing:0.02em;">{obras_sel[0].replace(" Ii"," II").upper()}</div>'
elif obras_activas_count >= 2:
    _sub = f'<div style="font-size:1.3rem;font-weight:700;color:{NARANJO};line-height:1.2;margin-top:4px;">VARIAS OBRAS</div>'
else:
    _sub = f'<div style="font-size:1.3rem;font-weight:700;color:{NARANJO};line-height:1.2;margin-top:4px;">TODOS LOS PROYECTOS</div>'
_header_placeholder.markdown(f"""<div style="background-color:white;padding:24px 32px;border-radius:12px;display:flex;align-items:center;gap:28px;border-bottom:4px solid {NARANJO};margin-bottom:28px;box-shadow:0 2px 12px rgba(26,24,70,0.08);">{logo_html}<div style="border-left:2px solid {GRIS};padding-left:24px;"><div style="font-size:2rem;font-weight:800;color:{AZUL};line-height:1.1;">Evaluación de Subcontratos</div>{_sub}</div></div>""", unsafe_allow_html=True)

# ── Filtro datos ───────────────────────────────────────────────────────────────
if not obras_sel:
    st.info("🔜 Este proyecto aún no tiene evaluaciones registradas.")
    st.stop()
df_f = df[df["OBRA"].isin(obras_sel) & df["ESTADO"].isin(["APROBADO","MEJORAR","REPROBADO"])]
if df_f.empty:
    st.info("🔜 Este proyecto aún no tiene evaluaciones registradas.")
    st.stop()

# ── KPIs ───────────────────────────────────────────────────────────────────────
total = len(df_f)
aprobados  = (df_f["ESTADO"] == "APROBADO").sum()
mejorar    = (df_f["ESTADO"] == "MEJORAR").sum()
reprobados = (df_f["ESTADO"] == "REPROBADO").sum()
nota_prom  = df_f["NOTA_FINAL"].mean()

k1,k2,k3,k4,k5 = st.columns(5)
for col, label, val, sub in [
    (k1,"Total evaluaciones",total,""),
    (k2,"✅ Aprobado",aprobados,f"{aprobados/total*100:.0f}%"),
    (k3,"⚠️ Mejorar",mejorar,f"{mejorar/total*100:.0f}%"),
    (k4,"❌ Reprobado",reprobados,f"{reprobados/total*100:.0f}%"),
    (k5,"Nota promedio",f"{nota_prom:.2f}","sobre 7.0"),
]:
    col.markdown(f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{val}</div><div class="kpi-sub">{sub}</div></div>', unsafe_allow_html=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 1 — VERIFICADOR AMPLIADO
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f'''
<div style="font-size:1.6rem;font-weight:800;color:{AZUL};margin-bottom:4px;">🔍 Verificador de Subcontratos</div>
<div class="seccion-sub">Busca un subcontrato — estado, historial y desglose completo por criterio</div>
''', unsafe_allow_html=True)

df_todos_ver = cargar_datos(excel_path.read_bytes())
nombres_lista = sorted(df_todos_ver["SUBCONTRATO"].dropna().unique().tolist())
ruts_lista    = sorted(df_todos_ver["RUT"].astype(str).dropna().unique().tolist()) if "RUT" in df_todos_ver.columns else []

col_b1, col_b2, col_b3 = st.columns([5,5,1])
if "ver_nombre" not in st.session_state: st.session_state["ver_nombre"] = ""
if "ver_rut"    not in st.session_state: st.session_state["ver_rut"]    = ""

with col_b3:
    st.markdown("<div style='padding-top:28px'>", unsafe_allow_html=True)
    if st.button("✕", help="Limpiar búsqueda"):
        st.session_state["ver_nombre"] = ""; st.session_state["ver_rut"] = ""; st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)
with col_b1:
    buscar_nombre = st.selectbox("Buscar por nombre", options=[""]+nombres_lista,
        index=0 if st.session_state["ver_nombre"]=="" else ([""]+nombres_lista).index(st.session_state["ver_nombre"]),
        format_func=lambda x: "Escribe o selecciona un nombre..." if x=="" else x, key="ver_nombre")
with col_b2:
    buscar_rut = st.selectbox("Buscar por RUT", options=[""]+ruts_lista,
        index=0 if st.session_state["ver_rut"]=="" else ([""]+ruts_lista).index(st.session_state["ver_rut"]),
        format_func=lambda x: "Escribe o selecciona un RUT..." if x=="" else x, key="ver_rut")

if buscar_nombre or buscar_rut:
    df_busq = df_todos_ver.copy()
    if buscar_nombre: df_busq = df_busq[df_busq["SUBCONTRATO"] == buscar_nombre]
    if buscar_rut:    df_busq = df_busq[df_busq["RUT"].astype(str) == buscar_rut]

    if df_busq.empty:
        st.info("No se encontraron resultados.")
    else:
        # Tarjeta de estado
        for _, row in df_busq.drop_duplicates(subset=["SUBCONTRATO"]).iterrows():
            estado = row["ESTADO"]
            cfg = {
                "APROBADO":      ("#d4edda","#155724","✅"),
                "MEJORAR":       ("#fff3cd","#856404","⚠️"),
                "REPROBADO":     ("#f8d7da","#721c24","❌"),
                "NO RECOMENDADO":("#fde8e8","#8B0000","🚫"),
            }.get(estado, ("#2a2a2a","#ffffff","⛔"))
            color_bg, color_txt, icono = cfg
            st.markdown(f"""<div style="background:{color_bg};border-left:4px solid {color_txt};color:{color_txt};padding:14px 20px;border-radius:8px;margin-bottom:8px;font-family:'Hanken Grotesk',sans-serif;">
                <div style="font-size:1.05rem;font-weight:700;">{icono} {row["SUBCONTRATO"]}</div>
                <div style="font-size:0.82rem;margin-top:4px;">RUT: {row["RUT"]} &nbsp;|&nbsp; Estado: <strong>{estado}</strong> &nbsp;|&nbsp; Actividad: {row.get("ACTIVIDAD","—")} &nbsp;|&nbsp; Obra: {row["OBRA"]}</div>
            </div>""", unsafe_allow_html=True)

        nombre_sel = df_busq["SUBCONTRATO"].iloc[0]

        # Evolución nota final
        df_hist = df_todos_ver[df_todos_ver["SUBCONTRATO"] == nombre_sel][["N_EVA","NOTA_FINAL","OBRA"]].dropna(subset=["N_EVA","NOTA_FINAL"]).sort_values("N_EVA")
        df_hist["N_EVA"] = df_hist["N_EVA"].astype(int)
        if not df_hist.empty:
            st.markdown(f'<div class="seccion-titulo" style="margin-top:16px;">📈 Evolución de nota final — {nombre_sel}</div>', unsafe_allow_html=True)
            fig_hist = px.line(df_hist, x="N_EVA", y="NOTA_FINAL", markers=True, text="NOTA_FINAL",
                               hover_data=["OBRA"], range_y=[1,7], color_discrete_sequence=[AZUL],
                               labels={"N_EVA":"N° Evaluación","NOTA_FINAL":"Nota"})
            fig_hist.update_traces(textposition="top center", texttemplate="%{text:.2f}")
            fig_hist.add_hline(y=5.5, line_dash="dash", line_color="#2ECC71", annotation_text="Aprobado ≥ 5.5")
            fig_hist.add_hline(y=4.0, line_dash="dash", line_color=NARANJO,   annotation_text="Mejorar ≥ 4.0")
            fig_hist.update_xaxes(tickmode="linear", dtick=1)
            fig_hist.update_layout(height=280)
            st.plotly_chart(plotly_layout(fig_hist), use_container_width=True)

        # ── NUEVO: Desglose por criterio (última evaluación disponible) ──
        df_sc = df_todos_ver[df_todos_ver["SUBCONTRATO"] == nombre_sel].sort_values("N_EVA")
        ultima_eva = df_sc["N_EVA"].max()
        df_ultima = df_sc[df_sc["N_EVA"] == ultima_eva]

        criterios_disponibles = [c for c in TODOS_CRITERIOS if c in df_sc.columns and df_sc[c].notna().any()]
        if criterios_disponibles:
            st.markdown(f'<div class="seccion-titulo" style="margin-top:20px;">📊 Desglose por criterio — Evaluación N°{int(ultima_eva)}</div>', unsafe_allow_html=True)
            st.markdown('<div class="seccion-sub">Nota promedio por criterio individual en la última evaluación disponible</div>', unsafe_allow_html=True)

            tab_t, tab_r, tab_s, tab_c = st.tabs(["🔵 Terreno","🟣 RRHH","🟠 SSOMA","🟡 Calidad"])
            for tab, area in [(tab_t,"TERRENO"),(tab_r,"RRHH"),(tab_s,"SSOMA"),(tab_c,"CALIDAD")]:
                with tab:
                    crits = [c for c in CRITERIOS[area] if c in df_ultima.columns]
                    vals  = [pd.to_numeric(df_ultima[c], errors="coerce").mean() for c in crits]
                    labels = [ETIQUETAS_CORTAS.get(c, c[:30]) for c in crits]
                    df_crit = pd.DataFrame({"Criterio": labels, "Nota": vals}).dropna()
                    if not df_crit.empty:
                        df_crit["Nota"] = df_crit["Nota"].round(2)
                        df_crit["Color"] = df_crit["Nota"].apply(
                            lambda v: "#2ECC71" if v >= 5.5 else (NARANJO if v >= 4.0 else TERRACOTA))
                        fig_crit = go.Figure(go.Bar(
                            x=df_crit["Nota"], y=df_crit["Criterio"],
                            orientation="h", text=df_crit["Nota"],
                            marker_color=df_crit["Color"],
                            textposition="outside",
                        ))
                        fig_crit.update_layout(height=max(200, len(df_crit)*38), xaxis_range=[0,7],
                                               xaxis_title="Nota (1-7)", showlegend=False)
                        st.plotly_chart(plotly_layout(fig_crit), use_container_width=True)
                    else:
                        st.info("Sin datos de criterio para esta área.")

            # ── Radar todos los criterios ──────────────────────────────────
            st.markdown(f'<div class="seccion-titulo" style="margin-top:20px;">🕸️ Radar completo — {nombre_sel}</div>', unsafe_allow_html=True)
            radar_labels, radar_vals = [], []
            for c in criterios_disponibles:
                v = pd.to_numeric(df_ultima[c], errors="coerce").mean()
                if not pd.isna(v):
                    radar_labels.append(ETIQUETAS_CORTAS.get(c, c[:20]))
                    radar_vals.append(round(v, 2))
            if radar_labels:
                fig_radar = go.Figure(go.Scatterpolar(
                    r=radar_vals + [radar_vals[0]],
                    theta=radar_labels + [radar_labels[0]],
                    fill="toself", line_color=AZUL,
                    fillcolor=f"rgba(26,24,70,0.15)",
                ))
                fig_radar.update_layout(
                    polar=dict(radialaxis=dict(visible=True, range=[0,7], tickfont=dict(size=10))),
                    height=420, showlegend=False,
                )
                st.plotly_chart(plotly_layout(fig_radar), use_container_width=True)

            # ── Evolución por criterio específico ─────────────────────────
            st.markdown(f'<div class="seccion-titulo" style="margin-top:20px;">📉 Evolución por criterio — {nombre_sel}</div>', unsafe_allow_html=True)
            crit_sel = st.selectbox("Selecciona un criterio para ver su evolución", 
                                    options=criterios_disponibles,
                                    format_func=lambda c: ETIQUETAS_CORTAS.get(c, c[:50]),
                                    key="crit_evol_ver")
            df_evol = df_sc[["N_EVA", crit_sel]].copy()
            df_evol[crit_sel] = pd.to_numeric(df_evol[crit_sel], errors="coerce")
            df_evol = df_evol.dropna().sort_values("N_EVA")
            df_evol["N_EVA"] = df_evol["N_EVA"].astype(int)
            if not df_evol.empty:
                fig_evol = px.line(df_evol, x="N_EVA", y=crit_sel, markers=True, text=crit_sel,
                                   range_y=[0,7], color_discrete_sequence=[NARANJO],
                                   labels={"N_EVA":"N° Evaluación", crit_sel:"Nota"})
                fig_evol.update_traces(textposition="top center", texttemplate="%{text:.1f}")
                fig_evol.add_hline(y=5.5, line_dash="dash", line_color="#2ECC71", annotation_text="Aprobado")
                fig_evol.add_hline(y=4.0, line_dash="dash", line_color=TERRACOTA,  annotation_text="Límite")
                fig_evol.update_xaxes(tickmode="linear", dtick=1)
                fig_evol.update_layout(height=260)
                st.plotly_chart(plotly_layout(fig_evol), use_container_width=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 2 — DISTRIBUCIÓN / EVOLUCIÓN
# ══════════════════════════════════════════════════════════════════════════════
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
    evo.columns = ["Evaluación","Nota promedio"]
    fig2 = px.line(evo, x="Evaluación", y="Nota promedio", markers=True,
                   range_y=[1,7], color_discrete_sequence=[NARANJO])
    fig2.add_hline(y=5.5, line_dash="dash", line_color="#2ECC71", annotation_text="Aprobado ≥ 5.5")
    fig2.add_hline(y=4.0, line_dash="dash", line_color=NARANJO,   annotation_text="Mejorar ≥ 4.0")
    st.plotly_chart(plotly_layout(fig2), use_container_width=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 3 — RADAR POR CRITERIO (vista global)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f'<div class="seccion-titulo">🕸️ Radar global por criterio</div>', unsafe_allow_html=True)
st.markdown('<div class="seccion-sub">Nota promedio de todos los subcontratos en cada criterio individual — identifica las áreas más débiles del conjunto</div>', unsafe_allow_html=True)

crits_disponibles_global = [c for c in TODOS_CRITERIOS if c in df_f.columns and df_f[c].notna().any()]
if crits_disponibles_global:
    radar_g_labels = [ETIQUETAS_CORTAS.get(c, c[:20]) for c in crits_disponibles_global]
    radar_g_vals   = [round(df_f[c].mean(), 2) for c in crits_disponibles_global]
    fig_radar_g = go.Figure(go.Scatterpolar(
        r=radar_g_vals + [radar_g_vals[0]],
        theta=radar_g_labels + [radar_g_labels[0]],
        fill="toself", line_color=NARANJO, fillcolor="rgba(225,132,38,0.12)",
    ))
    fig_radar_g.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0,7], tickfont=dict(size=10))),
        height=450, showlegend=False,
    )
    st.plotly_chart(plotly_layout(fig_radar_g), use_container_width=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 4 — MAPA DE CALOR
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="seccion-titulo">🌡️ Mapa de calor — subcontratos × criterios</div>', unsafe_allow_html=True)
st.markdown('<div class="seccion-sub">Verde = alto desempeño · Rojo = bajo desempeño · Gris = sin dato</div>', unsafe_allow_html=True)

col_hm1, col_hm2 = st.columns([2,1])
with col_hm1:
    area_hm = st.selectbox("Área", ["TERRENO","RRHH","SSOMA","CALIDAD"], key="hm_area")
with col_hm2:
    min_evals = st.slider("Mín. evaluaciones", 1, 4, 1, key="hm_min")

crits_hm = [c for c in CRITERIOS[area_hm] if c in df_f.columns]
if crits_hm:
    df_hm = df_f.groupby("SUBCONTRATO")[crits_hm].mean()
    df_hm = df_hm[df_f.groupby("SUBCONTRATO")["NOTA_FINAL"].count() >= min_evals]
    df_hm = df_hm.dropna(how="all").sort_index()
    df_hm.columns = [ETIQUETAS_CORTAS.get(c, c[:25]) for c in df_hm.columns]

    if not df_hm.empty:
        fig_hm = go.Figure(go.Heatmap(
            z=df_hm.values.round(1),
            x=df_hm.columns.tolist(),
            y=[s[:35] for s in df_hm.index.tolist()],
            colorscale=[[0,"#B04D2F"],[0.5,NARANJO],[1,"#2ECC71"]],
            zmin=1, zmax=7,
            text=df_hm.values.round(1),
            texttemplate="%{text}",
            textfont=dict(size=10),
            hoverongaps=False,
        ))
        fig_hm.update_layout(height=max(300, len(df_hm)*28), xaxis_tickangle=-30,
                              margin=dict(l=200, r=20, t=20, b=80))
        st.plotly_chart(plotly_layout(fig_hm), use_container_width=True)
    else:
        st.info("Sin datos suficientes para el filtro seleccionado.")

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 5 — ALERTAS POR CRITERIO CRÍTICO
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f'<div style="font-size:1.3rem;font-weight:800;color:{TERRACOTA};margin-bottom:4px;">⚠️ Alertas por criterio crítico</div>', unsafe_allow_html=True)
st.markdown('<div class="seccion-sub">Subcontratos con nota ≥ 4.0 en general pero con algún criterio individual bajo 4 — riesgo oculto en el promedio</div>', unsafe_allow_html=True)

umbral_alerta = st.slider("Umbral de alerta (nota criterio ≤)", 1.0, 5.0, 4.0, 0.5, key="umbral_alerta")

crits_alerta = [c for c in TODOS_CRITERIOS if c in df_f.columns]
alertas = []
for _, row in df_f[df_f["NOTA_FINAL"] >= 4.0].iterrows():
    for c in crits_alerta:
        val = pd.to_numeric(row.get(c), errors="coerce")
        if pd.notna(val) and val <= umbral_alerta:
            alertas.append({
                "Subcontrato": row["SUBCONTRATO"],
                "Obra": row["OBRA"],
                "N° Eva": int(row["N_EVA"]) if pd.notna(row["N_EVA"]) else "—",
                "Nota final": round(row["NOTA_FINAL"], 2),
                "Criterio": ETIQUETAS_CORTAS.get(c, c[:40]),
                "Área": next((a for a, cs in CRITERIOS.items() if c in cs), "—"),
                "Nota criterio": round(val, 1),
            })

if alertas:
    df_alertas = pd.DataFrame(alertas).sort_values("Nota criterio")
    st.markdown(f'<div style="font-size:0.9rem;color:{TERRACOTA};font-weight:600;margin-bottom:12px;">⚡ {len(df_alertas)} alertas encontradas</div>', unsafe_allow_html=True)

    # Agrupar por subcontrato para mostrar en tarjetas
    for sc, grupo in df_alertas.groupby("Subcontrato"):
        with st.expander(f"⚠️ {sc} — {len(grupo)} criterio(s) bajo umbral"):
            for _, a in grupo.iterrows():
                color_nota = "#2ECC71" if a["Nota criterio"] >= 5.5 else (NARANJO if a["Nota criterio"] >= 4.0 else TERRACOTA)
                st.markdown(f"""<div class="alerta-card">
                    <span style="font-weight:700;color:{AZUL};">{a['Área']} · {a['Criterio']}</span>
                    &nbsp;|&nbsp; Nota: <span style="font-weight:700;color:{color_nota};">{a['Nota criterio']}</span>
                    &nbsp;|&nbsp; Eva N°{a['N° Eva']} &nbsp;|&nbsp; Obra: {a['Obra']}
                    &nbsp;|&nbsp; Nota final: {a['Nota final']}
                </div>""", unsafe_allow_html=True)
else:
    st.success(f"✅ Sin criterios bajo {umbral_alerta} en subcontratos aprobados.")

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 6 — RANKING POR CRITERIO GRANULAR
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="seccion-titulo">🔬 Ranking por criterio individual</div>', unsafe_allow_html=True)
st.markdown('<div class="seccion-sub">¿Quién cumple mejor los plazos? ¿Quién tiene el peor EPP? Selecciona cualquier criterio</div>', unsafe_allow_html=True)

crits_rank = [c for c in TODOS_CRITERIOS if c in df_f.columns and df_f[c].notna().any()]
col_r1, col_r2 = st.columns([3,1])
with col_r1:
    crit_rank_sel = st.selectbox("Criterio", crits_rank,
        format_func=lambda c: f"[{next((a for a,cs in CRITERIOS.items() if c in cs),'?')}] {ETIQUETAS_CORTAS.get(c, c[:50])}",
        key="crit_rank")
with col_r2:
    n_rank = st.slider("Mostrar top/bottom", 5, 15, 10, key="n_rank")

if crit_rank_sel:
    df_rank_c = (df_f.groupby("SUBCONTRATO")[crit_rank_sel].mean().dropna()
                 .reset_index().rename(columns={crit_rank_sel:"Nota"}).sort_values("Nota", ascending=False))
    df_rank_c["Nota"] = df_rank_c["Nota"].round(2)

    col_top, col_bot = st.columns(2)
    with col_top:
        st.markdown(f"**🏆 Top {n_rank} — mejor desempeño**")
        top_df = df_rank_c.head(n_rank).sort_values("Nota", ascending=True)
        fig_top = px.bar(top_df, x="Nota", y="SUBCONTRATO", orientation="h", text="Nota",
                         range_x=[0,7], color_discrete_sequence=["#2ECC71"])
        fig_top.update_traces(textposition="outside")
        fig_top.update_layout(height=max(200, n_rank*32), showlegend=False)
        st.plotly_chart(plotly_layout(fig_top), use_container_width=True)
    with col_bot:
        st.markdown(f"**⚠️ Bottom {n_rank} — menor desempeño**")
        bot_df = df_rank_c.tail(n_rank).sort_values("Nota", ascending=False)
        fig_bot = px.bar(bot_df, x="Nota", y="SUBCONTRATO", orientation="h", text="Nota",
                         range_x=[0,7], color_discrete_sequence=[TERRACOTA])
        fig_bot.update_traces(textposition="outside")
        fig_bot.update_layout(height=max(200, n_rank*32), showlegend=False)
        st.plotly_chart(plotly_layout(fig_bot), use_container_width=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 7 — NOTAS POR ÁREA / REPROBADOS
# ══════════════════════════════════════════════════════════════════════════════
c3, c4 = st.columns(2)
with c3:
    st.markdown('<div class="seccion-titulo">Nota promedio por área evaluada</div>', unsafe_allow_html=True)
    areas = ["TERRENO","RRHH","SSOMA","CALIDAD"]
    notas_area = df_f[areas].mean().reset_index()
    notas_area.columns = ["Área","Nota"]
    fig3 = px.bar(notas_area, x="Área", y="Nota", color="Área",
                  color_discrete_map=COLOR_AREA, text=notas_area["Nota"].round(2), range_y=[1,7])
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

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 8 — TOP 5 POR ÁREA + TOP 5 POR ACTIVIDAD
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="seccion-titulo">🏆 Top 5 subcontratos por área evaluada</div>', unsafe_allow_html=True)
a1,a2,a3,a4 = st.columns(4)
for col, area, label in [(a1,"TERRENO","🔵 Terreno"),(a2,"RRHH","🟣 RRHH"),(a3,"SSOMA","🟠 SSOMA"),(a4,"CALIDAD","🟡 Calidad")]:
    with col:
        st.markdown(f"**{label}**")
        st.plotly_chart(top5_bar(df_f, area, COLOR_AREA[area]), use_container_width=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

st.markdown('<div class="seccion-titulo">🔧 Top 5 subcontratos por actividad</div>', unsafe_allow_html=True)
if "ACTIVIDAD" in df_f.columns:
    actividades = sorted(df_f["ACTIVIDAD"].dropna().unique())
    act_sel = st.selectbox("Selecciona una actividad", actividades)
    df_act = df_f[df_f["ACTIVIDAD"] == act_sel]
    top_act = (df_act.groupby("SUBCONTRATO")["NOTA_FINAL"].mean().dropna()
               .reset_index().rename(columns={"NOTA_FINAL":"Nota promedio"})
               .sort_values("Nota promedio", ascending=False).head(5).sort_values("Nota promedio", ascending=True))
    top_act["Nota promedio"] = top_act["Nota promedio"].round(2)
    fig_act = px.bar(top_act, x="Nota promedio", y="SUBCONTRATO", orientation="h",
                     text="Nota promedio", range_x=[1,7], color_discrete_sequence=[AZUL])
    fig_act.update_traces(textposition="outside")
    fig_act.update_layout(height=300)
    st.plotly_chart(plotly_layout(fig_act), use_container_width=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 9 — RANKING GENERAL
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="seccion-titulo">Ranking general de subcontratos</div>', unsafe_allow_html=True)
ranking = (df_f.groupby("SUBCONTRATO")
           .agg(Evaluaciones=("NOTA_FINAL","count"), Nota_promedio=("NOTA_FINAL","mean"), Ultimo_estado=("ESTADO","last"))
           .reset_index().sort_values("Nota_promedio", ascending=False))
ranking["Nota_promedio"] = ranking["Nota_promedio"].round(2)

def color_estado(val):
    _mapa = {"APROBADO":"#d4edda","MEJORAR":"#fff3cd","REPROBADO":"#f8d7da"}
    return f"background-color: {_mapa.get(val,'')}"
try:    styled = ranking.style.map(color_estado, subset=["Ultimo_estado"])
except: styled = ranking.style.applymap(color_estado, subset=["Ultimo_estado"])
st.dataframe(styled, use_container_width=True, height=400)

with st.expander("Ver detalle completo"):
    cols_show = [c for c in ["N_EVA","CODIGO_OBRA","OBRA","SUBCONTRATO","ACTIVIDAD","TERRENO","RRHH","SSOMA","CALIDAD","NOTA_FINAL","ESTADO"] if c in df_f.columns]
    st.dataframe(df_f[cols_show].sort_values(["OBRA","N_EVA"]), use_container_width=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 10 — NO RECOMENDADOS / NO AUTORIZADOS
# ══════════════════════════════════════════════════════════════════════════════
df_todos2 = cargar_datos(excel_path.read_bytes())
df_no_rec = df_todos2[df_todos2["ESTADO"]=="NO RECOMENDADO"].drop_duplicates(subset=["SUBCONTRATO"])
df_no_aut = df_todos2[df_todos2["ESTADO"]=="NO AUTORIZADO"].drop_duplicates(subset=["SUBCONTRATO"])

st.markdown(f'<div style="font-size:1.3rem;font-weight:800;color:#8B0000;margin-bottom:4px;">🚫 NO RECOMENDADOS ({len(df_no_rec)})</div>', unsafe_allow_html=True)
st.markdown('<div class="seccion-sub">Subcontratos con desempeño deficiente — usar solo con autorización expresa</div>', unsafe_allow_html=True)
if df_no_rec.empty:
    st.success("No hay subcontratos en esta categoría.")
else:
    filas_rec = df_no_rec.sort_values("SUBCONTRATO").reset_index(drop=True)
    for i in range(0, len(filas_rec), 2):
        cols = st.columns(2)
        for j, col in enumerate(cols):
            if i+j < len(filas_rec):
                row = filas_rec.iloc[i+j]
                col.markdown(f"""<div style="background:#fde8e8;border:1px solid #c0392b;border-radius:8px;padding:14px 16px;font-family:'Hanken Grotesk',sans-serif;height:100%;">
                    <div style="font-size:1rem;font-weight:700;color:#8B0000;">🚫 {row["SUBCONTRATO"]}</div>
                    <div style="font-size:0.8rem;color:#5a0000;margin-top:6px;"><b>RUT:</b> {row["RUT"]}<br><b>Actividad:</b> {row.get("ACTIVIDAD","—")}<br><b>Obra:</b> {row["OBRA"]}</div>
                </div>""", unsafe_allow_html=True)

st.markdown('<hr class="divider-naranja">', unsafe_allow_html=True)

st.markdown(f'<div style="font-size:1.3rem;font-weight:800;color:{AZUL};margin-bottom:4px;">⛔ NO AUTORIZADOS ({len(df_no_aut)})</div>', unsafe_allow_html=True)
st.markdown('<div class="seccion-sub">Subcontratos vetados — NO contratar bajo ninguna circunstancia</div>', unsafe_allow_html=True)
if df_no_aut.empty:
    st.success("No hay subcontratos en esta categoría.")
else:
    filas_aut = df_no_aut.sort_values("SUBCONTRATO").reset_index(drop=True)
    for i in range(0, len(filas_aut), 2):
        cols = st.columns(2)
        for j, col in enumerate(cols):
            if i+j < len(filas_aut):
                row = filas_aut.iloc[i+j]
                col.markdown(f"""<div style="background:#2a2a2a;border:1px solid #555;border-radius:8px;padding:14px 16px;font-family:'Hanken Grotesk',sans-serif;height:100%;">
                    <div style="font-size:1rem;font-weight:700;color:#ffffff;">⛔ {row["SUBCONTRATO"]}</div>
                    <div style="font-size:0.8rem;color:#bbbbbb;margin-top:6px;"><b>RUT:</b> {row["RUT"]}<br><b>Actividad:</b> {row.get("ACTIVIDAD","—")}<br><b>Obra:</b> {row["OBRA"]}</div>
                </div>""", unsafe_allow_html=True)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown(f"""<div style="text-align:center;padding:20px;color:{AZUL_MED};font-size:0.78rem;margin-top:30px;">
  Constructora Londres · Sistema de Evaluación de Subcontratos · {pd.Timestamp.now().year}
</div>""", unsafe_allow_html=True)
