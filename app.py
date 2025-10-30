# app.py — Dashboard Matriz de Obligaciones (MADR)
# Avance global: promedio simple. Incluye KPIs, tops, descargas y gráficos avanzados.

import io, re, unicodedata
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import re, unicodedata

# ---------- Config ----------
st.set_page_config(page_title="Matriz de Obligaciones • MADR", layout="wide")
LOGO_PATH = "assets/logo_agricultura.png"   # opcional

if LOGO_PATH:
    try: st.sidebar.image(LOGO_PATH, use_column_width=True)
    except: pass

st.title("✅ Dashboard — Matriz de Obligaciones Contractuales (MADR)")
st.caption("Carga un Excel (.xlsx) con obligaciones, columnas de '% avance' y 'Observaciones' por mes.")

# ---------- Constantes / helpers ----------
MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
         "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

CATEGORIAS_CANON = [
    "OBLIGACIONES EXCLUSIVAS DEL ENCARGO FIDUCIARIO, NUEVA FASE",
    "OBLIGACIONES A CARGO DEL ENCARGO FIDUCIARIO RESPECTO DE LOS PATRIMONIOS AUTÓNOMOS NUEVA FASE 2025",
    "OBLIGACIONES GENERALES DEL CONTRATISTA",
    "OBLIGACIONES DEL MINISTERIO",
]

# ========== AQUÍ VA EL TEMPLATE ==========
# ---------- Paleta de colores institucional ----------
# ---------- Paleta de colores institucional (mantener igual) ----------
COLOR_PALETTE = {
    'primary': '#2E7D32',      # Verde institucional MADR
    'secondary': '#1565C0',    # Azul
    'success': '#43A047',      # Verde éxito
    'warning': '#FB8C00',      # Naranja advertencia
    'danger': '#E53935',       # Rojo peligro
    'neutral': '#546E7A',      # Gris neutro
    'bg_light': '#F5F5F5',     # Fondo claro
}

# Paleta para categorías (gradiente verde institucional)
CATEGORIA_COLORS = [
    '#2E7D32',  # Verde principal
    '#43A047',  # Verde claro
    '#66BB6A',  # Verde más claro
    '#81C784',  # Verde pastel
    '#1565C0',  # Azul complementario
    '#42A5F5',  # Azul claro
    '#FB8C00',  # Naranja
    '#7B1FA2'   # Púrpura
]

# Colores para semáforo
SEMAFORO_COLORS = {
    'Alto': '#43A047',
    'Medio': '#FFB300',
    'Bajo': '#E53935',
    'Sin dato': '#BDBDBD'
}

# ---------- Template mejorado con bordes redondeados ----------
def get_plotly_template():
    """Template personalizado con estilo institucional mejorado y bordes redondeados"""
    return go.layout.Template(
        layout=go.Layout(
            # Tipografía
            font=dict(
                family="Segoe UI, Roboto, Arial, sans-serif",
                size=13,
                color="#1A252F"
            ),
            
            # Título estilizado
            title=dict(
                font=dict(
                    size=18, 
                    color="#1A252F", 
                    family="Segoe UI Semibold"
                ),
                x=0.5,
                xanchor='center',
                y=0.98,
                yanchor='top',
                pad=dict(t=20, b=20)
            ),
            
            # Fondo con gradiente sutil
            plot_bgcolor='#FAFBFC',
            paper_bgcolor='#FFFFFF',
            
            # Paleta de colores
            colorway=CATEGORIA_COLORS,
            
            # Interactividad mejorada
            hovermode='closest',
            hoverlabel=dict(
                bgcolor="white",
                font_size=13,
                font_family="Segoe UI",
                bordercolor="#E0E0E0",
                align="left"
            ),
            
            # Márgenes generosos para bordes redondeados
            margin=dict(l=70, r=70, t=100, b=70),
            
            # Ejes X mejorados
            xaxis=dict(
                showgrid=True,
                gridcolor='#E8EAF6',
                gridwidth=1,
                showline=True,
                linewidth=2,
                linecolor='#BDBDBD',
                mirror=False,
                ticks='outside',
                tickcolor='#BDBDBD',
                tickfont=dict(size=12, color='#546E7A'),
                title=dict(
                    font=dict(size=14, color='#37474F', family="Segoe UI Semibold"),
                    standoff=15
                )
            ),
            
            # Ejes Y mejorados
            yaxis=dict(
                showgrid=True,
                gridcolor='#E8EAF6',
                gridwidth=1,
                showline=True,
                linewidth=2,
                linecolor='#BDBDBD',
                mirror=False,
                ticks='outside',
                tickcolor='#BDBDBD',
                tickfont=dict(size=12, color='#546E7A'),
                title=dict(
                    font=dict(size=14, color='#37474F', family="Segoe UI Semibold"),
                    standoff=15
                ),
                zeroline=True,
                zerolinecolor='#CFD8DC',
                zerolinewidth=2
            ),
            
            # Leyenda estilizada
            legend=dict(
                bgcolor='rgba(255, 255, 255, 0.95)',
                bordercolor='#E0E0E0',
                borderwidth=1,
                font=dict(size=12, color='#37474F'),
                orientation='v',
                yanchor='top',
                y=0.99,
                xanchor='right',
                x=0.99
            ),
            
            # Sombras y efectos
            shapes=[],
            annotations=[]
        )
    )

# Crear el template una sola vez
PLOTLY_TEMPLATE = get_plotly_template()

# ---------- Función auxiliar para estilizar gráficas individuales ----------
def aplicar_estilo_grafica(fig, altura=450, mostrar_leyenda=True):
    """
    Aplica estilo consistente con bordes redondeados a cualquier gráfica Plotly.
    
    Args:
        fig: Figura de Plotly
        altura: Altura en píxeles
        mostrar_leyenda: Si mostrar o no la leyenda
    """
    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        height=altura,
        showlegend=mostrar_leyenda,
        
        # Bordes redondeados (simulados con shapes)
        shapes=[
            # Marco exterior con esquinas redondeadas
            dict(
                type="rect",
                xref="paper",
                yref="paper",
                x0=0,
                y0=0,
                x1=1,
                y1=1,
                line=dict(
                    color="#E0E0E0",
                    width=2
                ),
                fillcolor="rgba(0,0,0,0)"
            )
        ],
        
        # Animaciones suaves
        transition=dict(
            duration=300,
            easing='cubic-in-out'
        )
    )
    
    # Efecto de sombra en barras
    if fig.data:
        for trace in fig.data:
            if hasattr(trace, 'marker') and trace.marker:
                if isinstance(trace.marker, dict):
                    trace.marker['line'] = dict(
                        color='white',
                        width=2
                    )
                else:
                    trace.marker.line = dict(
                        color='white',
                        width=2
                    )
    
    return fig

# ---------- Estilos específicos para tipos de gráficas ----------

def estilo_linea_temporal(fig, altura=500):
    """Estilo específico para gráficas de línea temporal"""
    fig = aplicar_estilo_grafica(fig, altura, mostrar_leyenda=False)
    
    # Área de relleno bajo la línea con gradiente
    if fig.data and len(fig.data) > 0:
        fig.update_traces(
            fill='tozeroy',
            fillcolor='rgba(46, 125, 50, 0.1)',
            line=dict(
                width=4,
                shape='spline',  # Líneas suaves
                smoothing=1.3
            ),
            marker=dict(
                size=12,
                line=dict(width=3, color='white'),
                symbol='circle'
            )
        )
    
    return fig

def estilo_barras_horizontal(fig, altura=None, n_items=None):
    """Estilo específico para gráficas de barras horizontales"""
    if altura is None and n_items:
        altura = max(400, n_items * 60)
    
    fig = aplicar_estilo_grafica(fig, altura or 500, mostrar_leyenda=False)
    
    # Bordes redondeados en barras
    fig.update_traces(
        marker=dict(
            line=dict(color='white', width=2),
            cornerradius=8  # Esto funciona en versiones recientes de Plotly
        ),
        textfont=dict(
            size=13,
            family="Segoe UI Semibold",
            color='#37474F'
        )
    )
    
    return fig

def estilo_heatmap(fig, altura=None, n_categorias=None):
    """Estilo específico para heatmaps"""
    if altura is None and n_categorias:
        altura = max(400, n_categorias * 50)
    
    fig = aplicar_estilo_grafica(fig, altura or 500, mostrar_leyenda=False)
    
    # Colorbar estilizada
    fig.update_traces(
        colorbar=dict(
            thickness=20,
            len=0.7,
            bgcolor='rgba(255, 255, 255, 0.9)',
            bordercolor='#E0E0E0',
            borderwidth=2,
            tickfont=dict(size=11, color='#546E7A'),
            title=dict(
                font=dict(size=12, family="Segoe UI Semibold")
            )
        ),
        xgap=3,  # Espaciado entre celdas
        ygap=3
    )
    
    return fig

def estilo_boxplot(fig, altura=None, n_categorias=None):
    """Estilo específico para boxplots"""
    if altura is None and n_categorias:
        altura = max(400, n_categorias * 50)
    
    fig = aplicar_estilo_grafica(fig, altura or 500, mostrar_leyenda=False)
    
    # Cajas con color sólido y bordes
    fig.update_traces(
        marker=dict(
            color=COLOR_PALETTE['primary'],
            line=dict(color='#1B5E20', width=2),
            size=8
        ),
        line=dict(
            color='#1B5E20',
            width=2
        ),
        fillcolor='rgba(46, 125, 50, 0.3)'
    )
    
    return fig

# ========== FIN DEL TEMPLATE ==========

def normaliza(s):
    if s is None or (isinstance(s, float) and np.isnan(s)): return ""
    s = unicodedata.normalize("NFKD", str(s).strip()).encode("ascii","ignore").decode("ascii")
    return s

def detect_header_row(df0: pd.DataFrame):
    best_row, best_score = 0, -1
    for i in range(min(30, len(df0))):
        row = [normaliza(x).lower() for x in df0.iloc[i].tolist()]
        score = 0
        for m in MESES:
            mlo = m.lower()
            score += sum(1 for c in row if (mlo in c and ("avance" in c)) or (mlo in c and "observ" in c))
        score += 1 if any("obligac" in c for c in row) else 0
        if score > best_score: best_row, best_score = i, score
    return best_row

def map_month_columns(cols):
    mapping, cols_lo = {}, [normaliza(c).lower() for c in cols]
    for mes in MESES:
        mlo, avance_idx, obs_idx = mes.lower(), None, None
        for idx, c in enumerate(cols_lo):
            if mlo in c and "avance" in c: avance_idx = idx
            if mlo in c and "observ" in c: obs_idx = idx
        if avance_idx is not None or obs_idx is not None:
            mapping[mes] = (cols[avance_idx] if avance_idx is not None else None,
                            cols[obs_idx] if obs_idx is not None else None)
    return mapping

def to_number_percent(x):
    if pd.isna(x): return np.nan
    s = re.sub(r"[^0-9.\-]", "", str(x).replace(",", "."))
    if s in ("", "-", "."): return np.nan
    try: val = float(s)
    except: return np.nan
    return max(0, min(100, val*100 if 0 <= val <= 1 else val))

def tidy_from_excel(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    hoja = xls.sheet_names[0]
    df0 = xls.parse(hoja, header=None)
    header_row = detect_header_row(df0)
    df = pd.read_excel(file, sheet_name=hoja, header=header_row).dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    first_col = df.columns[0]
    map_mes = map_month_columns(df.columns)

    registros, categoria_actual = [], None
    for _, row in df.iterrows():
        texto = str(row[first_col]).strip() if first_col in row else ""
        if not texto or texto.lower() == "nan": continue

        texto_clean = normaliza(texto).upper()
        if any(texto_clean.startswith(normaliza(c).upper()) for c in CATEGORIAS_CANON):
            categoria_actual = next(c for c in CATEGORIAS_CANON if texto_clean.startswith(normaliza(c).upper()))
            continue

        obligacion = texto
        for mes, (avance_col, obs_col) in map_mes.items():
            avance = to_number_percent(row[avance_col]) if (avance_col in row) else np.nan
            obs = str(row[obs_col]).strip() if (obs_col in row and not pd.isna(row[obs_col])) else ""
            registros.append({"Categoria": categoria_actual or "Sin categoría",
                              "Obligacion": obligacion, "Mes": mes,
                              "Avance": avance, "Observacion": obs})
    tidy = pd.DataFrame(registros)
    if not tidy.empty:
        tidy = tidy.groupby(["Categoria","Obligacion","Mes"], as_index=False).agg(
            Avance=("Avance","mean"),
            Observacion=("Observacion", lambda x: " | ".join([s for s in x if s]))
        )
    return tidy

def mes_anterior(mes):
    return None if mes not in MESES or MESES.index(mes)==0 else MESES[MESES.index(mes)-1]

def kpis_from_tidy(tidy: pd.DataFrame):
    out = {}
    if tidy.empty: return out
    meses_reportados = tidy.groupby("Mes")["Avance"].apply(lambda s: s.notna().sum())
    meses_con_dato = [m for m in MESES if m in meses_reportados.index and meses_reportados[m]>0]
    ultimo_mes = meses_con_dato[-1] if meses_con_dato else None
    global_ultimo = tidy.loc[tidy["Mes"]==ultimo_mes, "Avance"].mean() if ultimo_mes else tidy["Avance"].mean()
    por_cat = (tidy[tidy["Mes"]==ultimo_mes].groupby("Categoria", as_index=False)["Avance"].mean()
               if ultimo_mes else tidy.groupby("Categoria", as_index=False)["Avance"].mean())
    por_mes = tidy.groupby("Mes", as_index=False)["Avance"].mean()
    por_mes["orden"] = por_mes["Mes"].apply(lambda m: MESES.index(m) if m in MESES else 999)
    por_mes = por_mes.sort_values("orden")
    mejor_mes = por_mes.loc[por_mes["Avance"].idxmax(),"Mes"] if not por_mes.empty else None
    # riesgos
    en_riesgo, obligaciones_riesgo = 0, []
    if ultimo_mes:
        piv = tidy.pivot_table(index=["Categoria","Obligacion"], columns="Mes", values="Avance", aggfunc="mean")
        if ultimo_mes in piv.columns:
            mask = piv[ultimo_mes] < 70
            en_riesgo = mask.fillna(False).sum()
            obligaciones_riesgo = [f"{c} | {o}" for (c,o) in piv[mask].index.tolist()]
    out.update(dict(
        ultimo_mes=ultimo_mes,
        avance_global=None if pd.isna(global_ultimo) else float(np.round(global_ultimo,2)),
        avance_por_categoria=por_cat,
        promedio_por_mes=por_mes[["Mes","Avance"]],
        mejor_mes=mejor_mes,
        obligaciones_en_riesgo=int(en_riesgo),
        lista_obligaciones_riesgo=obligaciones_riesgo
    ))
    return out

def semaforo(v, meta):
    if pd.isna(v): return "⚪ Sin dato"
    if v >= meta: return "🟢 Alto"
    if v >= 70:  return "🟡 Medio"
    return "🔴 Bajo"

def _norm_txt(x: str) -> str:
    if not isinstance(x, str): x = "" if x is None else str(x)
    x = x.lower().strip()
    x = unicodedata.normalize("NFKD", x).encode("ascii","ignore").decode("ascii")
    return x

def build_theme_index(theme_dict):
    """
    theme_dict: dict {'Tema': ['palabra clave', 'frase', ...], ...}
    Devuelve {'Tema': [regex_compilada, ...]}
    """
    idx = {}
    for tema, keywords in theme_dict.items():
        pats = []
        for kw in keywords:
            kw_norm = re.escape(_norm_txt(kw))
            # palabra/frase completa, tolerando signos
            pats.append(re.compile(rf'\b{kw_norm}\b', re.IGNORECASE))
        idx[tema] = pats
    return idx

def classify_observation(texto: str, theme_index) -> list:
    """
    Retorna lista de temas que matchean; puede ser multi-label.
    """
    t = _norm_txt(texto)
    if not t: return []
    hits = []
    for tema, patterns in theme_index.items():
        if any(p.search(t) for p in patterns):
            hits.append(tema)
    return hits

# Gancho IA (fallback) — no llama servicios externos; deja la interfaz para conectarlo luego.
def ai_fallback_classify(texto: str) -> str:
    """
    TODO: Conectar a tu proveedor de IA (Azure OpenAI / local).
    Por ahora devuelve '' para que solo se use si no matchea reglas.
    """
    return ""
# ---------- UI: carga ----------
st.sidebar.header("1) Cargar archivo")
file = st.sidebar.file_uploader("Excel (.xlsx)", type=["xlsx"])
if not file:
    st.info("Sube el archivo .xlsx para iniciar.")
    st.stop()

try:
    tidy = tidy_from_excel(file)
except Exception as e:
    st.error(f"No pude leer el Excel: {e}")
    st.stop()

if tidy.empty:
    st.warning("No se detectaron registros válidos. Revisa cabeceras de meses.")
    st.stop()

# ---------- Filtros ----------
st.sidebar.header("2) Filtros")
cat_opts = ["(Todas)"] + sorted(tidy["Categoria"].dropna().unique().tolist())
cat_sel = st.sidebar.selectbox("Categoría", cat_opts)
mes_opts = ["(Todos)"] + [m for m in MESES if m in tidy["Mes"].unique()]
mes_sel = st.sidebar.selectbox("Mes", mes_opts)

st.sidebar.header("3) Parámetros")
meta = st.sidebar.slider("Meta de cumplimiento (%)", 70, 100, 90, 1)
top_n = st.sidebar.slider("Top N (riesgos/caídas)", 3, 20, 10, 1)

# ---------- Análisis temático de observaciones ----------
st.sidebar.header("4) Análisis temático")
enable_themes = st.sidebar.toggle("Activar análisis temático (observaciones)", value=True)

# Diccionario base (puedes editar aquí las palabras por tema)
DEFAULT_THEMES = {
    "Administrativa": ["acta", "firma", "documento", "radicado", "tramite", "aprobacion", "resolucion"],
    "Financiera": ["pago", "transferencia", "cdp", "presupuesto", "anticipo", "reintegro", "factura"],
    "Logistica": ["transporte", "entrega", "retraso logistica", "bodega", "inventario", "despacho"],
    "Talento Humano": ["personal", "contratacion", "hoja de vida", "capacitacion", "ausencia"],
    "Tecnica/Operativa": ["sistema", "plataforma", "error", "incidencia", "validacion", "dato", "gps"],
}

# Permitir cargar un CSV con columnas: tema, keyword
csv_kw = st.sidebar.file_uploader("Diccionario de palabras clave (CSV: tema,keyword)", type=["csv"], key="kw_csv")
theme_dict = DEFAULT_THEMES.copy()
if csv_kw is not None:
    try:
        df_kw = pd.read_csv(csv_kw)
        if {"tema","keyword"}.issubset({c.lower() for c in df_kw.columns}):
            # normaliza nombres de columnas
            col_tema = [c for c in df_kw.columns if c.lower()=="tema"][0]
            col_kw   = [c for c in df_kw.columns if c.lower()=="keyword"][0]
            for tema, grupo in df_kw.groupby(col_tema):
                tema = str(tema).strip()
                kws = [str(x).strip() for x in grupo[col_kw].dropna().tolist() if str(x).strip()]
                if tema in theme_dict:
                    theme_dict[tema].extend(kws)
                else:
                    theme_dict[tema] = kws
            st.sidebar.success("Diccionario de temas cargado.")
        else:
            st.sidebar.warning("El CSV debe tener columnas: tema, keyword")
    except Exception as e:
        st.sidebar.error(f"No se pudo leer el CSV: {e}")

# Construye índice de regex
theme_index = build_theme_index(theme_dict)

# Permitir seleccionar qué temas mostrar en gráficos
tema_opts = list(theme_dict.keys())
tema_sel = st.sidebar.multiselect("Temas a mostrar (si vacía = todos)", options=tema_opts, default=tema_opts)


# Aplica filtros
df_view = tidy.copy()
if cat_sel != "(Todas)":
    df_view = df_view[df_view["Categoria"] == cat_sel]
if mes_sel != "(Todos)":
    df_view = df_view[df_view["Mes"] == mes_sel]

# Si no hay datos con los filtros, avisar y parar
if df_view.empty:
    st.warning("No hay datos para los filtros seleccionados.")
    st.stop()

# KPIs sobre la VISTA FILTRADA
kpis = kpis_from_tidy(df_view)

# fallback si último mes no se detecta tras filtrar
if not kpis.get("ultimo_mes"):
    meses_presentes = [m for m in MESES if m in df_view["Mes"].unique()]
    if meses_presentes:
        kpis["ultimo_mes"] = meses_presentes[-1]

#ultimo = kpis.get("ultimo_mes")
#c1,c2,c3,c4 = st.columns(4)
##c1.metric("Último mes con datos", kpis.get("ultimo_mes") or "—")
#c2.metric("Avance global (%)", f"{kpis.get('avance_global') if kpis.get('avance_global') is not None else '—'}")
#c3.metric("Obligaciones en riesgo", kpis.get("obligaciones_en_riesgo"))
#c4.metric("Mejor mes", kpis.get("mejor_mes") or "—")
#st.divider()

# --- Helpers visuales (añadir una vez) ---
def wrap_text(s: str, width: int = 32, max_lines: int = 3) -> str:
    if not isinstance(s, str): s = str(s)
    words, lines, line = s.split(), [], ""
    for w in words:
        if len(line + " " + w) <= width:
            line = (line + " " + w).strip()
        else:
            lines.append(line); line = w
        if len(lines) >= max_lines: break
    if line and len(lines) < max_lines: lines.append(line)
    out = "<br>".join(lines)
    if len(words) > 1 and (len(lines) == max_lines and " ".join(words).find(lines[-1]) < len(s)):
        out += "…"
    return out

def auto_height(n_items: int, row=30, base=280, max_h=900):
    return int(min(max_h, max(base, 140 + row * max(3, n_items))))

# ---------- KPIs ----------
kpis = kpis_from_tidy(tidy)
c1,c2,c3,c4 = st.columns(4)
c1.metric("Último mes con datos", kpis.get("ultimo_mes") or "—")
c2.metric("Avance global (%)", f"{kpis.get('avance_global') if kpis.get('avance_global') is not None else '—'}")
c3.metric("Obligaciones en riesgo", kpis.get("obligaciones_en_riesgo"))
c4.metric("Mejor mes", kpis.get("mejor_mes") or "—")
st.divider()
if not kpis.get("ultimo_mes"):
    # usa el último mes del año con cualquier dato
    meses_presentes = [m for m in MESES if m in tidy["Mes"].unique()]
    if meses_presentes: kpis["ultimo_mes"] = meses_presentes[-1]
ultimo = kpis.get("ultimo_mes")

# ========== GRÁFICA 1: EVOLUCIÓN MENSUAL ==========
# Busca "# ---------- Gráfico: serie mensual ----------" y reemplaza con:

st.subheader("📈 Evolución mensual del avance promedio")
serie = kpis.get("promedio_por_mes", pd.DataFrame(columns=["Mes","Avance"]))

if not serie.empty:
    fig1 = go.Figure()
    
    # Línea principal con gradiente
    fig1.add_trace(go.Scatter(
        x=serie["Mes"],
        y=serie["Avance"],
        mode='lines+markers',
        name='Avance',
        line=dict(
            color=COLOR_PALETTE['primary'], 
            width=4,
            shape='spline',
            smoothing=1.3
        ),
        marker=dict(
            size=14, 
            color=COLOR_PALETTE['primary'], 
            line=dict(width=3, color='white'),
            symbol='circle'
        ),
        fill='tozeroy',
        fillcolor='rgba(46, 125, 50, 0.1)',
        hovertemplate='<b>%{x}</b><br>Avance: <b>%{y:.1f}%</b><extra></extra>'
    ))
    
    # Línea de meta estilizada
    fig1.add_hline(
        y=meta, 
        line_dash="dash", 
        line_color=COLOR_PALETTE['success'],
        line_width=3,
        annotation_text=f"🎯 Meta: {meta}%",
        annotation_position="right",
        annotation=dict(
            font=dict(size=13, color=COLOR_PALETTE['success'], family="Segoe UI Semibold"),
            bgcolor="rgba(255, 255, 255, 0.9)",
            bordercolor=COLOR_PALETTE['success'],
            borderwidth=2,
            borderpad=6
        )
    )
    
    # Área sombreada de riesgo con gradiente
    fig1.add_hrect(
        y0=0, y1=70,
        fillcolor=COLOR_PALETTE['danger'],
        opacity=0.08,
        line_width=0,
        annotation_text="⚠️ Zona de riesgo",
        annotation_position="top left",
        annotation=dict(
            font=dict(size=11, color=COLOR_PALETTE['danger'])
        )
    )
    
    # Aplicar estilo
    fig1 = estilo_linea_temporal(fig1, altura=500)
    
    fig1.update_layout(
        title="Evolución del cumplimiento a lo largo del año",
        xaxis_title="Mes",
        yaxis_title="Avance Promedio (%)",
        yaxis_range=[0, 105]
    )
    
    st.plotly_chart(fig1, use_container_width=True)
else:
    st.info("Sin datos para la serie temporal.")

# ========== GRÁFICA 2: BARRAS POR CATEGORÍA ==========
# Busca "# ---------- Gráfico: barras por categoría ----------" y reemplaza con:

st.subheader("📊 Avance por categoría (último mes disponible)")
cat_df = kpis.get("avance_por_categoria", pd.DataFrame(columns=["Categoria","Avance"]))

if not cat_df.empty:
    cat_df = cat_df.sort_values("Avance", ascending=True)
    
    # Colores con gradiente según nivel
    colors = []
    for v in cat_df["Avance"]:
        if v < 70:
            colors.append(COLOR_PALETTE['danger'])
        elif v < meta:
            colors.append(COLOR_PALETTE['warning'])
        else:
            colors.append(COLOR_PALETTE['success'])
    
    fig2 = go.Figure()
    
    fig2.add_trace(go.Bar(
        y=cat_df["Categoria"],
        x=cat_df["Avance"],
        orientation='h',
        text=cat_df["Avance"].apply(lambda x: f'{x:.1f}%'),
        textposition='outside',
        textfont=dict(size=13, family="Segoe UI Semibold", color='#37474F'),
        marker=dict(
            color=colors,
            line=dict(color='white', width=2),
            cornerradius=8  # Bordes redondeados
        ),
        hovertemplate='<b>%{y}</b><br>Avance: <b>%{x:.1f}%</b><extra></extra>'
    ))
    
    # Línea de meta con anotación
    fig2.add_vline(
        x=meta,
        line_dash="dash",
        line_color=COLOR_PALETTE['neutral'],
        line_width=3,
        annotation_text=f"Meta {meta}%",
        annotation=dict(
            font=dict(size=12, family="Segoe UI Semibold"),
            bgcolor="rgba(255, 255, 255, 0.9)",
            bordercolor=COLOR_PALETTE['neutral'],
            borderwidth=2
        )
    )
    
    # Aplicar estilo
    fig2 = estilo_barras_horizontal(fig2, n_items=len(cat_df))
    
    fig2.update_layout(
        title=f"Comparativo de cumplimiento por categoría - {ultimo or 'último mes'}",
        xaxis_title="Avance (%)",
        yaxis_title="",
        xaxis_range=[0, 110]
    )
    
    st.plotly_chart(fig2, use_container_width=True)
else:
    st.info("Sin datos por categoría.")

# ---------- Heatmap observaciones (respeta filtros) ----------
# ========== GRÁFICA 3: HEATMAP ==========
# Busca "# ---------- Heatmap observaciones ----------" y reemplaza con:

st.subheader("🧭 Mapa de calor: Observaciones por categoría y mes")

obs = df_view.assign(
    TieneObs=df_view["Observacion"].fillna("").str.strip().ne("")
)

heat = (
    obs.groupby(["Categoria", "Mes"], as_index=False)["TieneObs"]
       .sum()
       .rename(columns={"TieneObs": "Conteo"})
)

if not heat.empty:
    heat["orden_mes"] = heat["Mes"].apply(lambda m: MESES.index(m) if m in MESES else 999)
    heat = heat.sort_values(["Categoria", "orden_mes"])

    pivot = (
        heat.pivot(index="Categoria", columns="Mes", values="Conteo")
            .reindex(columns=[m for m in MESES if m in heat["Mes"].unique()], fill_value=0)
            .fillna(0)
    )
    
    pivot = pivot.reindex(sorted(pivot.index), axis=0)

    # Heatmap con escala mejorada
    fig3 = go.Figure(data=go.Heatmap(
        z=pivot.values,
        x=pivot.columns,
        y=pivot.index,
        colorscale=[
            [0, '#E8F5E9'],      # Verde muy claro
            [0.25, '#C8E6C9'],   # Verde claro
            [0.5, '#FFE0B2'],    # Naranja claro
            [0.75, '#FFAB91'],   # Naranja
            [1, '#EF9A9A']       # Rojo claro
        ],
        text=pivot.values,
        texttemplate='<b>%{text}</b>',
        textfont={"size": 13, "color": "#1A252F", "family": "Segoe UI Semibold"},
        hovertemplate='<b>%{y}</b><br>%{x}: <b>%{z} observaciones</b><extra></extra>',
        xgap=3,  # Espaciado entre celdas
        ygap=3,
        colorbar=dict(
            title="Cantidad",
            titleside="right",
            titlefont=dict(size=13, family="Segoe UI Semibold"),
            thickness=20,
            len=0.7,
            bgcolor='rgba(255, 255, 255, 0.9)',
            bordercolor='#E0E0E0',
            borderwidth=2
        )
    ))
    
    # Aplicar estilo
    fig3 = estilo_heatmap(fig3, n_categorias=len(pivot))
    
    fig3.update_layout(
        title="Concentración de observaciones por período y responsable",
        xaxis_title="Mes",
        yaxis_title="Categoría",
        xaxis=dict(side='bottom')
    )
    
    st.plotly_chart(fig3, use_container_width=True)
    
    # Insight
    max_cat = pivot.sum(axis=1).idxmax()
    max_mes = pivot.sum(axis=0).idxmax()
    st.caption(f"💡 **Insight:** La categoría '{max_cat}' y el mes '{max_mes}' concentran más observaciones.")
else:
    st.info("No hay observaciones para graficar con los filtros actuales.")

st.divider()

# ---------- Tabla semaforizada ----------
if mes_sel == "(Todos)":
    tabla = df_view.copy()
    subtitulo_mes = "todos los meses"
else:
    tabla = df_view[df_view["Mes"] == mes_sel].copy()
    subtitulo_mes = mes_sel

st.subheader(f"📋 Estado por obligación ({subtitulo_mes})")

if not tabla.empty:
    # Semáforo por fila según la meta
    tabla["Estado"] = tabla["Avance"].apply(lambda v: semaforo(v, meta))
    # Ordenar meses cronológicamente
    tabla["orden_mes"] = tabla["Mes"].map(lambda m: MESES.index(m) if m in MESES else 999)

    mostrar = (
        tabla[["Categoria", "Obligacion", "Mes", "Avance", "Observacion", "Estado", "orden_mes"]]
        .sort_values(["Categoria", "Obligacion", "orden_mes"])
        .drop(columns=["orden_mes"])
    )
    st.dataframe(mostrar, use_container_width=True, height=460)
else:
    st.info("No hay datos para los filtros seleccionados.")

# ---------- TOPS ----------
st.subheader(f"🚨 Top {top_n} obligaciones en situación crítica")

if ultimo:
    base_last = (tidy[tidy["Mes"]==ultimo]
                 .groupby(["Categoria","Obligacion"], as_index=False)["Avance"].mean())
    top_riesgo = base_last.sort_values("Avance", ascending=True).head(top_n).copy()
    
    if not top_riesgo.empty:
        top_riesgo["Obligacion_wrapped"] = top_riesgo["Obligacion"].apply(
            lambda s: '<br>'.join([s[i:i+50] for i in range(0, len(s), 50)])
        )
        
        # Colores degradados según nivel de riesgo
        colors = [
            f'rgb({int(229 - (v/100)*80)}, {int(53 + (v/100)*100)}, 53)'
            for v in top_riesgo["Avance"]
        ]
        
        fig_top = go.Figure()
        
        fig_top.add_trace(go.Bar(
            y=top_riesgo["Obligacion_wrapped"],
            x=top_riesgo["Avance"],
            orientation='h',
            marker=dict(
                color=colors,
                line=dict(color='white', width=1)
            ),
            text=top_riesgo["Avance"].apply(lambda x: f'{x:.1f}%'),
            textposition='outside',
            hovertemplate='<b>%{y}</b><br>Avance: %{x:.1f}%<extra></extra>'
        ))
        
        # Zona de meta
        fig_top.add_vrect(
            x0=meta, x1=100,
            fillcolor=COLOR_PALETTE['success'],
            opacity=0.1,
            line_width=0
        )
        
        fig_top.update_layout(
            template=PLOTLY_TEMPLATE,
            title=f"Obligaciones con menor avance en {ultimo} (requieren atención inmediata)",
            xaxis_title="% Avance",
            yaxis_title="",
            xaxis_range=[0, 110],
            height=auto_height(len(top_riesgo), row=45),
            showlegend=False
        )
        
        st.plotly_chart(fig_top, use_container_width=True)
        
        # Tabla complementaria con detalles
        with st.expander("📋 Ver detalles de obligaciones en riesgo"):
            detalle = top_riesgo[["Categoria", "Obligacion", "Avance"]].copy()
            detalle["Gap vs Meta"] = (meta - detalle["Avance"]).round(1)
            st.dataframe(detalle, use_container_width=True)
    else:
        st.success("✅ No hay obligaciones en situación crítica.")
else:
    st.info("Se requiere un mes de referencia con datos.")

st.subheader("📉 Mayores caídas respecto al mes anterior (Δ)")
caidas = pd.DataFrame()
prev = mes_anterior(ultimo) if ultimo else None
if ultimo and prev:
    g = (tidy[tidy["Mes"].isin([prev, ultimo])]
         .groupby(["Categoria","Obligacion","Mes"], as_index=False)["Avance"].mean())
    wide = g.pivot(index=["Categoria","Obligacion"], columns="Mes", values="Avance")
    if prev in wide.columns and ultimo in wide.columns:
        wide["Delta"] = (wide[ultimo] - wide[prev]).round(2)
        caidas = (wide.sort_values("Delta", ascending=True)
                       .reset_index()
                       .head(top_n)
                       .rename(columns={"Delta":"Delta_%"}))
        caidas["Obligacion_wrapped"] = caidas["Obligacion"].apply(lambda s: wrap_text(s, 40, 3))
        fig_delta = px.bar(
            caidas.sort_values("Delta_%"),
            y="Obligacion_wrapped", x="Delta_%", color="Categoria",
            orientation="h",
            labels={"Delta_":"Δ %","Obligacion_wrapped":"Obligación"},
            title=f"Cambio {prev} → {ultimo} (negativo = cae)"
        )
        # margen a ambos lados para ver negativos/positivos
        dx = float(np.nanmax(np.abs(caidas["Delta_%"]))) if not caidas.empty else 0
        fig_delta.update_layout(
            height=auto_height(len(caidas), row=32),
            xaxis=dict(zeroline=True, zerolinewidth=2, range=[-dx*1.1, dx*1.1] if dx>0 else None),
            margin=dict(l=10, r=10, t=60, b=10), legend_title_text="Categoría"
        )
        fig_delta.update_traces(hovertemplate="<b>%{y}</b><br>Δ: %{x:.1f} p.p.<extra>%{legendgroup}</extra>")
        st.plotly_chart(fig_delta, use_container_width=True)
    else:
        st.info("No hay columnas suficientes para Δ.")
else:
    st.info("Se requieren dos meses consecutivos con datos.")


# ---------- Pareto de observaciones ----------
st.subheader("📊 Análisis de Pareto - Principio 80/20")

obs_cnt = (tidy.assign(TieneObs=tidy["Observacion"].fillna("").str.strip().ne(""))
           .groupby(["Categoria","Obligacion"], as_index=False)["TieneObs"].sum()
           .rename(columns={"TieneObs":"Observaciones"}))

if not obs_cnt.empty:
    obs_cnt = obs_cnt.sort_values("Observaciones", ascending=False).reset_index(drop=True)
    total_obs = int(obs_cnt["Observaciones"].sum())
    obs_cnt["Acumulado"] = obs_cnt["Observaciones"].cumsum()
    obs_cnt["%Acum"] = (obs_cnt["Acumulado"] / max(total_obs,1) * 100).round(1)
    
    top_p = obs_cnt.head(top_n).copy()
    top_p["Obligacion_short"] = top_p["Obligacion"].apply(
        lambda s: s[:40] + '...' if len(s) > 40 else s
    )
    
    # Crear gráfico con dos ejes Y
    fig_pareto = go.Figure()
    
    # Barras de frecuencia
    fig_pareto.add_trace(go.Bar(
        x=top_p["Obligacion_short"],
        y=top_p["Observaciones"],
        name='Observaciones',
        marker=dict(color=COLOR_PALETTE['primary']),
        yaxis='y',
        hovertemplate='<b>%{x}</b><br>Observaciones: %{y}<extra></extra>'
    ))
    
    # Línea de porcentaje acumulado
    fig_pareto.add_trace(go.Scatter(
        x=top_p["Obligacion_short"],
        y=top_p["%Acum"],
        name='% Acumulado',
        mode='lines+markers',
        line=dict(color=COLOR_PALETTE['danger'], width=3),
        marker=dict(size=10),
        yaxis='y2',
        hovertemplate='<b>%{x}</b><br>Acumulado: %{y:.1f}%<extra></extra>'
    ))
    
    # Línea de referencia 80%
    fig_pareto.add_hline(
        y=80,
        line_dash="dash",
        line_color=COLOR_PALETTE['warning'],
        annotation_text="80%",
        yref='y2'
    )
    
    fig_pareto.update_layout(
        template=PLOTLY_TEMPLATE,
        title=f"Pareto: {top_n} obligaciones concentran el {top_p['%Acum'].iloc[-1]:.0f}% de observaciones (Total: {total_obs})",
        xaxis_title="",
        yaxis=dict(
            title="Cantidad de Observaciones",
            side='left'
        ),
        yaxis2=dict(
            title="Porcentaje Acumulado (%)",
            overlaying='y',
            side='right',
            range=[0, 105]
        ),
        height=500,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        xaxis=dict(tickangle=-45)
    )
    
    st.plotly_chart(fig_pareto, use_container_width=True)
    
    # Cálculo automático del 80%
    cum_80 = obs_cnt[obs_cnt["%Acum"] <= 80]
    st.info(f"💡 **Principio 80/20:** Las primeras {len(cum_80)} obligaciones ({len(cum_80)/len(obs_cnt)*100:.1f}%) concentran el 80% de las observaciones.")
else:
    st.info("No hay observaciones registradas.")

# ========== GRÁFICA 4: BOXPLOT ==========
# Busca "# ---------- Boxplot por categoría ----------" y reemplaza con:

st.subheader("📦 Distribución de avance por categoría (boxplot)")
bx = df_view.dropna(subset=["Avance"])

if not bx.empty:
    med_orden = (bx.groupby("Categoria")["Avance"].median()
                   .sort_values(ascending=True).index.tolist())
    
    fig_box = go.Figure()
    
    for i, cat in enumerate(med_orden):
        datos_cat = bx[bx["Categoria"] == cat]["Avance"]
        
        fig_box.add_trace(go.Box(
            y=datos_cat,
            name=cat,
            marker=dict(
                color=CATEGORIA_COLORS[i % len(CATEGORIA_COLORS)],
                line=dict(color='white', width=2),
                size=8
            ),
            line=dict(color=CATEGORIA_COLORS[i % len(CATEGORIA_COLORS)], width=2),
            fillcolor=f'rgba{tuple(list(int(CATEGORIA_COLORS[i % len(CATEGORIA_COLORS)].lstrip("#")[i:i+2], 16) for i in (0, 2, 4)) + [0.3])}',
            boxmean='sd',  # Mostrar media y desviación estándar
            hovertemplate='<b>%{fullData.name}</b><br>Valor: <b>%{y:.1f}%</b><extra></extra>'
        ))
    
    # Aplicar estilo
    fig_box = estilo_boxplot(fig_box, n_categorias=len(med_orden))
    
    fig_box.update_layout(
        title="Dispersión de avances por categoría (año completo)",
        xaxis_title="Categoría",
        yaxis_title="% Avance",
        yaxis_range=[0, 105],
        showlegend=False
    )
    
    st.plotly_chart(fig_box, use_container_width=True)
else:
    st.info("No hay datos para boxplot.")

# ---------- Control chart ----------
st.subheader("📉 Gráfico de control estadístico (SPC)")

serie_cc = kpis.get("promedio_por_mes")
if serie_cc is not None and not serie_cc.empty:
    m = float(serie_cc["Avance"].mean())
    s = float(serie_cc["Avance"].std() or 0)
    ucl, lcl = m + 2*s, max(m - 2*s, 0)
    
    fig_cc = go.Figure()
    
    # Banda de control (±2σ)
    fig_cc.add_trace(go.Scatter(
        x=serie_cc["Mes"],
        y=[ucl] * len(serie_cc),
        mode='lines',
        name='UCL (+2σ)',
        line=dict(dash='dash', color=COLOR_PALETTE['warning'], width=2),
        showlegend=True
    ))
    
    fig_cc.add_trace(go.Scatter(
        x=serie_cc["Mes"],
        y=[lcl] * len(serie_cc),
        mode='lines',
        name='LCL (-2σ)',
        line=dict(dash='dash', color=COLOR_PALETTE['warning'], width=2),
        fill='tonexty',
        fillcolor='rgba(251, 140, 0, 0.1)',
        showlegend=True
    ))
    
    # Línea media
    fig_cc.add_trace(go.Scatter(
        x=serie_cc["Mes"],
        y=[m] * len(serie_cc),
        mode='lines',
        name=f'Media ({m:.1f}%)',
        line=dict(dash='dot', color=COLOR_PALETTE['neutral'], width=2),
        showlegend=True
    ))
    
    # Datos reales
    fig_cc.add_trace(go.Scatter(
        x=serie_cc["Mes"],
        y=serie_cc["Avance"],
        mode='lines+markers',
        name='Avance real',
        line=dict(color=COLOR_PALETTE['primary'], width=3),
        marker=dict(size=10, color=COLOR_PALETTE['primary'],
                   line=dict(width=2, color='white')),
        hovertemplate='<b>%{x}</b><br>Avance: %{y:.1f}%<extra></extra>',
        showlegend=True
    ))
    
    # Detectar puntos fuera de control
    outliers = serie_cc[(serie_cc["Avance"] > ucl) | (serie_cc["Avance"] < lcl)]
    if not outliers.empty:
        fig_cc.add_trace(go.Scatter(
            x=outliers["Mes"],
            y=outliers["Avance"],
            mode='markers',
            name='Fuera de control',
            marker=dict(size=15, color=COLOR_PALETTE['danger'],
                       symbol='x', line=dict(width=2)),
            showlegend=True
        ))
    
    fig_cc.update_layout(
        template=PLOTLY_TEMPLATE,
        title="Control estadístico del proceso - Detección de variaciones anormales",
        xaxis_title="Mes",
        yaxis_title="% Avance",
        height=500,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    st.plotly_chart(fig_cc, use_container_width=True)
    
    # Interpretación automática
    if not outliers.empty:
        st.warning(f"⚠️ **Alerta:** Se detectaron {len(outliers)} mes(es) con variación anormal: {', '.join(outliers['Mes'].tolist())}")
    else:
        st.success("✅ Proceso bajo control estadístico. No se detectan variaciones anormales.")
else:
    st.info("Sin datos para control chart.")

# ---------- Stack Cumplido vs Pendiente ----------
st.subheader("🧮 Cumplido vs Pendiente por mes (promedio)")
stack = tidy.groupby("Mes", as_index=False)["Avance"].mean()
if not stack.empty:
    stack["Cumplido"] = stack["Avance"].clip(0,100)
    stack["Pendiente"] = (100 - stack["Cumplido"]).clip(0,100)
    stack["orden"] = stack["Mes"].apply(lambda m: MESES.index(m) if m in MESES else 999)
    fig_stack = px.bar(stack.sort_values("orden"), x="Mes", y=["Cumplido","Pendiente"],
                       barmode="stack", labels={"value":"%", "variable":"Estado"},
                       title="Promedio mensual — Cumplido vs Pendiente")
    fig_stack.update_layout(yaxis_range=[0,100])
    st.plotly_chart(fig_stack, use_container_width=True)
else:
    st.info("Sin datos para el stack.")

st.divider()

# ---------- Descargas ----------
st.subheader("⬇️ Descargas")

usar_filtros = st.checkbox("Descargar usando filtros aplicados (recomendado)", value=True)

# Seleccionar dataset base según configuración del usuario
if usar_filtros:
    dataset_export = df_view.copy()
    kpis_export = kpis
else:
    dataset_export = tidy.copy()
    kpis_export = kpis_from_tidy(tidy)

# Preparar datasets auxiliares
serie_export = kpis_export.get("promedio_por_mes", pd.DataFrame())
cat_export   = kpis_export.get("avance_por_categoria", pd.DataFrame())

ultimo_export = kpis_export.get("ultimo_mes")
if ultimo_export:
    tabla_export = dataset_export[dataset_export["Mes"]==ultimo_export].copy()
else:
    tabla_export = dataset_export.copy()

# Exportar consolidado en Excel
with io.BytesIO() as buffer:
    with pd.ExcelWriter(buffer, engine="openpyxl") as w:
        # Base (según usar_filtros): df_view o tidy
        dataset_export.to_excel(w, "Consolidado_Tidy", index=False)

        # KPIs derivados (según usar_filtros)
        if not cat_export.empty:
            cat_export.to_excel(w, "Avance_por_Categoria", index=False)
        if not serie_export.empty:
            serie_export.to_excel(w, "Promedio_por_Mes", index=False)

        # Tabla de obligaciones (último mes del contexto exportado)
        if not tabla_export.empty:
            tabla_tmp = tabla_export.copy()
            tabla_tmp["Estado"] = tabla_tmp["Avance"].apply(lambda v: semaforo(v, meta))
            tabla_tmp[["Categoria","Obligacion","Mes","Avance","Observacion","Estado"]].to_excel(
                w, "Obligaciones_UltimoMes", index=False
            )

        # Tops (solo si se está exportando con filtros y existen)
        try:
            if usar_filtros and 'top_riesgo' in globals() and not top_riesgo.empty:
                top_riesgo.to_excel(w, "Top_Riesgo", index=False)
        except:
            pass
        try:
            if usar_filtros and 'caidas' in globals() and not caidas.empty:
                caidas.to_excel(w, "Top_Caidas", index=False)
        except:
            pass
        try:
            if usar_filtros and 'obs_cnt' in globals() and not obs_cnt.empty:
                # Si quieres todo el pareto, quita .head(top_n)
                obs_cnt.head(top_n).to_excel(w, "Pareto_Obs", index=False)
        except:
            pass

        # ---- NUEVO: Hojas de análisis temático (si está activado y hay datos) ----
        try:
            if 'enable_themes' in globals() and enable_themes and 'temas_export_mes' in locals() and not temas_export_mes.empty:
                temas_export_mes.to_excel(w, "Temas_por_Mes", index=False)
        except:
            pass
        try:
            if 'enable_themes' in globals() and enable_themes and 'temas_export_cat' in locals() and not temas_export_cat.empty:
                temas_export_cat.to_excel(w, "Temas_por_Categoria", index=False)
        except:
            pass
        try:
            if 'enable_themes' in globals() and enable_themes and 'temas_export_top' in locals() and not temas_export_top.empty:
                temas_export_top.to_excel(w, "Top_Tematicas", index=False)
        except:
            pass

    st.download_button(
        "📥 Descargar consolidado (Excel)",
        data=buffer.getvalue(),
        file_name="MatrizObligaciones_Consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



# KPIs CSV
k_df = pd.DataFrame([
    ["ultimo_mes", kpis.get("ultimo_mes")],
    ["avance_global", kpis.get("avance_global")],
    ["obligaciones_en_riesgo", kpis.get("obligaciones_en_riesgo")],
    ["mejor_mes", kpis.get("mejor_mes")],
    ["meta", meta],
], columns=["KPI","Valor"])
st.download_button(
    "📥 Descargar KPIs (CSV)",
    data=k_df.to_csv(index=False).encode("utf-8"),
    file_name="KPIs_MatrizObligaciones.csv",
    mime="text/csv"
)

# Reporte HTML imprimible
def generar_reporte_html_completo():
    """Genera un reporte HTML completo con gráficas interactivas y descripciones"""
    
    # CSS personalizado con estilos MADR
    css_styles = """
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            color: #2C3E50;
            padding: 20px;
            line-height: 1.6;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            padding: 40px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
        }
        
        .header {
            text-align: center;
            padding: 30px 0;
            border-bottom: 4px solid #2E7D32;
            margin-bottom: 40px;
            background: linear-gradient(135deg, #2E7D32 0%, #43A047 100%);
            color: white;
            border-radius: 10px;
            box-shadow: 0 5px 15px rgba(46, 125, 50, 0.3);
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }
        
        .header p {
            font-size: 1.1em;
            opacity: 0.95;
        }
        
        .metadata {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 30px;
            border-left: 5px solid #2E7D32;
        }
        
        .metadata p {
            margin: 8px 0;
            font-size: 1.05em;
        }
        
        .metadata strong {
            color: #2E7D32;
            font-weight: 600;
        }
        
        .kpi-section {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }
        
        .kpi-card {
            background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
            padding: 25px;
            border-radius: 12px;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0,0,0,0.08);
            border: 2px solid transparent;
            transition: all 0.3s ease;
        }
        
        .kpi-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.12);
            border-color: #2E7D32;
        }
        
        .kpi-label {
            font-size: 0.95em;
            color: #546E7A;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 10px;
            font-weight: 600;
        }
        
        .kpi-value {
            font-size: 2.5em;
            font-weight: bold;
            color: #2E7D32;
            margin: 10px 0;
        }
        
        .kpi-description {
            font-size: 0.85em;
            color: #78909C;
            margin-top: 8px;
        }
        
        .section {
            margin: 50px 0;
            page-break-inside: avoid;
        }
        
        .section-title {
            font-size: 1.8em;
            color: #2E7D32;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 3px solid #E8F5E9;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .section-description {
            background: #E8F5E9;
            padding: 15px 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            border-left: 4px solid #43A047;
            font-size: 1.05em;
            color: #1B5E20;
        }
        
        .chart-container {
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            margin: 20px 0;
        }
        
        .insight-box {
            background: linear-gradient(135deg, #FFF3E0 0%, #FFE0B2 100%);
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
            border-left: 5px solid #FB8C00;
        }
        
        .insight-box strong {
            color: #E65100;
            font-size: 1.1em;
        }
        
        .alert-box {
            background: linear-gradient(135deg, #FFEBEE 0%, #FFCDD2 100%);
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
            border-left: 5px solid #E53935;
        }
        
        .alert-box strong {
            color: #C62828;
            font-size: 1.1em;
        }
        
        .success-box {
            background: linear-gradient(135deg, #E8F5E9 0%, #C8E6C9 100%);
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
            border-left: 5px solid #43A047;
        }
        
        .success-box strong {
            color: #2E7D32;
            font-size: 1.1em;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background: white;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            border-radius: 8px;
            overflow: hidden;
        }
        
        th {
            background: linear-gradient(135deg, #2E7D32 0%, #43A047 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.9em;
            letter-spacing: 0.5px;
        }
        
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #E0E0E0;
        }
        
        tr:hover {
            background: #F5F5F5;
        }
        
        tr:last-child td {
            border-bottom: none;
        }
        
        .footer {
            margin-top: 60px;
            padding-top: 30px;
            border-top: 3px solid #E8F5E9;
            text-align: center;
            color: #78909C;
            font-size: 0.95em;
        }
        
        .page-break {
            page-break-after: always;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            
            .container {
                box-shadow: none;
                padding: 20px;
            }
            
            .kpi-card:hover {
                transform: none;
            }
            
            .section {
                page-break-inside: avoid;
            }
        }
    </style>
    """
    
    # Información de metadatos y filtros
    filtros_aplicados = []
    if cat_sel != "(Todas)":
        filtros_aplicados.append(f"Categoría: {cat_sel}")
    if mes_sel != "(Todos)":
        filtros_aplicados.append(f"Mes: {mes_sel}")
    
    filtros_texto = " | ".join(filtros_aplicados) if filtros_aplicados else "Sin filtros aplicados (vista completa)"
    
    # Construir HTML
    html_content = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Reporte Matriz de Obligaciones - MADR</title>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        {css_styles}
    </head>
    <body>
        <div class="container">
            <!-- HEADER -->
            <div class="header">
                <h1>📊 Reporte - Matriz de Obligaciones Contractuales</h1>
                <p>Ministerio de Agricultura y Desarrollo Rural (MADR)</p>
            </div>
            
            <!-- METADATA -->
            <div class="metadata">
                <p><strong>📅 Fecha de generación:</strong> {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
                <p><strong>🔍 Filtros aplicados:</strong> {filtros_texto}</p>
                <p><strong>🎯 Meta de cumplimiento:</strong> {meta}%</p>
                <p><strong>📈 Total de registros:</strong> {len(df_view):,}</p>
            </div>
            
            <!-- KPIs -->
            <div class="section">
                <h2 class="section-title">📌 Indicadores Clave de Desempeño (KPIs)</h2>
                <div class="kpi-section">
                    <div class="kpi-card">
                        <div class="kpi-label">Último Mes con Datos</div>
                        <div class="kpi-value">{kpis.get('ultimo_mes', '—')}</div>
                        <div class="kpi-description">Período más reciente reportado</div>
                    </div>
                    <div class="kpi-card">
                        <div class="kpi-label">Avance Global</div>
                        <div class="kpi-value">{f"{kpis.get('avance_global'):.1f}%" if kpis.get('avance_global') is not None else '—'}</div>
                        <div class="kpi-description">Promedio general de cumplimiento</div>
                    </div>
                    <div class="kpi-card">
                        <div class="kpi-label">Obligaciones en Riesgo</div>
                        <div class="kpi-value" style="color: #E53935;">{kpis.get('obligaciones_en_riesgo', 0)}</div>
                        <div class="kpi-description">Avance inferior al 70%</div>
                    </div>
                    <div class="kpi-card">
                        <div class="kpi-label">Mejor Mes</div>
                        <div class="kpi-value" style="color: #43A047;">{kpis.get('mejor_mes', '—')}</div>
                        <div class="kpi-description">Mayor avance registrado</div>
                    </div>
                </div>
            </div>
    """
    
    # Gráfico 1: Evolución mensual
    if not serie.empty:
        fig1_html = fig1.to_html(include_plotlyjs=False, div_id="grafico_evolucion")
        html_content += f"""
            <div class="section">
                <h2 class="section-title">📈 Evolución Mensual del Avance</h2>
                <div class="section-description">
                    <strong>¿Qué muestra este gráfico?</strong><br>
                    Presenta la tendencia del cumplimiento a lo largo del año, permitiendo identificar patrones estacionales, 
                    mejoras o deterioros en el desempeño. La línea punteada indica la meta establecida ({meta}%).
                </div>
                <div class="chart-container">
                    {fig1_html}
                </div>
                <div class="insight-box">
                    <strong>💡 Análisis:</strong> El promedio anual de cumplimiento es del {serie['Avance'].mean():.1f}%. 
                    {'✅ Se supera la meta en la mayoría de los meses.' if serie['Avance'].mean() >= meta else '⚠️ Se requiere atención para alcanzar la meta establecida.'}
                </div>
            </div>
        """
    
    # Gráfico 2: Avance por categoría
    if not cat_df.empty:
        fig2_html = fig2.to_html(include_plotlyjs=False, div_id="grafico_categoria")
        mejor_cat = cat_df.loc[cat_df['Avance'].idxmax(), 'Categoria']
        peor_cat = cat_df.loc[cat_df['Avance'].idxmin(), 'Categoria']
        
        html_content += f"""
            <div class="section">
                <h2 class="section-title">📊 Desempeño por Categoría</h2>
                <div class="section-description">
                    <strong>¿Qué muestra este gráfico?</strong><br>
                    Compara el nivel de cumplimiento entre las diferentes categorías de obligaciones para el {ultimo or 'último mes disponible'}. 
                    Las barras rojas indican categorías críticas (&lt;70%), naranjas en alerta (&lt;{meta}%), y verdes en cumplimiento (≥{meta}%).
                </div>
                <div class="chart-container">
                    {fig2_html}
                </div>
                <div class="insight-box">
                    <strong>💡 Análisis:</strong> 
                    <br>🏆 <strong>Mejor desempeño:</strong> {mejor_cat} con {cat_df.loc[cat_df['Avance'].idxmax(), 'Avance']:.1f}%
                    <br>⚠️ <strong>Requiere atención:</strong> {peor_cat} con {cat_df.loc[cat_df['Avance'].idxmin(), 'Avance']:.1f}%
                </div>
            </div>
        """
    
    # Gráfico 3: Heatmap de observaciones
    if not heat.empty:
        fig3_html = fig3.to_html(include_plotlyjs=False, div_id="grafico_heatmap")
        html_content += f"""
            <div class="section">
                <h2 class="section-title">🗺️ Mapa de Calor - Observaciones</h2>
                <div class="section-description">
                    <strong>¿Qué muestra este gráfico?</strong><br>
                    Visualiza la concentración de observaciones/incidencias por categoría y período. Los colores más intensos 
                    indican mayor cantidad de observaciones, señalando áreas que requieren seguimiento especial.
                </div>
                <div class="chart-container">
                    {fig3_html}
                </div>
                <div class="insight-box">
                    <strong>💡 Análisis:</strong> La categoría '{max_cat}' y el mes '{max_mes}' concentran la mayor cantidad de observaciones, 
                    lo que sugiere puntos críticos de atención para la gestión contractual.
                </div>
            </div>
            <div class="page-break"></div>
        """
    
    # Gráfico 4: Top obligaciones en riesgo
    if ultimo and not top_riesgo.empty:
        fig_top_html = fig_top.to_html(include_plotlyjs=False, div_id="grafico_top_riesgo")
        html_content += f"""
            <div class="section">
                <h2 class="section-title">🚨 Top {top_n} Obligaciones Críticas</h2>
                <div class="section-description">
                    <strong>¿Qué muestra este gráfico?</strong><br>
                    Identifica las obligaciones con menor porcentaje de avance que requieren acción inmediata. 
                    Estas son las prioridades de gestión para evitar incumplimientos contractuales.
                </div>
                <div class="chart-container">
                    {fig_top_html}
                </div>
                <div class="alert-box">
                    <strong>⚠️ Acción requerida:</strong> Se identificaron {len(top_riesgo)} obligaciones en situación crítica 
                    que requieren planes de acción inmediatos para alcanzar la meta del {meta}%.
                </div>
            </div>
        """
    
    # Gráfico 5: Caídas mes anterior
    if not caidas.empty:
        fig_delta_html = fig_delta.to_html(include_plotlyjs=False, div_id="grafico_caidas")
        html_content += f"""
            <div class="section">
                <h2 class="section-title">📉 Mayores Retrocesos vs Mes Anterior</h2>
                <div class="section-description">
                    <strong>¿Qué muestra este gráfico?</strong><br>
                    Muestra las obligaciones que experimentaron las mayores caídas en su porcentaje de avance comparado con 
                    el mes anterior ({prev} → {ultimo}). Los valores negativos indican retroceso en el cumplimiento.
                </div>
                <div class="chart-container">
                    {fig_delta_html}
                </div>
                <div class="alert-box">
                    <strong>⚠️ Tendencia negativa:</strong> Se detectaron {len(caidas)} obligaciones con caída significativa. 
                    Es necesario investigar las causas y establecer medidas correctivas.
                </div>
            </div>
        """
    
    # Gráfico 6: Pareto
    if not obs_cnt.empty:
        fig_pareto_html = fig_pareto.to_html(include_plotlyjs=False, div_id="grafico_pareto")
        cum_80 = obs_cnt[obs_cnt["%Acum"] <= 80]
        html_content += f"""
            <div class="section">
                <h2 class="section-title">📊 Análisis de Pareto (80/20)</h2>
                <div class="section-description">
                    <strong>¿Qué muestra este gráfico?</strong><br>
                    Aplica el principio de Pareto para identificar las obligaciones que concentran la mayoría de las observaciones. 
                    Típicamente, el 20% de las obligaciones generan el 80% de las incidencias. Permite priorizar esfuerzos de gestión.
                </div>
                <div class="chart-container">
                    {fig_pareto_html}
                </div>
                <div class="insight-box">
                    <strong>💡 Principio 80/20:</strong> Las primeras {len(cum_80)} obligaciones ({len(cum_80)/len(obs_cnt)*100:.1f}% del total) 
                    concentran el 80% de todas las observaciones registradas ({total_obs} en total). Focalizar recursos en estas áreas 
                    tendrá el mayor impacto en la gestión contractual.
                </div>
            </div>
            <div class="page-break"></div>
        """
    
    # Gráfico 7: Control estadístico
    if serie_cc is not None and not serie_cc.empty:
        fig_cc_html = fig_cc.to_html(include_plotlyjs=False, div_id="grafico_control")
        outliers_cc = serie_cc[(serie_cc["Avance"] > ucl) | (serie_cc["Avance"] < lcl)]
        
        html_content += f"""
            <div class="section">
                <h2 class="section-title">📉 Gráfico de Control Estadístico (SPC)</h2>
                <div class="section-description">
                    <strong>¿Qué muestra este gráfico?</strong><br>
                    Utiliza técnicas de control estadístico de procesos para detectar variaciones anormales en el cumplimiento. 
                    Los límites de control (±2σ) definen el rango esperado de variabilidad. Puntos fuera de estos límites 
                    señalan situaciones excepcionales que requieren investigación.
                </div>
                <div class="chart-container">
                    {fig_cc_html}
                </div>
        """
        
        if not outliers_cc.empty:
            html_content += f"""
                <div class="alert-box">
                    <strong>⚠️ Alerta de control:</strong> Se detectaron {len(outliers_cc)} mes(es) con variación anormal 
                    ({', '.join(outliers_cc['Mes'].tolist())}). Estas desviaciones requieren análisis de causas raíz.
                </div>
            """
        else:
            html_content += """
                <div class="success-box">
                    <strong>✅ Proceso bajo control:</strong> No se detectan variaciones estadísticamente anormales. 
                    El proceso mantiene estabilidad dentro de los límites esperados.
                </div>
            """
        
        html_content += """
            </div>
        """
    
    # Gráfico 8: Boxplot
    if not bx.empty:
        fig_box_html = fig_box.to_html(include_plotlyjs=False, div_id="grafico_boxplot")
        html_content += f"""
            <div class="section">
                <h2 class="section-title">📦 Distribución de Avances por Categoría</h2>
                <div class="section-description">
                    <strong>¿Qué muestra este gráfico?</strong><br>
                    Los diagramas de caja (boxplot) muestran la distribución estadística de los avances en cada categoría. 
                    Permiten identificar: la mediana (línea central), rango intercuartílico (caja), valores atípicos (puntos), 
                    y dispersión de datos. Útil para comparar la consistencia del desempeño entre categorías.
                </div>
                <div class="chart-container">
                    {fig_box_html}
                </div>
                <div class="insight-box">
                    <strong>💡 Análisis:</strong> Las categorías con cajas más estrechas muestran mayor consistencia en su desempeño, 
                    mientras que las cajas amplias indican alta variabilidad que requiere estandarización de procesos.
                </div>
            </div>
        """
    
    # Tabla resumen
    if not tabla_export.empty:
        tabla_tmp = tabla_export.copy()
        tabla_tmp["Estado"] = tabla_tmp["Avance"].apply(lambda v: semaforo(v, meta))
        tabla_html = tabla_tmp[["Categoria","Obligacion","Mes","Avance","Estado"]].head(20).to_html(
            index=False, 
            classes="tabla-datos",
            float_format=lambda x: f"{x:.1f}%" if pd.notna(x) else "—"
        )
        
        html_content += f"""
            <div class="section">
                <h2 class="section-title">📋 Detalle de Obligaciones (Top 20)</h2>
                <div class="section-description">
                    <strong>Vista detallada:</strong><br>
                    Listado de obligaciones con su estado actual. Para el reporte completo, consulte el archivo Excel descargable.
                </div>
                {tabla_html}
            </div>
        """
    
    # Footer
    html_content += f"""
            <!-- FOOTER -->
            <div class="footer">
                <p><strong>Ministerio de Agricultura y Desarrollo Rural (MADR)</strong></p>
                <p>Sistema de Seguimiento a Obligaciones Contractuales</p>
                <p>Reporte generado automáticamente el {pd.Timestamp.now().strftime('%d de %B de %Y a las %H:%M:%S')}</p>
                <p style="margin-top: 20px; font-size: 0.85em; color: #B0BEC5;">
                    Este documento contiene información de gestión contractual. 
                    Para exportar como PDF: Archivo → Imprimir → Guardar como PDF
                </p>
            </div>
        </div>
    </body>
    </html>
    """
    
    return html_content

# Generar y ofrecer descarga del reporte HTML mejorado
try:
    reporte_html = generar_reporte_html_completo()
    st.download_button(
        "📄 📊 Descargar Reporte Completo con Gráficas (HTML)",
        data=reporte_html.encode("utf-8"),
        file_name=f"Reporte_MatrizObligaciones_Completo_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.html",
        mime="text/html",
        help="Descarga un reporte HTML con todas las gráficas interactivas. Puedes abrirlo en el navegador e imprimirlo como PDF."
    )
    st.caption("💡 **Tip:** Abre el archivo HTML en tu navegador y usa 'Imprimir → Guardar como PDF' para obtener un reporte en PDF con todas las gráficas a color.")
except Exception as e:
    st.error(f"Error al generar el reporte HTML: {e}")
