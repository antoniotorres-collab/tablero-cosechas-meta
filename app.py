import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl
import os, re
from datetime import date, timedelta
import numpy as np

# ── configuración de página ───────────────────────────────────────────────────

st.set_page_config(
    page_title="Cosechas META",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 2rem; font-weight: 700; }
[data-testid="stMetricDelta"] { font-size: 0.85rem; }
.stTabs [data-baseweb="tab"]  { font-size: 0.95rem; font-weight: 600; }
div[data-testid="column"] > div { border-radius: 10px; }
.kpi-card { background:#f8f9fa; border-radius:12px; padding:16px 20px;
            border-left:4px solid #1F4E79; margin-bottom:8px; }
.alert-card { background:#FFF3CD; border-radius:8px; padding:10px 14px;
              border-left:4px solid #E26B0A; font-size:0.85rem; }
</style>
""", unsafe_allow_html=True)

# ── constantes ────────────────────────────────────────────────────────────────

DATA_DIR = os.path.join(os.path.dirname(__file__), "datos")
META     = 0.90
MESES    = {"Ene":1,"Feb":2,"Mar":3,"Abr":4,"May":5,"Jun":6,
            "Jul":7,"Ago":8,"Sep":9,"Oct":10,"Nov":11,"Dic":12}

VERDE   = "#375623"
AMARILLO= "#E07800"
ROJO    = "#C00000"
AZUL    = "#1F4E79"
GRIS    = "#595959"

def color_cosecha(v):
    if v is None or np.isnan(v): return "#E0E0E0"
    if v >= 0.90: return VERDE
    if v >= 0.80: return AMARILLO
    return ROJO

def semaforo(v):
    if v is None or np.isnan(v): return "⚪"
    if v >= 0.90: return "🟢"
    if v >= 0.80: return "🟡"
    return "🔴"

def tendencia_flecha(series):
    vals = [v for v in series if v is not None and not np.isnan(v)]
    if len(vals) < 2: return "—"
    diff = vals[-1] - vals[-2]
    if diff > 0.005:  return "↑"
    if diff < -0.005: return "↓"
    return "→"

# ── carga de datos (cached) ───────────────────────────────────────────────────

def parse_fecha(fname):
    name = re.sub(r"(CH|MC)$", "", fname.replace(".xlsx","").replace("Cosechas",""))
    m = re.match(r"^(\d+)([A-Za-z]+)(\d+)$", name)
    if not m: return None
    day, mes_str, yr = int(m.group(1)), m.group(2), 2000 + int(m.group(3))
    for k, v in MESES.items():
        if mes_str.lower() == k.lower(): return date(yr, v, day)
    return None

def col_label(d):
    meses_es = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
                7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
    return f"{d.day:02d}-{meses_es[d.month]}-{str(d.year)[2:]}"

def get_cols(header):
    col = {}
    for i, h in enumerate(header):
        if h in ("Clave Asesor","Clave_Asesor_Actual"):   col["id"]       = i
        elif h in ("Asesor","Asesor_Actual"):             col["nombre"]   = i
        elif h in ("Regional","RegionalComercial"):       col["regional"] = i
        elif h == "Sucursal":                             col["sucursal"] = i
        elif h == "Cosecha":                              col["cosecha"]  = i
        elif h in ("Estatus","Estatus_Asesor_Actual"):    col["estatus"]  = i
    return col

def load_sheet(filepath):
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb["Cosechas"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    col = get_cols(rows[0])
    data = {}
    for row in rows[1:]:
        if not any(v is not None for v in row): continue
        clave = row[col["id"]]
        if clave is None: continue
        clave = str(int(float(str(clave)))) if str(clave).replace(".","").isdigit() else str(clave).strip()
        data[clave] = {
            "nombre":   str(row[col["nombre"]] or "").strip(),
            "regional": str(row[col["regional"]] or "").strip(),
            "sucursal": str(row[col["sucursal"]] or "").strip(),
            "cosecha":  row[col["cosecha"]],
            "estatus":  str(row[col["estatus"]] or "") if "estatus" in col else "",
        }
    return data

@st.cache_data(ttl=300, show_spinner="Cargando datos...")
def load_all():
    all_files = [f for f in os.listdir(DATA_DIR) if f.endswith(".xlsx") and not f.startswith("~$")]

    today_d = date.today()
    current_week_mon = today_d - timedelta(days=today_d.weekday())

    weekly   = {}     # {week_monday: (file_date, filename, is_mc)} — un archivo por semana
    daily_now = None  # archivo más reciente no-lunes de la semana actual

    for f in all_files:
        d = parse_fecha(f)
        if d is None: continue
        is_mc     = "MC" in f.replace("Cosechas","").replace(".xlsx","")
        is_monday = (d.weekday() == 0)
        week_mon  = d - timedelta(days=d.weekday())

        # Representante semanal: lunes > no-lunes > más reciente > no-MC
        if week_mon not in weekly:
            weekly[week_mon] = (d, f, is_mc)
        else:
            cur, cur_is_mon = weekly[week_mon], (weekly[week_mon][0].weekday() == 0)
            replace = False
            if is_monday and not cur_is_mon:         replace = True
            elif is_monday == cur_is_mon:
                if d > cur[0]:                       replace = True
                elif d == cur[0] and cur[2] and not is_mc: replace = True
            if replace:
                weekly[week_mon] = (d, f, is_mc)

        # Columna "hoy": último no-lunes de la semana actual
        if week_mon == current_week_mon and not is_monday:
            if daily_now is None or d > daily_now[0]:
                daily_now = (d, f, is_mc)

    dated = {d: (f, mc) for _, (d, f, mc) in weekly.items()}
    if daily_now is not None and daily_now[0] not in dated:
        dated[daily_now[0]] = (daily_now[1], daily_now[2])

    ordered = sorted(dated.keys())
    all_data = {d: load_sheet(os.path.join(DATA_DIR, dated[d][0])) for d in ordered}

    base_date = max(ordered)
    base_data = all_data[base_date]
    activos   = {k: v for k, v in base_data.items() if v["estatus"].lower() == "activo"}

    rows = []
    for clave, info in activos.items():
        for d in ordered:
            ae = all_data[d].get(clave)
            cosecha = None
            if ae and ae["cosecha"] is not None:
                try: cosecha = float(ae["cosecha"])
                except: pass
            rows.append({
                "clave":    clave,
                "nombre":   info["nombre"],
                "regional": info["regional"],
                "sucursal": info["sucursal"],
                "semana":   col_label(d),
                "fecha":    d,
                "cosecha":  cosecha,
            })

    df = pd.DataFrame(rows)
    df["fecha"] = pd.to_datetime(df["fecha"])
    return df, [col_label(d) for d in ordered]

# ── cargar ────────────────────────────────────────────────────────────────────

df_full, semanas_ordenadas = load_all()

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — Filtros
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.image("https://img.icons8.com/color/96/bar-chart.png", width=48)
    st.title("Cosechas META")
    st.caption("Dashboard de evaluación de asesores")
    st.divider()

    # Regional
    regionales = sorted(df_full["regional"].unique())
    sel_reg = st.multiselect("Regional", regionales, placeholder="Todas las regionales")

    # Sucursal — cascada desde regional
    df_reg = df_full[df_full["regional"].isin(sel_reg)] if sel_reg else df_full
    sucursales = sorted(df_reg["sucursal"].unique())
    sel_suc = st.multiselect("Sucursal", sucursales, placeholder="Todas las sucursales")

    # Asesor — cascada desde sucursal
    df_suc = df_reg[df_reg["sucursal"].isin(sel_suc)] if sel_suc else df_reg
    asesores_opt = sorted(df_suc["nombre"].unique())
    sel_ase = st.multiselect("Asesor", asesores_opt, placeholder="Todos los asesores")

    st.divider()

    # Rango de semanas
    idx_min, idx_max = 0, len(semanas_ordenadas) - 1
    rango = st.select_slider(
        "Rango de semanas",
        options=semanas_ordenadas,
        value=(semanas_ordenadas[0], semanas_ordenadas[-1]),
    )
    i_ini = semanas_ordenadas.index(rango[0])
    i_fin = semanas_ordenadas.index(rango[1])
    semanas_sel = semanas_ordenadas[i_ini : i_fin + 1]

    st.divider()
    st.caption(f"📅 {semanas_sel[0]} → {semanas_sel[-1]}")
    st.caption(f"📊 {len(semanas_sel)} semanas")
    st.divider()
    if st.button("🔄 Recargar datos", use_container_width=True):
        load_all.clear()
        st.rerun()

# ── aplicar filtros ───────────────────────────────────────────────────────────

df = df_full.copy()
if sel_reg: df = df[df["regional"].isin(sel_reg)]
if sel_suc: df = df[df["sucursal"].isin(sel_suc)]
if sel_ase: df = df[df["nombre"].isin(sel_ase)]
df = df[df["semana"].isin(semanas_sel)]

# tabla wide (asesor × semana)
df_pivot = (df.pivot_table(index=["clave","nombre","regional","sucursal"],
                           columns="semana", values="cosecha", aggfunc="first")
              .reindex(columns=[s for s in semanas_sel if s in df["semana"].unique()])
              .reset_index())

asesores_visibles = df_pivot["clave"].tolist()
ultima_semana     = semanas_sel[-1] if semanas_sel[-1] in df.columns or semanas_sel else semanas_ordenadas[-1]

# ── KPIs globales ─────────────────────────────────────────────────────────────

df_ultima = df[df["semana"] == semanas_sel[-1]].dropna(subset=["cosecha"])
n_asesores   = df_pivot["clave"].nunique()
prom_actual  = df_ultima["cosecha"].mean() if not df_ultima.empty else 0
sobre_meta   = (df_ultima["cosecha"] >= META).sum()
pct_meta     = sobre_meta / len(df_ultima) * 100 if len(df_ultima) > 0 else 0
en_riesgo    = (df_ultima["cosecha"] < 0.80).sum()

# variación vs semana anterior
if len(semanas_sel) >= 2:
    df_ant   = df[df["semana"] == semanas_sel[-2]].dropna(subset=["cosecha"])
    prom_ant = df_ant["cosecha"].mean() if not df_ant.empty else prom_actual
    delta_prom = prom_actual - prom_ant
else:
    delta_prom = None

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════

col_t, col_s = st.columns([3, 1])
with col_t:
    filtro_texto = " · ".join(filter(None, [
        ", ".join(sel_reg) if sel_reg else "",
        ", ".join(sel_suc) if sel_suc else "",
        ", ".join(sel_ase) if sel_ase else "",
    ])) or "Portafolio completo"
    st.title(f"📊 {filtro_texto}")
    st.caption(f"Semanas: {semanas_sel[0]} → {semanas_sel[-1]}  |  {n_asesores} asesores activos")
with col_s:
    if en_riesgo > 0:
        st.markdown(f'<div class="alert-card">⚠️ <b>{en_riesgo} asesores</b> con cosecha &lt; 80%</div>',
                    unsafe_allow_html=True)

# KPIs
k1, k2, k3, k4 = st.columns(4)
k1.metric("Asesores activos",   n_asesores)
k2.metric("Promedio cosecha",   f"{prom_actual*100:.1f}%",
          f"{delta_prom*100:+.1f}pp" if delta_prom is not None else None)
k3.metric("Sobre meta (≥90%)",  f"{sobre_meta}",  f"{pct_meta:.0f}% del grupo")
k4.metric("En riesgo (<80%)",   f"{en_riesgo}",
          delta_color="inverse" if en_riesgo > 0 else "normal")

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# PESTAÑAS
# ══════════════════════════════════════════════════════════════════════════════

tab_port, tab_ranking, tab_regional, tab_detalle = st.tabs([
    "📈 Portafolio", "🏆 Ranking", "🗺 Por Regional", "🔍 Detalle Asesor"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — PORTAFOLIO
# ══════════════════════════════════════════════════════════════════════════════

with tab_port:

    # ── heatmap ──────────────────────────────────────────────────────────────

    st.subheader("Mapa de calor — Cosecha por asesor y semana")

    semanas_disp = [s for s in semanas_sel if s in df_pivot.columns]
    if semanas_disp and not df_pivot.empty:
        matrix_vals = df_pivot[semanas_disp].values.astype(float)
        y_labels    = [f"{row['clave']} · {row['nombre'][:26]}"
                       for _, row in df_pivot.iterrows()]

        text_matrix = []
        for row_v in matrix_vals:
            text_matrix.append([f"{v*100:.0f}%" if not np.isnan(v) else "" for v in row_v])

        fig_heat = go.Figure(go.Heatmap(
            z=matrix_vals,
            x=semanas_disp,
            y=y_labels,
            text=text_matrix,
            texttemplate="%{text}",
            textfont={"size": 9},
            colorscale=[
                [0.0,   "#C00000"],
                [0.5,   "#C00000"],
                [0.5,   "#E07800"],
                [0.75,  "#E07800"],
                [0.75,  "#375623"],
                [1.0,   "#375623"],
            ],
            zmin=0.60, zmax=1.0,
            colorbar=dict(tickformat=".0%", title="Cosecha"),
            hovertemplate="<b>%{y}</b><br>%{x}<br>Cosecha: %{z:.1%}<extra></extra>",
        ))
        fig_heat.update_layout(
            height=max(350, len(y_labels) * 22 + 80),
            margin=dict(l=0, r=0, t=10, b=0),
            xaxis=dict(tickangle=-40, side="top"),
            yaxis=dict(autorange="reversed"),
            plot_bgcolor="white",
        )
        st.plotly_chart(fig_heat, use_container_width=True)

    # ── evolución del promedio ────────────────────────────────────────────────

    st.subheader("Evolución semanal")
    c_evo1, c_evo2 = st.columns([3, 1])

    with c_evo1:
        prom_serie = (df.groupby("semana", sort=False)["cosecha"]
                       .mean()
                       .reindex(semanas_sel)
                       .reset_index())
        prom_serie.columns = ["semana", "cosecha"]

        fig_evo = go.Figure()

        # líneas individuales (finas, gris)
        for clave, grp in df.groupby("clave"):
            grp_s = grp.set_index("semana").reindex(semanas_sel)
            fig_evo.add_trace(go.Scatter(
                x=semanas_sel, y=grp_s["cosecha"],
                mode="lines",
                line=dict(color="rgba(180,180,180,0.35)", width=1),
                showlegend=False,
                hoverinfo="skip",
            ))

        # promedio grupal (gruesa)
        fig_evo.add_trace(go.Scatter(
            x=prom_serie["semana"], y=prom_serie["cosecha"],
            mode="lines+markers+text",
            name="Promedio",
            line=dict(color=AZUL, width=3),
            marker=dict(size=6),
            text=[f"{v*100:.1f}%" if pd.notna(v) else "" for v in prom_serie["cosecha"]],
            textposition="top center",
            textfont=dict(size=9, color=AZUL),
            hovertemplate="Semana: %{x}<br>Promedio: %{y:.1%}<extra></extra>",
        ))

        # línea meta
        fig_evo.add_hline(y=META, line_dash="dash", line_color=ROJO,
                          annotation_text="Meta 90%", annotation_position="right",
                          annotation_font_color=ROJO)

        fig_evo.update_layout(
            height=380, margin=dict(l=0, r=0, t=10, b=0),
            yaxis=dict(tickformat=".0%", range=[0.5, 1.05], title="Cosecha"),
            xaxis=dict(tickangle=-40),
            legend=dict(orientation="h", y=1.08),
            plot_bgcolor="white", paper_bgcolor="white",
            hovermode="x unified",
        )
        st.plotly_chart(fig_evo, use_container_width=True)

    with c_evo2:
        # donut de tendencias
        totales = {"Subió": 0, "Se mantuvo": 0, "Bajó": 0}
        for _, row in df_pivot.iterrows():
            vals = [row[s] for s in semanas_disp if s in row.index and pd.notna(row[s])]
            if len(vals) < 2: continue
            diff = vals[-1] - vals[0]
            if diff > 0.005:    totales["Subió"] += 1
            elif diff < -0.005: totales["Bajó"] += 1
            else:               totales["Se mantuvo"] += 1

        fig_don = go.Figure(go.Pie(
            labels=list(totales.keys()),
            values=list(totales.values()),
            hole=0.55,
            marker=dict(colors=[VERDE, AMARILLO, ROJO],
                        line=dict(color="white", width=2)),
            textinfo="label+value",
            hovertemplate="%{label}: %{value} asesores<extra></extra>",
        ))
        fig_don.add_annotation(text=f"<b>{n_asesores}</b><br>asesores",
                               x=0.5, y=0.5, showarrow=False, font_size=13)
        fig_don.update_layout(
            height=260, margin=dict(l=10, r=10, t=30, b=10),
            showlegend=True,
            legend=dict(orientation="h", y=-0.1),
            title=dict(text="Tendencia general", font_size=12, x=0.5),
        )
        st.plotly_chart(fig_don, use_container_width=True)

        # distribución última semana
        vals_dist = df_ultima["cosecha"].dropna().tolist()
        if vals_dist:
            fig_hist = go.Figure(go.Histogram(
                x=vals_dist, nbinsx=15,
                marker_color=AZUL, opacity=0.8,
                hovertemplate="Cosecha: %{x:.0%}<br>Asesores: %{y}<extra></extra>",
            ))
            fig_hist.add_vline(x=META, line_dash="dash", line_color=ROJO,
                               annotation_text="Meta", annotation_font_color=ROJO)
            fig_hist.update_layout(
                height=200, margin=dict(l=0, r=0, t=30, b=0),
                xaxis=dict(tickformat=".0%", title=""),
                yaxis=dict(title="Asesores"),
                title=dict(text=f"Distribución ({semanas_sel[-1]})", font_size=12, x=0.5),
                plot_bgcolor="white",
            )
            st.plotly_chart(fig_hist, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — RANKING
# ══════════════════════════════════════════════════════════════════════════════

with tab_ranking:
    st.subheader("Tabla de ranking — Asesores")

    # construir tabla de ranking
    ranking_rows = []
    for _, row in df_pivot.iterrows():
        vals = [row[s] for s in semanas_disp if s in row.index and pd.notna(row[s])]
        if not vals:
            continue
        cos_actual = vals[-1]
        cos_ini    = vals[0]
        variacion  = cos_actual - cos_ini
        tendencia  = tendencia_flecha(vals)
        sem_flag   = semaforo(cos_actual)

        # tendencia últimas 3 semanas
        ultimas3 = vals[-3:] if len(vals) >= 3 else vals
        tend_str = tendencia

        ranking_rows.append({
            "": sem_flag,
            "Clave":    row["clave"],
            "Asesor":   row["nombre"],
            "Sucursal": row["sucursal"],
            "Regional": row["regional"],
            f"Ini ({semanas_sel[0]})": cos_ini,
            f"Actual ({semanas_sel[-1]})": cos_actual,
            "Variación": variacion,
            "Tend.": tend_str,
            "_cos": cos_actual,
        })

    df_rank = pd.DataFrame(ranking_rows)
    if not df_rank.empty:
        df_rank = df_rank.sort_values("_cos", ascending=False).drop(columns=["_cos"])
        df_rank.index = range(1, len(df_rank) + 1)

        # formatos
        fmt_cols = {
            f"Ini ({semanas_sel[0]})": "{:.1%}",
            f"Actual ({semanas_sel[-1]})": "{:.1%}",
            "Variación": "{:+.1%}",
        }

        def color_row(row):
            v = row[f"Actual ({semanas_sel[-1]})"]
            if v >= META:    c = "#C6EFCE"
            elif v >= 0.80:  c = "#FFD7A0"
            else:            c = "#FFC7CE"
            return [f"background-color: {c}" if i > 3 else "" for i in range(len(row))]

        st.dataframe(
            df_rank.style
                   .format(fmt_cols)
                   .apply(color_row, axis=1),
            use_container_width=True,
            height=min(600, len(df_rank) * 38 + 50),
        )

        # top/bottom barras
        st.divider()
        c_top, c_bot = st.columns(2)

        with c_top:
            st.markdown("#### ⬆ Top 10 — mayor variación positiva")
            top10 = sorted(ranking_rows, key=lambda x: x["Variación"], reverse=True)[:10]
            fig_top = go.Figure(go.Bar(
                x=[r["Variación"] for r in top10],
                y=[f"{r['Clave']} {r['Asesor'].split()[0]}" for r in top10],
                orientation="h",
                marker_color=VERDE,
                text=[f"+{r['Variación']*100:.1f}%" for r in top10],
                textposition="outside",
                hovertemplate="%{y}<br>Variación: %{x:.1%}<extra></extra>",
            ))
            fig_top.update_layout(height=320, margin=dict(l=0,r=60,t=10,b=0),
                                  xaxis=dict(tickformat=".0%"),
                                  plot_bgcolor="white", yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig_top, use_container_width=True)

        with c_bot:
            st.markdown("#### ⬇ Top 10 — mayor caída")
            bot10 = sorted(ranking_rows, key=lambda x: x["Variación"])[:10]
            fig_bot = go.Figure(go.Bar(
                x=[r["Variación"] for r in bot10],
                y=[f"{r['Clave']} {r['Asesor'].split()[0]}" for r in bot10],
                orientation="h",
                marker_color=ROJO,
                text=[f"{r['Variación']*100:.1f}%" for r in bot10],
                textposition="outside",
                hovertemplate="%{y}<br>Variación: %{x:.1%}<extra></extra>",
            ))
            fig_bot.update_layout(height=320, margin=dict(l=0,r=60,t=10,b=0),
                                  xaxis=dict(tickformat=".0%"),
                                  plot_bgcolor="white", yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig_bot, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — POR REGIONAL
# ══════════════════════════════════════════════════════════════════════════════

with tab_regional:
    regionales_disp = sorted(df["regional"].unique())
    if not regionales_disp:
        st.info("No hay datos para el filtro actual.")
    else:
        for reg in regionales_disp:
            df_reg_tab = df[df["regional"] == reg]
            asesores_reg = df_reg_tab["clave"].unique()
            n_reg = len(asesores_reg)
            prom_reg_act = df_reg_tab[df_reg_tab["semana"] == semanas_sel[-1]]["cosecha"].mean()

            with st.expander(f"{'🟢' if prom_reg_act >= META else '🔴'} {reg}  —  {n_reg} asesores  |  Promedio actual: {prom_reg_act*100:.1f}%", expanded=True):
                fig_reg = go.Figure()

                for clave in asesores_reg:
                    grp = df_reg_tab[df_reg_tab["clave"] == clave].set_index("semana").reindex(semanas_sel)
                    nombre = df_reg_tab[df_reg_tab["clave"] == clave]["nombre"].iloc[0]
                    ultima_v = grp["cosecha"].dropna().iloc[-1] if not grp["cosecha"].dropna().empty else None
                    line_color = "rgba(55,86,35,0.6)" if ultima_v and ultima_v >= META else \
                                 "rgba(224,120,0,0.6)" if ultima_v and ultima_v >= 0.80 else \
                                 "rgba(192,0,0,0.6)"
                    fig_reg.add_trace(go.Scatter(
                        x=semanas_sel, y=grp["cosecha"],
                        mode="lines+markers", name=nombre,
                        line=dict(color=line_color, width=1.5),
                        marker=dict(size=4),
                        hovertemplate=f"<b>{nombre}</b><br>%{{x}}: %{{y:.1%}}<extra></extra>",
                    ))

                # promedio regional
                prom_r = (df_reg_tab.groupby("semana")["cosecha"]
                           .mean().reindex(semanas_sel))
                fig_reg.add_trace(go.Scatter(
                    x=semanas_sel, y=prom_r,
                    mode="lines+markers", name="★ Promedio regional",
                    line=dict(color=AZUL, width=3.5, dash="dot"),
                    marker=dict(size=7, symbol="diamond"),
                    hovertemplate="Promedio: %{y:.1%}<extra></extra>",
                ))
                fig_reg.add_hline(y=META, line_dash="dash", line_color=ROJO,
                                  annotation_text="Meta 90%",
                                  annotation_font_color=ROJO)
                fig_reg.update_layout(
                    height=320, margin=dict(l=0, r=0, t=10, b=0),
                    yaxis=dict(tickformat=".0%", range=[0.5, 1.05]),
                    xaxis=dict(tickangle=-40),
                    legend=dict(orientation="h", y=-0.25, font_size=10),
                    plot_bgcolor="white", paper_bgcolor="white",
                    hovermode="x unified",
                )
                st.plotly_chart(fig_reg, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — DETALLE ASESOR
# ══════════════════════════════════════════════════════════════════════════════

with tab_detalle:
    asesores_lista = sorted(df["nombre"].unique())
    asesor_sel = st.selectbox("Selecciona un asesor", asesores_lista,
                               index=0 if asesores_lista else None)

    if asesor_sel:
        df_a = df[df["nombre"] == asesor_sel].set_index("semana").reindex(semanas_sel).reset_index()
        info_a = df[df["nombre"] == asesor_sel].iloc[0]

        # meta-datos del asesor
        mc1, mc2, mc3 = st.columns(3)
        mc1.metric("Regional",  info_a["regional"])
        mc2.metric("Sucursal",  info_a["sucursal"])
        mc3.metric("Clave",     info_a["clave"])

        st.divider()

        vals_a = df_a["cosecha"].dropna()
        if not vals_a.empty:
            cos_actual_a = vals_a.iloc[-1]
            cos_ini_a    = vals_a.iloc[0]
            cos_max_a    = vals_a.max()
            cos_min_a    = vals_a.min()
            variacion_a  = cos_actual_a - cos_ini_a

            ka1, ka2, ka3, ka4 = st.columns(4)
            ka1.metric("Cosecha actual",  f"{cos_actual_a*100:.1f}%",
                       f"{variacion_a*100:+.1f}pp vs inicio")
            ka2.metric("Mejor semana",    f"{cos_max_a*100:.1f}%")
            ka3.metric("Peor semana",     f"{cos_min_a*100:.1f}%")
            ka4.metric("Semanas en riesgo", str((vals_a < 0.80).sum()))

        st.divider()

        # ── gráfica principal del asesor ──────────────────────────────────────

        df_suc_a   = df[df["sucursal"] == info_a["sucursal"]]
        prom_suc_a = df_suc_a.groupby("semana")["cosecha"].mean().reindex(semanas_sel)
        df_reg_a   = df[df["regional"] == info_a["regional"]]
        prom_reg_a = df_reg_a.groupby("semana")["cosecha"].mean().reindex(semanas_sel)

        fig_det = go.Figure()

        # área del asesor
        fig_det.add_trace(go.Scatter(
            x=df_a["semana"], y=df_a["cosecha"],
            mode="lines+markers", name=asesor_sel,
            line=dict(color=AZUL, width=3),
            marker=dict(size=8),
            fill="tozeroy", fillcolor="rgba(31,78,121,0.08)",
            hovertemplate="%{x}<br>Cosecha: %{y:.1%}<extra></extra>",
        ))

        # promedio sucursal
        fig_det.add_trace(go.Scatter(
            x=semanas_sel, y=prom_suc_a,
            mode="lines", name=f"Promedio {info_a['sucursal']}",
            line=dict(color=VERDE, width=2, dash="dash"),
            hovertemplate="Sucursal: %{y:.1%}<extra></extra>",
        ))

        # promedio regional
        fig_det.add_trace(go.Scatter(
            x=semanas_sel, y=prom_reg_a,
            mode="lines", name=f"Promedio {info_a['regional']}",
            line=dict(color=AMARILLO, width=2, dash="dot"),
            hovertemplate="Regional: %{y:.1%}<extra></extra>",
        ))

        fig_det.add_hline(y=META, line_dash="dash", line_color=ROJO,
                          annotation_text="Meta 90%", annotation_font_color=ROJO)

        fig_det.update_layout(
            height=400, margin=dict(l=0, r=0, t=10, b=0),
            yaxis=dict(tickformat=".0%", range=[0.5, 1.05], title="Cosecha"),
            xaxis=dict(tickangle=-40),
            legend=dict(orientation="h", y=1.08),
            plot_bgcolor="white", paper_bgcolor="white",
            hovermode="x unified",
        )
        st.plotly_chart(fig_det, use_container_width=True)

        # ── posición en sucursal ──────────────────────────────────────────────

        st.subheader(f"Posición dentro de {info_a['sucursal']}")
        df_suc_rank = df_suc_a[df_suc_a["semana"] == semanas_sel[-1]].dropna(subset=["cosecha"])
        df_suc_rank = df_suc_rank.sort_values("cosecha", ascending=False).reset_index(drop=True)
        df_suc_rank.index += 1

        pos = df_suc_rank[df_suc_rank["nombre"] == asesor_sel].index[0] if asesor_sel in df_suc_rank["nombre"].values else "—"
        st.caption(f"Posición: **#{pos}** de {len(df_suc_rank)} en {info_a['sucursal']}")

        fig_rank_suc = go.Figure(go.Bar(
            x=df_suc_rank["cosecha"],
            y=df_suc_rank["nombre"].str.split().str[0] + " " + df_suc_rank["nombre"].str.split().str[-1],
            orientation="h",
            marker_color=[AZUL if n == asesor_sel else "rgba(191,191,191,0.7)"
                          for n in df_suc_rank["nombre"]],
            text=df_suc_rank["cosecha"].apply(lambda v: f"{v*100:.1f}%"),
            textposition="outside",
            hovertemplate="%{y}: %{x:.1%}<extra></extra>",
        ))
        fig_rank_suc.add_vline(x=META, line_dash="dash", line_color=ROJO)
        fig_rank_suc.update_layout(
            height=max(250, len(df_suc_rank) * 28 + 60),
            margin=dict(l=0, r=60, t=10, b=0),
            xaxis=dict(tickformat=".0%", range=[0, 1.1]),
            plot_bgcolor="white",
            yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig_rank_suc, use_container_width=True)

        # ── tabla semana a semana ─────────────────────────────────────────────

        st.subheader("Histórico semana a semana")
        df_tabla = df_a[["semana","cosecha"]].copy()
        df_tabla.columns = ["Semana", "Cosecha"]
        df_tabla["Meta"] = META
        df_tabla["vs Meta"] = df_tabla["Cosecha"] - META
        df_tabla["Semáforo"] = df_tabla["Cosecha"].apply(semaforo)

        def color_cosecha_df(val):
            if pd.isna(val): return ""
            if val >= META:   return "background-color: #C6EFCE"
            if val >= 0.80:   return "background-color: #FFEB9C"
            return "background-color: #FFC7CE"

        st.dataframe(
            df_tabla.style
                    .format({"Cosecha": "{:.1%}", "Meta": "{:.0%}", "vs Meta": "{:+.1%}"})
                    .map(color_cosecha_df, subset=["Cosecha"]),
            use_container_width=True,
            height=min(500, len(df_tabla) * 36 + 50),
        )
