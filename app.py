"""
dashboard_edp.py — Tablero EDP + Contexto Operacional
Streamlit · Línea Mitre / SOFSE — Gerencia de Coordinación de Material Rodante
"""

import re, tempfile
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ─────────────────────────────────────────────
# CONSTANTES FIJAS DEL VEHÍCULO
# ─────────────────────────────────────────────
G               = 9.81
MU_DEFAULT      = 0.35

# Formación M27 — Línea Mitre
N_COCHES        = 6
MASA_TARA_COCHE = 45_000          # kg por coche
PLAZAS_MITRE    = 408             # 2×TC(60) + 4×M(72)
PLAZAS_SARMIENTO= 624
MASA_PAX_KG     = 75              # kg/pasajero (UIC)
PASTILLAS_X_COCHE = 16            # 2 bogies × 4 discos × 2 pastillas

MASA_TARA_FORM      = MASA_TARA_COCHE * N_COCHES
MASA_DISENO_FORM    = MASA_TARA_FORM + PLAZAS_MITRE * MASA_PAX_KG
MASA_PICO_FORM      = MASA_TARA_FORM + int(PLAZAS_MITRE * 1.5) * MASA_PAX_KG
MASA_SARMIENTO_FORM = MASA_TARA_FORM + PLAZAS_SARMIENTO * MASA_PAX_KG
N_PAST_TOTAL        = PASTILLAS_X_COCHE * N_COCHES   # 96

# Velocidades de ensayo EDP
V_SERV_KMH = 80
V_EMER_KMH = 30
TE_S       = 3.0    # tiempo respuesta equiv. estimado (s)

# Desaceleraciones de diseño del fabricante
A_DISEÑO_SERV = 1.0   # m/s²
A_DISEÑO_EMER = 1.2   # m/s²
A_CONFORT_MAX = 1.3   # m/s² — límite de confort (EN 12299 / norma europea)
A_EMERG_MAX   = 2.5   # m/s² — máximo en emergencia (EN 13452)

# Umbrales balance S/I
BALANCE_OK   = 10.0
BALANCE_WARN = 20.0
CV_OK        = 5.0
CV_WARN      = 10.0

# Colores — dark mode
AZUL   = "#1F4E79"
AZUL_C = "#4DA6FF"
VERDE  = "#5DBB63"
NARANJA= "#F4A31E"
ROJO   = "#FF4C4C"

BG_DARK   = "#0E1117"
BG_CARD   = "#1E2130"
TXT_WHITE = "#FAFAFA"

PLOTLY_LAYOUT = dict(
    plot_bgcolor =BG_DARK,
    paper_bgcolor=BG_DARK,
    font         =dict(family="Arial", color=TXT_WHITE),
    xaxis        =dict(gridcolor="#2A2F45", tickfont=dict(color=TXT_WHITE),
                       title_font=dict(color=TXT_WHITE)),
    yaxis        =dict(gridcolor="#2A2F45", tickfont=dict(color=TXT_WHITE),
                       title_font=dict(color=TXT_WHITE)),
    legend       =dict(bgcolor=BG_CARD, font=dict(color=TXT_WHITE)),
)

MAPEO_UD_RUEDA = {
    (1,1):2,(1,2):2,(1,3):1,(1,4):1,(1,5):4,(1,6):4,(1,7):3,(1,8):3,
    (2,1):6,(2,2):6,(2,3):5,(2,4):5,(2,5):8,(2,6):8,(2,7):7,(2,8):7,
}

# ─────────────────────────────────────────────
# FUNCIONES DE CÁLCULO FÍSICO
# ─────────────────────────────────────────────

def calcular_frenado(F_past_kgf, mu, masa_kg, v0_kmh):
    """Retorna (a m/s², distancia m) para frenado con te."""
    F_N = F_past_kgf * N_PAST_TOTAL * mu * G
    a   = F_N / masa_kg
    v0  = v0_kmh / 3.6
    s   = v0 * TE_S + v0**2 / (2 * a)
    return a, s

def estado_desaceleracion(a, cond):
    """Semáforo de desaceleración según condición."""
    if cond == "Servicio":
        if a < A_DISEÑO_SERV:      return "🔴 Bajo diseño"
        elif a <= A_CONFORT_MAX:   return "✅ OK"
        else:                       return "⚠️ Sobre confort"
    else:  # Emergencia
        if a < A_DISEÑO_EMER:      return "🔴 Bajo diseño"
        elif a <= A_EMERG_MAX:     return "✅ OK"
        else:                       return "⚠️ Excesivo"

# ─────────────────────────────────────────────
# CARGA DEL EXCEL
# ─────────────────────────────────────────────

def _extraer_bogie(nombre):
    m = re.search(r"[Bb](\d+)[_.]", str(nombre))
    return int(m.group(1)) if m else None

def _extraer_test(nombre):
    m = re.search(r"[Bb]\d+[_.](\d+)", str(nombre))
    return int(m.group(1)) if m else None

def cargar_excel(archivo):
    df = pd.read_excel(archivo, header=[0,1])
    df.columns = [c[1].strip() for c in df.columns]
    df = df.rename(columns={
        "Nombre de archivo":"archivo","Fecha ensayo":"fecha",
        "Formación":"formacion","Vehículo":"vehiculo",
        "Condición":"condicion","Presión aplicada (kg/cm²)":"presion",
        "Posición en formación":"posicion","Bogie":"bogie",
        "N° test":"nro_test","Unidad dinamométrica":"ud_label",
        "Fuerza Superior (kg)":"fuerza_sup","Fuerza Inferior (kg)":"fuerza_inf",
        "Total (kg)":"total",
    })
    df["ud_num"]   = df["ud_label"].str.extract(r"UD0?(\d+)").astype(int)
    df["bogie"]    = df.apply(lambda r: int(pd.to_numeric(r["bogie"], errors="coerce"))
                               if pd.notna(pd.to_numeric(r["bogie"], errors="coerce"))
                               else _extraer_bogie(r["archivo"]), axis=1)
    df["nro_test"] = df.apply(lambda r: int(pd.to_numeric(r["nro_test"], errors="coerce"))
                               if pd.notna(pd.to_numeric(r["nro_test"], errors="coerce"))
                               else _extraer_test(r["archivo"]), axis=1)
    df["rueda"]    = df.apply(lambda r: MAPEO_UD_RUEDA.get((r["bogie"], r["ud_num"]))
                               if r["bogie"] else None, axis=1)
    df["caliper"]  = df["ud_num"].apply(lambda n: "Superior" if n%2!=0 else "Inferior")
    df["balance_pct"] = (abs(df["fuerza_sup"]-df["fuerza_inf"])/
                         ((df["fuerza_sup"]+df["fuerza_inf"])/2)*100).round(2)
    df["condicion"] = df["condicion"].str.strip().str.title()
    df["estado_balance"] = df["balance_pct"].apply(
        lambda x: "OK" if x<BALANCE_OK else ("ATENCIÓN" if x<BALANCE_WARN else "FUERA"))
    return df

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def color_estado(e):
    return {"OK":"#1a9641","ATENCIÓN":"#f4a31e","FUERA":ROJO}.get(e,"#888")

# ─────────────────────────────────────────────
# SECCIÓN: CONTEXTO OPERACIONAL
# ─────────────────────────────────────────────

def seccion_contexto(df, mu):
    st.markdown("### 🚆 Contexto Operacional — Formación M27")

    # Fuerza promedio medida por condición
    promedios = df.groupby("condicion")["total"].mean()
    F_past_serv = promedios.get("Servicio",   1481.0)
    F_past_emer = promedios.get("Emergencia", 1718.0)

    # ── Info de la formación ──────────────────────────────
    st.markdown("#### Composición y masas")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Coches",              f"{N_COCHES}")
    c2.metric("Tara formación",      f"{MASA_TARA_FORM/1000:.0f} t")
    c3.metric("Con 408 pax (diseño)",f"{MASA_DISENO_FORM/1000:.1f} t")
    c4.metric("Pastillas totales",   f"{N_PAST_TOTAL}")

    st.caption(
        "Composición: TC1 (60 pl) · M1 · M2 · M3 · M4 (72 pl c/u) · TC2 (60 pl) = **408 plazas** · "
        f"μ utilizado: **{mu}** (ajustable en panel lateral)"
    )

    # ── Tabla de desaceleración por condición y carga ─────
    st.markdown("#### Desaceleración del freno neumático puro (sin regenerativo)")
    st.caption(
        "El EDP mide la capacidad del sistema **solo neumático**. "
        "El regenerativo es el freno principal para V > 20 km/h y garantiza los valores de diseño del fabricante. "
        "El neumático actúa como respaldo, emergencia y fase final (V ≤ 20 km/h)."
    )

    # Nota sobre desaceleración de diseño vs. calculada
    with st.info("**¿Por qué la desaceleración calculada supera el valor de diseño del fabricante?**"):
        st.markdown(f"""
El fabricante especifica **{A_DISEÑO_SERV} m/s²** (Servicio) y **{A_DISEÑO_EMER} m/s²** (Emergencia) 
como la desaceleración **del sistema combinado completo** (regenerativo + neumático), 
en condiciones normales de operación.

El freno neumático en sí está **sobredimensionado** por diseño, porque debe poder detener la 
formación por sí solo en el peor caso (emergencia total sin regenerativo). Esa reserva es la 
que explica los valores mayores a 1.0 m/s² que aparecen aquí.

- **V > 20 km/h normal:** el regenerativo regula la desaceleración a exactamente {A_DISEÑO_SERV} m/s². 
  El neumático no actúa o actúa levemente de apoyo.
- **V ≤ 20 km/h:** el neumático toma el control. A esa velocidad tan baja, 
  los valores altos de desaceleración no generan incomodidad perceptible.
- **Emergencia pura:** el neumático aplica su máxima capacidad. Los valores > {A_CONFORT_MAX} m/s² 
  están dentro de lo permitido (EN 13452: máx. {A_EMERG_MAX} m/s²) porque en emergencia 
  la prioridad es la detención, no el confort.
        """)

    filas = []
    for cond, F_past, v_kmh, a_dis in [
        ("Servicio (EDP)",   F_past_serv, V_SERV_KMH, A_DISEÑO_SERV),
        ("Emergencia (EDP)", F_past_emer, V_EMER_KMH, A_DISEÑO_EMER),
    ]:
        for label, masa_kg in [
            ("Tara (270 t)",       MASA_TARA_FORM),
            ("Diseño 408 pax (301 t)", MASA_DISENO_FORM),
            ("Pico ~612 pax (316 t)",  MASA_PICO_FORM),
        ]:
            a, s = calcular_frenado(F_past, mu, masa_kg, v_kmh)
            margen = a / a_dis
            cond_base = "Servicio" if "Serv" in cond else "Emergencia"
            estado = estado_desaceleracion(a, cond_base)
            filas.append({
                "Condición EDP": cond,
                "Carga": label,
                "a (m/s²)": round(a, 3),
                "Diseño (m/s²)": a_dis,
                "Margen": f"{margen:.2f}×",
                f"Dist. desde {v_kmh} km/h (m)": round(s, 1),
                "Estado": estado,
            })

    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

    # ── Gráfico de distancias ─────────────────────────────
    st.markdown("#### Distancia de parada — freno neumático puro")

    fig = go.Figure()
    cargas   = ["Tara\n(270 t)", "Diseño 408 pax\n(301 t)", "Pico\n(316 t)"]
    masas    = [MASA_TARA_FORM, MASA_DISENO_FORM, MASA_PICO_FORM]
    colores  = {("Servicio (EDP)", V_SERV_KMH): AZUL_C,
                ("Emergencia (EDP)", V_EMER_KMH): ROJO}

    for (cond, v_kmh), color in colores.items():
        F_past = F_past_serv if "Serv" in cond else F_past_emer
        distancias = [calcular_frenado(F_past, mu, m, v_kmh)[1] for m in masas]
        fig.add_trace(go.Bar(
            name=f"{cond} (V₀={v_kmh} km/h)",
            x=cargas, y=distancias,
            marker_color=color,
            text=[f"{d:.0f} m" for d in distancias],
            textposition="outside",
        ))

    fig.update_layout(**PLOTLY_LAYOUT, barmode="group", height=380, plot_bgcolor=BG_DARK,
        yaxis=dict(title="Distancia de parada (m)", gridcolor="#2A2F45"),
        xaxis_title="Condición de carga",
        font=dict(family="Arial"), legend_title="",
        title="Distancia de parada — freno neumático puro (con tₑ = 3 s)"
    )
    st.plotly_chart(fig, use_container_width=True)

    # ── Comparación Mitre vs Sarmiento ───────────────────
    st.markdown("#### Comparación operacional: Línea Mitre vs Línea Sarmiento")
    st.caption(
        "Mismo material rodante, diferente capacidad de pasajeros. "
        "A mayor carga, menor desaceleración disponible del freno neumático."
    )
    comp_data = []
    for linea, plazas, masa in [
        ("Mitre (408 pax)",     PLAZAS_MITRE,     MASA_DISENO_FORM),
        ("Sarmiento (624 pax)", PLAZAS_SARMIENTO, MASA_SARMIENTO_FORM),
    ]:
        for cond, F_past, v_kmh in [
            ("Servicio",   F_past_serv, V_SERV_KMH),
            ("Emergencia", F_past_emer, V_EMER_KMH),
        ]:
            a, s = calcular_frenado(F_past, mu, masa, v_kmh)
            comp_data.append({
                "Línea": linea, "Condición": cond,
                "Masa (t)": round(masa/1000, 1),
                "a (m/s²)": round(a, 3),
                f"Dist. (m)": round(s, 1),
            })
    st.dataframe(pd.DataFrame(comp_data), use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────
# SECCIONES DE MEDICIONES EDP
# ─────────────────────────────────────────────

def seccion_resumen(df, mu):
    st.markdown("### 📊 Resumen Ejecutivo — Mediciones EDP")
    for cond, a_dis, v_kmh in [
        ("Servicio",   A_DISEÑO_SERV, V_SERV_KMH),
        ("Emergencia", A_DISEÑO_EMER, V_EMER_KMH),
    ]:
        dfc = df[df["condicion"]==cond]
        if dfc.empty: continue
        F_past   = dfc["total"].mean()
        F_N      = F_past * N_PAST_TOTAL * mu * G
        F_req    = MASA_DISENO_FORM * a_dis
        margen   = F_N / F_req
        pct_ok   = (dfc["estado_balance"]=="OK").mean()*100
        presion  = dfc["presion"].iloc[0]
        with st.expander(
            f"**{cond}** — V₀: {v_kmh} km/h  |  Presión cilindro: {presion} kg/cm²  |  "
            f"Diseño fabricante: {a_dis} m/s²", expanded=True
        ):
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("F. apriete media/pastilla", f"{F_past:,.0f} kgf")
            c2.metric("F. frenante total estimada",
                      f"{F_N/G:,.0f} kgf",
                      help=f"F_apriete × {N_PAST_TOTAL} pastillas × μ={mu}")
            c3.metric("Margen s/ fuerza requerida", f"{margen:.2f}×",
                      delta=f"+{(margen-1)*100:.0f}% sobre lo necesario",
                      help=f"F requerida para {a_dis} m/s² con 301t: {F_req/G:,.0f} kgf")
            c4.metric("Calipers en rango (bal. S/I)", f"{pct_ok:.0f}%",
                      delta=f"Balance medio: {dfc['balance_pct'].mean():.1f}%")


def seccion_balance(df, bal_ok, bal_warn):
    st.markdown("### ⚖️ Balance Superior / Inferior por Caliper")
    st.caption("Balance = |F_sup − F_inf| / promedio × 100  ·  Ambas celdas actúan sobre el mismo disco: deberían ser iguales  |  ✅ < 10%   ⚠️ 10–20%   🔴 > 20%")
    tabs = st.tabs(sorted(df["condicion"].unique()))
    for tab, cond in zip(tabs, sorted(df["condicion"].unique())):
        with tab:
            dfc   = df[df["condicion"]==cond]
            bogs  = sorted(dfc["bogie"].dropna().unique())
            cols  = st.columns(len(bogs))
            for col, bogie in zip(cols, bogs):
                with col:
                    st.markdown(f"**Bogie {int(bogie)}**")
                    dfb = (dfc[dfc["bogie"]==bogie]
                           .groupby(["ud_num","rueda"])
                           .agg(fuerza_sup=("fuerza_sup","mean"),
                                fuerza_inf=("fuerza_inf","mean"))
                           .reset_index().sort_values("ud_num"))
                    uds   = dfb["ud_num"].tolist()
                    pares = [(uds[i],uds[i+1]) for i in range(0,len(uds)-1,2)]
                    for ud_s, ud_i in pares:
                        rs = dfb[dfb["ud_num"]==ud_s].iloc[0]
                        ri = dfb[dfb["ud_num"]==ud_i].iloc[0]
                        sup, inf = rs["fuerza_sup"], ri["fuerza_inf"]
                        bal    = abs(sup-inf)/((sup+inf)/2)*100
                        disco  = (ud_s+1)//2
                        estado = "OK" if bal<bal_ok else ("ATENCIÓN" if bal<bal_warn else "FUERA")
                        hx     = color_estado(estado)
                        emoji  = {"OK":"✅","ATENCIÓN":"⚠️","FUERA":"🔴"}[estado]
                        rueda  = int(rs["rueda"]) if rs["rueda"] else "?"
                        st.markdown(
                            f"<div style='border-left:4px solid {hx};padding:6px 10px;"
                            f"margin:5px 0;border-radius:4px;background:{BG_CARD};font-size:.9rem;color:{TXT_WHITE}'>"
                            f"<b>Disco {disco}</b> · Rueda {rueda}<br>"
                            f"↑ Sup: <b>{sup:.0f} kgf</b> &nbsp; ↓ Inf: <b>{inf:.0f} kgf</b><br>"
                            f"Balance: <b>{bal:.1f}%</b> &nbsp; {emoji} {estado}</div>",
                            unsafe_allow_html=True)


def seccion_ruedas(df):
    st.markdown("### 🚂 Fuerza Total por Rueda")
    st.caption("Promedio de ambos tests. Uniformidad entre ruedas = sistema bien regulado.")
    tabs = st.tabs(sorted(df["condicion"].unique()))
    for tab, cond in zip(tabs, sorted(df["condicion"].unique())):
        with tab:
            dfc  = df[df["condicion"]==cond].groupby(["bogie","rueda"])["total"].mean().reset_index()
            dfc["Rueda"] = "R"+dfc["rueda"].astype(int).astype(str)
            dfc["Bogie"] = "Bogie "+dfc["bogie"].astype(int).astype(str)
            prom = dfc["total"].mean()
            fig  = px.bar(dfc, x="Rueda", y="total", color="Bogie",
                          color_discrete_map={"Bogie 1":AZUL_C,"Bogie 2":VERDE},
                          barmode="group", text=dfc["total"].map(lambda x:f"{x:,.0f}"),
                          labels={"total":"Fuerza media (kgf)"}, title=f"Fuerza por rueda — {cond}")
            fig.add_hline(y=prom, line_dash="dash", line_color=ROJO,
                          annotation_text=f"Prom: {prom:,.0f} kgf", annotation_position="top right")
            fig.update_traces(textposition="outside")
            fig.update_layout(**PLOTLY_LAYOUT, height=420, plot_bgcolor=BG_DARK,
                              yaxis=dict(gridcolor="#2A2F45"), font=dict(family="Arial"), legend_title="")
            st.plotly_chart(fig, use_container_width=True)
            dfc["Desvío (%)"] = (dfc["total"]-prom)/prom*100
            dfc["Estado"] = dfc["Desvío (%)"].abs().apply(
                lambda x: "✅ OK" if x<10 else ("⚠️ Atención" if x<20 else "🔴 Fuera"))
            st.dataframe(
                dfc[["Bogie","Rueda","total","Desvío (%)","Estado"]]
                .rename(columns={"total":"Fuerza media (kgf)"})
                .style.format({"Fuerza media (kgf)":"{:,.0f}","Desvío (%)":"{:+.1f}%"}),
                use_container_width=True, hide_index=True)


def seccion_bogies(df):
    st.markdown("### 🔄 Comparación Bogie 1 vs Bogie 2")
    st.caption("Diferencia > 5% puede indicar problema neumático o desgaste diferencial.")
    medias = (df.groupby(["condicion","bogie","nro_test"])["total"].sum().reset_index()
                .groupby(["condicion","bogie"])["total"].mean().reset_index())
    medias.columns = ["condicion","bogie","fuerza"]
    for cond in sorted(df["condicion"].unique()):
        st.markdown(f"**{cond}**")
        dft = medias[medias["condicion"]==cond]
        if len(dft)<2: st.caption("Solo un bogie disponible."); continue
        b1   = dft[dft["bogie"]==1]["fuerza"].values[0]
        b2   = dft[dft["bogie"]==2]["fuerza"].values[0]
        diff = abs(b1-b2)/((b1+b2)/2)*100
        est  = "✅ Uniforme" if diff<5 else ("⚠️ Revisar" if diff<10 else "🔴 Investigar")
        c1,c2,c3 = st.columns(3)
        c1.metric("Bogie 1", f"{b1:,.0f} kgf")
        c2.metric("Bogie 2", f"{b2:,.0f} kgf")
        c3.metric("Diferencia", f"{diff:.1f}%", delta=est)
        st.divider()


def seccion_repetibilidad(df):
    st.markdown("### 🔁 Repetibilidad entre Tests")
    st.caption("CV = σ / media × 100  |  ✅ < 5%   ⚠️ 5–10%   🔴 > 10%")
    tabs = st.tabs(sorted(df["condicion"].unique()))
    for tab, cond in zip(tabs, sorted(df["condicion"].unique())):
        with tab:
            rep = (df[df["condicion"]==cond]
                   .groupby(["bogie","ud_num","ud_label"])["total"]
                   .agg(["mean","std"]).reset_index())
            rep["cv_pct"] = (rep["std"]/rep["mean"]*100).fillna(0)
            rep["estado"] = rep["cv_pct"].apply(
                lambda x:"OK" if x<CV_OK else("ATENCIÓN" if x<CV_WARN else"FUERA"))
            rep["Bogie"]  = "Bogie "+rep["bogie"].astype(int).astype(str)
            fig = px.bar(rep, x="ud_label", y="cv_pct", color="estado",
                         color_discrete_map={"OK":"#1a9641","ATENCIÓN":"#f4a31e","FUERA":ROJO},
                         facet_col="Bogie", text=rep["cv_pct"].map(lambda x:f"{x:.1f}%"),
                         labels={"cv_pct":"CV (%)","ud_label":"UD"}, title=f"Repetibilidad — {cond}")
            fig.add_hline(y=CV_OK,   line_dash="dash", line_color="#1a9641", line_width=1)
            fig.add_hline(y=CV_WARN, line_dash="dash", line_color=NARANJA,   line_width=1)
            fig.update_traces(textposition="outside")
            fig.update_layout(**PLOTLY_LAYOUT, height=380, plot_bgcolor=BG_DARK,
                              yaxis=dict(gridcolor="#2A2F45"),
                              font=dict(family="Arial"), showlegend=False)
            st.plotly_chart(fig, use_container_width=True)


def seccion_serv_vs_emer(df):
    st.markdown("### 📈 Servicio vs Emergencia")
    if df["condicion"].nunique()<2:
        st.info("El archivo debe contener datos de ambas condiciones."); return
    comp = df.groupby(["condicion","ud_num","ud_label"])["total"].mean().reset_index()
    fig  = px.line(comp, x="ud_label", y="total", color="condicion",
                   color_discrete_map={"Servicio":AZUL_C,"Emergencia":ROJO},
                   markers=True,
                   labels={"total":"Fuerza media (kgf)","ud_label":"UD","condicion":"Condición"},
                   title="Fuerza media por UD — Servicio vs Emergencia")
    fig.update_layout(**PLOTLY_LAYOUT, height=380, plot_bgcolor=BG_DARK,
                      yaxis=dict(gridcolor="#2A2F45"), font=dict(family="Arial"))
    st.plotly_chart(fig, use_container_width=True)
    pivot = comp.pivot(index="ud_label", columns="condicion", values="total").reset_index()
    if "Servicio" in pivot.columns and "Emergencia" in pivot.columns:
        pivot["ratio"] = pivot["Emergencia"]/pivot["Servicio"]
        fig2 = px.bar(pivot, x="ud_label", y="ratio",
                      text=pivot["ratio"].map(lambda x:f"{x:.2f}×"),
                      labels={"ratio":"Ratio Emer/Serv","ud_label":"UD"},
                      title="Ratio Emergencia / Servicio por UD",
                      color_discrete_sequence=[AZUL_C])
        ratio_presiones = 2.80/2.50
        fig2.add_hline(y=ratio_presiones, line_dash="dot", line_color=ROJO,
                       annotation_text=f"Ratio presiones: {ratio_presiones:.2f}×",
                       annotation_position="top right")
        fig2.update_traces(textposition="outside")
        fig2.update_layout(**PLOTLY_LAYOUT, height=360, plot_bgcolor=BG_DARK,
                           yaxis=dict(gridcolor="#2A2F45"), font=dict(family="Arial"))
        st.plotly_chart(fig2, use_container_width=True)
        st.caption(
            "El ratio de fuerza medida debería aproximarse al ratio de presiones aplicadas "
            f"(Emer {2.80} / Serv {2.50} = {ratio_presiones:.2f}×). "
            "Desvíos indican comportamiento no lineal del caliper o variación en la pastilla."
        )


def seccion_datos(df):
    st.markdown("### 📋 Datos Completos")
    mostrar = df.rename(columns={
        "archivo":"Archivo","fecha":"Fecha","formacion":"Formación","vehiculo":"Vehículo",
        "condicion":"Condición","presion":"Presión (kg/cm²)","posicion":"Posición",
        "bogie":"Bogie","nro_test":"N° Test","ud_label":"UD","rueda":"Rueda","caliper":"Caliper",
        "fuerza_sup":"F. Superior (kgf)","fuerza_inf":"F. Inferior (kgf)",
        "total":"Total (kgf)","balance_pct":"Balance (%)","estado_balance":"Estado",
    }).drop(columns=["ud_num"])
    st.dataframe(mostrar, use_container_width=True, hide_index=True)

    def to_excel(df_in):
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "EDP Análisis"
        thin = Side(style="thin",color="CCCCCC")
        borde = Border(left=thin,right=thin,top=thin,bottom=thin)
        hfill = PatternFill("solid",start_color="1F4E79")
        altfill = PatternFill("solid",start_color="D6E4F0")
        cols = list(df_in.columns)
        for c,h in enumerate(cols,1):
            cell = ws.cell(row=1,column=c,value=h)
            cell.font = Font(name="Arial",bold=True,color="FFFFFF",size=10)
            cell.fill = hfill
            cell.alignment = Alignment(horizontal="center",wrap_text=True)
            cell.border = borde
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = max(12,len(h)+2)
        ws.row_dimensions[1].height = 30
        for i,row in enumerate(df_in.itertuples(index=False),2):
            for c,val in enumerate(row,1):
                cell = ws.cell(row=i,column=c,value=val)
                cell.font = Font(name="Arial",size=10); cell.border = borde
                if i%2==0: cell.fill = altfill
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(cols))}1"
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False)
        wb.save(tmp.name); return tmp.name

    with open(to_excel(mostrar),"rb") as f:
        st.download_button("📥 Descargar Excel", data=f.read(),
                           file_name="edp_analisis.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ─────────────────────────────────────────────
# APP PRINCIPAL
# ─────────────────────────────────────────────


CSS_DARK = """
<style>
/* Tarjetas de balance */
div[data-testid="stMarkdownContainer"] div { color: #FAFAFA; }
/* Métricas */
[data-testid="stMetricValue"] { color: #FAFAFA; }
[data-testid="stMetricLabel"] { color: #A0A0B0; }
/* Tablas */
.stDataFrame { color: #FAFAFA; }
</style>
"""

st.set_page_config(page_title="Tablero EDP — Frenos M27", page_icon="🚆", layout="wide")
st.markdown(CSS_DARK, unsafe_allow_html=True)
st.markdown(
    f"<div style='background:{AZUL};padding:18px 24px;border-radius:8px;margin-bottom:20px'>"
    f"<h2 style='color:white;margin:0'>🚆 Tablero EDP — Ensayo Dinámico de Presión de Freno</h2>"
    f"<p style='color:#cde;margin:4px 0 0'>"
    f"Línea Mitre · Formación M27 · SOFSE / Trenes Argentinos · Coordinación de Material Rodante"
    f"</p></div>", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────
with st.sidebar:
    st.markdown(f"<h3 style='color:{AZUL}'>⚙️ Configuración</h3>", unsafe_allow_html=True)
    archivo = st.file_uploader("Cargar Excel EDP", type=["xlsx"],
                               help="Generado por extractor_edp.py")
    st.divider()

    st.markdown("**Parámetro de fricción**")
    mu_val = st.number_input(
        "Coef. fricción μ (pastilla-disco)",
        value=MU_DEFAULT, step=0.01, format="%.2f",
        help="Pastilla orgánica típica: 0.30–0.38 | Sinterizada: 0.30–0.40"
    )
    st.caption(
        f"Rango típico: 0.30–0.40  \n"
        f"El valor afecta los cálculos de fuerza frenante y desaceleración estimada."
    )
    st.divider()

    st.markdown("**Umbrales balance S/I**")
    bal_ok   = st.slider("Verde → Amarillo (%)",  5, 25, int(BALANCE_OK))
    bal_warn = st.slider("Amarillo → Rojo (%)",  10, 40, int(BALANCE_WARN))
    st.divider()

    st.markdown(f"""
    **Formación fija M27**
    - {N_COCHES} coches · {N_PAST_TOTAL} pastillas
    - Tara: {MASA_TARA_FORM/1000:.0f} t · Diseño: {MASA_DISENO_FORM/1000:.1f} t
    - V ensayo Serv.: {V_SERV_KMH} km/h
    - V ensayo Emer.: {V_EMER_KMH} km/h
    """)
    st.divider()
    st.markdown(
        "<small style='color:#888'>Ref.: UIC 541-3 · EN 15328 · UIC 544-1 · ETC FR v3.0 (AESF 2023)</small>",
        unsafe_allow_html=True)

# ── Contenido principal ───────────────────────
if not archivo:
    st.info("👈 Cargá el Excel generado por el extractor EDP para comenzar el análisis.")
    st.markdown(f"""
    **Sistema de freno — Formación M27 · Línea Mitre**

    | Parámetro | Valor |
    |---|---|
    | Coches | {N_COCHES} (TC1·M1·M2·M3·M4·TC2) |
    | Plazas | {PLAZAS_MITRE} |
    | Masa tara | {MASA_TARA_FORM/1000:.0f} t |
    | Masa con diseño de pax | {MASA_DISENO_FORM/1000:.1f} t |
    | Pastillas totales | {N_PAST_TOTAL} |
    | Freno principal (V > 20 km/h) | Regenerativo |
    | Freno de respaldo / emergencia | Neumático de caliper |
    | Velocidad ensayo EDP Servicio | {V_SERV_KMH} km/h |
    | Velocidad ensayo EDP Emergencia | {V_EMER_KMH} km/h |
    | a diseño Servicio | {A_DISEÑO_SERV} m/s² |
    | a diseño Emergencia | {A_DISEÑO_EMER} m/s² |
    | Límite de confort | {A_CONFORT_MAX} m/s² (norma europea) |
    """)
else:
    with st.spinner("Cargando datos..."):
        try:
            df = cargar_excel(archivo)
            df["estado_balance"] = df["balance_pct"].apply(
                lambda x: "OK" if x<bal_ok else ("ATENCIÓN" if x<bal_warn else "FUERA"))
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}"); st.stop()

    st.success(
        f"✅ {len(df)} mediciones · {df['archivo'].nunique()} archivos · "
        f"Condiciones: {', '.join(sorted(df['condicion'].unique()))} · "
        f"Vehículo: {df['vehiculo'].iloc[0]}"
    )

    tabs = st.tabs([
        "🚆 Contexto operacional",
        "📊 Resumen EDP",
        "⚖️ Balance S/I",
        "🚂 Por rueda",
        "🔄 Bogies",
        "🔁 Repetibilidad",
        "📈 Serv vs Emer",
        "📋 Datos",
    ])
    with tabs[0]: seccion_contexto(df, mu_val)
    with tabs[1]: seccion_resumen(df, mu_val)
    with tabs[2]: seccion_balance(df, bal_ok, bal_warn)
    with tabs[3]: seccion_ruedas(df)
    with tabs[4]: seccion_bogies(df)
    with tabs[5]: seccion_repetibilidad(df)
    with tabs[6]: seccion_serv_vs_emer(df)
    with tabs[7]: seccion_datos(df)
