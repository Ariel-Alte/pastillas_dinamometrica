"""
dashboard_edp.py — Tablero de Análisis de Frenos EDP
Streamlit · Línea Mitre / SOFSE — Gerencia de Coordinación de Material Rodante
"""

import re
import io
import tempfile
from pathlib import Path

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ─────────────────────────────────────────────
# CONSTANTES FÍSICAS Y DE DISEÑO
# ─────────────────────────────────────────────
MASA_KG         = 45_000
G               = 9.81
A_SERVICIO      = 1.0
A_EMERGENCIA    = 1.2
MU              = 0.35          # coef. fricción pastilla-disco (estimación conservadora)
PASTILLAS_TOTAL = 16            # 2 bogies × 4 discos × 2 pastillas
BALANCE_OK      = 10.0          # % umbral verde
BALANCE_WARN    = 20.0          # % umbral amarillo
CV_OK           = 5.0           # % CV repetibilidad verde
CV_WARN         = 10.0          # % CV repetibilidad amarillo

# Fuerza requerida por pastilla (distribución uniforme)
F_REQ_SERV_KGF  = (MASA_KG * A_SERVICIO  / G)
F_REQ_EMER_KGF  = (MASA_KG * A_EMERGENCIA / G)
F_MIN_PASTILLA_SERV = F_REQ_SERV_KGF  / MU / PASTILLAS_TOTAL
F_MIN_PASTILLA_EMER = F_REQ_EMER_KGF  / MU / PASTILLAS_TOTAL

# Mapeo UD → Rueda física
MAPEO_UD_RUEDA = {
    (1,1):2,(1,2):2,(1,3):1,(1,4):1,(1,5):4,(1,6):4,(1,7):3,(1,8):3,
    (2,1):6,(2,2):6,(2,3):5,(2,4):5,(2,5):8,(2,6):8,(2,7):7,(2,8):7,
}

# Colores corporativos
AZUL    = "#1F4E79"
AZUL_C  = "#2E75B6"
VERDE   = "#375623"
NARANJA = "#E26B0A"
ROJO    = "#C00000"
GRIS    = "#F2F2F2"

# ─────────────────────────────────────────────
# EXTRACCIÓN DE PDFs
# ─────────────────────────────────────────────

def parsear_encabezado(texto):
    r = {}
    for pat, key in [
        (r"Fecha:\s*(\d{1,2}\s+\w+\s+\d{4}\s+\d{2}:\d{2})", "fecha"),
        (r"Formación:\s*(\w+)", "formacion"),
        (r"Vehículo:\s*([\w\s]+?)(?:\n|Condición)", "vehiculo"),
        (r"Condición:\s*(\w+)", "condicion"),
        (r"Presión de Freno Aplicada:\s*([\d.,]+)\s*kg/cm", "presion"),
    ]:
        m = re.search(pat, texto)
        r[key] = m.group(1).strip() if m else ""
    return r

def parsear_nombre(nombre):
    stem = Path(nombre).stem
    r = {}
    m = re.search(r"_(TC1|TC2|M1|M2|M3|M4)_", stem, re.IGNORECASE)
    r["posicion"] = m.group(1).upper() if m else ""
    m = re.search(r"[_-]B(\d+)[_-]", stem)
    r["bogie"] = int(m.group(1)) if m else None
    m = re.search(r"[_-]B\d+[_-](\d+)[_-]", stem)
    r["nro_test"] = int(m.group(1)) if m else None
    m = re.search(r"_(SERV|EMER)$", stem, re.IGNORECASE)
    r["condicion_arch"] = m.group(1).upper() if m else ""
    return r

def procesar_pdfs(archivos):
    filas, errores = [], []
    for archivo in archivos:
        nombre = archivo.name
        try:
            with pdfplumber.open(archivo) as pdf:
                texto = pdf.pages[0].extract_text() or ""
            enc  = parsear_encabezado(texto)
            meta = parsear_nombre(nombre)
            patron = re.compile(r"(UD0?\d+\s+WIFI)\s+([\d]+)\s*kg\s+([\d]+)\s*kg\s+([\d]+)\s*kg")
            for m in patron.finditer(texto):
                ud_num = int(re.search(r"UD0?(\d+)", m.group(1)).group(1))
                label  = f"UD{ud_num:02d}"
                bogie  = meta["bogie"]
                rueda  = MAPEO_UD_RUEDA.get((bogie, ud_num)) if bogie else None
                sup, inf, tot = int(m.group(2)), int(m.group(3)), int(m.group(4))
                balance = abs(sup - inf) / ((sup + inf) / 2) * 100
                filas.append({
                    "archivo":       nombre,
                    "fecha":         enc["fecha"],
                    "formacion":     enc["formacion"],
                    "vehiculo":      enc["vehiculo"],
                    "condicion":     enc["condicion"],
                    "presion":       float(enc["presion"].replace(",", ".")) if enc["presion"] else None,
                    "posicion":      meta["posicion"],
                    "bogie":         bogie,
                    "nro_test":      meta["nro_test"],
                    "ud_num":        ud_num,
                    "ud_label":      label,
                    "rueda":         rueda,
                    "caliper":       "Superior" if ud_num % 2 != 0 else "Inferior",
                    "fuerza_sup":    sup,
                    "fuerza_inf":    inf,
                    "total":         tot,
                    "balance_pct":   round(balance, 2),
                    "estado_balance": "OK" if balance < BALANCE_OK else ("ATENCIÓN" if balance < BALANCE_WARN else "FUERA"),
                })
        except Exception as e:
            errores.append(f"❌ {nombre}: {e}")
    df = pd.DataFrame(filas) if filas else pd.DataFrame()
    return df, errores

# ─────────────────────────────────────────────
# HELPERS DE VISUALIZACIÓN
# ─────────────────────────────────────────────

def color_estado(estado):
    return {"OK": "#1a9641", "ATENCIÓN": "#f4a31e", "FUERA": ROJO}.get(estado, "#888")

def gauge_kpi(valor, titulo, sufijo="", referencia=None, color=AZUL_C):
    fig = go.Figure(go.Indicator(
        mode="number+delta" if referencia else "number",
        value=valor,
        number={"suffix": sufijo, "font": {"size": 32, "color": AZUL}},
        delta={"reference": referencia, "relative": False, "valueformat": ".1f"} if referencia else None,
        title={"text": titulo, "font": {"size": 13, "color": "#444"}},
    ))
    fig.update_layout(height=130, margin=dict(t=20, b=10, l=10, r=10),
                      paper_bgcolor="white")
    return fig

def semaforo_badge(estado):
    col = color_estado(estado)
    emoji = {"OK": "✅", "ATENCIÓN": "⚠️", "FUERA": "🔴"}.get(estado, "")
    return f'<span style="color:{col};font-weight:bold">{emoji} {estado}</span>'

# ─────────────────────────────────────────────
# SECCIONES DEL TABLERO
# ─────────────────────────────────────────────

def seccion_resumen(df):
    st.markdown("### 📊 Resumen Ejecutivo")

    cond_map = {"Servicio": "SERV", "Emergencia": "EMER",
                "servicio": "SERV", "emergencia": "EMER"}

    for cond_label, cond_key in [("Servicio", "Servicio"), ("Emergencia", "Emergencia")]:
        dfc = df[df["condicion"].str.lower() == cond_key.lower()]
        if dfc.empty:
            continue

        F_req = F_REQ_SERV_KGF if "serv" in cond_key.lower() else F_REQ_EMER_KGF
        a_ref = A_SERVICIO if "serv" in cond_key.lower() else A_EMERGENCIA
        F_min_past = F_MIN_PASTILLA_SERV if "serv" in cond_key.lower() else F_MIN_PASTILLA_EMER

        F_total_medida = dfc.groupby(["bogie", "nro_test"])["total"].sum().mean() * 2
        F_frenante = F_total_medida * MU
        eficiencia = F_frenante / F_req * 100
        bal_global = dfc["balance_pct"].mean()
        pct_ok = (dfc["estado_balance"] == "OK").mean() * 100

        with st.expander(f"🔵 **{cond_label}** — {a_ref} m/s² | Presión: {dfc['presion'].iloc[0]} kg/cm²", expanded=True):
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("Fuerza total medida", f"{F_total_medida:,.0f} kgf",
                          help="Suma de todas las pastillas, promedio de tests")
            with c2:
                st.metric("Fuerza frenante estimada (μ=0.35)", f"{F_frenante:,.0f} kgf",
                          help=f"Fuerza requerida: {F_req:,.0f} kgf")
            with c3:
                st.metric("Margen de frenado", f"{eficiencia:.1f}%",
                          delta=f"{eficiencia-100:.1f}% sobre requerido",
                          help="100% = exactamente lo requerido por diseño")
            with c4:
                st.metric("Calipers en rango", f"{pct_ok:.0f}%",
                          delta=f"Balance medio: {bal_global:.1f}%")


def seccion_balance(df):
    st.markdown("### ⚖️ Balance Superior / Inferior por Caliper")
    st.caption(
        "Índice de Balance = |F_sup − F_inf| / promedio × 100  |  "
        "✅ < 10%   ⚠️ 10–20%   🔴 > 20%"
    )

    condiciones = sorted(df["condicion"].unique())
    bogies      = sorted(df["bogie"].dropna().unique())

    tab_conds = st.tabs([c for c in condiciones])

    for tab, cond in zip(tab_conds, condiciones):
        with tab:
            dfc = df[df["condicion"] == cond]
            cols = st.columns(len(bogies))
            for col, bogie in zip(cols, bogies):
                with col:
                    st.markdown(f"**Bogie {int(bogie)}**")
                    dfb = dfc[dfc["bogie"] == bogie].groupby(["rueda", "ud_num"]).agg(
                        fuerza_sup=("fuerza_sup", "mean"),
                        fuerza_inf=("fuerza_inf", "mean"),
                        total=("total", "mean"),
                        balance_pct=("balance_pct", "mean"),
                        estado_balance=("estado_balance", lambda x: x.mode()[0]),
                    ).reset_index()

                    # Calcular disco por par de UDs consecutivos
                    discos = sorted(dfb["ud_num"].unique())
                    pares = [(discos[i], discos[i+1]) for i in range(0, len(discos)-1, 2)]

                    for ud_sup, ud_inf in pares:
                        row_s = dfb[dfb["ud_num"] == ud_sup].iloc[0]
                        row_i = dfb[dfb["ud_num"] == ud_inf].iloc[0]
                        balance = abs(row_s["fuerza_sup"] - row_i["fuerza_inf"]) / \
                                  ((row_s["fuerza_sup"] + row_i["fuerza_inf"]) / 2) * 100
                        disco_num = (ud_sup + 1) // 2
                        estado = "OK" if balance < BALANCE_OK else ("ATENCIÓN" if balance < BALANCE_WARN else "FUERA")
                        col_hex = color_estado(estado)
                        st.markdown(
                            f"<div style='border-left:4px solid {col_hex};"
                            f"padding:6px 10px;margin:4px 0;border-radius:4px;"
                            f"background:#fafafa'>"
                            f"<b>Disco {disco_num}</b> — Rueda {int(row_s['rueda'])}<br>"
                            f"↑ Sup: <b>{row_s['fuerza_sup']:.0f} kgf</b>  "
                            f"↓ Inf: <b>{row_i['fuerza_inf']:.0f} kgf</b><br>"
                            f"Balance: <b>{balance:.1f}%</b> {semaforo_badge(estado)}"
                            f"</div>",
                            unsafe_allow_html=True
                        )


def seccion_ruedas(df):
    st.markdown("### 🚂 Fuerza Total por Rueda")
    st.caption("Suma de pastilla superior + inferior por disco. Se espera uniformidad entre ruedas del mismo bogie y entre bogies.")

    condiciones = sorted(df["condicion"].unique())
    tab_conds = st.tabs([c for c in condiciones])

    for tab, cond in zip(tab_conds, condiciones):
        with tab:
            dfc = df[df["condicion"] == cond].groupby(
                ["bogie", "rueda"]
            )["total"].mean().reset_index()
            dfc["rueda_label"] = "Rueda " + dfc["rueda"].astype(int).astype(str)
            dfc["bogie_label"] = "Bogie " + dfc["bogie"].astype(int).astype(str)

            prom_global = dfc["total"].mean()

            fig = px.bar(
                dfc, x="rueda_label", y="total",
                color="bogie_label",
                color_discrete_map={"Bogie 1": AZUL_C, "Bogie 2": VERDE},
                barmode="group",
                text=dfc["total"].map(lambda x: f"{x:,.0f}"),
                labels={"total": "Fuerza por pastilla (kgf)", "rueda_label": "Rueda"},
                title=f"Fuerza por rueda — {cond}",
            )
            fig.add_hline(y=prom_global, line_dash="dash", line_color=ROJO,
                          annotation_text=f"Promedio global: {prom_global:,.0f} kgf",
                          annotation_position="top right")
            fig.update_traces(textposition="outside")
            fig.update_layout(
                height=420, plot_bgcolor="white",
                yaxis=dict(gridcolor="#e0e0e0"),
                legend_title="",
                font=dict(family="Arial"),
            )
            st.plotly_chart(fig, use_container_width=True)

            # Tabla de desvío de cada rueda vs promedio
            dfc["desvio_pct"] = (dfc["total"] - prom_global) / prom_global * 100
            dfc["estado"] = dfc["desvio_pct"].abs().apply(
                lambda x: "✅ OK" if x < 10 else ("⚠️ Atención" if x < 20 else "🔴 Fuera")
            )
            st.dataframe(
                dfc[["bogie_label", "rueda_label", "total", "desvio_pct", "estado"]]
                .rename(columns={
                    "bogie_label": "Bogie", "rueda_label": "Rueda",
                    "total": "Fuerza media (kgf)", "desvio_pct": "Desvío vs prom (%)",
                    "estado": "Estado"
                })
                .style.format({"Fuerza media (kgf)": "{:,.0f}", "Desvío vs prom (%)": "{:+.1f}%"}),
                use_container_width=True, hide_index=True
            )


def seccion_bogies(df):
    st.markdown("### 🔄 Comparación Bogie 1 vs Bogie 2")
    st.caption("Se espera un aporte similar de ambos bogies. Diferencias > 5% pueden indicar problemas en circuito neumático o desgaste diferencial.")

    resumen = df.groupby(["condicion", "bogie"]).agg(
        fuerza_total=("total", "sum"),
        fuerza_media_pastilla=("total", "mean"),
        balance_medio=("balance_pct", "mean"),
        n_calipers=("total", "count"),
    ).reset_index()

    # Normalizar por número de tests (promedio)
    tests = df.groupby(["condicion", "bogie", "nro_test"])["total"].sum().reset_index()
    tests_mean = tests.groupby(["condicion", "bogie"])["total"].mean().reset_index()
    tests_mean.columns = ["condicion", "bogie", "fuerza_bogie_mean"]

    for cond in sorted(df["condicion"].unique()):
        st.markdown(f"**{cond}**")
        dft = tests_mean[tests_mean["condicion"] == cond]
        if len(dft) < 2:
            continue
        b1 = dft[dft["bogie"] == 1]["fuerza_bogie_mean"].values[0]
        b2 = dft[dft["bogie"] == 2]["fuerza_bogie_mean"].values[0]
        diff_pct = abs(b1 - b2) / ((b1 + b2) / 2) * 100

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Bogie 1", f"{b1:,.0f} kgf")
        with c2:
            st.metric("Bogie 2", f"{b2:,.0f} kgf")
        with c3:
            estado = "✅ Uniforme" if diff_pct < 5 else ("⚠️ Diferencia" if diff_pct < 10 else "🔴 Investigar")
            st.metric("Diferencia B1 vs B2", f"{diff_pct:.1f}%", delta=estado)
        st.divider()


def seccion_repetibilidad(df):
    st.markdown("### 🔁 Repetibilidad entre Tests")
    st.caption(
        "Coeficiente de Variación (CV) = desvío estándar / promedio × 100  |  "
        "✅ CV < 5%  ⚠️ 5–10%  🔴 > 10%"
    )

    for cond in sorted(df["condicion"].unique()):
        st.markdown(f"**{cond}**")
        dfc = df[df["condicion"] == cond]

        rep = dfc.groupby(["bogie", "ud_num", "ud_label", "rueda"])["total"].agg(
            ["mean", "std", "count"]
        ).reset_index()
        rep.columns = ["bogie", "ud_num", "ud_label", "rueda", "media", "std", "n"]
        rep["cv_pct"] = rep["std"] / rep["media"] * 100
        rep["estado"] = rep["cv_pct"].apply(
            lambda x: "OK" if x < CV_OK else ("ATENCIÓN" if x < CV_WARN else "FUERA")
        )
        rep["bogie_label"] = "Bogie " + rep["bogie"].astype(int).astype(str)

        fig = px.bar(
            rep, x="ud_label", y="cv_pct",
            color="estado",
            color_discrete_map={"OK": "#1a9641", "ATENCIÓN": "#f4a31e", "FUERA": ROJO},
            facet_col="bogie_label",
            text=rep["cv_pct"].map(lambda x: f"{x:.1f}%"),
            labels={"cv_pct": "CV (%)", "ud_label": "Unidad dinamométrica"},
            title=f"Coeficiente de Variación entre tests — {cond}",
        )
        fig.add_hline(y=CV_OK,   line_dash="dash", line_color="#1a9641", line_width=1)
        fig.add_hline(y=CV_WARN, line_dash="dash", line_color=NARANJA,   line_width=1)
        fig.update_traces(textposition="outside")
        fig.update_layout(height=380, plot_bgcolor="white",
                          yaxis=dict(gridcolor="#e0e0e0"), font=dict(family="Arial"),
                          showlegend=False)
        st.plotly_chart(fig, use_container_width=True)


def seccion_serv_vs_emer(df):
    st.markdown("### 📈 Servicio vs Emergencia")
    st.caption("Mayor presión de freno → mayor fuerza de apriete. Se verifica que la progresión sea coherente con la presión aplicada.")

    if df["condicion"].nunique() < 2:
        st.info("Cargá PDFs de ambas condiciones (SERV y EMER) para activar esta comparación.")
        return

    comp = df.groupby(["condicion", "ud_num", "ud_label"])["total"].mean().reset_index()

    fig = px.line(
        comp, x="ud_label", y="total", color="condicion",
        color_discrete_map={"Servicio": AZUL_C, "Emergencia": ROJO},
        markers=True,
        labels={"total": "Fuerza media (kgf)", "ud_label": "UD", "condicion": "Condición"},
        title="Fuerza media por UD — Servicio vs Emergencia",
    )
    fig.update_layout(height=380, plot_bgcolor="white",
                      yaxis=dict(gridcolor="#e0e0e0"), font=dict(family="Arial"))
    st.plotly_chart(fig, use_container_width=True)

    # Ratio emergencia/servicio por UD
    pivot = comp.pivot(index="ud_label", columns="condicion", values="total").reset_index()
    if "Servicio" in pivot.columns and "Emergencia" in pivot.columns:
        pivot["ratio"] = pivot["Emergencia"] / pivot["Servicio"]
        st.caption(f"Ratio Emergencia/Servicio esperado: {A_EMERGENCIA/A_SERVICIO:.2f}×  (por diseño de presión)")
        fig2 = px.bar(pivot, x="ud_label", y="ratio",
                      text=pivot["ratio"].map(lambda x: f"{x:.2f}×"),
                      labels={"ratio": "Ratio Emer/Serv", "ud_label": "UD"},
                      title="Ratio de fuerza Emergencia / Servicio por UD",
                      color_discrete_sequence=[AZUL_C])
        fig2.add_hline(y=A_EMERGENCIA/A_SERVICIO, line_dash="dot", line_color=ROJO,
                       annotation_text=f"Esperado: {A_EMERGENCIA/A_SERVICIO:.2f}×")
        fig2.update_traces(textposition="outside")
        fig2.update_layout(height=340, plot_bgcolor="white",
                           yaxis=dict(gridcolor="#e0e0e0"), font=dict(family="Arial"))
        st.plotly_chart(fig2, use_container_width=True)


def seccion_datos(df):
    st.markdown("### 📋 Datos Completos")
    mostrar = df.rename(columns={
        "archivo": "Archivo", "fecha": "Fecha", "formacion": "Formación",
        "vehiculo": "Vehículo", "condicion": "Condición", "presion": "Presión (kg/cm²)",
        "posicion": "Posición", "bogie": "Bogie", "nro_test": "N° Test",
        "ud_label": "UD", "rueda": "Rueda", "caliper": "Caliper",
        "fuerza_sup": "F. Superior (kgf)", "fuerza_inf": "F. Inferior (kgf)",
        "total": "Total (kgf)", "balance_pct": "Balance (%)", "estado_balance": "Estado",
    })
    st.dataframe(mostrar.drop(columns=["ud_num"]), use_container_width=True, hide_index=True)

    # Descarga Excel
    def generar_excel(df_in):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "EDP Análisis"
        AZUL_X = "1F4E79"; GRIS_X = "D6E4F0"
        hf = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        hfill = PatternFill("solid", start_color=AZUL_X)
        thin = Side(style="thin", color="CCCCCC")
        borde = Border(left=thin, right=thin, top=thin, bottom=thin)
        cols = list(mostrar.drop(columns=["ud_num"]).columns)
        for c, h in enumerate(cols, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = hf; cell.fill = hfill
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = borde
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = max(12, len(h)+2)
        ws.row_dimensions[1].height = 30
        for i, row in enumerate(mostrar.drop(columns=["ud_num"]).itertuples(index=False), 2):
            alt = PatternFill("solid", start_color=GRIS_X) if i % 2 == 0 else None
            for c, val in enumerate(row, 1):
                cell = ws.cell(row=i, column=c, value=val)
                cell.font = Font(name="Arial", size=10)
                cell.border = borde
                if alt: cell.fill = alt
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(cols))}1"
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        wb.save(tmp.name)
        return tmp.name

    ruta = generar_excel(df)
    with open(ruta, "rb") as f:
        st.download_button(
            "📥 Descargar Excel",
            data=f.read(),
            file_name="edp_analisis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ─────────────────────────────────────────────
# LAYOUT PRINCIPAL
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Tablero EDP — Frenos Ferroviarios",
    page_icon="🚆",
    layout="wide",
)

# Header
st.markdown(
    f"""
    <div style='background:{AZUL};padding:18px 24px;border-radius:8px;margin-bottom:20px'>
      <h2 style='color:white;margin:0'>🚆 Tablero de Análisis — Ensayo Dinámico de Presión de Freno</h2>
      <p style='color:#cde;margin:4px 0 0'>
        Línea Mitre · SOFSE / Trenes Argentinos · Gerencia de Coordinación de Material Rodante
      </p>
    </div>
    """,
    unsafe_allow_html=True
)

# Sidebar — carga de archivos y parámetros
with st.sidebar:
    st.markdown(f"<h3 style='color:{AZUL}'>⚙️ Configuración</h3>", unsafe_allow_html=True)
    archivos = st.file_uploader(
        "Cargar PDFs EDP",
        type="pdf",
        accept_multiple_files=True,
        help="Formato: NNNNN-EDP-006-LM-FORM-VEH_POS_-B#_#_COND.pdf"
    )
    st.divider()
    st.markdown("**Parámetros del vehículo**")
    masa    = st.number_input("Masa (kg)", value=MASA_KG, step=500)
    a_serv  = st.number_input("Desaceleración Servicio (m/s²)", value=A_SERVICIO, step=0.1, format="%.2f")
    a_emer  = st.number_input("Desaceleración Emergencia (m/s²)", value=A_EMERGENCIA, step=0.1, format="%.2f")
    mu_val  = st.number_input("Coef. fricción μ (pastilla)", value=MU, step=0.01, format="%.2f",
                               help="Pastilla orgánica: ~0.35 | Sinterizada: ~0.32–0.40")
    st.divider()
    st.markdown("**Umbrales de balance S/I**")
    bal_ok   = st.slider("Verde → Amarillo (%)", 5, 25, int(BALANCE_OK))
    bal_warn = st.slider("Amarillo → Rojo (%)",  10, 40, int(BALANCE_WARN))

    st.divider()
    st.markdown(
        "<small style='color:#888'>"
        "Normativa de referencia:<br>UIC 541-3 · EN 15328 · UIC 544-1"
        "</small>",
        unsafe_allow_html=True
    )

# Procesar y mostrar
if not archivos:
    st.info("👈 Cargá los PDFs de ensayo en el panel lateral para comenzar el análisis.")
    st.markdown("""
    **¿Qué analiza este tablero?**
    - Fuerza de apriete de pastillas de freno por caliper (Superior / Inferior)
    - Balance entre sensores Superior e Inferior por disco
    - Comparación de fuerza entre ruedas y entre bogies
    - Repetibilidad entre tests del mismo ensayo
    - Comparación Servicio vs Emergencia con progresión esperada
    - Margen de eficiencia de frenado respecto a desaceleración de diseño
    """)
else:
    with st.spinner("Procesando PDFs..."):
        df, errores = procesar_pdfs(archivos)

    if errores:
        for e in errores:
            st.warning(e)

    if df.empty:
        st.error("No se pudieron extraer datos de los archivos cargados.")
    else:
        # Actualizar constantes con parámetros del sidebar
        F_req_s = (masa * a_serv / G)
        F_req_e = (masa * a_emer / G)
        F_min_s = F_req_s / mu_val / PASTILLAS_TOTAL
        F_min_e = F_req_e / mu_val / PASTILLAS_TOTAL

        st.success(f"✅ {len(df)} mediciones cargadas de {len(archivos)} archivo(s)")

        secciones = [
            "📊 Resumen ejecutivo",
            "⚖️ Balance S/I por caliper",
            "🚂 Fuerza por rueda",
            "🔄 Comparación de bogies",
            "🔁 Repetibilidad",
            "📈 Servicio vs Emergencia",
            "📋 Datos completos",
        ]
        tab_list = st.tabs(secciones)

        with tab_list[0]: seccion_resumen(df)
        with tab_list[1]: seccion_balance(df)
        with tab_list[2]: seccion_ruedas(df)
        with tab_list[3]: seccion_bogies(df)
        with tab_list[4]: seccion_repetibilidad(df)
        with tab_list[5]: seccion_serv_vs_emer(df)
        with tab_list[6]: seccion_datos(df)
