"""
dashboard_edp.py — Tablero de Análisis de Frenos EDP
Streamlit · Línea Mitre / SOFSE — Gerencia de Coordinación de Material Rodante
Entrada: Excel generado por extractor_edp.py
"""

import re, tempfile
import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

MASA_KG=45_000; G=9.81; A_SERVICIO=1.0; A_EMERGENCIA=1.2; MU=0.35; PASTILLAS_TOTAL=16
BALANCE_OK=10.0; BALANCE_WARN=20.0; CV_OK=5.0; CV_WARN=10.0
AZUL="#1F4E79"; AZUL_C="#2E75B6"; VERDE="#375623"; NARANJA="#E26B0A"; ROJO="#C00000"

MAPEO_UD_RUEDA = {
    (1,1):2,(1,2):2,(1,3):1,(1,4):1,(1,5):4,(1,6):4,(1,7):3,(1,8):3,
    (2,1):6,(2,2):6,(2,3):5,(2,4):5,(2,5):8,(2,6):8,(2,7):7,(2,8):7,
}

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
        "Nombre de archivo":"archivo","Fecha ensayo":"fecha","Formación":"formacion",
        "Vehículo":"vehiculo","Condición":"condicion","Presión aplicada (kg/cm²)":"presion",
        "Posición en formación":"posicion","Bogie":"bogie","N° test":"nro_test",
        "Unidad dinamométrica":"ud_label","Fuerza Superior (kg)":"fuerza_sup",
        "Fuerza Inferior (kg)":"fuerza_inf","Total (kg)":"total",
    })
    df["ud_num"]  = df["ud_label"].str.extract(r"UD0?(\d+)").astype(int)
    # Bogie y test: desde columna si existe, sino desde nombre de archivo
    df["bogie"]   = df.apply(lambda r: int(pd.to_numeric(r["bogie"],   errors="coerce"))
                              if pd.notna(pd.to_numeric(r["bogie"],   errors="coerce"))
                              else _extraer_bogie(r["archivo"]), axis=1)
    df["nro_test"]= df.apply(lambda r: int(pd.to_numeric(r["nro_test"],errors="coerce"))
                              if pd.notna(pd.to_numeric(r["nro_test"],errors="coerce"))
                              else _extraer_test(r["archivo"]), axis=1)
    df["rueda"]   = df.apply(lambda r: MAPEO_UD_RUEDA.get((r["bogie"], r["ud_num"])) if r["bogie"] else None, axis=1)
    df["caliper"] = df["ud_num"].apply(lambda n: "Superior" if n%2!=0 else "Inferior")
    df["balance_pct"] = (abs(df["fuerza_sup"]-df["fuerza_inf"]) / ((df["fuerza_sup"]+df["fuerza_inf"])/2)*100).round(2)
    df["condicion"]   = df["condicion"].str.strip().str.title()
    df["estado_balance"] = df["balance_pct"].apply(lambda x: "OK" if x<BALANCE_OK else ("ATENCIÓN" if x<BALANCE_WARN else "FUERA"))
    return df

def color_estado(e):
    return {"OK":"#1a9641","ATENCIÓN":"#f4a31e","FUERA":ROJO}.get(e,"#888")

# ── SECCIONES ────────────────────────────────

def seccion_resumen(df, masa, a_serv, a_emer, mu):
    st.markdown("### 📊 Resumen Ejecutivo")
    for cond, a_ref in [("Servicio",a_serv),("Emergencia",a_emer)]:
        dfc = df[df["condicion"]==cond]
        if dfc.empty: continue
        F_req   = masa*a_ref/G
        F_total = dfc.groupby(["bogie","nro_test"])["total"].sum().mean()*2
        F_fren  = F_total*mu; margen=F_fren/F_req*100
        pct_ok  = (dfc["estado_balance"]=="OK").mean()*100
        presion = dfc["presion"].iloc[0]
        with st.expander(f"**{cond}** — {a_ref} m/s²  |  Presión: {presion} kg/cm²", expanded=True):
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Fuerza total apriete",         f"{F_total:,.0f} kgf")
            c2.metric(f"Fuerza frenante est. (μ={mu})",f"{F_fren:,.0f} kgf", help=f"Requerida: {F_req:,.0f} kgf")
            c3.metric("Margen sobre diseño",           f"{margen:.1f}%", delta=f"+{margen-100:.1f}% sobre lo requerido")
            c4.metric("Calipers en rango (bal.)",      f"{pct_ok:.0f}%", delta=f"Balance medio: {dfc['balance_pct'].mean():.1f}%")


def seccion_balance(df, bal_ok, bal_warn):
    st.markdown("### ⚖️ Balance Superior / Inferior por Caliper")
    st.caption("Balance = |F_sup − F_inf| / promedio × 100  ·  Ambas celdas actúan sobre el mismo disco: deberían ser iguales  |  ✅ < 10%   ⚠️ 10–20%   🔴 > 20%")
    tabs = st.tabs(sorted(df["condicion"].unique()))
    for tab, cond in zip(tabs, sorted(df["condicion"].unique())):
        with tab:
            dfc  = df[df["condicion"]==cond]
            bogs = sorted(dfc["bogie"].dropna().unique())
            cols = st.columns(len(bogs))
            for col, bogie in zip(cols, bogs):
                with col:
                    st.markdown(f"**Bogie {int(bogie)}**")
                    dfb = (dfc[dfc["bogie"]==bogie]
                           .groupby(["ud_num","rueda"])
                           .agg(fuerza_sup=("fuerza_sup","mean"),fuerza_inf=("fuerza_inf","mean"))
                           .reset_index().sort_values("ud_num"))
                    uds   = dfb["ud_num"].tolist()
                    pares = [(uds[i],uds[i+1]) for i in range(0,len(uds)-1,2)]
                    for ud_s, ud_i in pares:
                        rs = dfb[dfb["ud_num"]==ud_s].iloc[0]
                        ri = dfb[dfb["ud_num"]==ud_i].iloc[0]
                        sup,inf = rs["fuerza_sup"], ri["fuerza_inf"]
                        bal   = abs(sup-inf)/((sup+inf)/2)*100
                        disco = (ud_s+1)//2
                        estado= "OK" if bal<bal_ok else ("ATENCIÓN" if bal<bal_warn else "FUERA")
                        hx    = color_estado(estado)
                        emoji = {"OK":"✅","ATENCIÓN":"⚠️","FUERA":"🔴"}[estado]
                        rueda = int(rs["rueda"]) if rs["rueda"] else "?"
                        st.markdown(
                            f"<div style='border-left:4px solid {hx};padding:6px 10px;"
                            f"margin:5px 0;border-radius:4px;background:#fafafa;font-size:.9rem'>"
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
            dfc = df[df["condicion"]==cond].groupby(["bogie","rueda"])["total"].mean().reset_index()
            dfc["Rueda"] = "R"+dfc["rueda"].astype(int).astype(str)
            dfc["Bogie"] = "Bogie "+dfc["bogie"].astype(int).astype(str)
            prom = dfc["total"].mean()
            fig = px.bar(dfc,x="Rueda",y="total",color="Bogie",
                         color_discrete_map={"Bogie 1":AZUL_C,"Bogie 2":VERDE},
                         barmode="group",text=dfc["total"].map(lambda x:f"{x:,.0f}"),
                         labels={"total":"Fuerza media (kgf)"},title=f"Fuerza por rueda — {cond}")
            fig.add_hline(y=prom,line_dash="dash",line_color=ROJO,
                          annotation_text=f"Prom: {prom:,.0f} kgf",annotation_position="top right")
            fig.update_traces(textposition="outside")
            fig.update_layout(height=420,plot_bgcolor="white",
                              yaxis=dict(gridcolor="#e0e0e0"),font=dict(family="Arial"),legend_title="")
            st.plotly_chart(fig,use_container_width=True)
            dfc["Desvío (%)"] = (dfc["total"]-prom)/prom*100
            dfc["Estado"] = dfc["Desvío (%)"].abs().apply(lambda x:"✅ OK" if x<10 else("⚠️ Atención" if x<20 else"🔴 Fuera"))
            st.dataframe(
                dfc[["Bogie","Rueda","total","Desvío (%)","Estado"]]
                .rename(columns={"total":"Fuerza media (kgf)"})
                .style.format({"Fuerza media (kgf)":"{:,.0f}","Desvío (%)":"{:+.1f}%"}),
                use_container_width=True,hide_index=True)


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
        b1 = dft[dft["bogie"]==1]["fuerza"].values[0]
        b2 = dft[dft["bogie"]==2]["fuerza"].values[0]
        diff = abs(b1-b2)/((b1+b2)/2)*100
        estado = "✅ Uniforme" if diff<5 else ("⚠️ Revisar" if diff<10 else "🔴 Investigar")
        c1,c2,c3 = st.columns(3)
        c1.metric("Bogie 1",f"{b1:,.0f} kgf"); c2.metric("Bogie 2",f"{b2:,.0f} kgf")
        c3.metric("Diferencia",f"{diff:.1f}%",delta=estado)
        st.divider()


def seccion_repetibilidad(df):
    st.markdown("### 🔁 Repetibilidad entre Tests")
    st.caption("CV = σ / media × 100  ·  Estabilidad del equipo al repetir el ensayo  |  ✅ < 5%   ⚠️ 5–10%   🔴 > 10%")
    tabs = st.tabs(sorted(df["condicion"].unique()))
    for tab, cond in zip(tabs, sorted(df["condicion"].unique())):
        with tab:
            rep = (df[df["condicion"]==cond]
                   .groupby(["bogie","ud_num","ud_label"])["total"]
                   .agg(["mean","std"]).reset_index())
            rep["cv_pct"]  = (rep["std"]/rep["mean"]*100).fillna(0)
            rep["estado"]  = rep["cv_pct"].apply(lambda x:"OK" if x<CV_OK else("ATENCIÓN" if x<CV_WARN else"FUERA"))
            rep["Bogie"]   = "Bogie "+rep["bogie"].astype(int).astype(str)
            fig = px.bar(rep,x="ud_label",y="cv_pct",color="estado",
                         color_discrete_map={"OK":"#1a9641","ATENCIÓN":"#f4a31e","FUERA":ROJO},
                         facet_col="Bogie",text=rep["cv_pct"].map(lambda x:f"{x:.1f}%"),
                         labels={"cv_pct":"CV (%)","ud_label":"UD"},title=f"Repetibilidad — {cond}")
            fig.add_hline(y=CV_OK,  line_dash="dash",line_color="#1a9641",line_width=1)
            fig.add_hline(y=CV_WARN,line_dash="dash",line_color=NARANJA,  line_width=1)
            fig.update_traces(textposition="outside")
            fig.update_layout(height=380,plot_bgcolor="white",
                              yaxis=dict(gridcolor="#e0e0e0"),font=dict(family="Arial"),showlegend=False)
            st.plotly_chart(fig,use_container_width=True)


def seccion_serv_vs_emer(df):
    st.markdown("### 📈 Servicio vs Emergencia")
    if df["condicion"].nunique()<2:
        st.info("El archivo debe contener datos de ambas condiciones."); return
    comp = df.groupby(["condicion","ud_num","ud_label"])["total"].mean().reset_index()
    fig = px.line(comp,x="ud_label",y="total",color="condicion",
                  color_discrete_map={"Servicio":AZUL_C,"Emergencia":ROJO},markers=True,
                  labels={"total":"Fuerza media (kgf)","ud_label":"UD","condicion":"Condición"},
                  title="Fuerza media por UD — Servicio vs Emergencia")
    fig.update_layout(height=380,plot_bgcolor="white",yaxis=dict(gridcolor="#e0e0e0"),font=dict(family="Arial"))
    st.plotly_chart(fig,use_container_width=True)
    pivot = comp.pivot(index="ud_label",columns="condicion",values="total").reset_index()
    if "Servicio" in pivot.columns and "Emergencia" in pivot.columns:
        pivot["ratio"] = pivot["Emergencia"]/pivot["Servicio"]
        fig2 = px.bar(pivot,x="ud_label",y="ratio",
                      text=pivot["ratio"].map(lambda x:f"{x:.2f}×"),
                      labels={"ratio":"Ratio Emer/Serv","ud_label":"UD"},
                      title="Ratio Emergencia / Servicio por UD",
                      color_discrete_sequence=[AZUL_C])
        fig2.add_hline(y=A_EMERGENCIA/A_SERVICIO,line_dash="dot",line_color=ROJO,
                       annotation_text=f"Teórico: {A_EMERGENCIA/A_SERVICIO:.2f}×",
                       annotation_position="top right")
        fig2.update_traces(textposition="outside")
        fig2.update_layout(height=360,plot_bgcolor="white",
                           yaxis=dict(gridcolor="#e0e0e0"),font=dict(family="Arial"))
        st.plotly_chart(fig2,use_container_width=True)


def seccion_datos(df):
    st.markdown("### 📋 Datos Completos")
    mostrar = df.rename(columns={
        "archivo":"Archivo","fecha":"Fecha","formacion":"Formación","vehiculo":"Vehículo",
        "condicion":"Condición","presion":"Presión (kg/cm²)","posicion":"Posición",
        "bogie":"Bogie","nro_test":"N° Test","ud_label":"UD","rueda":"Rueda","caliper":"Caliper",
        "fuerza_sup":"F. Superior (kgf)","fuerza_inf":"F. Inferior (kgf)","total":"Total (kgf)",
        "balance_pct":"Balance (%)","estado_balance":"Estado",
    }).drop(columns=["ud_num"])
    st.dataframe(mostrar,use_container_width=True,hide_index=True)
    def to_excel(df_in):
        wb=openpyxl.Workbook(); ws=wb.active; ws.title="EDP Análisis"
        thin=Side(style="thin",color="CCCCCC"); borde=Border(left=thin,right=thin,top=thin,bottom=thin)
        hfill=PatternFill("solid",start_color="1F4E79"); altfill=PatternFill("solid",start_color="D6E4F0")
        cols=list(df_in.columns)
        for c,h in enumerate(cols,1):
            cell=ws.cell(row=1,column=c,value=h)
            cell.font=Font(name="Arial",bold=True,color="FFFFFF",size=10)
            cell.fill=hfill; cell.alignment=Alignment(horizontal="center",wrap_text=True); cell.border=borde
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width=max(12,len(h)+2)
        ws.row_dimensions[1].height=30
        for i,row in enumerate(df_in.itertuples(index=False),2):
            for c,val in enumerate(row,1):
                cell=ws.cell(row=i,column=c,value=val)
                cell.font=Font(name="Arial",size=10); cell.border=borde
                if i%2==0: cell.fill=altfill
        ws.freeze_panes="A2"
        ws.auto_filter.ref=f"A1:{openpyxl.utils.get_column_letter(len(cols))}1"
        tmp=tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False); wb.save(tmp.name); return tmp.name
    with open(to_excel(mostrar),"rb") as f:
        st.download_button("📥 Descargar Excel",data=f.read(),file_name="edp_analisis.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── APP ──────────────────────────────────────

st.set_page_config(page_title="Tablero EDP",page_icon="🚆",layout="wide")
st.markdown(
    f"<div style='background:{AZUL};padding:18px 24px;border-radius:8px;margin-bottom:20px'>"
    f"<h2 style='color:white;margin:0'>🚆 Tablero EDP — Frenos Ferroviarios</h2>"
    f"<p style='color:#cde;margin:4px 0 0'>Línea Mitre · SOFSE / Trenes Argentinos · Coordinación de Material Rodante</p>"
    f"</div>",unsafe_allow_html=True)

with st.sidebar:
    st.markdown(f"<h3 style='color:{AZUL}'>⚙️ Configuración</h3>",unsafe_allow_html=True)
    archivo = st.file_uploader("Cargar Excel EDP",type=["xlsx"],help="Generado por extractor_edp.py")
    st.divider()
    st.markdown("**Parámetros del vehículo**")
    masa   = st.number_input("Masa (kg)",                        value=MASA_KG,      step=500)
    a_serv = st.number_input("Desaceleración Servicio (m/s²)",   value=A_SERVICIO,   step=0.1,format="%.2f")
    a_emer = st.number_input("Desaceleración Emergencia (m/s²)", value=A_EMERGENCIA, step=0.1,format="%.2f")
    mu_val = st.number_input("Coef. fricción μ",                 value=MU,           step=0.01,format="%.2f")
    st.divider()
    st.markdown("**Umbrales balance S/I**")
    bal_ok   = st.slider("Verde → Amarillo (%)",5, 25,int(BALANCE_OK))
    bal_warn = st.slider("Amarillo → Rojo (%)", 10,40,int(BALANCE_WARN))
    st.divider()
    st.markdown("<small style='color:#888'>Ref.: UIC 541-3 · EN 15328 · UIC 544-1</small>",unsafe_allow_html=True)

if not archivo:
    st.info("👈 Cargá el Excel generado por el extractor EDP para comenzar.")
    st.markdown("""
    **Análisis disponibles**
    - Margen de frenado vs desaceleración de diseño
    - Balance Superior / Inferior por caliper con semáforo ✅⚠️🔴
    - Fuerza por rueda con desvío porcentual
    - Comparación Bogie 1 vs Bogie 2
    - Repetibilidad entre tests (Coef. de Variación)
    - Progresión Servicio vs Emergencia vs ratio teórico
    """)
else:
    with st.spinner("Cargando datos..."):
        try:
            df = cargar_excel(archivo)
            BALANCE_OK   = bal_ok
            BALANCE_WARN = bal_warn
            df["estado_balance"] = df["balance_pct"].apply(
                lambda x:"OK" if x<BALANCE_OK else("ATENCIÓN" if x<BALANCE_WARN else"FUERA"))
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}"); st.stop()

    st.success(
        f"✅ {len(df)} mediciones · {df['archivo'].nunique()} archivos · "
        f"Condiciones: {', '.join(sorted(df['condicion'].unique()))} · "
        f"Vehículo: {df['vehiculo'].iloc[0]}")

    tabs = st.tabs(["📊 Resumen","⚖️ Balance S/I","🚂 Por rueda",
                    "🔄 Bogies","🔁 Repetibilidad","📈 Serv vs Emer","📋 Datos"])
    with tabs[0]: seccion_resumen(df,masa,a_serv,a_emer,mu_val)
    with tabs[1]: seccion_balance(df,bal_ok,bal_warn)
    with tabs[2]: seccion_ruedas(df)
    with tabs[3]: seccion_bogies(df)
    with tabs[4]: seccion_repetibilidad(df)
    with tabs[5]: seccion_serv_vs_emer(df)
    with tabs[6]: seccion_datos(df)
