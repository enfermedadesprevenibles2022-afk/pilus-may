from __future__ import annotations

from datetime import date, time
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st

from excel_generator import MAX_ALIMENTOS, MAX_PERSONAS, MAX_SINTOMAS, PERIODOS, build_excel
from report_generator import analyze, build_reports

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "plantilla_eta.xlsx"
ANEXO3_ORIGINAL = BASE_DIR / "Anexo 3.FORMATO_BROTES_ETA_JULIO 2022.xls"

st.set_page_config(page_title="ETA - Investigación de brotes", page_icon="📋", layout="wide")

DEFAULT_SYMPTOMS = ["Vómito", "Diarrea", "Náuseas", "Dolor abdominal", "Fiebre"]


def parse_list(text: str, maximum: int) -> list[str]:
    raw = [x.strip() for line in text.splitlines() for x in line.split(",")]
    out, seen = [], set()
    for item in raw:
        if item and item.casefold() not in seen:
            seen.add(item.casefold())
            out.append(item)
    return out[:maximum]


def fmt_date(v):
    if not v:
        return ""
    return v.strftime("%d/%m/%Y") if hasattr(v, "strftime") else str(v)


def fmt_time(v):
    if not v:
        return ""
    return v.strftime("%H:%M") if hasattr(v, "strftime") else str(v)


def empty_person():
    return {
        "Nombres y apellidos": "", "Identificación": "", "Edad": "", "Sexo": "",
        "Dirección y teléfono": "", "Día síntomas": None, "Hora síntomas": None,
        "selected_symptoms": [], "Enfermo": False, "Consulta": False,
        "Hospitalizado": False, "Muestra": "",
    }


def empty_consumption():
    return {p: {"Día": None, "Hora": None, "Lugar de consumo": "", "selected_foods": []} for p in PERIODOS}


def init_state():
    if "num_personas" not in st.session_state:
        st.session_state.num_personas = 15
    if "symptoms" not in st.session_state:
        st.session_state.symptoms = DEFAULT_SYMPTOMS.copy()
    if "foods" not in st.session_state:
        st.session_state.foods = []
    if "people" not in st.session_state:
        st.session_state.people = [empty_person() for _ in range(st.session_state.num_personas)]
    if "consumptions" not in st.session_state:
        st.session_state.consumptions = [empty_consumption() for _ in range(st.session_state.num_personas)]
    if "general" not in st.session_state:
        st.session_state.general = {
            "departamento": "Santander", "municipio": "", "localidad": "", "lugar_brote": "",
            "direccion_brote": "", "telefono_brote": "", "fecha_ocurrencia": None,
            "fecha_deteccion": None, "fecha_notificacion": None, "fecha_investigacion": None,
            "encuestador": "", "telefono_encuestador": "",
        }
    if "report_fields" not in st.session_state:
        st.session_state.report_fields = {}


def resize_records(n: int):
    while len(st.session_state.people) < n:
        st.session_state.people.append(empty_person())
        st.session_state.consumptions.append(empty_consumption())
    if len(st.session_state.people) > n:
        st.session_state.people = st.session_state.people[:n]
        st.session_state.consumptions = st.session_state.consumptions[:n]
    st.session_state.num_personas = n


def consumer_payload():
    persons = []
    for p in st.session_state.people:
        row = {
            "Nombres y apellidos": p["Nombres y apellidos"],
            "Identificación": p["Identificación"],
            "Edad": p["Edad"], "Sexo": p["Sexo"],
            "Dirección y teléfono": p["Dirección y teléfono"],
            "Día síntomas": fmt_date(p["Día síntomas"]),
            "Hora síntomas": fmt_time(p["Hora síntomas"]),
            "Enfermo": p["Enfermo"], "Consulta": p["Consulta"],
            "Hospitalizado": p["Hospitalizado"], "Muestra": p["Muestra"],
        }
        for s in st.session_state.symptoms:
            row[s] = s in p["selected_symptoms"]
        persons.append(row)

    food_rows = []
    for i, cons in enumerate(st.session_state.consumptions, start=1):
        ident = st.session_state.people[i - 1]["Identificación"]
        for period in PERIODOS:
            c = cons[period]
            row = {
                "No.": i, "Identificación": ident, "Periodo": period,
                "Día": fmt_date(c["Día"]), "Hora": fmt_time(c["Hora"]),
                "Lugar de consumo": c["Lugar de consumo"],
            }
            for f in st.session_state.foods:
                row[f] = f in c["selected_foods"]
            food_rows.append(row)

    g = st.session_state.general.copy()
    for k in ["fecha_ocurrencia", "fecha_deteccion", "fecha_notificacion", "fecha_investigacion"]:
        g[k] = fmt_date(g.get(k))
    return {"general": g, "symptoms": st.session_state.symptoms, "foods": st.session_state.foods,
            "persons": persons, "food_rows": food_rows}


def active_count():
    return sum(1 for p in st.session_state.people if p["Nombres y apellidos"].strip() or p["Identificación"].strip())


init_state()

st.title("📋 Investigación de brotes de Enfermedades Transmitidas por Alimentos - ETA")
st.caption("Versión optimizada: hasta 100 personas, calendarios, guardado por bloques y generación de encuesta + informes 24 h, 72 h y final.")

# Navegación simple y rápida
page = st.sidebar.radio(
    "Módulos",
    ["1. Configuración", "2. Personas", "3. Consumo de alimentos", "4. Informes y análisis", "5. Descargar"],
)
st.sidebar.metric("Personas configuradas", st.session_state.num_personas)
st.sidebar.metric("Personas diligenciadas", active_count())
st.sidebar.caption("Los formularios solo recalculan la app al pulsar Guardar, evitando los bloqueos de la versión anterior.")

if page == "1. Configuración":
    st.header("1. Datos generales y configuración")
    with st.form("config_form"):
        c1, c2, c3 = st.columns(3)
        n = c1.number_input("Número máximo de personas de este brote", 1, MAX_PERSONAS, st.session_state.num_personas, 1)
        departamento = c2.text_input("Departamento", st.session_state.general.get("departamento", ""))
        municipio = c3.text_input("Municipio", st.session_state.general.get("municipio", ""))
        c4, c5, c6 = st.columns(3)
        localidad = c4.text_input("Barrio / corregimiento / vereda", st.session_state.general.get("localidad", ""))
        lugar = c5.text_input("Lugar o institución del brote", st.session_state.general.get("lugar_brote", ""))
        direccion = c6.text_input("Dirección del brote", st.session_state.general.get("direccion_brote", ""))
        c7, c8 = st.columns(2)
        tel_brote = c7.text_input("Teléfono del lugar", st.session_state.general.get("telefono_brote", ""))
        encuestador = c8.text_input("Nombre del encuestador", st.session_state.general.get("encuestador", ""))
        tel_enc = st.text_input("Teléfono del encuestador", st.session_state.general.get("telefono_encuestador", ""))

        st.subheader("Fechas")
        d1, d2, d3, d4 = st.columns(4)
        fecha_oc = d1.date_input("Inicio / ocurrencia", value=st.session_state.general.get("fecha_ocurrencia"), format="DD/MM/YYYY")
        fecha_det = d2.date_input("Detección", value=st.session_state.general.get("fecha_deteccion"), format="DD/MM/YYYY")
        fecha_not = d3.date_input("Notificación", value=st.session_state.general.get("fecha_notificacion"), format="DD/MM/YYYY")
        fecha_inv = d4.date_input("Inicio investigación", value=st.session_state.general.get("fecha_investigacion"), format="DD/MM/YYYY")

        st.subheader("Variables de la encuesta")
        l, rr = st.columns(2)
        symptom_text = l.text_area(f"Signos y síntomas (máx. {MAX_SINTOMAS})", "\n".join(st.session_state.symptoms), height=180)
        food_text = rr.text_area(f"Alimentos a investigar (máx. {MAX_ALIMENTOS})", "\n".join(st.session_state.foods), height=180,
                                 placeholder="Un alimento por línea")
        submitted = st.form_submit_button("💾 Guardar configuración", type="primary", use_container_width=True)
    if submitted:
        resize_records(int(n))
        st.session_state.symptoms = parse_list(symptom_text, MAX_SINTOMAS)
        st.session_state.foods = parse_list(food_text, MAX_ALIMENTOS)
        st.session_state.general.update({
            "departamento": departamento, "municipio": municipio, "localidad": localidad,
            "lugar_brote": lugar, "direccion_brote": direccion, "telefono_brote": tel_brote,
            "encuestador": encuestador, "telefono_encuestador": tel_enc,
            "fecha_ocurrencia": fecha_oc, "fecha_deteccion": fecha_det,
            "fecha_notificacion": fecha_not, "fecha_investigacion": fecha_inv,
        })
        # Eliminar selecciones que ya no estén en las variables configuradas.
        for p in st.session_state.people:
            p["selected_symptoms"] = [x for x in p["selected_symptoms"] if x in st.session_state.symptoms]
        for cons in st.session_state.consumptions:
            for period in PERIODOS:
                cons[period]["selected_foods"] = [x for x in cons[period]["selected_foods"] if x in st.session_state.foods]
        st.success("Configuración guardada. Los calendarios ya están activos.")

elif page == "2. Personas":
    st.header("2. Personas, signos y síntomas")
    page_size = 10
    total_blocks = (st.session_state.num_personas + page_size - 1) // page_size
    block = st.selectbox("Bloque de personas", list(range(1, total_blocks + 1)), format_func=lambda x: f"Personas {(x-1)*page_size+1} a {min(x*page_size, st.session_state.num_personas)}")
    start = (block - 1) * page_size
    end = min(start + page_size, st.session_state.num_personas)
    st.info("Completa este bloque y pulsa **Guardar bloque**. Cambiar una casilla o una fecha ya no recarga toda la aplicación.")
    with st.form(f"people_{block}"):
        new_rows = []
        for idx in range(start, end):
            p = st.session_state.people[idx]
            with st.expander(f"Persona {idx+1} — {p['Nombres y apellidos'] or p['Identificación'] or 'sin diligenciar'}", expanded=(idx == start)):
                c1, c2, c3, c4 = st.columns([2, 1.2, .7, .8])
                name = c1.text_input("Nombres y apellidos", p["Nombres y apellidos"], key=f"name_{block}_{idx}")
                ident = c2.text_input("Identificación", p["Identificación"], key=f"id_{block}_{idx}")
                age = c3.text_input("Edad", str(p["Edad"]), key=f"age_{block}_{idx}")
                sex_options = ["", "F", "M", "Otro"]
                sex = c4.selectbox("Sexo", sex_options, index=sex_options.index(p["Sexo"]) if p["Sexo"] in sex_options else 0, key=f"sex_{block}_{idx}")
                address = st.text_input("Dirección y teléfono", p["Dirección y teléfono"], key=f"addr_{block}_{idx}")
                d1, d2 = st.columns(2)
                symptom_date = d1.date_input("Fecha de inicio de síntomas", value=p["Día síntomas"], format="DD/MM/YYYY", key=f"date_{block}_{idx}")
                symptom_time = d2.time_input("Hora de inicio de síntomas", value=p["Hora síntomas"], step=300, key=f"time_{block}_{idx}")
                selected = st.multiselect("Signos y síntomas", st.session_state.symptoms, default=[x for x in p["selected_symptoms"] if x in st.session_state.symptoms], key=f"sym_{block}_{idx}")
                q1, q2, q3 = st.columns(3)
                sick = q1.checkbox("Enfermo", p["Enfermo"], key=f"sick_{block}_{idx}")
                consult = q2.checkbox("Consultó", p["Consulta"], key=f"consult_{block}_{idx}")
                hosp = q3.checkbox("Hospitalizado", p["Hospitalizado"], key=f"hosp_{block}_{idx}")
                sample = st.text_input("Muestra tomada / tipo de muestra", p["Muestra"], key=f"sample_{block}_{idx}")
                new_rows.append({
                    "Nombres y apellidos": name, "Identificación": ident, "Edad": age, "Sexo": sex,
                    "Dirección y teléfono": address, "Día síntomas": symptom_date, "Hora síntomas": symptom_time,
                    "selected_symptoms": selected, "Enfermo": sick, "Consulta": consult,
                    "Hospitalizado": hosp, "Muestra": sample,
                })
        save = st.form_submit_button("💾 Guardar bloque de personas", type="primary", use_container_width=True)
    if save:
        for offset, row in enumerate(new_rows):
            st.session_state.people[start + offset] = row
        st.success(f"Personas {start+1} a {end} guardadas.")

elif page == "3. Consumo de alimentos":
    st.header("3. Consumo de alimentos")
    if not st.session_state.foods:
        st.warning("Primero define los alimentos en **1. Configuración**.")
    else:
        page_size = 5
        total_blocks = (st.session_state.num_personas + page_size - 1) // page_size
        block = st.selectbox("Bloque de consumo", list(range(1, total_blocks + 1)), format_func=lambda x: f"Personas {(x-1)*page_size+1} a {min(x*page_size, st.session_state.num_personas)}")
        start = (block - 1) * page_size
        end = min(start + page_size, st.session_state.num_personas)
        st.info("Por cada persona aparecen los tres periodos del Anexo 2. Usa el calendario, la hora y selecciona los alimentos consumidos.")
        with st.form(f"food_{block}"):
            saved = []
            for idx in range(start, end):
                person = st.session_state.people[idx]
                title = f"Persona {idx+1} — {person['Nombres y apellidos'] or person['Identificación'] or 'sin diligenciar'}"
                with st.expander(title, expanded=(idx == start)):
                    person_out = {}
                    for pno, period in enumerate(PERIODOS):
                        c = st.session_state.consumptions[idx][period]
                        st.markdown(f"**{period}**")
                        c1, c2, c3 = st.columns([1, 1, 2])
                        d = c1.date_input("Fecha", value=c["Día"], format="DD/MM/YYYY", key=f"fd_{block}_{idx}_{pno}")
                        t = c2.time_input("Hora", value=c["Hora"], step=300, key=f"ft_{block}_{idx}_{pno}")
                        place = c3.text_input("Lugar de consumo", c["Lugar de consumo"], key=f"fp_{block}_{idx}_{pno}")
                        selected = st.multiselect("Alimentos consumidos", st.session_state.foods,
                                                  default=[x for x in c["selected_foods"] if x in st.session_state.foods],
                                                  key=f"ff_{block}_{idx}_{pno}")
                        person_out[period] = {"Día": d, "Hora": t, "Lugar de consumo": place, "selected_foods": selected}
                        if pno < 2:
                            st.divider()
                    saved.append(person_out)
            save = st.form_submit_button("💾 Guardar bloque de consumos", type="primary", use_container_width=True)
        if save:
            for offset, row in enumerate(saved):
                st.session_state.consumptions[start + offset] = row
            st.success(f"Consumos de las personas {start+1} a {end} guardados.")

elif page == "4. Informes y análisis":
    st.header("4. Informes de 24 horas, 72 horas y final")
    payload = consumer_payload()
    a = analyze(payload)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Expuestos diligenciados", a["total_exposed"])
    k2.metric("Casos / enfermos", a["total_cases"])
    k3.metric("Tasa de ataque", f"{a['attack_rate']:.1f}%")
    k4.metric("Hospitalizados", a["hospitalized"])

    with st.expander("Vista rápida del análisis automático", expanded=True):
        st.markdown("**Signos y síntomas**")
        if a["symptom_counts"]:
            st.dataframe([{"Signo/síntoma": s, "Casos": n, "%": round(p, 1)} for s, n, p in a["symptom_counts"]], hide_index=True, use_container_width=True)
        st.markdown("**Análisis por alimento**")
        if a["food_analysis"]:
            st.dataframe([{
                "Alimento": x["food"], "Caso exp.": x["a"], "Sano exp.": x["b"],
                "Caso no exp.": x["c"], "Sano no exp.": x["d"],
                "TA exp. %": round((x["attack_exposed"] or 0)*100, 1),
                "TA no exp. %": round((x["attack_unexposed"] or 0)*100, 1),
                "RR*": round(x["rr"], 2) if x["rr"] is not None else None,
                "OR*": round(x["or"], 2) if x["or"] is not None else None,
            } for x in a["food_analysis"]], hide_index=True, use_container_width=True)

    rf = st.session_state.report_fields
    with st.form("reports_form"):
        tab24, tab72, tabf = st.tabs(["Informe preliminar / 24 h", "Informe 72 h", "Informe final"])
        with tab24:
            c1, c2 = st.columns(2)
            casos_upgd = c1.number_input("Casos identificados en UPGD", min_value=0, value=int(rf.get("casos_upgd", 0) or 0))
            casos_bac = c2.number_input("Casos identificados por BAC", min_value=0, value=int(rf.get("casos_bac", 0) or 0))
            antecedentes = st.text_area("Antecedentes del brote", rf.get("antecedentes", ""), height=110)
            posibles = st.text_area("Posibles alimentos/agua o mecanismos de transmisión", rf.get("posibles_alimentos", ", ".join(st.session_state.foods)), height=90)
            hip = st.text_area("Hipótesis inicial", rf.get("hipotesis_inicial", ""), height=100)
            medidas = st.text_area("Medidas iniciales de control", rf.get("medidas_control", ""), height=110)
            otra = st.text_area("Otra información relevante", rf.get("otra_informacion", ""), height=90)
        with tab72:
            definicion = st.text_area("Definición operacional de caso", rf.get("definicion_caso", ""), height=100)
            manejo = st.text_area("Manejo y tratamiento clínico / complicaciones", rf.get("manejo_clinico", ""), height=100)
            lab = st.text_area("Resultados de laboratorio disponibles", rf.get("resultados_laboratorio", ""), height=100)
            hallazgos = st.text_area("Hallazgos ambientales / factores de riesgo", rf.get("hallazgos_ambientales", ""), height=110)
            analisis = st.text_area("Análisis de resultados e hipótesis", rf.get("analisis_resultados", ""), height=110)
            medidas72 = st.text_area("Medidas de control implementadas", rf.get("medidas_control_72", ""), height=100)
            rec72 = st.text_area("Recomendaciones - 72 horas", rf.get("recomendaciones_72", ""), height=100)
            con72 = st.text_area("Conclusiones - 72 horas", rf.get("conclusiones_72", ""), height=100)
        with tabf:
            fecha_cierre = st.date_input("Fecha de cierre del brote", value=rf.get("fecha_cierre"), format="DD/MM/YYYY")
            estados = ["", "Abierto", "Cerrado con agente identificado", "Cerrado sin agente identificado"]
            est_old = rf.get("estado_brote", "")
            estado = st.selectbox("Estado del brote", estados, index=estados.index(est_old) if est_old in estados else 0)
            agente = st.text_input("Agente identificado", rf.get("agente_identificado", ""))
            fuente = st.text_input("Fuente / alimento implicado", rf.get("fuente_implicada", ""))
            modo = st.text_input("Modo de transmisión", rf.get("modo_transmision", ""))
            resumen = st.text_area("Resumen de la situación", rf.get("resumen_final", ""), height=110)
            descripcion = st.text_area("Descripción del brote", rf.get("descripcion_brote", ""), height=110)
            labf = st.text_area("Resultados de laboratorio finales", rf.get("resultados_laboratorio_final", ""), height=100)
            factores = st.text_area("Factores determinantes", rf.get("factores_determinantes", ""), height=100)
            recf = st.text_area("Recomendaciones finales", rf.get("recomendaciones_final", ""), height=100)
            conf = st.text_area("Conclusiones finales", rf.get("conclusiones_final", ""), height=100)
            plan = st.text_area("Plan de mejoramiento / seguimiento", rf.get("plan_mejoramiento", ""), height=110)
        responsable = st.text_input("Responsable de elaboración de los informes", rf.get("responsable", st.session_state.general.get("encuestador", "")))
        save = st.form_submit_button("💾 Guardar información de los informes", type="primary", use_container_width=True)
    if save:
        st.session_state.report_fields = {
            "casos_upgd": casos_upgd, "casos_bac": casos_bac, "antecedentes": antecedentes,
            "posibles_alimentos": posibles, "hipotesis_inicial": hip, "medidas_control": medidas,
            "otra_informacion": otra, "definicion_caso": definicion, "manejo_clinico": manejo,
            "resultados_laboratorio": lab, "hallazgos_ambientales": hallazgos,
            "analisis_resultados": analisis, "medidas_control_72": medidas72,
            "recomendaciones_72": rec72, "conclusiones_72": con72,
            "fecha_cierre": fecha_cierre, "estado_brote": estado, "agente_identificado": agente,
            "fuente_implicada": fuente, "modo_transmision": modo, "resumen_final": resumen,
            "descripcion_brote": descripcion, "resultados_laboratorio_final": labf,
            "factores_determinantes": factores, "recomendaciones_final": recf,
            "conclusiones_final": conf, "plan_mejoramiento": plan, "responsable": responsable,
        }
        st.success("Información de los informes guardada.")

else:
    st.header("5. Generar y descargar")
    st.write("Los archivos se generan **solo cuando pulsas el botón**, para que la app no se trabe mientras diligencias.")
    payload = consumer_payload()
    a = analyze(payload)
    st.info(f"Se usarán {a['total_exposed']} personas diligenciadas, {a['total_cases']} casos y {len(st.session_state.foods)} alimentos definidos.")

    if st.button("⚙️ Generar archivos del brote", type="primary", use_container_width=True):
        try:
            # La encuesta conserva hasta el número configurado de filas; los vacíos no cuentan en los análisis.
            survey_bytes, survey_name = build_excel(payload, TEMPLATE_PATH)
            reports_bytes, reports_name = build_reports(payload, st.session_state.report_fields)
            zip_io = BytesIO()
            with ZipFile(zip_io, "w", ZIP_DEFLATED) as z:
                z.writestr(survey_name, survey_bytes)
                z.writestr(reports_name, reports_bytes)
                if ANEXO3_ORIGINAL.exists():
                    z.write(ANEXO3_ORIGINAL, arcname=ANEXO3_ORIGINAL.name)
            st.session_state.generated = (survey_bytes, survey_name, reports_bytes, reports_name, zip_io.getvalue())
            st.success("Archivos generados correctamente.")
        except Exception as exc:
            st.error(f"No fue posible generar los archivos: {exc}")

    if "generated" in st.session_state:
        survey_bytes, survey_name, reports_bytes, reports_name, zip_bytes = st.session_state.generated
        c1, c2 = st.columns(2)
        c1.download_button("⬇️ Encuesta de consumidores (Anexo 2)", survey_bytes, survey_name,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        c2.download_button("⬇️ Informes 24 h + 72 h + final", reports_bytes, reports_name,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.download_button("📦 Descargar expediente completo", zip_bytes, "ETA_expediente_completo.zip", "application/zip",
                           type="primary", use_container_width=True)
        if ANEXO3_ORIGINAL.exists():
            st.download_button("📄 Descargar Anexo 3 original aportado", ANEXO3_ORIGINAL.read_bytes(), ANEXO3_ORIGINAL.name,
                               "application/vnd.ms-excel", use_container_width=True)

st.divider()
st.caption("Privacidad: la aplicación trabaja localmente. No envía por sí sola identificaciones ni datos clínicos a servicios externos.")
