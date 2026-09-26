from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st

from excel_generator import MAX_ALIMENTOS, MAX_PERSONAS, MAX_SINTOMAS, PERIODOS, build_excel
from importer import parse_consumer_excel, person_key
from report_generator import analyze, build_reports

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "plantilla_eta.xlsx"
ANEXO3_ORIGINAL = BASE_DIR / "Anexo 3.FORMATO_BROTES_ETA_JULIO 2022.xls"

st.set_page_config(page_title="APP ETA - Investigación de brotes", page_icon="📋", layout="wide")

DEFAULT_SYMPTOMS = ["Vómito", "Diarrea", "Náuseas", "Dolor abdominal", "Fiebre"]
MODULES = [
    "1. Inicio / Importar",
    "2. Encuesta en campo",
    "3. Configuración",
    "4. Revisar registros",
    "5. Informes y análisis",
    "6. Descargar",
]


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
    if "module" not in st.session_state:
        st.session_state.module = MODULES[0]


def resize_records(n: int):
    n = max(1, min(int(n), MAX_PERSONAS))
    while len(st.session_state.people) < n:
        st.session_state.people.append(empty_person())
        st.session_state.consumptions.append(empty_consumption())
    if len(st.session_state.people) > n:
        st.session_state.people = st.session_state.people[:n]
        st.session_state.consumptions = st.session_state.consumptions[:n]
    st.session_state.num_personas = n


def active_count():
    return sum(1 for p in st.session_state.people if p["Nombres y apellidos"].strip() or p["Identificación"].strip())


def first_empty_index() -> int:
    for i, p in enumerate(st.session_state.people, start=1):
        if not (p["Nombres y apellidos"].strip() or p["Identificación"].strip()):
            return i
    return st.session_state.num_personas


def union_keep_order(a: list[str], b: list[str], maximum: int) -> list[str]:
    out, seen = [], set()
    for item in list(a) + list(b):
        text = str(item).strip()
        if text and text.casefold() not in seen:
            seen.add(text.casefold())
            out.append(text)
    return out[:maximum]


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
    return {
        "general": g,
        "symptoms": st.session_state.symptoms,
        "foods": st.session_state.foods,
        "persons": persons,
        "food_rows": food_rows,
    }


def apply_import(parsed: dict, mode: str) -> dict:
    imported_people = parsed.get("people", [])[:MAX_PERSONAS]
    imported_cons = parsed.get("consumptions", [])[:len(imported_people)]
    imported_symptoms = parsed.get("symptoms", [])
    imported_foods = parsed.get("foods", [])

    # Conservar variables detectadas en el Excel para que las X importadas sigan teniendo significado.
    if mode == "Reemplazar registros actuales":
        st.session_state.symptoms = imported_symptoms or st.session_state.symptoms
        st.session_state.foods = imported_foods or st.session_state.foods
        st.session_state.people = [dict(p) for p in imported_people]
        st.session_state.consumptions = [dict(c) for c in imported_cons]
        n = max(1, len(imported_people))
        resize_records(n)
        added = len(imported_people)
        duplicates = 0
    else:
        st.session_state.symptoms = union_keep_order(st.session_state.symptoms, imported_symptoms, MAX_SINTOMAS)
        st.session_state.foods = union_keep_order(st.session_state.foods, imported_foods, MAX_ALIMENTOS)

        existing_pairs = [
            (p, st.session_state.consumptions[i])
            for i, p in enumerate(st.session_state.people)
            if p["Nombres y apellidos"].strip() or p["Identificación"].strip()
        ]
        seen = {person_key(p) for p, _ in existing_pairs if person_key(p)}
        added = duplicates = 0
        for p, c in zip(imported_people, imported_cons):
            key = person_key(p)
            if key and key in seen:
                duplicates += 1
                continue
            if len(existing_pairs) >= MAX_PERSONAS:
                break
            existing_pairs.append((dict(p), dict(c)))
            if key:
                seen.add(key)
            added += 1

        target_n = min(MAX_PERSONAS, max(st.session_state.num_personas, len(existing_pairs), 1))
        st.session_state.people = [p for p, _ in existing_pairs]
        st.session_state.consumptions = [c for _, c in existing_pairs]
        st.session_state.num_personas = len(st.session_state.people)
        resize_records(target_n)

    # Importar solo valores generales existentes, sin borrar lo ya diligenciado con vacíos.
    for key, value in parsed.get("general", {}).items():
        if value not in (None, ""):
            st.session_state.general[key] = value

    st.session_state.pop("generated", None)
    return {"added": added, "duplicates": duplicates, "total": active_count()}


init_state()

st.title("📋 APP ETA - Investigación de brotes y Encuesta de Consumidores")
st.caption("Dos formas de captura: importar el Anexo 2 ya diligenciado o realizar las encuestas directamente en campo desde la app.")

page = st.sidebar.radio("Módulos", MODULES, key="module")
st.sidebar.metric("Cupos habilitados", st.session_state.num_personas)
st.sidebar.metric("Encuestas diligenciadas", active_count())
st.sidebar.caption("Máximo: 100 personas por investigación.")

if page == "1. Inicio / Importar":
    st.header("1. ¿Cómo quieres ingresar la información?")
    left, right = st.columns(2)
    with left:
        st.subheader("📤 Opción A — Cargar Excel")
        st.write("Carga una **Encuesta de Consumidores (Anexo 2) en .xlsx** que ya esté diligenciada. La app leerá las dos pestañas, personas, síntomas y alimentos.")
        st.info("Puedes reemplazar los registros actuales o agregar el Excel a las encuestas que ya llevas en la app. Al agregar, se revisan duplicados por identificación y, si no existe, por nombre.")
    with right:
        st.subheader("🧑‍⚕️ Opción B — Encuestar en la app")
        st.write("Ideal para trabajo de campo. Registras **una persona a la vez**, incluidos síntomas y consumo de alimentos de los tres periodos del Anexo 2.")
        st.success("Para iniciar en campo, selecciona **2. Encuesta en campo** en el menú lateral.")

    st.divider()
    st.subheader("Importar Encuesta de Consumidores")
    uploaded = st.file_uploader("Selecciona el Anexo 2 diligenciado", type=["xlsx"], help="Debe ser el formato de Encuesta de Consumidores con sus dos pestañas.")
    if uploaded is not None:
        upload_key = f"{uploaded.name}:{uploaded.size}"
        if st.session_state.get("import_preview_key") != upload_key:
            try:
                st.session_state.import_preview = parse_consumer_excel(uploaded.getvalue())
                st.session_state.import_preview_key = upload_key
                st.session_state.import_preview_name = uploaded.name
            except Exception as exc:
                st.session_state.import_preview = None
                st.error(str(exc))

        parsed = st.session_state.get("import_preview")
        if parsed:
            pcases = sum(1 for p in parsed["people"] if p.get("Enfermo"))
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Personas encontradas", len(parsed["people"]))
            m2.metric("Enfermos", pcases)
            m3.metric("Síntomas detectados", len(parsed["symptoms"]))
            m4.metric("Alimentos detectados", len(parsed["foods"]))
            if parsed.get("warnings"):
                for w in parsed["warnings"]:
                    st.warning(w)

            with st.expander("Vista previa de personas", expanded=False):
                preview_rows = [{
                    "No.": i + 1,
                    "Nombre": p["Nombres y apellidos"],
                    "Identificación": p["Identificación"],
                    "Edad": p["Edad"],
                    "Sexo": p["Sexo"],
                    "Enfermo": "Sí" if p["Enfermo"] else "No",
                } for i, p in enumerate(parsed["people"])]
                st.dataframe(preview_rows, hide_index=True, width="stretch")

            mode = st.radio(
                "¿Qué deseas hacer con los datos cargados?",
                ["Reemplazar registros actuales", "Agregar sin duplicar"],
                horizontal=True,
            )
            if st.button("✅ Incorporar Excel a la investigación", type="primary", width="stretch"):
                result = apply_import(parsed, mode)
                st.session_state.last_import = {
                    "archivo": uploaded.name,
                    "modo": mode,
                    **result,
                }
                st.success(
                    f"Excel incorporado. Se agregaron {result['added']} personas; "
                    f"duplicados omitidos: {result['duplicates']}. Total diligenciado: {result['total']}."
                )

    if st.session_state.get("last_import"):
        li = st.session_state.last_import
        st.caption(f"Última importación: {li['archivo']} — {li['modo']} — total actual: {li['total']} personas.")

elif page == "2. Encuesta en campo":
    st.header("2. Encuesta de Consumidores en campo")
    st.write("Diligencia una persona a la vez. Al guardar, la encuesta queda integrada con cualquier Excel que hayas importado.")

    if not st.session_state.foods:
        st.warning("Aún no has definido alimentos a investigar. Puedes registrar la persona y los síntomas, pero para seleccionar alimentos ve a **3. Configuración** y agrega el listado del brote.")

    target = int(st.session_state.pop("field_nav_target", first_empty_index()))
    target = max(1, min(target, st.session_state.num_personas))
    person_no = st.selectbox(
        "Encuesta / persona a diligenciar",
        list(range(1, st.session_state.num_personas + 1)),
        index=target - 1,
        format_func=lambda x: f"Persona {x} — {st.session_state.people[x-1]['Nombres y apellidos'] or st.session_state.people[x-1]['Identificación'] or 'NUEVA'}",
    )
    idx = person_no - 1
    p = st.session_state.people[idx]
    cons = st.session_state.consumptions[idx]

    with st.form(f"field_survey_{person_no}"):
        st.subheader(f"Identificación — Persona {person_no}")
        c1, c2, c3, c4 = st.columns([2, 1.2, .7, .8])
        name = c1.text_input("Nombres y apellidos", p["Nombres y apellidos"])
        ident = c2.text_input("Identificación", p["Identificación"])
        age = c3.text_input("Edad", str(p["Edad"]))
        sex_options = ["", "F", "M", "Otro"]
        sex = c4.selectbox("Sexo", sex_options, index=sex_options.index(p["Sexo"]) if p["Sexo"] in sex_options else 0)
        address = st.text_input("Dirección y teléfono", p["Dirección y teléfono"])

        st.subheader("Condición clínica")
        q1, q2, q3 = st.columns(3)
        sick = q1.checkbox("Enfermo", p["Enfermo"])
        consult = q2.checkbox("Consultó", p["Consulta"])
        hosp = q3.checkbox("Hospitalizado", p["Hospitalizado"])
        d1, d2 = st.columns(2)
        symptom_date = d1.date_input("Fecha de inicio de síntomas", value=p["Día síntomas"], format="DD/MM/YYYY")
        symptom_time = d2.time_input("Hora de inicio de síntomas", value=p["Hora síntomas"], step=300)
        selected_symptoms = st.multiselect(
            "Signos y síntomas",
            st.session_state.symptoms,
            default=[x for x in p["selected_symptoms"] if x in st.session_state.symptoms],
        )
        sample = st.text_input("Muestra tomada / tipo de muestra", p["Muestra"])

        st.subheader("Consumo de alimentos")
        st.caption("Registra los tres periodos del Anexo 2 para esta misma persona.")
        new_cons = {}
        for pno, period in enumerate(PERIODOS):
            old = cons[period]
            with st.expander(period, expanded=(pno == 0)):
                c1, c2, c3 = st.columns([1, 1, 2])
                food_date = c1.date_input("Fecha", value=old["Día"], format="DD/MM/YYYY", key=f"field_fd_{person_no}_{pno}")
                food_time = c2.time_input("Hora", value=old["Hora"], step=300, key=f"field_ft_{person_no}_{pno}")
                place = c3.text_input("Lugar de consumo", old["Lugar de consumo"], key=f"field_fp_{person_no}_{pno}")
                selected_foods = st.multiselect(
                    "Alimentos consumidos",
                    st.session_state.foods,
                    default=[x for x in old["selected_foods"] if x in st.session_state.foods],
                    key=f"field_ff_{person_no}_{pno}",
                )
                new_cons[period] = {
                    "Día": food_date,
                    "Hora": food_time,
                    "Lugar de consumo": place,
                    "selected_foods": selected_foods,
                }

        b1, b2 = st.columns(2)
        save = b1.form_submit_button("💾 Guardar encuesta", type="primary", width="stretch")
        save_next = b2.form_submit_button("💾 Guardar y continuar con la siguiente", width="stretch")

    if save or save_next:
        st.session_state.people[idx] = {
            "Nombres y apellidos": name,
            "Identificación": ident,
            "Edad": age,
            "Sexo": sex,
            "Dirección y teléfono": address,
            "Día síntomas": symptom_date,
            "Hora síntomas": symptom_time,
            "selected_symptoms": selected_symptoms,
            "Enfermo": sick,
            "Consulta": consult,
            "Hospitalizado": hosp,
            "Muestra": sample,
        }
        st.session_state.consumptions[idx] = new_cons
        st.session_state.pop("generated", None)
        if save_next:
            if person_no == st.session_state.num_personas and st.session_state.num_personas < MAX_PERSONAS:
                resize_records(st.session_state.num_personas + 1)
            st.session_state.field_nav_target = min(person_no + 1, st.session_state.num_personas)
            st.success(f"Persona {person_no} guardada. Abriendo la siguiente encuesta...")
            st.rerun()
        else:
            st.success(f"Encuesta de la persona {person_no} guardada correctamente.")

elif page == "3. Configuración":
    st.header("3. Datos generales y variables de la investigación")
    with st.form("config_form"):
        c1, c2, c3 = st.columns(3)
        n = c1.number_input("Cupos de personas habilitados", 1, MAX_PERSONAS, st.session_state.num_personas, 1)
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
        symptom_text = l.text_area(f"Signos y síntomas (máx. {MAX_SINTOMAS})", "\n".join(st.session_state.symptoms), height=220)
        food_text = rr.text_area(
            f"Alimentos a investigar (máx. {MAX_ALIMENTOS})",
            "\n".join(st.session_state.foods),
            height=220,
            placeholder="Un alimento por línea",
        )
        submitted = st.form_submit_button("💾 Guardar configuración", type="primary", width="stretch")
    if submitted:
        resize_records(int(n))
        new_symptoms = parse_list(symptom_text, MAX_SINTOMAS)
        new_foods = parse_list(food_text, MAX_ALIMENTOS)
        st.session_state.symptoms = new_symptoms
        st.session_state.foods = new_foods
        st.session_state.general.update({
            "departamento": departamento, "municipio": municipio, "localidad": localidad,
            "lugar_brote": lugar, "direccion_brote": direccion, "telefono_brote": tel_brote,
            "encuestador": encuestador, "telefono_encuestador": tel_enc,
            "fecha_ocurrencia": fecha_oc, "fecha_deteccion": fecha_det,
            "fecha_notificacion": fecha_not, "fecha_investigacion": fecha_inv,
        })
        for person in st.session_state.people:
            person["selected_symptoms"] = [x for x in person["selected_symptoms"] if x in new_symptoms]
        for all_cons in st.session_state.consumptions:
            for period in PERIODOS:
                all_cons[period]["selected_foods"] = [x for x in all_cons[period]["selected_foods"] if x in new_foods]
        st.session_state.pop("generated", None)
        st.success("Configuración guardada.")

elif page == "4. Revisar registros":
    st.header("4. Revisar registros consolidados")
    payload = consumer_payload()
    a = analyze(payload)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Personas diligenciadas", a["total_exposed"])
    k2.metric("Enfermos", a["total_cases"])
    k3.metric("No enfermos", max(0, a["total_exposed"] - a["total_cases"]))
    k4.metric("Tasa de ataque", f"{a['attack_rate']:.1f}%")

    rows = []
    for i, p in enumerate(st.session_state.people, start=1):
        if not (p["Nombres y apellidos"].strip() or p["Identificación"].strip()):
            continue
        rows.append({
            "No.": i,
            "Nombres y apellidos": p["Nombres y apellidos"],
            "Identificación": p["Identificación"],
            "Edad": p["Edad"],
            "Sexo": p["Sexo"],
            "Inicio síntomas": f"{fmt_date(p['Día síntomas'])} {fmt_time(p['Hora síntomas'])}".strip(),
            "Enfermo": "Sí" if p["Enfermo"] else "No",
            "Consultó": "Sí" if p["Consulta"] else "No",
            "Hospitalizado": "Sí" if p["Hospitalizado"] else "No",
        })
    st.dataframe(rows, hide_index=True, width="stretch")
    st.caption("Para corregir una persona, entra a **2. Encuesta en campo** y selecciona su número.")

    with st.expander("Limpiar una encuesta equivocada", expanded=False):
        clear_no = st.selectbox("Persona a limpiar", list(range(1, st.session_state.num_personas + 1)), key="clear_person")
        if st.button("🗑️ Limpiar persona seleccionada"):
            st.session_state.people[clear_no - 1] = empty_person()
            st.session_state.consumptions[clear_no - 1] = empty_consumption()
            st.session_state.pop("generated", None)
            st.success(f"Se limpiaron los datos de la persona {clear_no}.")

elif page == "5. Informes y análisis":
    st.header("5. Informes de 24 horas, 72 horas y final")
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
            st.dataframe(
                [{"Signo/síntoma": s, "Casos": n, "%": round(pct, 1)} for s, n, pct in a["symptom_counts"]],
                hide_index=True,
                width="stretch",
            )
        st.markdown("**Análisis por alimento**")
        if a["food_analysis"]:
            st.dataframe([{
                "Alimento": x["food"], "Caso exp.": x["a"], "Sano exp.": x["b"],
                "Caso no exp.": x["c"], "Sano no exp.": x["d"],
                "TA exp. %": round((x["attack_exposed"] or 0) * 100, 1),
                "TA no exp. %": round((x["attack_unexposed"] or 0) * 100, 1),
                "RR*": round(x["rr"], 2) if x["rr"] is not None else None,
                "OR*": round(x["or"], 2) if x["or"] is not None else None,
            } for x in a["food_analysis"]], hide_index=True, width="stretch")

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
        save_reports = st.form_submit_button("💾 Guardar información de los informes", type="primary", width="stretch")
    if save_reports:
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
    st.header("6. Generar y descargar")
    st.write("Los archivos se generan solo al pulsar el botón. Tanto los Excel importados como las encuestas realizadas en la app quedan consolidados en el mismo resultado.")
    payload = consumer_payload()
    a = analyze(payload)
    st.info(f"Se usarán {a['total_exposed']} personas diligenciadas, {a['total_cases']} casos y {len(st.session_state.foods)} alimentos definidos.")

    if st.button("⚙️ Generar archivos del brote", type="primary", width="stretch"):
        try:
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
        c1.download_button(
            "⬇️ Encuesta de consumidores (Anexo 2)", survey_bytes, survey_name,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
        )
        c2.download_button(
            "⬇️ Informes 24 h + 72 h + final", reports_bytes, reports_name,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
        )
        st.download_button(
            "📦 Descargar expediente completo", zip_bytes, "ETA_expediente_completo.zip", "application/zip",
            type="primary", width="stretch",
        )
        if ANEXO3_ORIGINAL.exists():
            st.download_button(
                "📄 Descargar Anexo 3 original aportado", ANEXO3_ORIGINAL.read_bytes(), ANEXO3_ORIGINAL.name,
                "application/vnd.ms-excel", width="stretch",
            )

st.divider()
st.caption(
    "Privacidad: si ejecutas la app localmente, los datos permanecen en tu equipo durante la sesión. "
    "Si la publicas en Streamlit Community Cloud, la información se procesa en infraestructura en la nube; "
    "para datos reales identificables usa únicamente un entorno institucional autorizado y con controles de acceso."
)
