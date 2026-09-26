from __future__ import annotations

from datetime import date, datetime, time
from io import BytesIO
from typing import Any
import re

from openpyxl import load_workbook

from excel_generator import (
    ALIMENTO_COLS,
    MAX_ALIMENTOS,
    MAX_PERSONAS,
    MAX_SINTOMAS,
    PERIODOS,
    SINTOMA_COLS,
)


def _clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value != value:
        return ""
    return str(value).strip()


def _truthy(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return _clean(value).casefold() in {"x", "si", "sí", "s", "1", "true", "verdadero", "yes"}


def _parse_date(value: Any):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = _clean(value)
    # Tolerar fecha con hora al final.
    text = text.split(" ")[0]
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def _parse_time(value: Any):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)
    # Excel puede representar una hora como fracción de día.
    if isinstance(value, (int, float)) and 0 <= float(value) < 1:
        total_minutes = int(round(float(value) * 24 * 60)) % (24 * 60)
        return time(total_minutes // 60, total_minutes % 60)
    text = _clean(value).lower().replace(".", "")
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p"):
        try:
            return datetime.strptime(text.upper(), fmt).time().replace(second=0, microsecond=0)
        except ValueError:
            pass
    return None


def _three_part_date(ws, row: int):
    """Lee el patrón día/mes/año usado por el Anexo 2 en G/I/K."""
    d = ws[f"G{row}"].value
    m = ws[f"I{row}"].value
    y = ws[f"K{row}"].value
    try:
        if d not in (None, "") and m not in (None, "") and y not in (None, ""):
            return date(int(float(y)), int(float(m)), int(float(d)))
    except (TypeError, ValueError):
        pass
    return None


def _find_sheet(wb, preferred: str, fallback_index: int):
    if preferred in wb.sheetnames:
        return wb[preferred]
    if len(wb.worksheets) > fallback_index:
        return wb.worksheets[fallback_index]
    raise ValueError(f"No se encontró la hoja requerida: {preferred}.")


def _is_signature_row(ws, row: int) -> bool:
    texts = " ".join(_clean(ws.cell(row, c).value) for c in range(1, min(ws.max_column, 30) + 1)).casefold()
    return "firma del encuestador" in texts or "firma del encuest" in texts


def parse_consumer_excel(file_bytes: bytes) -> dict:
    """
    Lee un Anexo 2 ETA (.xlsx) ya diligenciado y lo convierte al mismo modelo
    que usa la app. Está optimizado para la plantilla oficial y para los Excel
    generados por esta aplicación (15 a 100 personas).
    """
    try:
        wb = load_workbook(BytesIO(file_bytes), data_only=False)
    except Exception as exc:
        raise ValueError(
            "No fue posible leer el archivo. Carga la Encuesta de Consumidores en formato .xlsx."
        ) from exc

    ws1 = _find_sheet(wb, "Hoja1", 0)
    # En el archivo oficial la segunda pestaña útil suele llamarse Hoja3.
    if "Hoja3" in wb.sheetnames:
        ws3 = wb["Hoja3"]
    elif len(wb.worksheets) >= 2:
        ws3 = wb.worksheets[-1]
    else:
        raise ValueError("El archivo no contiene la segunda pestaña de alimentos.")

    symptoms = [_clean(ws1[f"{col}10"].value) for col in SINTOMA_COLS]
    symptoms = [x for x in symptoms if x][:MAX_SINTOMAS]
    foods = [_clean(ws3[f"{col}10"].value) for col in ALIMENTO_COLS]
    foods = [x for x in foods if x][:MAX_ALIMENTOS]

    persons: list[dict] = []
    max_person_row = min(ws1.max_row, 12 + MAX_PERSONAS + 10)
    for row in range(12, max_person_row + 1):
        if _is_signature_row(ws1, row):
            break

        no = ws1[f"A{row}"].value
        name = _clean(ws1[f"B{row}"].value)
        ident = _clean(ws1[f"C{row}"].value)
        age = ws1[f"D{row}"].value
        sex = _clean(ws1[f"E{row}"].value)
        address = _clean(ws1[f"F{row}"].value)
        symptom_date = _parse_date(ws1[f"G{row}"].value)
        symptom_time = _parse_time(ws1[f"H{row}"].value)
        enfermo_raw = ws1[f"AA{row}"].value
        consulta_raw = ws1[f"AB{row}"].value
        hosp_raw = ws1[f"AC{row}"].value
        sample = _clean(ws1[f"AD{row}"].value)

        symptom_values = []
        for symptom, col in zip(symptoms, SINTOMA_COLS):
            if _truthy(ws1[f"{col}{row}"].value):
                symptom_values.append(symptom)

        meaningful = any([
            name, ident, _clean(age), sex, address, symptom_date, symptom_time,
            symptom_values, _clean(enfermo_raw), _clean(consulta_raw), _clean(hosp_raw), sample,
        ])
        # Filas numeradas vacías del formato no cuentan como encuestas diligenciadas.
        if not meaningful:
            continue

        persons.append({
            "Nombres y apellidos": name,
            "Identificación": ident,
            "Edad": "" if age is None else age,
            "Sexo": sex,
            "Dirección y teléfono": address,
            "Día síntomas": symptom_date,
            "Hora síntomas": symptom_time,
            "selected_symptoms": symptom_values,
            "Enfermo": _truthy(enfermo_raw),
            "Consulta": _truthy(consulta_raw),
            "Hospitalizado": _truthy(hosp_raw),
            "Muestra": sample,
            "_source_no": int(float(no)) if no not in (None, "") and str(no).replace(".", "", 1).isdigit() else len(persons) + 1,
        })
        if len(persons) >= MAX_PERSONAS:
            break

    # Construye un mapa por número original para mantener consumo y persona alineados.
    source_to_new = {int(p.pop("_source_no", i + 1)): i for i, p in enumerate(persons)}
    consumptions = [
        {period: {"Día": None, "Hora": None, "Lugar de consumo": "", "selected_foods": []} for period in PERIODOS}
        for _ in persons
    ]

    current_no = None
    max_food_row = min(ws3.max_row, 12 + MAX_PERSONAS * 3 + 20)
    for row in range(12, max_food_row + 1):
        if _is_signature_row(ws3, row):
            break
        raw_no = ws3[f"A{row}"].value
        if raw_no not in (None, ""):
            try:
                current_no = int(float(raw_no))
            except (TypeError, ValueError):
                current_no = None
        if current_no is None or current_no not in source_to_new:
            continue

        period_text = _clean(ws3[f"C{row}"].value)
        period = next((p for p in PERIODOS if p.casefold() == period_text.casefold()), None)
        if period is None:
            # Tolerar pequeñas variaciones del texto del formato oficial.
            t = period_text.casefold()
            if "día de los síntomas" in t or "dia de los sintomas" in t:
                period = PERIODOS[0]
            elif "un día antes" in t or "un dia antes" in t:
                period = PERIODOS[1]
            elif "dos días antes" in t or "dos dias antes" in t or "2 días antes" in t or "2 dias antes" in t:
                period = PERIODOS[2]
            else:
                continue

        selected_foods = []
        for food, col in zip(foods, ALIMENTO_COLS):
            if _truthy(ws3[f"{col}{row}"].value):
                selected_foods.append(food)

        idx = source_to_new[current_no]
        consumptions[idx][period] = {
            "Día": _parse_date(ws3[f"D{row}"].value),
            "Hora": _parse_time(ws3[f"E{row}"].value),
            "Lugar de consumo": _clean(ws3[f"F{row}"].value),
            "selected_foods": selected_foods,
        }

    # Datos generales que realmente están disponibles en el Anexo 2.
    general = {
        "fecha_ocurrencia": _three_part_date(ws1, 6),
        "fecha_notificacion": _three_part_date(ws1, 7),
        "fecha_investigacion": _three_part_date(ws1, 8),
    }

    # Intentar recuperar nombre/teléfono del encuestador de la fila de firma.
    for row in range(12, ws1.max_row + 1):
        if _is_signature_row(ws1, row):
            sig = _clean(ws1[f"B{row}"].value)
            tel = _clean(ws1[f"H{row}"].value)
            if ":" in sig:
                general["encuestador"] = sig.split(":", 1)[1].strip()
            if ":" in tel:
                general["telefono_encuestador"] = tel.split(":", 1)[1].strip()
            break

    warnings: list[str] = []
    if not persons:
        warnings.append("No se encontraron personas diligenciadas en la primera pestaña.")
    if not symptoms:
        warnings.append("No se identificaron encabezados de signos y síntomas en la fila 10.")
    if not foods:
        warnings.append("No se identificaron alimentos en la segunda pestaña.")

    return {
        "general": general,
        "symptoms": symptoms,
        "foods": foods,
        "people": persons,
        "consumptions": consumptions,
        "warnings": warnings,
        "sheet_names": wb.sheetnames,
    }


def person_key(person: dict) -> str:
    ident = re.sub(r"\s+", "", _clean(person.get("Identificación"))).casefold()
    if ident:
        return f"id:{ident}"
    name = re.sub(r"\s+", " ", _clean(person.get("Nombres y apellidos"))).strip().casefold()
    return f"name:{name}" if name else ""
