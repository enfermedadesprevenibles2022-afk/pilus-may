from __future__ import annotations

from datetime import date, datetime, time
from io import BytesIO
from typing import Any
import re
import unicodedata

from openpyxl import load_workbook

from excel_generator import MAX_ALIMENTOS, MAX_PERSONAS, MAX_SINTOMAS, PERIODOS


def _clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value != value:
        return ""
    return str(value).strip()


def _norm(value: Any) -> str:
    text = _clean(value)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"\s+", " ", text).strip().casefold()
    return text


def _truthy(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return _norm(value) in {"x", "si", "s", "1", "true", "verdadero", "yes"}


def _normalize_sex(value: Any) -> str:
    n = _norm(value)
    if n in {"f", "femenino", "femenina", "mujer"}:
        return "F"
    if n in {"m", "masculino", "masculina", "hombre"}:
        return "M"
    if n in {"otro", "otra", "intersexual", "no binario", "no binaria"}:
        return "Otro"
    return _clean(value)


def _parse_time(value: Any):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)
    if isinstance(value, (int, float)) and 0 <= float(value) < 1:
        total_minutes = int(round(float(value) * 24 * 60)) % (24 * 60)
        return time(total_minutes // 60, total_minutes % 60)
    text = _clean(value).lower().replace(";", ":")
    text = text.replace("a. m.", "am").replace("p. m.", "pm")
    text = text.replace("a.m.", "am").replace("p.m.", "pm")
    text = re.sub(r"\s+", "", text)
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M%p", "%I%p"):
        try:
            return datetime.strptime(text.upper(), fmt).time().replace(second=0, microsecond=0)
        except ValueError:
            pass
    return None


def _parse_date(value: Any, base_date: date | None = None):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        number = float(value)
        # En el Anexo 2 diligenciado con frecuencia se escribe solamente el día del mes.
        if number.is_integer() and 1 <= number <= 31 and base_date:
            day = int(number)
            year, month = base_date.year, base_date.month
            # Tolerar cruce de mes alrededor de la fecha de ocurrencia.
            if day < base_date.day - 15:
                month += 1
                if month == 13:
                    month = 1
                    year += 1
            elif day > base_date.day + 15:
                month -= 1
                if month == 0:
                    month = 12
                    year -= 1
            try:
                return date(year, month, day)
            except ValueError:
                return None
    text = _clean(value).split(" ")[0]
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def _three_part_date(ws, row: int):
    """Lee la fecha tripartita del encabezado sin depender de columnas fijas.

    El Anexo 2 circula en variantes donde día/mes/año quedan en G-H-J,
    G-I-K u otras posiciones por celdas combinadas. Se toman los primeros
    tres componentes numéricos plausibles después de la etiqueta "Día".
    """
    values = [ws.cell(row, c).value for c in range(1, min(ws.max_column, 20) + 1)]
    nums = []
    for value in values:
        if isinstance(value, (date, datetime)):
            return value.date() if isinstance(value, datetime) else value
        try:
            if value in (None, ""):
                continue
            n = int(float(value))
        except (TypeError, ValueError):
            continue
        nums.append(n)

    # Buscar una combinación día-mes-año válida preservando el orden.
    for i in range(len(nums)):
        d = nums[i]
        if not 1 <= d <= 31:
            continue
        for j in range(i + 1, len(nums)):
            m = nums[j]
            if not 1 <= m <= 12:
                continue
            for k in range(j + 1, len(nums)):
                y = nums[k]
                if 0 <= y < 100:
                    y += 2000
                if not 1900 <= y <= 2100:
                    continue
                try:
                    return date(y, m, d)
                except ValueError:
                    continue
    return None


def _find_sheet(wb, preferred: str, fallback_index: int):
    if preferred in wb.sheetnames:
        return wb[preferred]
    if len(wb.worksheets) > fallback_index:
        return wb.worksheets[fallback_index]
    raise ValueError(f"No se encontró la hoja requerida: {preferred}.")


def _is_signature_row(ws, row: int) -> bool:
    texts = " ".join(_clean(ws.cell(row, c).value) for c in range(1, min(ws.max_column, 40) + 1)).casefold()
    return "firma del encuestador" in texts or "firma del encuest" in texts or "nombre y apellidos- telefono" in texts


def _find_header_col(ws, row: int, names: tuple[str, ...]) -> int | None:
    wanted = tuple(_norm(x) for x in names)
    for col in range(1, ws.max_column + 1):
        val = _norm(ws.cell(row, col).value)
        if val and any(w == val or w in val for w in wanted):
            return col
    return None


def _col_letter(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _detect_person_layout(ws):
    # Columnas base estables del Anexo 2.
    no_col = _find_header_col(ws, 10, ("no.", "no")) or 1
    name_col = _find_header_col(ws, 10, ("nombres y apellidos",)) or 2
    id_col = _find_header_col(ws, 10, ("identificacion",)) or 3
    age_col = _find_header_col(ws, 10, ("edad",)) or 4
    sex_col = _find_header_col(ws, 10, ("sexo",)) or 5
    address_col = _find_header_col(ws, 10, ("direccion y telefono", "direccion")) or 6
    day_col = _find_header_col(ws, 10, ("dia",)) or 7
    hour_col = _find_header_col(ws, 10, ("hora",)) or 8

    # Detectar los bloques de estado por los títulos de la fila 9. Esto permite leer
    # tanto el formato con U:X como el formato ampliado con AA:AD.
    enfermo_col = _find_header_col(ws, 9, ("enfermo",))
    consulta_col = _find_header_col(ws, 9, ("consulta", "consulto", "consultó"))
    hosp_col = _find_header_col(ws, 9, ("hospitalizado", "hospitalizacion", "hospitalización"))
    muestra_col = _find_header_col(ws, 9, ("muestra",))
    if not enfermo_col:
        # Fallback por encabezados Si/No continuos al final del bloque.
        for c in range(hour_col + 1, ws.max_column + 1):
            if _norm(ws.cell(10, c).value) in {"si/no", "si no"}:
                enfermo_col = c
                consulta_col = c + 1
                hosp_col = c + 2
                muestra_col = c + 3
                break
    enfermo_col = enfermo_col or max(hour_col + 2, ws.max_column - 7)
    consulta_col = consulta_col or enfermo_col + 1
    hosp_col = hosp_col or enfermo_col + 2
    muestra_col = muestra_col or enfermo_col + 3

    symptoms = []
    symptom_cols = []
    for c in range(hour_col + 1, enfermo_col):
        label = _clean(ws.cell(10, c).value)
        if label and _norm(label) not in {"si/no", "si no"}:
            symptoms.append(label)
            symptom_cols.append(c)
        if len(symptoms) >= MAX_SINTOMAS:
            break

    return {
        "no": no_col, "name": name_col, "id": id_col, "age": age_col, "sex": sex_col,
        "address": address_col, "day": day_col, "hour": hour_col,
        "enfermo": enfermo_col, "consulta": consulta_col, "hosp": hosp_col, "muestra": muestra_col,
        "symptoms": symptoms, "symptom_cols": symptom_cols,
    }


def _detect_food_layout(ws):
    no_col = _find_header_col(ws, 10, ("no.", "no")) or 1
    id_col = _find_header_col(ws, 10, ("identificacion",)) or 2
    period_col = 3
    day_col = _find_header_col(ws, 10, ("dia",)) or 4
    hour_col = _find_header_col(ws, 10, ("hora",)) or 5
    place_col = 6

    foods, cols = [], []
    # En algunos archivos los encabezados continúan en la fila 11 (p. ej. TINTO/AROMÁTICA).
    for c in range(place_col + 1, ws.max_column + 1):
        label = _clean(ws.cell(10, c).value) or _clean(ws.cell(11, c).value)
        if label:
            n = _norm(label)
            if n not in {"logo et", "logo", "alimentos consumidos"}:
                foods.append(label)
                cols.append(c)
        if len(foods) >= MAX_ALIMENTOS:
            break
    return {
        "no": no_col, "id": id_col, "period": period_col, "day": day_col,
        "hour": hour_col, "place": place_col, "foods": foods, "food_cols": cols,
    }


def parse_consumer_excel(file_bytes: bytes) -> dict:
    """Convierte un Anexo 2 ETA diligenciado al modelo interno de la app."""
    try:
        wb = load_workbook(BytesIO(file_bytes), data_only=False)
    except Exception as exc:
        raise ValueError("No fue posible leer el archivo. Carga la Encuesta de Consumidores en formato .xlsx.") from exc

    ws1 = _find_sheet(wb, "Hoja1", 0)
    ws3 = wb["Hoja3"] if "Hoja3" in wb.sheetnames else (wb.worksheets[-1] if len(wb.worksheets) >= 2 else None)
    if ws3 is None:
        raise ValueError("El archivo no contiene la segunda pestaña de alimentos.")

    general = {
        "fecha_ocurrencia": _three_part_date(ws1, 6),
        "fecha_notificacion": _three_part_date(ws1, 7),
        "fecha_investigacion": _three_part_date(ws1, 8),
    }
    base_date = general.get("fecha_ocurrencia") or general.get("fecha_investigacion") or general.get("fecha_notificacion")

    p_layout = _detect_person_layout(ws1)
    symptoms = p_layout["symptoms"][:MAX_SINTOMAS]
    persons: list[dict] = []
    source_nos: list[int] = []

    for row in range(12, ws1.max_row + 1):
        if _is_signature_row(ws1, row):
            break
        name = _clean(ws1.cell(row, p_layout["name"]).value)
        ident = _clean(ws1.cell(row, p_layout["id"]).value)
        age = ws1.cell(row, p_layout["age"]).value
        sex = _normalize_sex(ws1.cell(row, p_layout["sex"]).value)
        address = _clean(ws1.cell(row, p_layout["address"]).value)
        symptom_date = _parse_date(ws1.cell(row, p_layout["day"]).value, base_date)
        symptom_time = _parse_time(ws1.cell(row, p_layout["hour"]).value)
        enfermo_raw = ws1.cell(row, p_layout["enfermo"]).value
        consulta_raw = ws1.cell(row, p_layout["consulta"]).value
        hosp_raw = ws1.cell(row, p_layout["hosp"]).value
        sample = _clean(ws1.cell(row, p_layout["muestra"]).value)

        selected_symptoms = []
        for symptom, col in zip(symptoms, p_layout["symptom_cols"]):
            if _truthy(ws1.cell(row, col).value):
                selected_symptoms.append(symptom)

        meaningful = any([
            name, ident, _clean(age), sex, address, symptom_date, symptom_time,
            selected_symptoms, _clean(enfermo_raw), _clean(consulta_raw), _clean(hosp_raw), sample,
        ])
        if not meaningful:
            continue

        raw_no = ws1.cell(row, p_layout["no"]).value
        try:
            source_no = int(float(raw_no))
        except (TypeError, ValueError):
            source_no = len(persons) + 1

        persons.append({
            "Nombres y apellidos": name,
            "Identificación": ident,
            "Edad": "" if age is None else age,
            "Sexo": sex,
            "Dirección y teléfono": address,
            "Día síntomas": symptom_date,
            "Hora síntomas": symptom_time,
            "selected_symptoms": selected_symptoms,
            "Enfermo": _truthy(enfermo_raw),
            "Consulta": _truthy(consulta_raw),
            "Hospitalizado": _truthy(hosp_raw),
            "Muestra": sample,
        })
        source_nos.append(source_no)
        if len(persons) >= MAX_PERSONAS:
            break

    source_to_new = {no: i for i, no in enumerate(source_nos)}
    consumptions = [
        {period: {"Día": None, "Hora": None, "Lugar de consumo": "", "selected_foods": []} for period in PERIODOS}
        for _ in persons
    ]

    f_layout = _detect_food_layout(ws3)
    foods = f_layout["foods"][:MAX_ALIMENTOS]
    current_no = None
    for row in range(12, ws3.max_row + 1):
        if _is_signature_row(ws3, row):
            break
        raw_no = ws3.cell(row, f_layout["no"]).value
        if raw_no not in (None, ""):
            try:
                current_no = int(float(raw_no))
            except (TypeError, ValueError):
                current_no = None
        if current_no is None or current_no not in source_to_new:
            continue

        period_text = _norm(ws3.cell(row, f_layout["period"]).value)
        period = None
        if "dia de los sintomas" in period_text:
            period = PERIODOS[0]
        elif "un dia antes" in period_text:
            period = PERIODOS[1]
        elif "dos dias antes" in period_text or "2 dias antes" in period_text:
            period = PERIODOS[2]
        elif period_text:
            period = next((p for p in PERIODOS if _norm(p) == period_text), None)
        if period is None:
            continue

        selected_foods = []
        for food, col in zip(foods, f_layout["food_cols"]):
            if _truthy(ws3.cell(row, col).value):
                selected_foods.append(food)

        idx = source_to_new[current_no]
        consumptions[idx][period] = {
            "Día": _parse_date(ws3.cell(row, f_layout["day"]).value, base_date),
            "Hora": _parse_time(ws3.cell(row, f_layout["hour"]).value),
            "Lugar de consumo": _clean(ws3.cell(row, f_layout["place"]).value),
            "selected_foods": selected_foods,
        }

    # Firma del encuestador si existe.
    for row in range(12, ws1.max_row + 1):
        if _is_signature_row(ws1, row):
            for col in range(1, min(ws1.max_column, 40) + 1):
                text = _clean(ws1.cell(row, col).value)
                if ":" in text:
                    left, right = text.split(":", 1)
                    if "nombre" in _norm(left):
                        general["encuestador"] = right.strip()
                    elif "tele" in _norm(left):
                        general["telefono_encuestador"] = right.strip()
            break

    warnings: list[str] = []
    if not persons:
        warnings.append("No se encontraron personas diligenciadas en la primera pestaña.")
    if not symptoms:
        warnings.append("No se identificaron encabezados de signos y síntomas.")
    if not foods:
        warnings.append("No se identificaron alimentos en la segunda pestaña.")

    incomplete = sum(1 for p in persons if not _clean(p.get("Edad")) or not _clean(p.get("Sexo")))
    if incomplete:
        warnings.append(f"Hay {incomplete} persona(s) sin edad o sexo; el informe de 72 horas puede mostrar diferencias en la tabla por edad y género hasta completar esos datos.")

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
    ident = _norm(person.get("Identificación"))
    if ident:
        return "id:" + ident
    name = _norm(person.get("Nombres y apellidos"))
    return "name:" + name if name else ""
