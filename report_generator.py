from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime, time, timedelta
from io import BytesIO
from pathlib import Path
from statistics import mean, median
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile
import math
import re
import unicodedata
import xml.etree.ElementTree as ET

AGE_GROUPS = [
    ("Menor de 1 año", None, 0),
    ("De 1 a 4 años", 1, 4),
    ("De 5 a 9 años", 5, 9),
    ("De 10 a 19 años", 10, 19),
    ("De 20 a 49 años", 20, 49),
    ("De 50 a 74 años", 50, 74),
    ("De 75 años y más", 75, None),
]

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)


def _clean(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _norm(v: Any) -> str:
    s = unicodedata.normalize("NFKD", _clean(v)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip().casefold()


def _bool(v: Any) -> bool:
    if isinstance(v, bool):
        return v
    return _norm(v) in {"si", "s", "1", "true", "x", "yes", "verdadero"}


def _safe_float(v: Any):
    try:
        if v in (None, ""):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def _parse_date(v: Any):
    if not v:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    text = _clean(v)
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text[:10], fmt).date()
        except ValueError:
            pass
    return None


def _parse_time(v: Any):
    if not v:
        return None
    if isinstance(v, datetime):
        return v.time().replace(second=0, microsecond=0)
    if isinstance(v, time):
        return v.replace(second=0, microsecond=0)
    if isinstance(v, (int, float)) and 0 <= float(v) < 1:
        minutes = int(round(float(v) * 24 * 60)) % (24 * 60)
        return time(minutes // 60, minutes % 60)
    text = _clean(v).lower().replace(";", ":").replace(" ", "")
    text = text.replace("a.m.", "am").replace("p.m.", "pm")
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M%p"):
        try:
            return datetime.strptime(text.upper(), fmt).time().replace(second=0, microsecond=0)
        except ValueError:
            pass
    return None


def _fmt_date(v: Any) -> str:
    d = _parse_date(v)
    return d.strftime("%d/%m/%Y") if d else _clean(v)


def _fmt_time(v: Any) -> str:
    t = _parse_time(v)
    return t.strftime("%H:%M") if t else _clean(v)


def _active_person(p: dict) -> bool:
    return bool(_clean(p.get("Nombres y apellidos")) or _clean(p.get("Identificación")))


def _age_group(age) -> str:
    a = _safe_float(age)
    if a is None:
        return "Sin dato"
    if a < 1:
        return AGE_GROUPS[0][0]
    for label, low, high in AGE_GROUPS[1:]:
        if high is None and a >= low:
            return label
        if low <= a <= high:
            return label
    return "Sin dato"


def _sex_code(v: Any) -> str:
    n = _norm(v)
    if n in {"m", "masculino", "hombre", "male"}:
        return "M"
    if n in {"f", "femenino", "mujer", "female"}:
        return "F"
    return ""


def _datetime_pair(dv: Any, tv: Any):
    d = _parse_date(dv)
    t = _parse_time(tv)
    if not d:
        return None
    return datetime.combine(d, t or time(0, 0))


def _rr_ci(a: float, b: float, c: float, d: float, z: float = 1.96):
    try:
        rr = (a / (a + b)) / (c / (c + d))
        se = math.sqrt((1 / a) - (1 / (a + b)) + (1 / c) - (1 / (c + d)))
        return rr, math.exp(math.log(rr) - z * se), math.exp(math.log(rr) + z * se)
    except (ZeroDivisionError, ValueError):
        return None, None, None


def _or_ci(a: float, b: float, c: float, d: float, z: float = 1.96):
    try:
        orr = (a * d) / (b * c)
        se = math.sqrt(1 / a + 1 / b + 1 / c + 1 / d)
        return orr, math.exp(math.log(orr) - z * se), math.exp(math.log(orr) + z * se)
    except (ZeroDivisionError, ValueError):
        return None, None, None


def analyze(payload: dict) -> dict:
    all_persons = payload.get("persons", [])
    active_by_no = {i: p for i, p in enumerate(all_persons, start=1) if _active_person(p)}
    persons = list(active_by_no.values())
    cases = [p for p in persons if _bool(p.get("Enfermo"))]
    noncases = [p for p in persons if not _bool(p.get("Enfermo"))]
    foods = [_clean(f) for f in payload.get("foods", []) if _clean(f)]
    symptoms = [_clean(s) for s in payload.get("symptoms", []) if _clean(s)]
    total = len(persons)
    ncases = len(cases)

    symptom_counts = []
    for symptom in symptoms:
        n = sum(1 for p in cases if _bool(p.get(symptom)))
        symptom_counts.append((symptom, n, (n / ncases * 100) if ncases else 0.0))
    symptom_counts.sort(key=lambda x: (-x[1], x[0].casefold()))

    age_sex = {
        label: {"M_exp": 0, "M_cases": 0, "F_exp": 0, "F_cases": 0}
        for label, _, _ in AGE_GROUPS
    }
    missing_age = missing_sex = 0
    for p in persons:
        grp = _age_group(p.get("Edad"))
        sx = _sex_code(p.get("Sexo"))
        if grp == "Sin dato":
            missing_age += 1
        if not sx:
            missing_sex += 1
        if grp in age_sex and sx:
            age_sex[grp][f"{sx}_exp"] += 1
            if _bool(p.get("Enfermo")):
                age_sex[grp][f"{sx}_cases"] += 1

    consumed_by_no = defaultdict(set)
    consumption_events = defaultdict(list)
    for row in payload.get("food_rows", []):
        try:
            no = int(float(row.get("No.")))
        except (TypeError, ValueError):
            continue
        dt = _datetime_pair(row.get("Día"), row.get("Hora"))
        for food in foods:
            if _bool(row.get(food)):
                consumed_by_no[no].add(food)
                if dt:
                    consumption_events[(no, food)].append(dt)

    food_analysis = []
    for food in foods:
        a = b = c = d = 0
        for no, p in active_by_no.items():
            sick = _bool(p.get("Enfermo"))
            exposed = food in consumed_by_no.get(no, set())
            if sick and exposed:
                a += 1
            elif (not sick) and exposed:
                b += 1
            elif sick and (not exposed):
                c += 1
            else:
                d += 1

        ta_exp = a / (a + b) if (a + b) else None
        ta_unexp = c / (c + d) if (c + d) else None
        corrected = any(x == 0 for x in (a, b, c, d))
        aa, bb, cc, dd = (a + 1, b + 1, c + 1, d + 1) if corrected else (a, b, c, d)
        ta_exp_corr = aa / (aa + bb) if (aa + bb) else None
        ta_unexp_corr = cc / (cc + dd) if (cc + dd) else None
        rr, rr_lo, rr_hi = _rr_ci(aa, bb, cc, dd)
        orr, or_lo, or_hi = _or_ci(aa, bb, cc, dd)
        food_analysis.append({
            "food": food,
            "a": a, "b": b, "c": c, "d": d,
            "input_a": aa, "input_b": bb, "input_c": cc, "input_d": dd,
            "attack_exposed": ta_exp,
            "attack_unexposed": ta_unexp,
            "attack_exposed_official": ta_exp_corr,
            "attack_unexposed_official": ta_unexp_corr,
            "risk_difference": (ta_exp - ta_unexp) if ta_exp is not None and ta_unexp is not None else None,
            "risk_difference_official": (ta_exp_corr - ta_unexp_corr) if ta_exp_corr is not None and ta_unexp_corr is not None else None,
            "corrected": corrected,
            "rr": rr, "rr_low": rr_lo, "rr_high": rr_hi,
            "or": orr, "or_low": or_lo, "or_high": or_hi,
        })

    food_analysis.sort(key=lambda x: (x["rr"] is not None, x["rr"] or -1, x["or"] or -1), reverse=True)
    top_food = food_analysis[0] if food_analysis else None

    onset_dts = []
    onset_dates = []
    for p in cases:
        dt = _datetime_pair(p.get("Día síntomas"), p.get("Hora síntomas"))
        if dt:
            onset_dts.append(dt)
            onset_dates.append(dt.date())

    incubation_hours = []
    if top_food:
        top_name = top_food["food"]
        for no, p in active_by_no.items():
            if not _bool(p.get("Enfermo")):
                continue
            onset = _datetime_pair(p.get("Día síntomas"), p.get("Hora síntomas"))
            if not onset:
                continue
            candidates = [x for x in consumption_events.get((no, top_name), []) if x <= onset]
            if candidates:
                exposure = max(candidates)
                hours = (onset - exposure).total_seconds() / 3600
                if 0 <= hours <= 240:
                    incubation_hours.append(hours)

    incubation = None
    if incubation_hours:
        incubation = {
            "min": min(incubation_hours),
            "max": max(incubation_hours),
            "mean": mean(incubation_hours),
            "median": median(incubation_hours),
            "values": incubation_hours,
        }

    return {
        "persons": persons,
        "cases": cases,
        "noncases": noncases,
        "total_exposed": total,
        "total_cases": ncases,
        "attack_rate": (ncases / total * 100) if total else 0.0,
        "hospitalized": sum(1 for p in cases if _bool(p.get("Hospitalizado"))),
        "consulted": sum(1 for p in cases if _bool(p.get("Consulta"))),
        "with_sample": sum(1 for p in cases if _clean(p.get("Muestra"))),
        "symptom_counts": symptom_counts,
        "age_sex": age_sex,
        "missing_age": missing_age,
        "missing_sex": missing_sex,
        "food_analysis": food_analysis,
        "top_food": top_food,
        "first_onset": min(onset_dates) if onset_dates else None,
        "last_onset": max(onset_dates) if onset_dates else None,
        "incubation": incubation,
    }


def build_suggestions(payload: dict, analysis: dict | None = None) -> dict:
    a = analysis or analyze(payload)
    g = payload.get("general", {})
    place = _clean(g.get("lugar_brote")) or _clean(g.get("localidad")) or "el lugar investigado"
    top_sym = [f"{s} ({n}/{a['total_cases']}, {pct:.1f}%)" for s, n, pct in a["symptom_counts"][:5] if n]
    sym_text = ", ".join(top_sym) if top_sym else "sin patrón clínico consolidado"

    summary = (
        f"Se consolidaron {a['total_exposed']} personas expuestas y {a['total_cases']} personas enfermas, "
        f"para una tasa de ataque general de {a['attack_rate']:.1f}%. "
        f"Los signos y síntomas de mayor frecuencia fueron: {sym_text}."
    )

    if a["top_food"]:
        x = a["top_food"]
        ta_e = x["attack_exposed"] * 100 if x["attack_exposed"] is not None else 0
        ta_u = x["attack_unexposed"] * 100 if x["attack_unexposed"] is not None else 0
        rr_txt = f"{x['rr']:.2f}" if x.get("rr") is not None else "no estimable"
        ci_txt = (f"IC95% {x['rr_low']:.2f}–{x['rr_high']:.2f}"
                  if x.get("rr_low") is not None and x.get("rr_high") is not None else "IC95% no estimable")
        raw_counts = (
            f"{x['a']} de {x['a'] + x['b']} consumidores enfermaron ({ta_e:.1f}%) y "
            f"{x['c']} de {x['c'] + x['d']} no consumidores enfermaron ({ta_u:.1f}%)."
        )
        correction = ""
        if x["corrected"]:
            ce = x["attack_exposed_official"] * 100 if x.get("attack_exposed_official") is not None else 0
            cu = x["attack_unexposed_official"] * 100 if x.get("attack_unexposed_official") is not None else 0
            correction = (
                f" Como una celda de la tabla 2×2 fue cero, el Anexo 3 indica sumar 1 a las cuatro celdas; "
                f"con esa corrección, las tasas usadas por el formato son {ce:.1f}% y {cu:.1f}%, con RR {rr_txt} ({ci_txt})."
            )
        else:
            correction = f" El RR estimado es {rr_txt} ({ci_txt})."

        precision = ""
        if x.get("rr_low") is not None and x.get("rr_high") is not None:
            if x["rr_low"] <= 1 <= x["rr_high"]:
                precision = " El intervalo de confianza incluye 1, por lo que la asociación no es concluyente por sí sola y debe interpretarse con cautela."
            else:
                precision = " El intervalo de confianza no incluye 1; aun así, la causalidad debe confirmarse con la investigación de campo y laboratorio."

        hypothesis = (
            f"El alimento con mayor asociación epidemiológica observada es {x['food']}. {raw_counts}"
            + correction + precision
            + " Este resultado orienta la hipótesis y debe correlacionarse con IVC, trazabilidad, patrón clínico, periodo de incubación y resultados de laboratorio antes de atribuir causalidad."
        )

        top3 = []
        for item in a["food_analysis"][:3]:
            te = item["attack_exposed"] * 100 if item["attack_exposed"] is not None else 0
            tu = item["attack_unexposed"] * 100 if item["attack_unexposed"] is not None else 0
            rr = f"{item['rr']:.2f}" if item.get("rr") is not None else "no estimable"
            top3.append(f"{item['food']}: TA consumidores {te:.1f}%, TA no consumidores {tu:.1f}%, RR {rr}")
        food_analysis_text = "Análisis por alimento: " + "; ".join(top3) + ". " + hypothesis
        possible_foods = ", ".join(item["food"] for item in a["food_analysis"][:3])
    else:
        hypothesis = "No hay información suficiente de consumo para orientar un alimento probable. Complete la encuesta de consumidores y revise las exposiciones comunes."
        food_analysis_text = hypothesis
        possible_foods = ", ".join(payload.get("foods", []))

    first = a.get("first_onset")
    last = a.get("last_onset")
    period = ""
    if first:
        period = f" con inicio de síntomas desde {first.strftime('%d/%m/%Y')}"
        if last and last != first:
            period += f" hasta {last.strftime('%d/%m/%Y')}"
    definition = (
        f"Persona que estuvo expuesta a alimentos o bebidas en {place}{period} y presentó uno o más signos o síntomas "
        f"compatibles con el patrón clínico identificado ({', '.join(s for s, n, _ in a['symptom_counts'][:5] if n)}). "
        "Definición sugerida para revisión y ajuste por el equipo investigador."
    )

    inc = a.get("incubation")
    inc_text = ""
    if inc:
        inc_text = (
            f" Para el alimento con mayor asociación, el período de incubación calculable en los registros disponibles fue "
            f"mínimo {inc['min']:.1f} h, máximo {inc['max']:.1f} h, promedio {inc['mean']:.1f} h y mediana {inc['median']:.1f} h."
        )

    analysis_text = summary + " " + food_analysis_text + inc_text
    final_summary = summary + " " + (hypothesis if a.get("top_food") else "")
    final_conclusion = (
        "La investigación epidemiológica permite orientar la fuente probable con base en la comparación de tasas de ataque entre consumidores y no consumidores. "
        "La conclusión etiológica y causal definitiva debe integrar resultados de laboratorio, inspección sanitaria, trazabilidad y demás hallazgos de campo."
    )

    return {
        "resumen": summary,
        "posibles_alimentos": possible_foods,
        "hipotesis_inicial": hypothesis,
        "definicion_caso": definition,
        "analisis_resultados": analysis_text,
        "resumen_final": final_summary,
        "descripcion_brote": analysis_text,
        "conclusiones_72": final_conclusion,
        "conclusiones_final": final_conclusion,
    }


class XlsxTemplatePatcher:
    """Edita únicamente valores de celdas del XLSX oficial, conservando formato, imágenes, fórmulas y gráficos."""

    def __init__(self, template_path: Path):
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise FileNotFoundError(f"No se encontró la plantilla oficial: {self.template_path.name}")
        with ZipFile(self.template_path, "r") as zin:
            self.files = {name: zin.read(name) for name in zin.namelist()}
        self.sheet_paths = self._sheet_paths()
        self.trees: dict[str, ET.Element] = {}

    def _sheet_paths(self):
        wb = ET.fromstring(self.files["xl/workbook.xml"])
        rels = ET.fromstring(self.files["xl/_rels/workbook.xml.rels"])
        rmap = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
        out = {}
        for sheet in wb.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet"):
            rid = sheet.attrib[f"{{{NS_REL}}}id"]
            target = rmap[rid].lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            out[sheet.attrib["name"]] = target
        return out

    @staticmethod
    def _split_ref(ref: str):
        m = re.fullmatch(r"([A-Z]+)(\d+)", ref.upper())
        if not m:
            raise ValueError(f"Referencia de celda inválida: {ref}")
        col, row = m.group(1), int(m.group(2))
        n = 0
        for ch in col:
            n = n * 26 + ord(ch) - 64
        return n, row

    def _root(self, sheet_name: str):
        if sheet_name not in self.trees:
            self.trees[sheet_name] = ET.fromstring(self.files[self.sheet_paths[sheet_name]])
        return self.trees[sheet_name]

    def _cell(self, sheet_name: str, ref: str):
        root = self._root(sheet_name)
        sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
        col_num, row_num = self._split_ref(ref)
        row = sheet_data.find(f"{{{NS_MAIN}}}row[@r='{row_num}']")
        if row is None:
            row = ET.Element(f"{{{NS_MAIN}}}row", {"r": str(row_num)})
            rows = list(sheet_data)
            pos = next((i for i, x in enumerate(rows) if int(x.attrib.get("r", 0)) > row_num), len(rows))
            sheet_data.insert(pos, row)
        cell = row.find(f"{{{NS_MAIN}}}c[@r='{ref.upper()}']")
        if cell is None:
            cell = ET.Element(f"{{{NS_MAIN}}}c", {"r": ref.upper()})
            cells = list(row)
            pos = len(cells)
            for i, other in enumerate(cells):
                other_col, _ = self._split_ref(other.attrib["r"])
                if other_col > col_num:
                    pos = i
                    break
            row.insert(pos, cell)
        return cell

    @staticmethod
    def _clear_children(cell):
        for child in list(cell):
            cell.remove(child)

    def set_text(self, sheet: str, ref: str, value: Any):
        cell = self._cell(sheet, ref)
        self._clear_children(cell)
        cell.set("t", "inlineStr")
        isel = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
        t = ET.SubElement(isel, f"{{{NS_MAIN}}}t")
        text = _clean(value)
        if text.startswith(" ") or text.endswith(" ") or "\n" in text:
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text

    def set_number(self, sheet: str, ref: str, value: Any):
        cell = self._cell(sheet, ref)
        self._clear_children(cell)
        cell.attrib.pop("t", None)
        v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
        v.text = str(int(value) if isinstance(value, (int, float)) and float(value).is_integer() else value)

    def set(self, sheet: str, ref: str, value: Any):
        if value is None or value == "":
            self.set_text(sheet, ref, "")
        elif isinstance(value, bool):
            self.set_text(sheet, ref, "X" if value else "")
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            self.set_number(sheet, ref, value)
        else:
            self.set_text(sheet, ref, value)

    def mark(self, sheet: str, ref: str, selected: bool):
        self.set_text(sheet, ref, "X" if selected else "")

    def save_bytes(self) -> bytes:
        # Pedir recálculo completo al abrir en Excel.
        wb = ET.fromstring(self.files["xl/workbook.xml"])
        calc = wb.find(f"{{{NS_MAIN}}}calcPr")
        if calc is None:
            calc = ET.SubElement(wb, f"{{{NS_MAIN}}}calcPr")
        calc.set("calcMode", "auto")
        calc.set("fullCalcOnLoad", "1")
        calc.set("forceFullCalc", "1")
        calc.set("calcId", "0")
        self.files["xl/workbook.xml"] = ET.tostring(wb, encoding="utf-8", xml_declaration=True)

        for name, root in self.trees.items():
            self.files[self.sheet_paths[name]] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        out = BytesIO()
        with ZipFile(out, "w", ZIP_DEFLATED) as zout:
            for name, data in self.files.items():
                zout.writestr(name, data)
        return out.getvalue()


def _top_foods_text(a: dict, n=3):
    return ", ".join(x["food"] for x in a["food_analysis"][:n])


def _epi_week(g: dict, rf: dict):
    if rf.get("semana_epidemiologica") not in (None, "", 0):
        return rf.get("semana_epidemiologica")
    d = _parse_date(g.get("fecha_ocurrencia"))
    return d.isocalendar().week if d else ""


def _fill_preliminary(x: XlsxTemplatePatcher, payload: dict, a: dict, rf: dict, sug: dict):
    s = "Inf. Preliminar"
    g = payload.get("general", {})
    x.set_text(s, "F12", g.get("departamento", ""))
    x.set_text(s, "F13", g.get("municipio", ""))
    x.set_text(s, "F14", g.get("localidad", ""))
    immediate = rf.get("notificacion_inmediata", "")
    x.mark(s, "I16", _norm(immediate) == "si")
    x.mark(s, "K16", _norm(immediate) == "no")
    x.set_text(s, "H18", _fmt_date(g.get("fecha_ocurrencia")))
    x.set_text(s, "H20", _fmt_date(g.get("fecha_deteccion")))
    x.set_text(s, "H22", _fmt_date(rf.get("fecha_notificacion_sat") or g.get("fecha_notificacion")))
    x.set_text(s, "H24", _fmt_date(rf.get("fecha_notificacion_sivigila") or g.get("fecha_notificacion")))
    x.set(s, "H26", _epi_week(g, rf))
    x.set_text(s, "E28", g.get("lugar_brote", ""))
    x.set_text(s, "E30", g.get("direccion_brote", ""))
    x.set_text(s, "K30", g.get("telefono_brote", ""))
    x.set_text(s, "H32", _fmt_date(a.get("first_onset")))
    x.set_number(s, "E34", a["total_exposed"])

    upgd = int(rf.get("casos_upgd", 0) or 0)
    bac = rf.get("casos_bac")
    bac = a["total_cases"] - upgd if bac in (None, "") else int(bac or 0)
    if upgd + bac != a["total_cases"]:
        # No inventar casos: si la distribución UPGD/BAC no cuadra, dejar el total real en BAC como valor de apoyo.
        bac = max(a["total_cases"] - upgd, 0)
    x.set_number(s, "H34", upgd)
    x.set_number(s, "K34", bac)

    deaths = int(rf.get("casos_muertos", 0) or 0)
    alive = max(a["total_cases"] - deaths, 0)
    home = max(a["total_cases"] - a["hospitalized"] - deaths, 0)
    x.set_number(s, "E38", alive)
    x.set_number(s, "H38", deaths)
    x.set_number(s, "K38", a["hospitalized"])
    x.set_number(s, "N38", home)

    x.set_text(s, "F40", rf.get("posibles_alimentos") or sug["posibles_alimentos"])
    symptom_text = ", ".join(f"{name}: {n} ({pct:.1f}%)" for name, n, pct in a["symptom_counts"] if n)
    x.set_text(s, "F42", symptom_text)
    x.set_text(s, "E44", rf.get("antecedentes") or sug["resumen"])
    x.set_text(s, "E47", rf.get("hipotesis_inicial") or sug["hipotesis_inicial"])
    x.set_text(s, "E49", rf.get("medidas_control", ""))
    x.set_text(s, "E51", rf.get("otra_informacion", ""))

    x.mark(s, "F53", bool(rf.get("muestras_biologicas")) or a["with_sample"] > 0)
    x.mark(s, "I53", bool(rf.get("muestras_superficies")))
    x.mark(s, "F54", bool(rf.get("muestras_alimentos")))
    x.mark(s, "I54", bool(rf.get("muestras_manipuladores")))

    industrial = rf.get("alimento_industrializado", "")
    x.mark(s, "F56", _norm(industrial) == "si")
    x.mark(s, "F57", _norm(industrial) == "no")
    x.set_text(s, "J56", rf.get("datos_industrializado", ""))
    support = rf.get("requiere_apoyo", "")
    x.mark(s, "F59", _norm(support) == "si")
    x.mark(s, "F60", _norm(support) == "no")
    x.set_text(s, "J59", rf.get("apoyo_instancias", ""))

    x.set_text(s, "H62", rf.get("responsable") or g.get("encuestador", ""))
    x.set_text(s, "H63", rf.get("revisor", ""))
    x.set_text(s, "H64", rf.get("telefono_responsable") or g.get("telefono_encuestador", ""))
    x.set_text(s, "H65", rf.get("email_responsable", ""))


def _fill_age_sex(x: XlsxTemplatePatcher, a: dict):
    s = "Inf. 72 horas"
    for idx, (label, _, _) in enumerate(AGE_GROUPS, start=49):
        d = a["age_sex"].get(label, {})
        x.set_number(s, f"F{idx}", d.get("M_exp", 0))
        x.set_number(s, f"G{idx}", d.get("M_cases", 0))
        x.set_number(s, f"I{idx}", d.get("F_exp", 0))
        x.set_number(s, f"K{idx}", d.get("F_cases", 0))


def _fill_curve(x: XlsxTemplatePatcher, a: dict):
    s = "Inf. 72 horas"
    # El formato original trae intervalos de dos horas entre 0 y 60 h.
    x.mark(s, "D63", False)
    x.mark(s, "F63", True)
    x.mark(s, "H63", False)
    for row in range(68, 98):
        x.set_number(s, f"E{row}", 0)
    inc = a.get("incubation")
    if not inc:
        return
    for hours in inc["values"]:
        # Filas 68..97 representan: 0-2; 2.1-4; ...; 58.1-60.
        if 0 <= hours <= 60:
            bin_idx = min(int(math.ceil(hours / 2.0)) - 1 if hours > 0 else 0, 29)
            row = 68 + bin_idx
            # La plantilla se recalculará; acumulamos aquí el valor de entrada.
            # Como el patcher no lee valores, acumulamos temporalmente en atributo externo.
            key = f"_freq_{row}"
            current = getattr(x, key, 0)
            setattr(x, key, current + 1)
    for row in range(68, 98):
        x.set_number(s, f"E{row}", getattr(x, f"_freq_{row}", 0))


def _fill_food_table(x: XlsxTemplatePatcher, a: dict):
    s = "Inf. 72 horas"
    # Cohorte: contamos expuestos y no expuestos, enfermos y sanos.
    x.mark(s, "E104", True)
    x.mark(s, "E105", False)
    # IC 95 %.
    x.mark(s, "N103", False)
    x.mark(s, "N104", True)
    x.mark(s, "N105", False)

    rows = list(range(110, 125))
    for row in rows:
        for ref in (f"C{row}", f"D{row}", f"E{row}", f"H{row}", f"I{row}"):
            x.set_text(s, ref, "")

    for row, item in zip(rows, a["food_analysis"][: len(rows)]):
        x.set_text(s, f"C{row}", item["food"])
        # El Anexo 3 indica sumar 1 a las cuatro celdas cuando alguna es cero.
        x.set_number(s, f"D{row}", item["input_a"])
        x.set_number(s, f"E{row}", item["input_b"])
        x.set_number(s, f"H{row}", item["input_c"])
        x.set_number(s, f"I{row}", item["input_d"])


def _fill_72h(x: XlsxTemplatePatcher, payload: dict, a: dict, rf: dict, sug: dict):
    s = "Inf. 72 horas"
    g = payload.get("general", {})
    x.set_text(s, "G16", _fmt_date(g.get("fecha_investigacion") or g.get("fecha_notificacion")))
    x.set_text(s, "E22", rf.get("definicion_caso") or sug["definicion_caso"])
    x.set_text(s, "E24", rf.get("manejo_clinico", ""))
    x.set_text(s, "G30", _fmt_date(a.get("first_onset")))
    x.set_text(s, "G31", _fmt_date(a.get("last_onset")))
    x.set_number(s, "K30", a["total_cases"])
    x.set_number(s, "N30", a["total_exposed"])

    for row in range(34, 44):
        x.set_text(s, f"C{row}", "")
        x.set_text(s, f"D{row}", "")
    for row, (name, n, _) in zip(range(34, 44), a["symptom_counts"][:10]):
        x.set_text(s, f"C{row}", name)
        x.set_number(s, f"D{row}", n)

    _fill_age_sex(x, a)
    _fill_curve(x, a)
    _fill_food_table(x, a)

    # Sección de tasa de exposición: solo se usa cuando no existe grupo de no enfermos.
    for row in range(133, 145):
        for ref in (f"C{row}", f"F{row}"):
            x.set_text(s, ref, "")
    if a["total_cases"] and not a["noncases"]:
        for row, item in zip(range(133, 145), a["food_analysis"][:12]):
            x.set_text(s, f"C{row}", item["food"])
            x.set_number(s, f"F{row}", item["a"])

    # Tipo de establecimiento. Si no coincide con una categoría, marcar Otro.
    place_type = _norm(rf.get("tipo_establecimiento", ""))
    check_map = {
        "hogar": "F152", "establecimiento educativo": "J152", "establecimiento penitenciario": "F153",
        "casino institucional": "J153", "establecimiento militar": "F154", "hogar geriatrico": "J154",
        "hogar de bienestar familiar": "F155", "club social": "J155", "restaurante comercial": "F156", "otro": "J156",
    }
    for ref in check_map.values():
        x.mark(s, ref, False)
    if place_type in check_map:
        x.mark(s, check_map[place_type], True)
    elif place_type:
        x.mark(s, check_map["otro"], True)

    # Bloques narrativos del formato oficial.
    x.set_text(s, "C160", rf.get("hallazgos_ambientales", ""))
    x.set_text(s, "C201", rf.get("recomendaciones_72", ""))
    x.set_text(s, "C214", rf.get("conclusiones_72") or sug["conclusiones_72"])
    x.set_text(s, "G226", rf.get("responsable") or g.get("encuestador", ""))
    x.set_text(s, "G227", rf.get("revisor", ""))
    x.set_text(s, "G228", rf.get("telefono_responsable") or g.get("telefono_encuestador", ""))
    x.set_text(s, "G229", rf.get("email_responsable", ""))


def _fill_final(x: XlsxTemplatePatcher, payload: dict, a: dict, rf: dict, sug: dict):
    s = "Inf. final"
    g = payload.get("general", {})
    x.set_text(s, "G16", _fmt_date(rf.get("fecha_agente")))
    x.set_text(s, "G18", _fmt_date(rf.get("fecha_cierre")))
    x.set_number(s, "F23", a["total_exposed"])
    x.set_number(s, "J23", a["total_cases"])
    deaths = int(rf.get("casos_muertos", 0) or 0)
    x.set_number(s, "D25", max(a["total_cases"] - deaths, 0))
    x.set_number(s, "G25", deaths)
    x.set_number(s, "K25", a["hospitalized"])
    x.set_number(s, "P25", max(a["total_cases"] - a["hospitalized"] - deaths, 0))

    source = _norm(rf.get("fuente_transmision", "alimentos"))
    source_map = {"agua": "G27", "alimentos": "G28", "persona a persona": "G29", "contaminacion medio ambiental": "K27", "otro": "K28", "desconocido": "K29"}
    for ref in source_map.values():
        x.mark(s, ref, False)
    x.mark(s, source_map.get(source, "G28"), True)

    mode = _norm(rf.get("modo_transmision", "oral"))
    mode_map = {"oral": "Q27", "oral - fecal": "Q28", "oral-fecal": "Q28", "cruzada": "Q29"}
    for ref in mode_map.values():
        x.mark(s, ref, False)
    if mode in mode_map:
        x.mark(s, mode_map[mode], True)

    state = _norm(rf.get("estado_brote", ""))
    with_agent = "con agente" in state or "con identificacion" in state
    without_agent = "sin agente" in state or "sin identificacion" in state
    x.mark(s, "I31", with_agent)
    x.mark(s, "O31", without_agent)
    x.set_text(s, "I32", rf.get("agente_identificado", "") if with_agent else "")

    x.set_text(s, "E34", rf.get("resumen_final") or sug["resumen_final"])
    x.set_text(s, "E37", rf.get("descripcion_brote") or sug["descripcion_brote"])
    x.set_text(s, "E49", rf.get("factores_determinantes") or rf.get("hallazgos_ambientales", ""))
    x.set_text(s, "C52", rf.get("plan_mejoramiento", ""))
    x.set_text(s, "C55", rf.get("recomendaciones_final") or rf.get("recomendaciones_72", ""))
    x.set_text(s, "C68", rf.get("conclusiones_final") or sug["conclusiones_final"])
    x.set_text(s, "G80", rf.get("responsable") or g.get("encuestador", ""))
    x.set_text(s, "G81", rf.get("revisor", ""))
    x.set_text(s, "G82", rf.get("telefono_responsable") or g.get("telefono_encuestador", ""))
    x.set_text(s, "G83", rf.get("email_responsable", ""))


def build_reports(payload: dict, report_fields: dict | None = None, template_path: Path | None = None) -> tuple[bytes, str]:
    report_fields = report_fields or {}
    if template_path is None:
        template_path = Path(__file__).resolve().parent / "Anexo_3_ETA_OFICIAL.xlsx"
    a = analyze(payload)
    sug = build_suggestions(payload, a)
    x = XlsxTemplatePatcher(Path(template_path))
    _fill_preliminary(x, payload, a, report_fields, sug)
    _fill_72h(x, payload, a, report_fields, sug)
    _fill_final(x, payload, a, report_fields, sug)
    g = payload.get("general", {})
    dt = _fmt_date(g.get("fecha_ocurrencia")).replace("/", "-") or "sin_fecha"
    return x.save_bytes(), f"Anexo_3_ETA_24_72_Final_{dt}.xlsx"
