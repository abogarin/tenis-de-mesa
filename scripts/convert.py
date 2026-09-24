#!/usr/bin/env python3
"""Convert ranking workbooks into the JSON the website reads.

For every .xlsx under inbox/ (including inbox/2025/, inbox/2026/, ...):
  1. Find the ranking table: the "Ranking Final" sheet, or sheets named "Ranking ...",
     or any other sheet with Nombre + Puntos columns.
  2. Work out circuit, season, division (U9…U19, Mayor, Open A/B, Master, PTT),
     gender and stage number. Sources, most trusted first:
       a) "<file>.json" override written by a person
       b) the workbook itself (category title on "Inscripciones", dates on "Grupo" sheets)
       c) "<file>.hint.json" written by fetch_fecoteme.py from the FECOTEME link label
       d) the file name and the inbox/<year>/ folder
     Disagreements between (b) and (c) are reported as warnings.
  3. Save data/stages/<id>.json (a re-upload of the same stage replaces it) and move
     the workbook to archive/<season>/.
Finally all stages are combined into site/data/ranking.json.
"""
import datetime as dt
import json
import re
import shutil
import sys
import unicodedata
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
INBOX, ARCHIVE = ROOT / "inbox", ROOT / "archive"
STAGES, OUT = ROOT / "data" / "stages", ROOT / "site" / "data" / "ranking.json"
WARN_FILE = ROOT / "data" / "warnings.json"

ROMAN = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
ORD_WORDS = {"PRIMER": 1, "PRIMERO": 1, "PRIMERA": 1, "1ER": 1, "1RO": 1, "1RN": 1, "SEGUNDO": 2, "SEGUNDA": 2, "2DO": 2, "2RN": 2,
             "TERCER": 3, "TERCERO": 3, "TERCERA": 3, "3ER": 3, "3RO": 3, "3CER": 3, "3RN": 3, "CUARTO": 4, "CUARTA": 4, "4TO": 4, "4RN": 4,
             "QUINTO": 5, "QUINTA": 5, "5TO": 5, "5RN": 5, "SEXTO": 6, "6TO": 6}
SKIP_SHEETS = re.compile(r"^(INSCRIP|GRUPO|LLAVE|SORTEO|PLANILLA|PREMIA|TABULA)", re.I)


def fold(s):
    s = unicodedata.normalize("NFD", str(s if s is not None else ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").upper()


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", fold(s).lower()).strip("-")


# ---------- reading the ranking table ----------

def table_from_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    for h, r in enumerate(rows[:20]):
        hdr = [fold(c).strip() for c in r]
        iN = next((i for i, c in enumerate(hdr) if "NOMBRE" in c or c in ("JUGADOR", "ATLETA")), None)
        iP = next((i for i, c in enumerate(hdr) if "PUNTO" in c or c in ("PTS", "TOTAL")), None)
        if iN is None or iP is None:
            continue
        iR = next((i for i, c in enumerate(hdr) if c.startswith("RANK") or c in ("POS", "POSICION", "#")), None)
        iC = next((i for i, c in enumerate(hdr) if c.startswith("CARN") or c in ("ID", "CEDULA", "LICENCIA")), None)
        iK = next((i for i, c in enumerate(hdr) if "CLUB" in c or "EQUIPO" in c), None)
        out = []
        for r2 in rows[h + 1:]:
            if len(r2) <= max(iN, iP):
                continue
            name, pts = r2[iN], r2[iP]
            if name is None or not isinstance(pts, (int, float)) or not str(name).strip():
                continue
            name = " ".join(str(name).split())
            carne = None
            if iC is not None and r2[iC] is not None:
                carne = str(r2[iC]).strip()
                carne = carne[:-2] if carne.endswith(".0") else carne
            if not carne:
                carne = "n-" + slug(name)  # no carné column: fall back to the name
            rank = int(r2[iR]) if iR is not None and isinstance(r2[iR], (int, float)) else None
            club = str(r2[iK]).strip() if iK is not None and r2[iK] is not None else ""
            out.append([rank, carne, name, club, round(float(pts), 2) if pts % 1 else int(pts)])
        if out:
            # fill missing ranks from order (ties share the rank)
            if any(r[0] is None for r in out):
                out.sort(key=lambda r: -r[4])
                for i, r in enumerate(out):
                    r[0] = out[i - 1][0] if i and out[i - 1][4] == r[4] else i + 1
            return out
    return None


def read_tables(wb):
    """Return [(sheet_name_or_None, rows)]. More than one entry = several divisions in one file."""
    if "Ranking Final" in wb.sheetnames:
        t = table_from_sheet(wb["Ranking Final"])
        if t:
            return [(None, t)]
    named = [n for n in wb.sheetnames if re.match(r"\s*RANK", fold(n)) and wb[n].sheet_state == "visible"]
    tables = [(n, t) for n in named if (t := table_from_sheet(wb[n]))]
    if tables:
        return [(None, tables[0][1])] if len(tables) == 1 else tables
    cands = [(n, t) for n in wb.sheetnames
             if not SKIP_SHEETS.match(fold(n)) and wb[n].sheet_state == "visible" and (t := table_from_sheet(wb[n]))]
    if len(cands) > 1 and all(bracket_like(n) for n, _ in cands):
        return cands  # one sheet per division, e.g. Master 30-39 / 40-49 / 50-59 / 60+
    return [(None, cands[0][1])] if cands else []


# ---------- metadata ----------

def text_division(t):
    m = re.search(r"\b(?:U|SUB)\s?-?\s?(\d{1,2})\b", t)
    if m:
        return f"U{m.group(1)}"
    if re.search(r"\bPTT\b|PARA\s*TENIS|PARALIMP", t):        # before Master: one PTT label says "Master PTT"
        return "PTT"
    if "MASTER" in t:
        m = re.search(r"\b([3-9]\d)\s*(?:-|\s|A)\s*([3-9]\d)\b", t)
        if m:
            return f"Master {m.group(1)}-{m.group(2)}"
        m = re.search(r"\b([3-9]\d)\s*\+|\+\s*([3-9]\d)\b|MAS\s*DE\s*([3-9]\d)", t)
        if m:
            return f"Master {next(g for g in m.groups() if g)}+"
        return "Master"
    if re.search(r"OPEN|FEMENIN|\bFEM\b", t):
        # Open Femenino divisions: A = Primera, B = Segunda (both names are used on FECOTEME pages)
        m = re.search(r"OPEN\s*(?:FEM\w*\s*)?-?\s*\b([AB])\b|FEMENIN\w*\s*\b([AB])\b|\b([AB])\s*$", t)
        if m:
            return "Open " + next(g for g in m.groups() if g)
        if re.search(r"\bPRIMERA\b", t):
            return "Open A"
        if re.search(r"\bSEGUNDA\b", t):
            return "Open B"
        if re.search(r"OPEN", t):
            return "Open"
    if re.search(r"\bMAYOR\b|\bSENIOR\b|PRIMERA DIVISION", t):
        return "Mayor"
    return None


def bracket_like(sheet_name):
    """A sheet whose name is a division, e.g. '30-39', '60+', 'U13 Fem', 'Open A'."""
    t = fold(sheet_name)
    return bool(re.fullmatch(r"\s*(RANKING\s*)?(MASTER\s*)?(\d{2}\s*-\s*\d{2}|\d{2}\s*\+|\+\s*\d{2})\s*", t)) or bool(text_division(t))


def text_gender(t):
    if re.search(r"\bFEM|FEMENINO|DAMAS|MUJER", t):
        return "F"
    if re.search(r"\bMASC|MASCULINO|VARON|CABALLERO", t):
        return "M"
    return None


def text_circuit(t, division):
    if division and division.startswith("U"):
        return "Liga Menor"
    if division and division.startswith("Open"):
        return "Open Femenino"
    if division == "PTT":
        return "PTT"
    if division and division.startswith("Master"):
        return "Master"
    if division == "PTT":
        return "PTT"
    if re.search(r"FEMENIN|\bOPEN\b", t):
        return "Open Femenino"
    if re.search(r"MENOR", t):
        return "Liga Menor"
    if re.search(r"MAYOR", t):
        return "Liga Mayor"
    return None


def text_stage(name):
    toks = [t for t in re.split(r"[-_\s.]+", fold(name)) if t]
    for t in toks:
        if t in ROMAN:
            return ROMAN[t]
    for t in toks:
        if t in ORD_WORDS:
            return ORD_WORDS[t]
    m = re.search(r"(\d)\s*(?:ER|DO|TO|RO|CER)?\s*(?:RANKING|ETAPA|FECHA|RN)", fold(name))
    return int(m.group(1)) if m else None


def workbook_title(wb):
    """Category text printed inside the workbook (e.g. 'U13 MASCULINO')."""
    parts = []
    if "Inscripciones" in wb.sheetnames:
        for r in wb["Inscripciones"].iter_rows(min_row=1, max_row=4, values_only=True):
            parts += [fold(c) for c in r if isinstance(c, str)]
    for sn in wb.sheetnames:
        if fold(sn).startswith("GRUPO"):
            for r in wb[sn].iter_rows(min_row=1, max_row=12, values_only=True):
                cells = list(r)
                for i, c in enumerate(cells):
                    if fold(c).strip() == "CATEGORIA" and i + 2 < len(cells):
                        parts += [fold(x) for x in cells[i + 1:i + 3] if isinstance(x, str)]
            break
    return " ".join(parts)


def workbook_date(wb):
    for sn in wb.sheetnames:
        if not fold(sn).startswith("GRUPO"):
            continue
        for r in wb[sn].iter_rows(min_row=1, max_row=12, values_only=True):
            for c in r:
                if isinstance(c, dt.datetime):
                    return c.date().isoformat()
    return None


def load_json(p):
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}


def metadata(path, wb, warnings):
    override = load_json(path.with_name(path.name + ".json"))
    hint = load_json(path.with_name(path.name + ".hint.json"))
    name_text = fold(path.stem.replace("-", " ").replace("_", " "))
    inside = workbook_title(wb)
    date = workbook_date(wb)

    w_div, w_gen = text_division(inside), text_gender(inside)
    n_div, n_gen = text_division(name_text), text_gender(name_text)
    for key, wv, hv, nv in (("category", w_div, hint.get("division"), n_div), ("gender", w_gen, hint.get("gender"), n_gen)):
        if wv and hv and wv != hv:
            warnings.append(f"{path.name}: FECOTEME link says {key} {hv} but the workbook says {wv}; using the workbook.")
        elif not wv and hv and nv and hv != nv:
            warnings.append(f"{path.name}: FECOTEME link says {key} {hv} but the file name says {nv}; using the link. Check this file.")

    division = override.get("division") or w_div or hint.get("division") or n_div
    gender = override.get("gender") or w_gen or hint.get("gender") or n_gen
    circuit = override.get("circuit") or hint.get("circuit") or text_circuit(inside + " " + name_text, division)
    if circuit == "Liga Mayor" and not division:
        division = "Mayor"
    if circuit == "Master" and not (division or "").startswith("Master"):
        division = "Master"
    if circuit == "PTT":
        division = division if division == "PTT" else "PTT"
    if circuit == "Open Femenino":
        gender = "F"
        division = division if (division or "").startswith("Open") else "Open"
    if not gender:
        gender = "M" if circuit == "Liga Mayor" else "X"
    if circuit == "Open Femenino" and division == "Open":
        warnings.append(f"{path.name}: women's file without division A/B (Primera/Segunda); filed as 'Open'. "
                        f"Add {path.name}.json with {{\"division\": \"Open A\"}} or \"Open B\" to fix.")

    folder_year = next((int(p.name) for p in path.parents if re.fullmatch(r"20\d\d", p.name)), None)
    name_year = re.search(r"\b(20\d\d)\b", path.stem.replace("-", " ").replace("_", " "))
    season = override.get("season") or hint.get("season") or folder_year or (int(date[:4]) if date else None) \
        or (int(name_year.group(1)) if name_year else dt.date.today().year)
    stage = override.get("stage") or hint.get("stage") or text_stage(path.stem.split("-", 1)[-1] if re.match(r"^[A-Za-z0-9]{12}-", path.stem) else path.stem) or text_stage(path.stem)
    event = override.get("event") or hint.get("sourceLabel") or path.stem
    return {"circuit": circuit, "season": int(season), "division": division, "gender": gender, "stage": stage,
            "event": event, "date": override.get("date") or date or "", "sourceUrl": hint.get("sourceUrl", "")}


def stage_id(d):
    return f"{d['season']}__{slug(d['circuit'])}__{slug(d['division'])}-{d['gender']}__e{d['stage']}"


def main():
    STAGES.mkdir(parents=True, exist_ok=True)
    errors, warnings = [], []
    for path in sorted(INBOX.rglob("*.xls*")):
        if path.name.startswith("~$") or path.suffix.lower() not in (".xlsx", ".xlsm"):
            continue
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{path.relative_to(ROOT)}: could not open ({e})")
            continue
        tables = read_tables(wb)
        if not tables:
            errors.append(f"{path.relative_to(ROOT)}: no ranking table found (needs Nombre and Puntos columns)")
            continue
        meta = metadata(path, wb, warnings)
        saved = []
        for sheet, rows in tables:
            d = dict(meta)
            if sheet:  # several divisions in one workbook, e.g. Master age brackets
                d["division"] = text_division(fold(sheet)) or re.sub(r"(?i)^\s*ranking\s*", "", sheet).replace(" ", "").strip() or d["division"]
                if d["circuit"] == "Master" and not d["division"].startswith("Master"):
                    d["division"] = "Master " + d["division"]
            missing = [k for k in ("circuit", "division", "stage") if not d.get(k)]
            if missing:
                errors.append(f"{path.relative_to(ROOT)}{' / ' + sheet if sheet else ''}: could not detect {', '.join(missing)}. "
                              f"Add {path.name}.json next to it, e.g. {{\"stage\": 2, \"division\": \"U13\", \"gender\": \"F\"}}")
                continue
            d["stage"] = int(d["stage"])
            d.update(sourceFile=path.name, rows=rows)
            sid = stage_id(d)
            (STAGES / f"{sid}.json").write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
            saved.append(sid)
            print(f"OK  {path.relative_to(ROOT)}{' / ' + sheet if sheet else ''} -> {sid} ({len(rows)} athletes)")
        if saved:
            dest = ARCHIVE / str(meta["season"])
            dest.mkdir(parents=True, exist_ok=True)
            for p in (path, path.with_name(path.name + ".json"), path.with_name(path.name + ".hint.json")):
                if p.exists():
                    shutil.move(str(p), dest / p.name)

    stages = []
    for f in sorted(STAGES.glob("*.json")):
        d = json.loads(f.read_text(encoding="utf-8"))
        d["id"] = f.stem
        stages.append(d)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps({"generatedAt": dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
                               "stages": stages}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"Wrote {OUT.relative_to(ROOT)} with {len(stages)} stages")

    old = load_json(WARN_FILE).get("warnings", []) if WARN_FILE.exists() else []
    WARN_FILE.write_text(json.dumps({"warnings": (old + warnings)[-200:]}, ensure_ascii=False, indent=1), encoding="utf-8")
    for w in warnings:
        print("WARNING", w)
    if errors:
        print("\nFILES NOT PROCESSED (left in inbox/):", *errors, sep="\n  ")
        sys.exit(1)


if __name__ == "__main__":
    main()
