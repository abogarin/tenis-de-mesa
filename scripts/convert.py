#!/usr/bin/env python3
"""Convert ranking workbooks into the JSON the website reads.

Every run rebuilds all stages from archive/<year>/ (files already processed) plus
inbox/<year>/ (new files), so improvements to this script also apply to older files.
For every workbook:
  1. Find the ranking table: the "Ranking Final" sheet, or sheets named "Ranking ...",
     or any other sheet with Nombre + Puntos columns.
  2. Work out circuit, season, division (U9…U19, Mayor, Open A/B, Master, PTT),
     gender and stage number. Sources, most trusted first:
       a) "<file>.json" override written by a person
       b) the workbook itself (category title on "Inscripciones", dates on "Grupo" sheets)
       c) "<file>.hint.json" written by fetch_fecoteme.py from the FECOTEME link label
       d) the file name and the inbox/<year>/ folder
     Disagreements between (b) and (c) are reported as warnings.
  3. Save data/stages/<id>.json (a newer upload of the same stage replaces it); new
     workbooks are moved from inbox/ to archive/<season>/.
To remove a stage, delete its workbook from archive/.
Finally all stages are combined into site/data/ranking.json.
Workbooks that include the tournament itself ("Grupo N" and "LLAVE" sheets) also give the
individual matches, saved to site/data/matches.json.
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
MATCHES_OUT = ROOT / "site" / "data" / "matches.json"

ROMAN = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
ORD_WORDS = {"PRIMER": 1, "PRIMERO": 1, "PRIMERA": 1, "1ER": 1, "1RO": 1, "1RN": 1, "SEGUNDO": 2, "SEGUNDA": 2, "2DO": 2, "2RN": 2,
             "TERCER": 3, "TERCERO": 3, "TERCERA": 3, "3ER": 3, "3RO": 3, "3CER": 3, "3RN": 3, "CUARTO": 4, "CUARTA": 4, "4TO": 4, "4RN": 4,
             "QUINTO": 5, "QUINTA": 5, "5TO": 5, "5RN": 5, "SEXTO": 6, "6TO": 6}
SKIP_SHEETS = re.compile(r"^(INSCRIP|GRUPO|LLAVE|SORTEO|PLANILLA|PREMIA)", re.I)  # hidden sheets are always skipped


def fold(s):
    s = unicodedata.normalize("NFD", str(s if s is not None else ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").upper()


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", fold(s).lower()).strip("-")


# ---------- reading the ranking table ----------

def rank_by_points(rows):
    rows.sort(key=lambda r: -r[4])
    for i, r in enumerate(rows):
        r[0] = rows[i - 1][0] if i and rows[i - 1][4] == r[4] else i + 1
    return rows


def mark_movement_block(rows):
    """Some files append the players moving to this division next stage after the ranking, with
    positions restarting at 1 (Open A: 'Ascenso'). Keep them after the ranking, sharing the next
    position, flagged 'asc' (or 'desc'), instead of duplicating positions 1, 2, 3…"""
    start = next((i for i in range(1, len(rows)) if rows[i][0] is not None and rows[i - 1][0] is not None
                  and rows[i][0] == 1 and rows[i - 1][0] > 1), None)
    if start is None:
        return rows
    block = rows[start:]
    nxt = max(r[0] for r in rows[:start]) + 1
    kind = "desc" if any(r[5] == "DESC" for r in block) else "asc"
    for r in block:
        r[0], r[5] = nxt, kind
    return rows


def rerank_by_order(rows):
    """Renumber 1..n following the existing ranks (ties keep sharing a position)."""
    rows.sort(key=lambda r: r[0])
    prev, out = None, []
    for i, r in enumerate(rows):
        orig = r[0]
        r[0] = rows[i - 1][0] if i and orig == prev else i + 1
        prev = orig
    return rows


MAYOR_DIVS = ["Primera", "Segunda", "Tercera", "Cuarta", "Quinta", "Sexta"]


AGE_BRACKET = re.compile(r"^\s*(\d{2}\s*-\s*\d{2}|\d{2}\s*\+|\+\s*\d{2}|(?:U|SUB)\s*-?\s*\d{1,2})\s*$", re.I)
SECTION = re.compile(r"CATEGOR(?:I|Í)A\s*[:\-]?\s*(\S.*)$", re.I)


def norm_sub(v):
    v = re.sub(r"\s+", "", str(v).strip().upper())
    return re.sub(r"\bC-(\d)", r"C\1", v)          # PTT classes: "C-3" -> "C3"


def table_from_sheet(ws):
    """Return {sub_division_or_None: rows}. Sub-divisions come from a 'Categoría' column
    (Liga Mayor Primera…Sexta, PTT classes, Open-Primera/Segunda) or from section title rows
    ("CATEGORÍA C1-C2" above each block). Age-bracket values (Master) are ignored: the sheet
    already is the bracket, and those columns are often auto-filled wrongly (40-49, 40-50, 40-51…)."""
    rows = list(ws.iter_rows(values_only=True))

    def section_of(r):
        cells = [c for c in r if c is not None and str(c).strip()]
        if len(cells) == 1 and isinstance(cells[0], str):
            m = SECTION.search(cells[0].strip())
            if m:
                return norm_sub(m.group(1))
        return None

    for h, r in enumerate(rows[:20]):
        hdr = [fold(c).strip() for c in r]
        iN = next((i for i, c in enumerate(hdr) if "NOMBRE" in c or c in ("JUGADOR", "ATLETA")), None)
        iP = next((i for i, c in enumerate(hdr) if "PUNTO" in c or "POINT" in c or c in ("PTS", "TOTAL")), None)
        if iN is None or iP is None:
            continue
        iR = next((i for i, c in enumerate(hdr) if c.startswith("RANK") or c in ("POS", "POSICION", "#")), None)
        iC = next((i for i, c in enumerate(hdr) if c.startswith("CARN") or c in ("ID", "CEDULA", "LICENCIA")), None)
        iK = next((i for i, c in enumerate(hdr) if "CLUB" in c or "EQUIPO" in c), None)
        iG = next((i for i, c in enumerate(hdr) if c.startswith("CATEGOR") or c.startswith("CLASE")), None)
        if iG is not None:
            vals = [str(x[iG]) for x in rows[h + 1:] if len(x) > iG and x[iG] is not None and str(x[iG]).strip()]
            if not vals or all(AGE_BRACKET.match(v) for v in vals):
                iG = None
        section = next((sec for x in reversed(rows[:h]) if (sec := section_of(x))), None)
        groups = {}
        for r2 in rows[h + 1:]:
            sec = section_of(r2)
            if sec:
                section = sec
                continue
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
            sub = section
            if iG is not None and len(r2) > iG and r2[iG] is not None and str(r2[iG]).strip():
                sub = norm_sub(r2[iG])
            note = " ".join(fold(c) for c in r2 if isinstance(c, str))
            groups.setdefault(sub, []).append([rank, carne, name, club, round(float(pts), 2) if pts % 1 else int(pts),
                                               "DESC" if "DESCIEND" in note else "ASC" if "ASCEN" in note else ""])
        for out in groups.values():
            mark_movement_block(out)
            for x in out:
                if x[5] not in ("asc", "desc"):
                    x.pop()  # plain rows stay [rank, carne, name, club, points]
        if groups:
            for out in groups.values():
                if any(x[0] is None for x in out):
                    rank_by_points(out)   # no ranking column (e.g. PTT with a class column)
            return groups
    return None


def read_tables(wb):
    """Return [(sub_division_or_None, rows)]. More than one entry = several divisions in one file."""
    def flat(sheet_label, groups):
        if len(groups) == 1 and None in groups:
            return [(sheet_label, groups[None])]
        return [((f"{sheet_label} {k}" if sheet_label else k) if k else sheet_label, v) for k, v in groups.items()]

    if "Ranking Final" in wb.sheetnames:
        g = table_from_sheet(wb["Ranking Final"])
        if g:
            return flat(None, g)
    visible = [n for n in wb.sheetnames if wb[n].sheet_state == "visible"]
    named = [n for n in visible if re.match(r"\s*RANK", fold(n))]
    tables = [(n, g) for n in named if (g := table_from_sheet(wb[n]))]
    if tables:
        if len(tables) == 1:
            return flat(None, tables[0][1])
        return [x for n, g in tables for x in flat(n, g)]
    cands = [(n, g) for n in visible if not SKIP_SHEETS.match(fold(n)) and (g := table_from_sheet(wb[n]))]
    if len(cands) > 1 and all(bracket_like(n) for n, _ in cands):
        return [x for n, g in cands for x in flat(n, g)]  # one sheet per division, e.g. Master 30-39 / 60+
    return flat(None, cands[0][1]) if cands else []


# ---------- matches (group and knockout sheets) ----------

ROUND_NAMES = {"1/32": "Ronda de 64", "1/16": "Ronda de 32", "1/8": "Octavos", "OCTAVOS": "Octavos", "QF": "Cuartos", "1/4": "Cuartos",
               "CUARTOS": "Cuartos", "SF": "Semifinal", "1/2": "Semifinal", "SEMIFINAL": "Semifinal", "SEMIFINALES": "Semifinal",
               "F": "Final", "FINAL": "Final", "3Y4": "3.er lugar", "3ERLUGAR": "3.er lugar", "TERCERLUGAR": "3.er lugar"}


def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return int(v)
    if isinstance(v, str) and v.strip().isdigit():
        return int(v.strip())
    return None


def _carne(v):
    n = _num(v)
    return str(n) if n is not None else None


def _sets(vals):
    """[a1, b1, a2, b2, ...] -> [[a1, b1], ...] for the sets actually played."""
    out = []
    for a, b in zip(vals[0::2], vals[1::2]):
        a, b = _num(a), _num(b)
        if a is None or b is None or (a == 0 and b == 0):
            continue
        out.append([a, b])
    return out


def _grid(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _find(grid, pred):
    for i, row in enumerate(grid):
        for j, v in enumerate(row):
            if isinstance(v, str) and pred(fold(v).strip()):
                return i, j
    return None


def _right(row, j):
    for v in row[j + 1:]:
        if v not in (None, ""):
            return v
    return None


ENTRY_POINTS = {}  # carné -> points at the start of the stage, filled while reading group sheets


def group_matches(ws):
    g = _grid(ws)
    hdr = next((i for i, r in enumerate(g) if {"ORDEN", "CARNE", "JUGADORES"} <= {fold(v).strip() for v in r if isinstance(v, str)}), None)
    part = _find(g, lambda t: t.startswith("PARTIDA"))
    if hdr is None or not part:
        return []
    cols = {fold(v).strip(): j for j, v in enumerate(g[hdr]) if isinstance(v, str)}
    players = {}
    for r in g[hdr + 1:]:
        o = _num(r[cols["ORDEN"]]) if cols["ORDEN"] < len(r) else None
        if o is None:
            break
        players[o] = [_carne(r[cols["CARNE"]]), str(r[cols["JUGADORES"]] or "").strip(), str(r[cols.get("LUGAR", cols["JUGADORES"] + 1)] or "").strip()]
        # points when the stage started: the "Ranking"/"Puntos" columns hold position and points (their labels are swapped in some files)
        nums = [_num(r[cols[k]]) for k in ("RANKING", "PUNTOS") if k in cols and cols[k] < len(r)]
        nums = [n for n in nums if n is not None]
        if nums and max(nums) >= 100:
            ENTRY_POINTS[players[o][0]] = max(nums)
    gp = _find(g, lambda t: t.startswith("GRUPO N"))
    group = _num(_right(g[gp[0]], gp[1])) if gp else _num(re.sub(r"\D", "", ws.title) or None)
    fp = _find(g, lambda t: t.startswith("FECHA"))
    date = _right(g[fp[0]], fp[1]) if fp else None
    date = date.date().isoformat() if isinstance(date, dt.datetime) else None
    pi, pj = part
    head = [fold(v).strip() if isinstance(v, str) else "" for v in g[pi]]
    set_cols = [j for j, t in enumerate(head) if "SET" in t]
    win_col = next((j for j, t in enumerate(head) if t.startswith("GANADOR")), None)
    out, i = [], pi + 1
    while i + 1 < len(g):
        r1, r2 = g[i], g[i + 1]
        if _num(r1[pj]) is None:
            i += 1
            if i > pi + 40:
                break
            continue
        oa, ob = _num(r1[pj + 1]), _num(r2[pj + 1])
        if oa in players and ob in players:
            vals = []
            for j in set_cols:
                vals += [r1[j], r2[j]]
            sets = _sets(vals)
            if sets:
                won_a = sum(a > b for a, b in sets)
                w = _num(r1[win_col]) if win_col is not None else None
                win = 0 if w == oa else 1 if w == ob else (0 if won_a * 2 > len(sets) else 1)
                out.append({"r": "Grupo", "g": group, "d": date, "a": players[oa], "b": players[ob], "s": sets, "w": win})
        i += 2
    return out


def knockout_matches(ws):
    g = _grid(ws)
    hdr = next((i for i, r in enumerate(g) if {"MATCH", "ROUND", "PLA 1", "PLA 2"} <= {fold(v).strip() for v in r if isinstance(v, str)}), None)
    if hdr is None:
        return []
    head = [fold(v).strip() if isinstance(v, str) else "" for v in g[hdr]]
    c = {t: j for j, t in reversed(list(enumerate(head))) if t}
    ass = [j for j, t in enumerate(head) if t == "ASS"]
    ga, gx = c.get("G A"), c.get("G X")
    if not ass or ga is None or "A" not in c or "B" not in c:
        return []
    set_cols = list(range(ass[-1] + 1, ga))
    out = []
    for r in g[hdr + 1:]:
        if _num(r[c["MATCH"]]) is None:
            continue
        ca, cb = _carne(r[c["A"]]), _carne(r[c["B"]])
        na, nb = str(r[c["PLA 1"]] or "").strip(), str(r[c["PLA 2"]] or "").strip()
        rnd = re.sub(r"\s+", "", fold(str(r[c["ROUND"]] or "")))
        rname = ROUND_NAMES.get(rnd, str(r[c["ROUND"]] or "").strip())
        bye_a, bye_b = (not ca or fold(na) == "BYE"), (not cb or fold(nb) == "BYE")
        if bye_a != bye_b:  # one real player against a BYE: record that the player passed the round
            p = [cb, nb, str(r[ass[-1]] or "").strip()] if bye_a else [ca, na, str(r[ass[0]] or "").strip()]
            if p[0]:
                out.append({"r": rname, "g": None, "d": None, "a": p, "b": None, "s": [], "w": 0, "bye": True})
            continue
        if bye_a or bye_b:
            continue
        sets = _sets([r[j] for j in set_cols])
        if not sets:
            continue
        a_won, b_won = _num(r[ga]), _num(r[gx])
        win = (0 if a_won > b_won else 1) if a_won is not None and b_won is not None and a_won != b_won \
            else (0 if sum(x > y for x, y in sets) * 2 > len(sets) else 1)
        out.append({"r": rname, "g": None, "d": None,
                    "a": [ca, na, str(r[ass[0]] or "").strip()], "b": [cb, nb, str(r[ass[-1]] or "").strip()], "s": sets, "w": win})
    return out


def read_matches(wb):
    """Individual matches from 'Grupo N' and 'LLAVE ..' sheets, in playing order (groups first)."""
    groups, ko = [], []
    ENTRY_POINTS.clear()
    for ws in wb.worksheets:
        t = fold(ws.title).strip()
        try:
            if re.match(r"GRUPO\s*\d+", t):
                groups += group_matches(ws)
            elif t.startswith("LLAVE"):
                ko += knockout_matches(ws)
        except Exception as e:  # noqa: BLE001  a layout we don't know: skip the sheet, keep the ranking
            print(f"WARNING {ws.title}: could not read matches ({e})")
    date = next((m["d"] for m in groups if m["d"]), None)
    for m in ko:
        m["d"] = date
    ms = groups + ko
    return {"pts": dict(ENTRY_POINTS), "m": ms} if ms else None


# ---------- metadata ----------

def text_division(t):
    m = re.search(r"(?<![A-Z0-9])(?:U|SUB)\s?-?\s?(\d{1,2})(?!\d)", t)
    if m:
        return f"U{m.group(1)}"
    if re.search(r"\bPTT\b|PARA\s*TENIS|PARALIMP", t):        # before Master: one PTT label says "Master PTT"
        return "PTT"
    if "MASTER" in t:
        m = re.search(r"\b([3-9]\d)\s*(?:-|\s|A)\s*([3-9]\d)\b", t)
        if m:
            return f"Master {m.group(1)}-{m.group(2)}"
        m = re.search(r"\b([3-9]\d)\s*\+|\+\s*([3-9]\d)\b|MAS\s*DE\s*([3-9]\d)|MASTER\w*\s*-?\s*\b(6\d|7\d)\b(?!\s*-)", t)
        if m:
            return f"Master {next(g for g in m.groups() if g)}+"
        return "Master"
    if re.search(r"OPEN|FEMENIN|\bFEM\b", t):
        # Open Femenino divisions: A = Primera, B = Segunda (both names are used on FECOTEME pages)
        m = re.search(r"OPEN\s*(?:FEM\w*\s*)?-?\s*\b([AB])\b|FEMENIN\w*\s*\b([AB])\b|\b([AB])\s*$", t)
        if m:
            return "Open " + next(g for g in m.groups() if g)
        m = re.search(r"\b(I{1,2})\s*CATEGOR", t)   # "I Categoría" = A, "II Categoría" = B
        if m:
            return "Open A" if m.group(1) == "I" else "Open B"
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
    if re.search(r"(?<![A-Z])FEM|FEMENINO|DAMAS|MUJER", t):
        return "F"
    if re.search(r"(?<![A-Z])MASC|MASCULINO|VARON|CABALLERO", t):
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
    m = re.search(r"(?:^|[^IVX])([IVX]{1,4})RN", fold(name))   # "U13-Masculino-IIIRN-menor", "U13MascVRNmenor"
    if m and m.group(1) in ROMAN:
        return ROMAN[m.group(1)]
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
        gender = "X"  # Liga Mayor, Master and PTT are mixed
    if circuit == "Open Femenino" and division == "Open":
        warnings.append(f"{path.name}: women's file without division A/B (Primera/Segunda); filed as 'Open'. "
                        f"Add {path.name}.json with {{\"division\": \"Open A\"}} or \"Open B\" to fix.")

    folder_year = next((int(p.name) for p in path.parents if re.fullmatch(r"20\d\d", p.name)), None)
    name_year = re.search(r"\b(20\d\d)\b", path.stem.replace("-", " ").replace("_", " "))
    season = override.get("season") or hint.get("season") or folder_year or (int(date[:4]) if date else None) \
        or (int(name_year.group(1)) if name_year else dt.date.today().year)
    stage = override.get("stage") or hint.get("stage") or text_stage(path.stem.split("-", 1)[-1] if re.match(r"^[A-Za-z0-9]{12}-", path.stem) else path.stem) or text_stage(path.stem)
    label = hint.get("sourceLabel") or ""
    event = override.get("event") or (label if re.search(r"ranking", label, re.I) else "") or re.sub(r"^[A-Za-z0-9]{12}-", "", path.stem).replace("-", " ")
    return {"circuit": circuit, "season": int(season), "division": division, "gender": gender, "stage": stage,
            "event": event, "date": override.get("date") or date or "", "sourceUrl": hint.get("sourceUrl", "")}


def stage_id(d):
    return f"{d['season']}__{slug(d['circuit'])}__{slug(d['division'])}-{d['gender']}__e{d['stage']}"


def mayor_thresholds():
    """Minimum points for each Liga Mayor division, from sources.json ("ligaMayorDivisions")."""
    default = {"Primera": 1900, "Segunda": 1700, "Tercera": 1500, "Cuarta": 1300, "Quinta": 1100, "Sexta": 0}
    try:
        cfg = json.loads((ROOT / "sources.json").read_text(encoding="utf-8")).get("ligaMayorDivisions")
        return {k: float(cfg[k]) for k in MAYOR_DIVS} if cfg else default
    except Exception:  # noqa: BLE001
        return default


def combine_open_divisions(meta, tables):
    """Open Femenino files are one per division (A or B). When such a file also has a category
    column (Open-Primera / Open-Segunda), that column says where each player goes next stage:
    keep one ranking and flag the players moving up ('asc') or down ('desc')."""
    if meta.get("circuit") != "Open Femenino" or meta.get("division") not in ("Open A", "Open B") or len(tables) < 2:
        return tables
    divs = [(label, text_division(fold(str(label).replace("-", " "))) if label else None, rows) for label, rows in tables]
    if not all(d in ("Open A", "Open B") for _, d, _ in divs):
        return tables
    merged = []
    for _, d, rows in divs:
        for r in rows:
            r = list(r[:5])
            if d != meta["division"]:
                r.append("asc" if d == "Open A" else "desc")
            merged.append(r)
    merged.sort(key=lambda r: r[0])
    return [(None, merged)]


def infer_mayor_divisions(built, warnings):
    """Liga Mayor has six rankings (Primera…Sexta). When a stage file is one overall list with no
    division column, each player's division is set by their points (thresholds in sources.json).
    Players may play one division up: if, in the nearest stage of the same season whose file
    does list divisions, a player was exactly one division above their points division, they are
    kept there. Marked so the site can say so."""
    limits = sorted(mayor_thresholds().items(), key=lambda kv: -kv[1])
    def by_points(p):
        return next((name for name, low in limits if p >= low), limits[-1][0])

    known = {}  # season -> stage -> {carne: division}
    for d, _ in built.values():
        if d["circuit"] == "Liga Mayor" and d["division"] in MAYOR_DIVS and not d.get("divisionBy"):
            known.setdefault(d["season"], {}).setdefault(d["stage"], {}).update({r[1]: (d["division"], r[4]) for r in d["rows"]})

    for sid in [x for x, (d, _) in built.items() if d["circuit"] == "Liga Mayor" and d["division"] not in MAYOR_DIVS]:
        d, f = built.pop(sid)
        ks = known.get(d["season"], {})
        ref = min(ks, key=lambda st: (abs(st - d["stage"]), st > d["stage"])) if ks else None
        parts, up = {}, 0
        for r in sorted(d["rows"], key=lambda r: r[0]):
            div = by_points(r[4])
            prev = ks[ref].get(r[1]) if ref else None
            if prev:
                prev_div, prev_pts = prev
                played_up = MAYOR_DIVS.index(prev_div) == MAYOR_DIVS.index(by_points(prev_pts)) - 1
                if played_up and MAYOR_DIVS.index(prev_div) == MAYOR_DIVS.index(div) - 1:
                    div, up = prev_div, up + 1  # was playing one division up there: keep them up
            parts.setdefault(div, []).append(list(r))
        for div, rows in parts.items():
            nd = dict(d, division=div, rows=rerank_by_order(rows), divisionBy="points", playUpFrom=ref)
            built[stage_id(nd)] = (nd, f)
        rule = ", ".join(f"{n} {int(v)}+" for n, v in limits)
        warnings.append(f"Liga Mayor {d['season']} stage {d['stage']}: the file has no division column; divisions by points ({rule})"
                        + (f"; {up} player(s) kept one division up as in stage {ref}." if ref else "."))


def main():
    """Rebuild every stage from all workbooks: archive/ (already seen) + inbox/ (new).
    Rebuilding each time means improvements to this script also apply to older files."""
    STAGES.mkdir(parents=True, exist_ok=True)
    errors, warnings = [], []
    archived = sorted(p for p in ARCHIVE.rglob("*.xls*") if p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$"))
    incoming = sorted(p for p in INBOX.rglob("*.xls*") if p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$"))
    built = {}  # stage id -> (doc, file)
    found_matches = {}  # stage id -> {file: [matches]}
    undated = []  # files with no stage in the name: accepted only if identical to another file
    for path in archived + incoming:
        is_new = path in incoming
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
        except Exception as e:  # noqa: BLE001
            errors.append(f"{path.relative_to(ROOT)}: could not open ({e})")
            continue
        tables = read_tables(wb)
        if not tables:
            errors.append(f"{path.relative_to(ROOT)}: no ranking table found (needs Nombre and Puntos columns)")
            continue
        meta = metadata(path, wb, warnings)
        file_matches = read_matches(wb) if len(tables) == 1 else None
        tables = combine_open_divisions(meta, tables)
        saved = []
        for sheet, rows in tables:
            if sheet and len(tables) > 1:
                rerank_by_order(rows)  # one ranking per division, numbered within the division
            d = dict(meta)
            if sheet:  # several divisions in one workbook, e.g. Master age brackets or PTT classes
                d["division"] = text_division(fold(sheet)) or re.sub(r"(?i)^\s*ranking\s*", "", sheet).replace(" ", "").strip() or d["division"]
                if d["circuit"] == "Liga Mayor":
                    d["division"] = re.sub(r"\s+", " ", sheet).strip().title()
                for circ in ("Master", "PTT"):
                    if d["circuit"] == circ and not d["division"].startswith(circ):
                        d["division"] = f"{circ} {d['division']}"
            missing = [k for k in ("circuit", "division", "stage") if not d.get(k)]
            if missing == ["stage"]:
                undated.append((path, sheet, d, rows, is_new))
                continue
            if missing:
                errors.append(f"{path.relative_to(ROOT)}{' / ' + sheet if sheet else ''}: could not detect {', '.join(missing)}. "
                              f"Add {path.name}.json next to it, e.g. {{\"stage\": 2, \"division\": \"U13\", \"gender\": \"F\"}}")
                continue
            d["stage"] = int(d["stage"])
            rows.sort(key=lambda r: r[0])
            # Some files number players by their finish in that stage's tournament instead of by
            # ranking points (e.g. 2025 U13 Masc stage 5). If many positions contradict the points,
            # re-rank by points.
            inv = sum(1 for x, y in zip(rows, rows[1:]) if y[4] > x[4])
            if len(rows) >= 5 and inv / (len(rows) - 1) >= 0.2:
                rank_by_points(rows)
                warnings.append(f"{path.name}{' / ' + sheet if sheet else ''}: positions did not follow the points "
                                f"({inv} of {len(rows) - 1}); re-ranked by points.")
            d.update(sourceFile=path.name, rows=rows)
            sid = stage_id(d)
            if file_matches:
                found_matches.setdefault(sid, {})[path.name] = file_matches
            if sid in built and built[sid][1] != path.name:
                prev = built[sid][1]
                keep_prev = bool(re.search(r"REV|CORREC", fold(prev))) and not re.search(r"REV|CORREC", fold(path.name)) and not is_new
                warnings.append(f"{sid}: found in two files ({prev} and {path.name}); using {prev if keep_prev else path.name}.")
                if keep_prev:
                    continue
            built[sid] = (d, path.name)
            saved.append(sid)
            if is_new:
                print(f"NEW {path.relative_to(ROOT)}{' / ' + sheet if sheet else ''} -> {sid} ({len(rows)} athletes)")
        if saved and is_new:
            dest = ARCHIVE / str(meta["season"])
            dest.mkdir(parents=True, exist_ok=True)
            for p in (path, path.with_name(path.name + ".json"), path.with_name(path.name + ".hint.json")):
                if p.exists():
                    shutil.move(str(p), dest / p.name)

    # A file with no stage in its name is usually a second download of a stage we already have
    # (e.g. "U13-Masculino.xlsx" = "U13-Masculino-IIRN-menor.xlsx"). Skip exact copies.
    for path, sheet, d, rows, is_new in undated:
        same = lambda r: {x[1]: x[4] for x in r}
        twin = next((f for sid, (b, f) in built.items() if b["season"] == d["season"] and b["circuit"] == d["circuit"]
                     and b["division"] == d["division"] and b["gender"] == d["gender"] and same(b["rows"]) == same(rows)), None)
        if twin:
            warnings.append(f"{path.name}: no stage in the name, but identical to {twin}; skipped as a duplicate.")
            if is_new:
                dest = ARCHIVE / str(d["season"])
                dest.mkdir(parents=True, exist_ok=True)
                if path.exists():
                    shutil.move(str(path), dest / path.name)
        else:
            errors.append(f"{path.relative_to(ROOT)}: could not detect the stage. Rename it with the stage (e.g. "
                          f"II-{path.name}) or add {path.name}.json next to it with {{\"stage\": 2}}")

    infer_mayor_divisions(built, warnings)

    for f in STAGES.glob("*.json"):
        f.unlink()
    for sid, (d, _) in built.items():
        (STAGES / f"{sid}.json").write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
    print(f"Rebuilt {len(built)} stages from {len(archived)} archived + {len(incoming)} new file(s)")

    stages = []
    for f in sorted(STAGES.glob("*.json")):
        d = json.loads(f.read_text(encoding="utf-8"))
        d["id"] = f.stem
        stages.append(d)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps({"generatedAt": dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
                               "ligaMayorDivisions": mayor_thresholds(),
                               "stages": stages}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"Wrote {OUT.relative_to(ROOT)} with {len(stages)} stages")

    # Matches: from the file used for the stage, or from another copy of the same stage that has them
    matches = {}
    for sid, by_file in found_matches.items():
        if sid in built:
            matches[sid] = by_file.get(built[sid][1]) or next(iter(by_file.values()))
    MATCHES_OUT.write_text(json.dumps(matches, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"Wrote {MATCHES_OUT.relative_to(ROOT)} with {sum(len(v['m']) for v in matches.values())} matches from {len(matches)} stages")

    WARN_FILE.write_text(json.dumps({"warnings": warnings}, ensure_ascii=False, indent=1), encoding="utf-8")
    for w in warnings:
        print("WARNING", w)
    if errors:
        print("\nFILES NOT PROCESSED (left in inbox/):", *errors, sep="\n  ")
        sys.exit(1)


if __name__ == "__main__":
    main()
