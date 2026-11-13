# -*- coding: utf-8 -*-
"""
tsct_insert.py - write material quantities into the IStructE Structural Carbon Tool (TSCT v3.1)
===============================================================================================

Takes a dictionary / JSON / CSV (e.g. exported from Grasshopper) and fills in
'Project Info' and the 'Scheme 1..6' input rows (B12:M43) of TSCT-v3-01.xlsm.

The workbook is edited at XML level, so macros (vbaProject.bin), dropdowns,
charts, comments and conditional formatting are all preserved (openpyxl would
silently strip several of these). Excel recalculates everything when the
output file is opened.

Accepted input (dict, JSON string/file, or CSV file)
----------------------------------------------------
Full form:
  {
    "project": {"name": "Fellenoord", "number": "AEC-01", "typology": "Residential",
                "design_stage": "Concept", "design_life": 50,
                "reinforcement": "Global avg", "empty_site": "Yes", "landfill_km": 50},
    "schemes": {
      "1": {"name": "Baseline - RC frame", "gia_new": 4500,
            "rows": [
              {"material": "concrete C32/40", "element": "slab", "quantity": 420, "rebar": "General Slab"},
              {"material": "steel sections", "element": "beam", "quantity": 38000, "unit": "kg"}
            ]},
      "2": {"name": "Timber hybrid", "gia_new": 4500,
            "rows": {"CLT": {"volume": 310, "element": "slab"}, "glulam": {"volume": 45, "element": "column"}}}
    }
  }
Short forms:
  {"CLT": 310, "concrete C32/40": 150}                     -> Scheme 1, volumes in m3
  {"Scheme 1": {...}, "Scheme 2": {...}}                     -> per scheme
  {"old": {...}, "new": {...}}                               -> Scheme 1 = old, Scheme 2 = new
CSV columns: scheme, material, element, quantity, unit, rebar, lifespan, description, aspect

Row fields: material (full TSCT name, alias or keywords), quantity | volume | mass,
unit ("m3" default, or "kg"), element (TSCT name or alias: slab, floor, beam, column,
wall, core, bracing, foundation, pile, ground slab, basement wall, stairs, roof, other),
rebar (rate name like "General Slab", or a number in kg/m3), lifespan (years, optional),
description, aspect ("New Build" | "Extension" | "Retained").

Usage
-----
  python tsct_insert.py input.json TSCT-v3-01.xlsm TSCT-filled.xlsm
  python tsct_insert.py input.csv  TSCT-v3-01.xlsm TSCT-filled.xlsm --keep   (don't clear old rows)

From Python / Grasshopper:
  import tsct_insert
  log = tsct_insert.insert(data, r"C:\\...\\TSCT-v3-01.xlsm", r"C:\\...\\TSCT-filled.xlsm")

Needs gh_lca_delta.py and tsct_materials.json in the same folder (material aliases).
Close the output workbook in Excel before writing to it.
"""
from __future__ import division, print_function

import csv
import io
import json
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
if HERE not in sys.path:
    sys.path.append(HERE)
import gh_lca_delta as lca  # noqa: E402  (material alias resolution)

FIRST_ROW, LAST_ROW = 12, 43
INPUT_COLS = "BCDEFHIJKLM"          # G (lifespan) is a formula, only overwritten if given

ELEMENT_ALIASES = {
    "pile": "1.1. Piling", "piles": "1.1. Piling", "piling": "1.1. Piling",
    "foundation": "1.1 Foundations (incl. pile caps)", "foundations": "1.1 Foundations (incl. pile caps)",
    "pile cap": "1.1 Foundations (incl. pile caps)", "raft": "1.1 Foundations (incl. pile caps)",
    "ground slab": "1.2.1 Lowest slab", "lowest slab": "1.2.1 Lowest slab",
    "basement slab": "1.2.2 Suspended slabs", "basement wall": "1.2.3 Basement retaining walls",
    "retaining wall": "1.2.3 Basement retaining walls",
    "column": "2.1.1 Frame (vertical) - columns/structural walls & braces",
    "columns": "2.1.1 Frame (vertical) - columns/structural walls & braces",
    "wall": "2.1.1 Frame (vertical) - columns/structural walls & braces",
    "core": "2.1.1 Frame (vertical) - columns/structural walls & braces",
    "bracing": "2.1.1 Frame (vertical) - columns/structural walls & braces",
    "stability": "2.1.1 Frame (vertical) - columns/structural walls & braces",
    "vertical": "2.1.1 Frame (vertical) - columns/structural walls & braces",
    "beam": "2.1.2 Frame (Horizontal) - beams, joists & braces",
    "beams": "2.1.2 Frame (Horizontal) - beams, joists & braces",
    "joist": "2.1.2 Frame (Horizontal) - beams, joists & braces",
    "horizontal": "2.1.2 Frame (Horizontal) - beams, joists & braces",
    "slab": "2.2.1 Upper floor and roof - structural slabs",
    "floor": "2.2.1 Upper floor and roof - structural slabs",
    "roof": "2.2.1 Upper floor and roof - structural slabs",
    "topping": "2.2.2 Upper floor and roof - non-structural slabs",
    "screed": "2.2.2 Upper floor and roof - non-structural slabs",
    "stairs": "2.4.1 Stairs", "stair": "2.4.1 Stairs", "ramp": "2.4.2 Ramps",
    "modular": "6 Pre-fabricated buildings and building units",
    "other": "Other",
    # TSCT has no facade/envelope category -> 'Other' (report it separately from the structure)
    "facade": "Other", "façade": "Other", "envelope": "Other", "cladding": "Other",
    "curtain wall": "Other", "glazing": "Other", "window": "Other", "green facade": "Other",
}
# plausible thickness range (volume / area, in m) per material keyword, for sanity warnings
THICKNESS_CHECK = {"glass": (0.004, 0.06), "aluminium": (0.0005, 0.02), "clt": (0.06, 0.4),
                   "concrete": (0.05, 1.5), "brick": (0.05, 0.5), "timber": (0.01, 0.5)}
# default rebar rate per element group when concrete has none given
REBAR_BY_ELEMENT = {"1.1. Piling": "Pile", "1.1 Foundations": "Raft", "1.2.1": "General Slab",
                    "1.2.2": "General Slab", "1.2.3": "Retaining Wall", "2.1.1": "General Wall",
                    "2.1.2": "General Beam", "2.2.1": "General Slab", "2.2.2": "Mass Concrete",
                    "2.4": "General Slab"}
PROJECT_CELLS = {"number": "C5", "name": "C6", "typology": "C7", "design_stage": "C8",
                 "project_value": "C9", "secondary_functional_unit": "C10", "empty_site": "C11",
                 "landfill_km": "C14", "reinforcement": "C15", "excavation_removed": "C16",
                 "soil_density": "C17", "design_life": "C18"}


# ============================================================================ xlsx helpers
def _esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def _unesc(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
            .replace("&apos;", "'").replace("&amp;", "&"))


class Workbook(object):
    """Minimal zip/XML editor for an existing .xlsx/.xlsm."""

    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.infos = z.infolist()
            self.files = dict((i.filename, z.read(i.filename)) for i in self.infos)
        wb = self.files["xl/workbook.xml"].decode("utf-8")
        rels = self.files["xl/_rels/workbook.xml.rels"].decode("utf-8")
        rid_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
        rid_target.update(dict((a, b) for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)))
        self.sheets = {}
        for name, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb):
            t = rid_target[rid]
            self.sheets[_unesc(name)] = "xl/" + t.lstrip("/").replace("xl/", "") if not t.startswith("/xl/") else t[1:]
        ss = self.files.get("xl/sharedStrings.xml", b"").decode("utf-8")
        self.shared = [re.sub(r"<[^>]+>", "", _unesc(si)) for si in re.findall(r"<si>(.*?)</si>", ss, re.S)]
        self._xml = {}

    def xml(self, sheet):
        if sheet not in self._xml:
            self._xml[sheet] = self.files[self.sheets[sheet]].decode("utf-8")
        return self._xml[sheet]

    def _cell_re(self, ref):
        return re.compile(r'<c r="%s"(?:\s[^>]*?)?(?:/>|>.*?</c>)' % ref, re.S)

    def read(self, sheet, ref):
        m = self._cell_re(ref).search(self.xml(sheet))
        if not m:
            return None
        c = m.group(0)
        v = re.search(r"<v>(.*?)</v>", c, re.S)
        t = re.search(r'\st="(\w+)"', c.split(">")[0])
        t = t.group(1) if t else None
        if t == "inlineStr":
            return _unesc(re.sub(r"<[^>]+>", "", re.search(r"<is>(.*?)</is>", c, re.S).group(1)))
        if not v:
            return None
        if t == "s":
            return self.shared[int(v.group(1))]
        if t in ("str", "e"):
            return _unesc(v.group(1))
        try:
            return float(v.group(1))
        except ValueError:
            return v.group(1)

    def read_col(self, sheet, col, r1, r2):
        return [self.read(sheet, "%s%d" % (col, r)) for r in range(r1, r2 + 1)]

    def write(self, sheet, ref, value):
        """Write a value (str/number/None) keeping the cell style; removes any formula."""
        x = self.xml(sheet)
        m = self._cell_re(ref).search(x)
        style = ""
        if m:
            s = re.search(r'\ss="(\d+)"', m.group(0).split(">")[0])
            style = ' s="%s"' % s.group(1) if s else ""
        if value is None or value == "":
            new = '<c r="%s"%s/>' % (ref, style)
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            new = '<c r="%s"%s><v>%s</v></c>' % (ref, style, repr(float(value)) if value != int(value) else int(value))
        else:
            new = '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, style, _esc(value))
        if m:
            x = x[:m.start()] + new + x[m.end():]
        else:
            x = self._insert_cell(x, ref, new)
        self._xml[sheet] = x

    def write_formula(self, sheet, ref, formula):
        x = self.xml(sheet)
        m = self._cell_re(ref).search(x)
        s = re.search(r'\ss="(\d+)"', m.group(0).split(">")[0]) if m else None
        new = '<c r="%s"%s><f>%s</f></c>' % (ref, ' s="%s"' % s.group(1) if s else "", formula)
        self._xml[sheet] = x[:m.start()] + new + x[m.end():] if m else self._insert_cell(x, ref, new)

    def formula(self, sheet, ref):
        m = self._cell_re(ref).search(self.xml(sheet))
        f = re.search(r"<f>(.*?)</f>", m.group(0), re.S) if m else None
        return f.group(1) if f else None

    @staticmethod
    def _colnum(ref):
        n = 0
        for ch in re.match(r"[A-Z]+", ref).group(0):
            n = n * 26 + ord(ch) - 64
        return n

    def _insert_cell(self, x, ref, new):
        row = int(re.search(r"\d+", ref).group(0))
        m = re.search(r'<row r="%d"[^>]*?(?:/>|>(.*?)</row>)' % row, x, re.S)
        if not m:
            raise ValueError("Row %d not found - cannot insert %s" % (row, ref))
        if m.group(0).endswith("/>"):
            rowxml = m.group(0)[:-2] + ">" + new + "</row>"
            return x[:m.start()] + rowxml + x[m.end():]
        body = m.group(1)
        pos = len(body)
        for cm in re.finditer(r'<c r="([A-Z]+)%d"' % row, body):
            if self._colnum(cm.group(1)) > self._colnum(ref):
                pos = cm.start()
                break
        start = m.start(1)
        return x[:start + pos] + new + x[start + pos:]

    def save(self, out_path):
        for sheet, x in self._xml.items():
            self.files[self.sheets[sheet]] = x.encode("utf-8")
        # force Excel to recalculate everything; drop the (now stale) calc chain
        wb = self.files["xl/workbook.xml"].decode("utf-8")
        if "fullCalcOnLoad" not in wb:
            wb = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1)
        self.files["xl/workbook.xml"] = wb.encode("utf-8")
        if "xl/calcChain.xml" in self.files:
            del self.files["xl/calcChain.xml"]
            r = self.files["xl/_rels/workbook.xml.rels"].decode("utf-8")
            r = re.sub(r'<Relationship [^>]*Target="calcChain.xml"[^>]*/>', "", r)
            self.files["xl/_rels/workbook.xml.rels"] = r.encode("utf-8")
            ct = self.files["[Content_Types].xml"].decode("utf-8")
            ct = re.sub(r'<Override [^>]*PartName="/xl/calcChain.xml"[^>]*/>', "", ct)
            self.files["[Content_Types].xml"] = ct.encode("utf-8")
        tmp = out_path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for info in self.infos:
                if info.filename in self.files:
                    z.writestr(info, self.files[info.filename], compress_type=info.compress_type)
        if os.path.exists(out_path):
            os.remove(out_path)
        os.rename(tmp, out_path)


# ============================================================================ input parsing
def _load(data):
    if isinstance(data, (dict, list)):
        return data
    s = str(data)
    if os.path.isfile(s):
        if s.lower().endswith(".csv"):
            with io.open(s, encoding="utf-8-sig") as f:
                return _from_csv(f.read())
        with io.open(s, encoding="utf-8-sig") as f:
            s = f.read()
    s = s.strip()
    if s.startswith("{") or s.startswith("["):
        return json.loads(s)
    return _from_csv(s)


# header names of the simple element list (.txt/.csv) -> internal field names
COLUMN_ALIASES = {
    "element nr": "element_nr", "element no": "element_nr", "element number": "element_nr", "nr": "element_nr",
    "element type": "element", "element": "element",
    "material type": "material", "material": "material",
    "unit": "unit", "area": "area", "area_m2": "area",
    "volume": "volume", "volume_m3": "volume", "mass": "mass", "mass_kg": "mass",
    "old/new constructed": "status", "old/new": "status", "status": "status",
    "scheme": "scheme", "rebar": "rebar", "lifespan": "lifespan", "description": "description",
    "quantity": "quantity", "aspect": "aspect",
}
STATUS_TO_ASPECT = {"old": "Retained", "existing": "Retained", "retained": "Retained",
                    "new": "New Build", "new build": "New Build", "extension": "Extension"}


def _from_csv(text):
    text = "\n".join(ln for ln in text.splitlines() if ln.strip() and not ln.strip().startswith("#"))
    first = text.split("\n")[0]
    delim = ";" if first.count(";") > first.count(",") else ("\t" if "\t" in first else ",")
    rows = list(csv.DictReader(io.StringIO(text) if sys.version_info[0] >= 3 else io.BytesIO(text),
                               delimiter=delim, skipinitialspace=True))
    schemes = {}
    for r in rows:
        r = dict((COLUMN_ALIASES.get(k.strip().lower(), k.strip().lower()), (v or "").strip())
                 for k, v in r.items() if k)
        if r.get("status"):
            st = r.pop("status").lower()
            if st not in STATUS_TO_ASPECT:
                raise ValueError("Unknown old/new value '%s' (use old, new, retained or extension)" % st)
            r.setdefault("aspect", STATUS_TO_ASPECT[st])
        if not r.get("description"):
            bits = ["Element %s" % r["element_nr"] if r.get("element_nr") else "",
                    r.get("element", ""), r.get("material", ""),
                    "(%s m2)" % r["area"] if r.get("area") else ""]
            r["description"] = " ".join(b for b in bits if b)
        # 'unit' in the element list describes the area column; quantities come from volume/mass
        if r.get("unit", "").lower() in ("m2", "m²"):
            r["unit"] = "kg" if r.get("mass") and not r.get("volume") else "m3"
        sid = r.pop("scheme", "") or "1"
        schemes.setdefault(re.sub(r"\D", "", sid) or "1", {"rows": []})["rows"].append(r)
    return {"schemes": schemes}


def _rows_from(obj):
    """dict {material: qty | {..}} or list of row dicts -> list of row dicts."""
    if isinstance(obj, dict) and "rows" in obj:
        obj = obj["rows"]
    if isinstance(obj, dict):
        out = []
        for mat, v in obj.items():
            r = dict(v) if isinstance(v, dict) else {"quantity": v}
            r.setdefault("material", mat)
            out.append(r)
        return out
    return [dict(r) for r in obj]


def normalise(data):
    data = _load(data)
    if isinstance(data, list):
        return {"project": {}, "schemes": {"1": {"rows": _rows_from(data)}}}
    low = dict((str(k).lower(), k) for k in data)
    if "schemes" in low or "project" in low:
        sch = data.get(low.get("schemes"), {}) or {}
        if isinstance(sch, list):
            sch = dict((str(i + 1), s) for i, s in enumerate(sch))
        return {"project": data.get(low.get("project"), {}) or {},
                "schemes": dict((re.sub(r"\D", "", str(k)) or "1", v) for k, v in sch.items())}
    if set(low) == {"old", "new"}:
        return {"project": {}, "schemes": {"1": {"name": "Baseline (old)", "rows": _rows_from(data[low["old"]])},
                                           "2": {"name": "Alternative (new)", "rows": _rows_from(data[low["new"]])}}}
    if all(re.match(r"scheme\s*\d$", str(k).lower()) for k in data):
        return {"project": {}, "schemes": dict((re.sub(r"\D", "", k), v if isinstance(v, dict) and "rows" in v
                                                else {"rows": _rows_from(v)}) for k, v in data.items())}
    return {"project": {}, "schemes": {"1": {"rows": _rows_from(data)}}}


# ============================================================================ main
class Inserter(object):
    def __init__(self, xlsm, db_path=None):
        self.wb = Workbook(xlsm)
        self.db = lca.load_db(db_path)
        self.log = []
        # live lists from the workbook, so values always match its dropdowns/lookups
        raw = [k for k in self.wb.read_col("LIFECYCLE MATERIAL DATA", "A", 5, 104) if k and not k.startswith("Custom_EPD")]
        self.raw_keys = dict((k.strip().lower(), k) for k in raw)
        self.elements = [e for e in self.wb.read_col("MATERIAL DATA", "Q", 5, 25) if e]
        self.rebar_labels = {}
        self.rebar_rows = {}
        for r in range(16, 46):
            name = self.wb.read("Custom Data", "S%d" % r)
            rate = self.wb.read("Custom Data", "T%d" % r)
            if name and isinstance(rate, float):
                self.rebar_labels[name] = rate
                self.rebar_rows[name] = r
        self.g_template = self.wb.formula("Scheme 1", "G12")

    # ---------------------------------------------------------------- lookups
    def material(self, label):
        warn = []
        key = lca.resolve_material(label, self.db, warn)
        self.log.extend(warn)
        raw = self.raw_keys.get(key.strip().lower())
        if raw is None:
            raise KeyError("'%s' is not in this workbook's LIFECYCLE MATERIAL DATA" % key)
        b, c, d = raw.split(" - ", 2)
        return raw, b, c, d

    def element(self, label):
        if not label:
            self.log.append("No element given - using 'Other'")
            return "Other"
        s = str(label).strip()
        for e in self.elements:
            if e.lower() == s.lower():
                return e
        if s.lower() in ELEMENT_ALIASES:
            return ELEMENT_ALIASES[s.lower()]
        hits = [e for e in self.elements if s.lower() in e.lower()]
        if len(hits) == 1:
            return hits[0]
        raise KeyError("Unknown structural element '%s'. Use one of: %s" % (s, "; ".join(ELEMENT_ALIASES)))

    def rebar(self, raw_key, element, given, custom_slots):
        b = raw_key.split(" - ")[0]
        if b != "Concrete" or "hollow core" in raw_key:
            return None
        if given in (None, ""):
            name = next((v for k, v in REBAR_BY_ELEMENT.items() if element.startswith(k)), "General Slab")
            if "Precast" in raw_key:
                name = "Precast Plank" if element.startswith("2.2") else "Precast Column"
            if "Mortar" in raw_key:
                name = "Mass Concrete"
            self.log.append("Rebar for %s / %s not given - assumed '%s' (%g kg/m3)"
                            % (raw_key.split(' - ')[-1], element.split(' ')[0], name, self.rebar_labels[name]))
            return "%s (%g)" % (name, self.rebar_labels[name])
        try:
            rate = float(given)
        except (TypeError, ValueError):
            rate = None
        if rate is None:
            g = re.sub(r"\s*\(.*\)$", "", str(given)).strip()
            for n in self.rebar_labels:
                if n.lower() == g.lower():
                    return "%s (%g)" % (n, self.rebar_labels[n])
            raise KeyError("Unknown rebar rate '%s'. Options: %s" % (given, ", ".join(self.rebar_labels)))
        for n, v in self.rebar_labels.items():
            if abs(v - rate) < 1e-9 and not n.startswith("Custom") and n != "N/A":
                return "%s (%g)" % (n, v)
        # numeric rate not in the table -> put it in a free 'Custom n' slot of the Custom Data sheet
        if rate in custom_slots:
            return custom_slots[rate]
        free = [n for n in sorted(self.rebar_labels) if n.startswith("Custom") and n not in custom_slots.values()
                and self.rebar_labels[n] == 0 and n not in [s.split(" (")[0] for s in custom_slots.values()]]
        if not free:
            raise ValueError("No free 'Custom' rebar slot left for %g kg/m3" % rate)
        n = free[0]
        self.wb.write("Custom Data", "T%d" % self.rebar_rows[n], rate)
        self.rebar_labels[n] = rate
        custom_slots[rate] = "%s (%g)" % (n, rate)
        self.log.append("Rebar %g kg/m3 stored in Custom Data as '%s'" % (rate, n))
        return custom_slots[rate]

    # ---------------------------------------------------------------- writing
    def write_project(self, p):
        for k, v in p.items():
            k2 = k.lower().replace(" ", "_")
            if k2 not in PROJECT_CELLS:
                self.log.append("Project field '%s' ignored (known: %s)" % (k, ", ".join(PROJECT_CELLS)))
                continue
            if k2 == "reinforcement" and not str(v).startswith("Steel"):
                v = "Steel - Rebar - " + ("UK" if str(v).strip().lower() == "uk" else "Global avg")
            if k2 == "typology" and str(v).lower() == "bridge":
                self.log.append("WARNING: typology 'Bridge' changes the element lists - aliases here are for buildings")
            if k2 == "design_stage" and str(v).lower() not in ("concept", "detailed", "as built"):
                raise ValueError("design_stage must be Concept, Detailed or As Built")
            self.wb.write("Project Info", PROJECT_CELLS[k2], v)

    def clear_scheme(self, sheet):
        for r in range(FIRST_ROW, LAST_ROW + 1):
            for col in INPUT_COLS:
                self.wb.write(sheet, "%s%d" % (col, r), None)
            if self.wb.formula(sheet, "G%d" % r) is None and self.g_template:
                self.wb.write_formula(sheet, "G%d" % r,
                                      re.sub(r"(\$[A-Z]{1,3})12(?!\d)", r"\g<1>%d" % r, self.g_template))

    def write_scheme(self, sid, scheme, clear=True):
        sheet = "Scheme %s" % sid
        if sheet not in self.wb.sheets:
            raise KeyError("Workbook has no sheet '%s' (use schemes 1-6)" % sheet)
        rows = _rows_from(scheme) if not (isinstance(scheme, dict) and "rows" in scheme) else _rows_from(scheme["rows"])
        if len(rows) > LAST_ROW - FIRST_ROW + 1:
            raise ValueError("%s: %d rows, TSCT allows max %d" % (sheet, len(rows), LAST_ROW - FIRST_ROW + 1))
        if clear:
            self.clear_scheme(sheet)
        if isinstance(scheme, dict):
            if scheme.get("name"):
                self.wb.write(sheet, "F3", scheme["name"])
            for key, ref in (("gia_demolished", "C9"), ("gia_retained", "D9"), ("gia_new", "E9"), ("gia", "E9")):
                if scheme.get(key) is not None:
                    self.wb.write(sheet, ref, float(scheme[key]))
        custom_slots = {}
        for i, r in enumerate(rows):
            n = FIRST_ROW + i
            r = dict((str(k).strip().lower(), v) for k, v in r.items())
            raw, b, c, d = self.material(r["material"])
            el = self.element(r.get("element"))
            unit = str(r.get("unit") or ("kg" if r.get("mass") not in (None, "") else "m3")).lower()
            qty = next((r[k] for k in ("quantity", "volume", "mass") if r.get(k) not in (None, "")), None)
            if qty is None:
                raise ValueError("%s row %d (%s): no volume or mass given%s" % (
                    sheet, n, r["material"], " - an area alone can't be converted to carbon" if r.get("area") else ""))
            qty = float(str(qty).replace(",", "."))
            if r.get("area") not in (None, "") and unit.startswith("m3"):
                area = float(str(r["area"]).replace(",", "."))
                if area > 0:
                    t = qty / area
                    for kw, (lo, hi) in THICKNESS_CHECK.items():
                        if kw in str(r["material"]).lower() and not lo <= t <= hi:
                            self.log.append("CHECK %s row %d: %s volume/area = %.3f m thick (expected %g-%g m)"
                                            % (sheet, n, r["material"], t, lo, hi))
            if qty < 0:
                raise ValueError("%s row %d (%s): negative quantity %g. TSCT only takes absolute quantities - "
                                 "put the old design in one scheme and the new one in another." % (sheet, n, r["material"], qty))
            self.wb.write(sheet, "B%d" % n, b)
            self.wb.write(sheet, "C%d" % n, c)
            self.wb.write(sheet, "D%d" % n, d)
            self.wb.write(sheet, "E%d" % n, el)
            self.wb.write(sheet, "F%d" % n, r.get("description") or r["material"])
            if r.get("lifespan") not in (None, ""):
                self.wb.write(sheet, "G%d" % n, float(r["lifespan"]))
            self.wb.write(sheet, "H%d" % n, r.get("aspect") or "New Build")
            tw = str(r.get("temporary_works") or "No").capitalize()
            self.wb.write(sheet, "I%d" % n, "Yes" if tw.startswith("Y") else "No")
            if tw.startswith("Y"):
                self.wb.write(sheet, "J%d" % n, float(r.get("temp_uses") or 1))
            self.wb.write(sheet, "K%d" % n, "Mass [kg]" if unit.startswith("kg") or unit == "mass" else "Volume [m3]")
            self.wb.write(sheet, "L%d" % n, qty)
            self.wb.write(sheet, "M%d" % n, self.rebar(raw, el, r.get("rebar"), custom_slots))
        self.log.append("%s: %d rows written%s" % (sheet, len(rows),
                                                   (" ('%s')" % scheme["name"]) if isinstance(scheme, dict) and scheme.get("name") else ""))


def insert(data, xlsm_in, xlsm_out=None, clear=True, db_path=None):
    """Fill the TSCT workbook from a dict / JSON / CSV. Returns a list of log messages."""
    if xlsm_out is None:
        root, ext = os.path.splitext(xlsm_in)
        xlsm_out = root + "_filled" + ext
    spec = normalise(data)
    ins = Inserter(xlsm_in, db_path)
    if spec.get("project"):
        ins.write_project(spec["project"])
    design_life = ins.wb.read("Project Info", "C18")
    if not design_life and "design_life" not in [k.lower() for k in spec.get("project", {})]:
        ins.write_project({"design_life": 50})
        ins.log.append("Design life was empty - set to 50 years")
    for sid in sorted(spec["schemes"], key=lambda s: int(s)):
        ins.write_scheme(sid, spec["schemes"][sid], clear=clear)
    ins.wb.save(xlsm_out)
    ins.log.append("Saved: %s  (Excel recalculates on opening - enable macros if asked)" % xlsm_out)
    return ins.log


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Write material quantities into the TSCT workbook")
    ap.add_argument("input", help="JSON or CSV file")
    ap.add_argument("workbook", help="TSCT .xlsm to read")
    ap.add_argument("output", nargs="?", help="output .xlsm (default: <workbook>_filled.xlsm)")
    ap.add_argument("--keep", action="store_true", help="don't clear existing rows in the target schemes")
    ap.add_argument("--db", default=None, help="path to tsct_materials.json")
    a = ap.parse_args()
    for line in insert(a.input, a.workbook, a.output, clear=not a.keep, db_path=a.db):
        print(line)
