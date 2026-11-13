# -*- coding: utf-8 -*-
"""
gh_lca_delta.py - Grasshopper material-change -> embodied carbon (LCA) calculator
=================================================================================

Replicates the per-element maths of the IStructE "Structural Carbon Tool" (TSCT v3.1)
for a *change* in material quantities coming out of Grasshopper.

Input  : a dictionary  {material: changed_volume_m3}  (negative = material removed)
Output : embodied carbon per life-cycle module (A1-A3, A4, A5.3, B4, C2, C3/C4,
         biogenic, D) for every material and in total, in kgCO2e.

Works in CPython 3 (Rhino 8 / command line) and IronPython 2.7 (Rhino 7).

Accepted input formats (all may also be passed as a JSON string):
  1. {"Concrete - Insitu - UK C32/40 (25% GGBS)": -12.5, "CLT": 14.0}
  2. {"concrete C32/40": {"volume": -12.5, "rebar": "General Slab"},
      "reused steel":     {"mass": 800}}                    # kg instead of m3
  3. {"old": {...}, "new": {...}}       -> delta = new - old (per material)
  4. ["concrete C32/40: -12.5", "CLT: 14"]                   # GH panel lines
  5. [["concrete C32/40", -12.5], ["CLT", 14]]
  6. calculate(keys=[...], values=[...])                      # two GH lists

Material names may be the full TSCT name, an alias from the JSON database,
or a close match (fuzzy matching, reported in the warnings).

Grasshopper (GhPython / Rhino 8 Python 3 component)
---------------------------------------------------
  Inputs : data (str or dict), gia (float, optional), db (str, path to tsct_materials.json)
  Code   :
      import sys
      folder = r"C:\\Users\\you\\AEC_Hackathon\\scripts"
      if folder not in sys.path: sys.path.append(folder)
      import gh_lca_delta as lca
      try:
          import importlib; importlib.reload(lca)   # pick up edits while developing
      except Exception:
          reload(lca)                               # IronPython
      res = lca.calculate(data, db_path=db, gia=gia)
      report  = lca.format_report(res)              # -> Panel
      names, a_c, a1a3 = lca.gh_lists(res)          # -> charts / lists
      total_kg = res["totals"]["A_C_decarb"]

Command line
------------
  python gh_lca_delta.py change.json --gia 4500
  python gh_lca_delta.py --demo
"""
from __future__ import division, print_function

import difflib
import json
import math
import os

HERE = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
DEFAULT_DB = os.path.join(HERE, "tsct_materials.json")

MODULES = ["A1A3", "A4", "A5_3", "B4", "C2", "C34", "biogenic", "D"]

# Rebar defaults when Grasshopper gives no rate (TSCT "Reinforcement Rate Table", kg/m3)
DEFAULT_REBAR = {"Concrete - Insitu": "General Slab", "Concrete - Precast_concrete": "Precast Plank"}


# ----------------------------------------------------------------------------- database
def load_db(db_path=None):
    path = db_path or DEFAULT_DB
    with open(path) as f:
        db = json.load(f)
    db.setdefault("aliases", {})
    db["_lookup"] = dict((k.lower(), k) for k in db["materials"])
    return db


def resolve_material(name, db, warnings):
    """Map a Grasshopper material label to a TSCT material key."""
    raw = str(name).strip()
    key = raw.lower()
    if key in db["_lookup"]:
        return db["_lookup"][key]
    aliases = dict((k.lower(), v) for k, v in db["aliases"].items() if not k.startswith("_"))
    if key in aliases:
        return aliases[key]
    # substring match (e.g. "UK C32/40" or "Glulam UK")
    hits = [v for k, v in db["_lookup"].items() if key in k]
    if len(hits) == 1:
        warnings.append("'%s' matched by substring to '%s'" % (raw, hits[0]))
        return hits[0]
    # every word of the label appears in exactly one material name
    tokens = [t for t in key.replace("-", " ").split() if t]
    hits = [v for k, v in db["_lookup"].items() if all(t in k for t in tokens)]
    if len(hits) == 1:
        warnings.append("'%s' matched by keywords to '%s'" % (raw, hits[0]))
        return hits[0]
    if len(hits) > 1:
        raise KeyError("'%s' is ambiguous (%d matches, e.g. %s). Be more specific or add an alias."
                       % (raw, len(hits), "; ".join(hits[:3])))
    close = difflib.get_close_matches(key, list(db["_lookup"]) + list(aliases), n=1, cutoff=0.85)
    if close:
        target = db["_lookup"].get(close[0]) or aliases[close[0]]
        warnings.append("'%s' fuzzy-matched to '%s' - CHECK THIS" % (raw, target))
        return target
    raise KeyError("Unknown material '%s'. Add it to 'aliases' or 'materials' in the JSON database." % raw)


# ----------------------------------------------------------------------------- input parsing
def _num(v):
    return float(str(v).replace(",", ".")) if isinstance(v, str) else float(v)


def _normalise_item(v):
    """Return dict(volume=..|None, mass=..|None, rebar=..|None, element=..|None)."""
    if isinstance(v, dict):
        lv = dict((str(k).lower(), x) for k, x in v.items())
        vol = lv.get("volume", lv.get("volume_m3", lv.get("delta_volume", lv.get("m3"))))
        mass = lv.get("mass", lv.get("mass_kg", lv.get("kg")))
        return {"volume": None if vol is None else _num(vol),
                "mass": None if mass is None else _num(mass),
                "rebar": lv.get("rebar", lv.get("reinforcement")),
                "element": lv.get("element")}
    return {"volume": _num(v), "mass": None, "rebar": None, "element": None}


def parse_gh_input(data=None, keys=None, values=None):
    """Convert any supported Grasshopper output into {label: item-dict} (signed deltas)."""
    if keys is not None and values is not None:
        data = dict(zip(list(keys), list(values)))
    if data is None:
        raise ValueError("No input data")
    # GH often hands over .NET strings / lists
    if not isinstance(data, (dict, list, tuple)):
        s = str(data).strip()
        if s.startswith("{") or s.startswith("["):
            data = json.loads(s)
        else:
            data = [ln for ln in s.splitlines() if ln.strip()]
    if isinstance(data, (list, tuple)):
        d = {}
        for row in data:
            if isinstance(row, (list, tuple)) and len(row) >= 2:
                k, v = row[0], row[1]
            else:
                txt = str(row)
                sep = ":" if ":" in txt else ("=" if "=" in txt else ";")
                k, v = txt.rsplit(sep, 1)
            k = str(k).strip()
            d[k] = d.get(k, 0.0) + _num(str(v).strip()) if k in d else _num(str(v).strip())
        data = d
    # old / new comparison
    low = dict((str(k).lower(), k) for k in data)
    if "old" in low and "new" in low and len(data) == 2:
        old, new = data[low["old"]], data[low["new"]]
        out = {}
        for k in set(old) | set(new):
            o = _normalise_item(old.get(k, 0.0)) if k in old else None
            n = _normalise_item(new.get(k, 0.0)) if k in new else None
            base = n or o
            item = dict(base)
            for q in ("volume", "mass"):
                ov = (o or {}).get(q) or 0.0
                nv = (n or {}).get(q) or 0.0
                item[q] = (nv - ov) if ((o and o.get(q) is not None) or (n and n.get(q) is not None)) else None
            out[k] = item
        return out
    return dict((str(k), _normalise_item(v)) for k, v in data.items())


# ----------------------------------------------------------------------------- core maths
def _rebar_rate(mat_key, rebar, db, warnings, label):
    """kg rebar per m3 of concrete, following TSCT column BM/BO logic."""
    if not mat_key.startswith("Concrete") or "hollow core" in mat_key:
        return 0.0
    if rebar is None:
        group = " - ".join(mat_key.split(" - ")[:2])
        rebar = DEFAULT_REBAR.get(group)
        if rebar is None:
            return 0.0
        warnings.append("'%s': no rebar rate given, assumed '%s' (%g kg/m3)"
                        % (label, rebar, db["rebar_rates_kg_m3"][rebar]))
    if isinstance(rebar, (int, float)):
        return float(rebar)
    try:
        return float(rebar)
    except ValueError:
        pass
    if rebar not in db["rebar_rates_kg_m3"]:
        raise KeyError("Unknown rebar rate '%s'. Options: %s" % (rebar, ", ".join(db["rebar_rates_kg_m3"])))
    return db["rebar_rates_kg_m3"][rebar]


def element_carbon(mat_key, volume_m3=None, mass_kg=None, rebar_kg_m3=0.0, db=None, design_life=50):
    """
    Embodied carbon (kgCO2e) of one element, TSCT v3.1 scheme-sheet formulas
    (columns BQ..CD, O..X). Signed quantities are allowed (delta calculation).
    """
    m = db["materials"][mat_key]
    rb = db["materials"][db["settings"]["rebar_material"]]
    dens = m["density_t_m3"] * 1000.0                       # kg/m3
    if mass_kg is None:
        mass_kg = volume_m3 * dens                          # BQ
    if volume_m3 is None:
        volume_m3 = mass_kg / dens if dens else 0.0
    rebar_kg = rebar_kg_m3 * volume_m3                      # BP

    def f(field, with_rebar=True):
        return m[field] * mass_kg + (rb[field] * rebar_kg if with_rebar else 0.0)

    A1A3 = f("A1A3")                                        # BX
    A4 = f("A4")                                            # BY
    C2 = f("C2")                                            # BR
    C34 = f("C34_1") + f("C34_2", with_rebar=False)         # BS + BT (TSCT omits rebar in C34.2)
    C2C4 = f("C2C4")                                        # CB
    is_bio = mat_key.startswith("Timber") or "Bamboo" in mat_key
    bio = m["biogenic"] * mass_kg if is_bio else 0.0        # CC
    A5_3 = m["WF"] * (A1A3 + A4 + C2C4 + bio)               # BZ  waste incl. its own life cycle

    life = m["lifespan"]
    n_rep = int(math.ceil(design_life / life - 1 - 1e-12)) if 0 < life <= design_life else 0  # CI
    B4 = (A1A3 + A4 + A5_3 + C2 + C34 + bio) * n_rep        # S
    D = (m["D1"] * mass_kg + rb["D1"] * rebar_kg) * (n_rep + 1)  # W

    return {"mass_kg": mass_kg, "rebar_kg": rebar_kg, "volume_m3": volume_m3, "replacements": n_rep,
            "A1A3": A1A3, "A4": A4, "A5_3": A5_3, "B4": B4, "C2": C2, "C34": C34,
            "biogenic": bio, "D": D}


def _totals(rows, settings):
    t = dict((k, sum(r[k] for r in rows)) for k in MODULES)
    t["A1A5_material"] = t["A1A3"] + t["A4"] + t["A5_3"]
    t["A_C"] = t["A1A5_material"] + t["B4"] + t["C2"] + t["C34"] + t["biogenic"]
    dec = settings.get("decarbonisation_factor", 0.5)
    # TSCT: future grid decarbonisation applied to module B and C2 transport (and D)
    t["A_C_decarb"] = t["A1A5_material"] + dec * (t["B4"] + t["C2"]) + t["C34"] + t["biogenic"]
    t["D_decarb"] = dec * t["D"]
    uf = settings.get("uncertainty_factor", 0.15)
    t["A_C_decarb_withUF"] = t["A_C_decarb"] * (1 + uf)
    t["A1A5_withUF"] = t["A1A5_material"] * (1 + uf)
    return t


def calculate(data=None, keys=None, values=None, db_path=None, gia=None,
              design_life=None, uncertainty_factor=None, rebar_material=None):
    """Main entry point. Returns dict with 'rows', 'totals', 'per_m2', 'warnings'."""
    db = load_db(db_path)
    s = dict(db["settings"])
    if design_life:
        s["design_life"] = float(design_life)
    if uncertainty_factor is not None:
        s["uncertainty_factor"] = float(uncertainty_factor)
    if rebar_material:
        s["rebar_material"] = resolve_material(rebar_material, db, [])
    db["settings"] = s

    warnings = []
    items = parse_gh_input(data, keys, values)
    rows = []
    for label, it in sorted(items.items()):
        key = resolve_material(label, db, warnings)
        if it["volume"] is None and it["mass"] is None:
            continue
        rate = _rebar_rate(key, it.get("rebar"), db, warnings, label)
        r = element_carbon(key, it["volume"], it["mass"], rate, db, s["design_life"])
        r.update({"input": label, "material": key, "element": it.get("element"), "rebar_kg_m3": rate})
        r["A_C"] = r["A1A3"] + r["A4"] + r["A5_3"] + r["B4"] + r["C2"] + r["C34"] + r["biogenic"]
        rows.append(r)

    totals = _totals(rows, s)
    out = {"rows": rows, "totals": totals, "settings": s, "warnings": warnings, "unit": "kgCO2e"}
    if gia:
        out["per_m2"] = dict((k, v / float(gia)) for k, v in totals.items())
        out["gia_m2"] = float(gia)
    return out


# ----------------------------------------------------------------------------- outputs
def format_report(res):
    L = []
    w = L.append
    w("EMBODIED CARBON OF MATERIAL CHANGE  (kgCO2e, negative = saving)")
    w("Method: IStructE Structural Carbon Tool v3.1 factors, design life %g yr" % res["settings"]["design_life"])
    w("-" * 96)
    w("%-46s %9s %9s %9s %9s %9s" % ("Material", "dV [m3]", "A1-A3", "A1-A5", "A-C", "D"))
    for r in res["rows"]:
        a15 = r["A1A3"] + r["A4"] + r["A5_3"]
        w("%-46s %9.2f %9.0f %9.0f %9.0f %9.0f" % (r["material"][:46], r["volume_m3"], r["A1A3"], a15, r["A_C"], r["D"]))
        if r["rebar_kg"]:
            w("    incl. rebar %.0f kg (%g kg/m3)" % (r["rebar_kg"], r["rebar_kg_m3"]))
    t = res["totals"]
    w("-" * 96)
    for lbl, k in [("A1-A3 product", "A1A3"), ("A4 transport", "A4"), ("A5.3 waste", "A5_3"),
                   ("B4 replacement", "B4"), ("C2 transport EoL", "C2"), ("C3/C4 processing", "C34"),
                   ("Biogenic carbon (stored)", "biogenic"),
                   ("TOTAL A1-A5 (upfront)", "A1A5_material"), ("TOTAL A-C (decarbonised)", "A_C_decarb"),
                   ("TOTAL A-C incl. uncertainty factor", "A_C_decarb_withUF"),
                   ("Module D (reported separately)", "D_decarb")]:
        extra = ("   (%.1f /m2)" % res["per_m2"][k]) if "per_m2" in res else ""
        w("%-38s %12.0f kgCO2e%s" % (lbl, t[k], extra))
    if res["warnings"]:
        w("-" * 96)
        w("WARNINGS:")
        for x in res["warnings"]:
            w("  - " + x)
    return "\n".join(L)


def gh_lists(res, module="A_C"):
    """Parallel lists for Grasshopper charts: names, A-C, A1-A3."""
    names = [r["material"] for r in res["rows"]]
    return names, [r[module] for r in res["rows"]], [r["A1A3"] for r in res["rows"]]


DEMO = {
    "old": {"concrete C32/40": {"volume": 420.0, "rebar": "General Slab"},
            "steel sections": {"mass": 38000}},
    "new": {"concrete C32/40": {"volume": 150.0, "rebar": "General Slab"},
            "CLT": 310.0, "glulam": 45.0,
            "steel sections": {"mass": 12000}, "reused steel": {"mass": 20000}},
}

if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    p.add_argument("input", nargs="?", help="JSON file with the Grasshopper dictionary")
    p.add_argument("--db", default=None)
    p.add_argument("--gia", type=float, default=None, help="floor area for kgCO2e/m2")
    p.add_argument("--life", type=float, default=None, help="design life (years)")
    p.add_argument("--json", action="store_true", help="print raw JSON result")
    p.add_argument("--demo", action="store_true")
    a = p.parse_args()
    data = DEMO if a.demo or not a.input else json.load(open(a.input))
    res = calculate(data, db_path=a.db, gia=a.gia, design_life=a.life)
    print(json.dumps(res, indent=1) if a.json else format_report(res))
