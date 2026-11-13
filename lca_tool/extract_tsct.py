"""Extract the TSCT v3.1 material database into tsct_materials.json (run with CPython + openpyxl)."""
import json, sys, warnings
warnings.filterwarnings("ignore")
import openpyxl

def extract(xlsm, out):
    wb = openpyxl.load_workbook(xlsm, data_only=True, read_only=True)
    ws = wb["LIFECYCLE MATERIAL DATA"]
    mats = {}
    for r in ws.iter_rows(min_row=5, max_row=104, max_col=17, values_only=True):
        name = r[0]
        if not name or str(name).startswith("Custom_EPD"):
            continue
        f = lambda v: float(v) if v not in (None, "") else 0.0
        mats[name.strip()] = {
            "density_t_m3": f(r[3]), "biogenic": f(r[4]), "A1A3": f(r[5]), "A4": f(r[6]),
            "WF": f(r[7]), "C2": f(r[8]), "C34_1": f(r[9]), "C34_2": f(r[10]),
            "C34": f(r[11]), "C2C4": f(r[12]), "D1": f(r[13]), "lifespan": f(r[16]),
        }
    cd = wb["Custom Data"]
    rebar = {}
    for r in cd.iter_rows(min_row=16, max_row=45, min_col=19, max_col=20, values_only=True):
        if r[0] and isinstance(r[1], (int, float)):
            rebar[r[0]] = float(r[1])
    data = {
        "_source": "IStructE The Structural Carbon Tool v3.1 (April 2025), LIFECYCLE MATERIAL DATA. Factors in kgCO2e/kg.",
        "settings": {"rebar_material": "Steel - Rebar - Global avg", "design_life": 50,
                     "uncertainty_factor": 0.15, "decarbonisation_factor": 0.5},
        "rebar_rates_kg_m3": rebar,
        "materials": mats,
        "aliases": {},
    }
    json.dump(data, open(out, "w"), indent=1)
    print(len(mats), "materials,", len(rebar), "rebar rates ->", out)

if __name__ == "__main__":
    extract(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "tsct_materials.json")
