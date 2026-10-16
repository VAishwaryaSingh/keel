"""Generates model/snapshots.json: engine-run values behind Lever_Analysis, Sensitivity and Business_Review.

Usage: python snapshots.py model/keel_saas_model.xlsx model/snap_spec.json model/snapshots.json
The workbook has no Excel data tables (not portable to Sheets), so each lever / grid point is produced by
changing inputs on the Custom scenario column, recalculating the whole workbook with pycel, and recording outputs.
build_model.py then pastes these values (grey italic) and the Checks tab flags them if they go stale.
Requires: pip install openpyxl pycel   (run inside a virtualenv)
"""
import json
import sys
import openpyxl
from pycel import ExcelCompiler
from pycel.excelwrapper import ExcelOpxWrapper


def _defined_names(self):
    if self._defined_names is None:
        self._defined_names = {}
        for name, dn in self.workbook.defined_names.items():
            dest = [(alias, sheet) for sheet, alias in dn.destinations if sheet in self.workbook]
            if dest:
                self._defined_names[str(name)] = dest
    return self._defined_names


ExcelOpxWrapper.defined_names = property(_defined_names)

path, spec_path, out_path = sys.argv[1:4]
spec = json.load(open(spec_path))
wb = openpyxl.load_workbook(path)
xl = ExcelCompiler(filename=path)
SEL = "Inputs_Drivers!" + spec["selector"].replace("$", "")
cells = {n: "Inputs_Scenarios!" + c for n, c in spec["scen_cells"].items()}
scen_ws = wb["Inputs_Scenarios"]
base_val = {n: scen_ws[c].value for n, c in spec["scen_cells"].items()}   # column F holds Custom; C holds Base
base_val = {n: scen_ws["C" + c[1:]].value for n, c in spec["scen_cells"].items()}


def setv(addr, value):
    xl.evaluate(addr)            # pycel only lets you overwrite cells it has already placed in its map
    xl.set_value(addr, value)


def metrics():
    out = {}
    for mid, ref in spec["metrics"].items():
        v = xl.evaluate(ref)
        out[mid] = v if isinstance(v, (int, float, str)) else None
    return out


def reset_custom():
    for n, addr in cells.items():
        setv(addr, base_val[n])


snap = {"scen": {}, "levers": {}}
for sc in ("Base", "Upside", "Downside"):
    setv(SEL, sc)
    snap["scen"][sc] = metrics()
snap["base"] = snap["scen"]["Base"]

setv(SEL, "Custom")
reset_custom()
chk = metrics()
for k in ("arr24", "cash24", "ebitda_fy2", "mincash"):
    assert abs(chk[k] - snap["base"][k]) < 1e-6, ("Custom != Base", k, chk[k], snap["base"][k])

for lv in spec["levers"]:
    d, st = lv["driver"], lv["step"]
    snap["levers"][d] = {}
    for tag, sign in (("dn", -1), ("up", 1)):
        reset_custom()
        setv(cells[d], base_val[d] + sign * st)
        m = metrics()
        snap["levers"][d][tag] = {k: m[k] for k in spec["lever_metrics"]}
    print("lever", d, flush=True)

grid, arr_scale = [], []
for cs in spec["churn_shifts"]:
    row = []
    for sc in spec["scales"]:
        reset_custom()
        for seg in ("ent", "mid"):
            setv(cells[f"inp_churn_{seg}"], base_val[f"inp_churn_{seg}"] + cs)
            setv(cells[f"inp_logos_{seg}"], base_val[f"inp_logos_{seg}"] * sc)
        m = metrics()
        row.append(m["mincash"])
        if abs(cs) < 1e-12:
            arr_scale.append(m["arr24"])
    grid.append(row)
    print("grid row", cs, flush=True)
snap["grid_mincash"] = grid
snap["arr24_by_scale"] = arr_scale
setv(SEL, "Base")
json.dump(snap, open(out_path, "w"), indent=1)
print("wrote", out_path)
