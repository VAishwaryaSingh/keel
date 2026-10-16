"""Reference answers for the website's acceptance tests 5.9 #3 and #4, computed by recalculating the whole workbook
independently (pycel): a manual +/-10% one-at-a-time test and one two-way grid cell.
Usage: python export_sensitivity_fixtures.py model/keel_saas_model.xlsx model/snap_spec.json web/tests/fixtures
"""
import json
import os
import re
import sys

import openpyxl
from pycel import ExcelCompiler
from pycel.excelwrapper import ExcelOpxWrapper

path, spec_path, outdir = sys.argv[1:4]
spec = json.load(open(spec_path))


def _dn(self):
    if self._defined_names is None:
        self._defined_names = {}
        for name, dn in self.workbook.defined_names.items():
            dest = [(a, sh) for sh, a in dn.destinations if sh in self.workbook]
            if dest:
                self._defined_names[str(name)] = dest
    return self._defined_names


ExcelOpxWrapper.defined_names = property(_dn)
wb = openpyxl.load_workbook(path)
dws, sws = wb["Inputs_Drivers"], wb["Inputs_Scenarios"]
xl = ExcelCompiler(filename=path)
SEL = "Inputs_Drivers!" + spec["selector"].replace("$", "")
xl.evaluate(SEL)
xl.set_value(SEL, "Custom")

# input cells: scenario-driven inputs live in the Custom column; static inputs on Inputs_Drivers
info = {}
for nm, d in wb.defined_names.items():
    if not nm.startswith("inp_") or nm in ("inp_scenario", "inp_scen_idx"):
        continue
    r = int(re.sub(r"\D", "", d.attr_text.split("!")[1]))
    v = dws[f"C{r}"].value
    if isinstance(v, str) and v.startswith("=INDEX"):
        sr = int(re.search(r"\$C(\d+):", v).group(1))
        info[nm] = dict(kind="scenario", cell=f"Inputs_Scenarios!F{sr}", base=sws[f"C{sr}"].value, min=dws[f"D{r}"].value, max=dws[f"E{r}"].value)
    else:
        info[nm] = dict(kind="static", cell=f"Inputs_Drivers!C{r}", base=v, min=dws[f"D{r}"].value, max=dws[f"E{r}"].value, unit=dws[f"B{r}"].value)
for nm, i in info.items():
    xl.evaluate(i["cell"])


def reset():
    for nm, i in info.items():
        xl.set_value(i["cell"], i["base"])


def outputs(ids):
    return {k: xl.evaluate(spec["metrics"][k]) for k in ids}


IDS = ["arr24", "cash24", "ebitda_fy2", "mincash", "ma_value", "rev_fy1", "ebitda_fy1", "pf_cash24"]
EXTRA = ["inp_marketing_m", "inp_sal_rd", "inp_sal_sm", "inp_host_ent", "inp_ma_offer_mult", "inp_ma_cost_syn", "inp_ma_integration_total", "inp_ma_wacc"]
drivers = [n for n, i in info.items() if (i["kind"] == "scenario" or n in EXTRA) and not (i["min"] == 0 and i["max"] == 1 and n in ("inp_ma_include", "inp_capitalise_commission"))]
reset()
base = outputs(IDS)
torn = {}
for n in drivers:
    i = info[n]
    v = i["base"]
    if v == 0:
        continue
    lo, hi = sorted((v * 0.9, v * 1.1))
    lo, hi = max(lo, i["min"]), min(hi, i["max"])
    reset()
    xl.set_value(i["cell"], lo)
    a = outputs(IDS)
    reset()
    xl.set_value(i["cell"], hi)
    b = outputs(IDS)
    torn[n] = dict(base=v, lo=lo, hi=hi, atLo=a, atHi=b)
reset()
# one grid cell: churn_ent x logos_ent
xl.set_value(info["inp_churn_ent"]["cell"], 0.08)
xl.set_value(info["inp_logos_ent"]["cell"], 2.5)
cell = outputs(IDS)
reset()
os.makedirs(outdir, exist_ok=True)
json.dump(dict(base=base, tornado=torn, grid_cell=dict(x="inp_logos_ent", xv=2.5, y="inp_churn_ent", yv=0.08, outputs=cell)), open(os.path.join(outdir, "sensitivity_reference.json"), "w"), indent=1)
print("tornado drivers:", len(torn))
