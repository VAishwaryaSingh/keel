"""Acceptance test 5.9 #5, the independent half: type the goal-seek answer found by the website into the workbook
(recalculated by pycel, not by HyperFormula) and confirm the output reaches the target.
Usage: python verify_goalseek.py model/keel_saas_model.xlsx model/snap_spec.json web/tests/fixtures/goalseek_result.json
"""
import json
import re
import sys

import openpyxl
from pycel import ExcelCompiler
from pycel.excelwrapper import ExcelOpxWrapper

path, spec_path, res_path = sys.argv[1:4]
spec, res = json.load(open(spec_path)), json.load(open(res_path))


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
xl = ExcelCompiler(filename=path)
sel = "Inputs_Drivers!" + spec["selector"].replace("$", "")
xl.evaluate(sel)
xl.set_value(sel, "Custom")
cell = "Inputs_Scenarios!" + spec["scen_cells"][res["driver"]]     # scenario-driven inputs are typed into the Custom column
xl.evaluate(cell)
xl.set_value(cell, res["x"])
got = xl.evaluate(spec["metrics"][res["output"]])
gap = got - res["target"]
print(f"{res['driver']} = {res['x']:.10f} typed into the workbook -> {res['output']} = {got:.6f} (target {res['target']}, gap {gap:.2e})")
sys.exit(0 if abs(gap) <= 0.01 else 1)
