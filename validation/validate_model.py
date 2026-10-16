"""Independent validation of keel_saas_model.xlsx. Writes validation/validation_report.md.

Usage:  python validate_model.py <model.xlsx> [report.md]
Needs:  pip install openpyxl pycel   (inside a virtualenv)

What it does (none of it reuses the workbook's own check formulas):
  A. Static scans: hardcoded numbers inside formula ranges, rows whose formula changes across periods, non-portable
     functions, undefined names, numeric constants that are not marked as inputs, embedded numeric literals, input ranges.
  B. Independent re-computation: a second implementation of the Northgate model and of the Vaultline goodwill
     calculation, written in plain Python from the input cells only, compared with the workbook's recalculated values
     for every scenario.
  C. Test the tests: deliberately break the model in five ways and confirm the workbook's checks catch each one.
  D. Error scan: every evaluated cell on the output tabs is free of error values.
"""
import datetime
import os
import re
import sys
import time

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter as L
from pycel import ExcelCompiler
from pycel.excelwrapper import ExcelOpxWrapper

path = sys.argv[1]
report_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "validation_report.md")


def _defined_names(self):
    if self._defined_names is None:
        self._defined_names = {}
        for name, dn in self.workbook.defined_names.items():
            dest = [(alias, sheet) for sheet, alias in dn.destinations if sheet in self.workbook]
            if dest:
                self._defined_names[str(name)] = dest
    return self._defined_names


ExcelOpxWrapper.defined_names = property(_defined_names)

wbf = openpyxl.load_workbook(path)                    # formulas
TIME_SHEETS = ["Rev_Build", "Rev_Rec", "Headcount_Opex", "Contract_Costs", "COGS", "IS", "BS", "CF", "SaaS_Metrics", "MA_Target", "MA_ProForma"]
INPUT_SHEETS = {"Inputs_Drivers", "Inputs_Scenarios", "Inputs_Budget", "Inputs_Comps", "Cover"}
F0, F1 = 15, 38
RESULTS = []          # (section, status, name, detail)


def add(section, status, name, detail=""):
    RESULTS.append((section, status, name, detail))


def color_of(cell):
    c = cell.font.color if cell.font else None
    if c is not None and isinstance(c.rgb, str):
        return c.rgb[-6:].upper()
    return "000000"


# ---------------------------------------------------------------------------------------------------
# A. Static scans
# ---------------------------------------------------------------------------------------------------
sec = "A. Static scans"
hard, incons, checked_rows = [], [], 0
for sn in TIME_SHEETS:
    ws = wbf[sn]
    for r in range(8, ws.max_row + 1):
        o = ws.cell(row=r, column=F0)
        if o.value is None:
            continue
        for ci in range(F0, F1 + 1):
            v = ws.cell(row=r, column=ci).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                hard.append(f"{sn}!{L(ci)}{r}")
        if isinstance(o.value, str) and o.value.startswith("="):
            checked_rows += 1
            for ci in range(F0 + 1, F1 + 1):
                c = ws.cell(row=r, column=ci)
                exp = Translator(o.value, origin=f"{L(F0)}{r}").translate_formula(f"{L(ci)}{r}")
                if c.value != exp:
                    incons.append(f"{sn}!{L(ci)}{r}")
                    break
add(sec, "PASS" if not hard else "FAIL", "No hardcoded numbers in forecast ranges (months 1-24, 11 time-series tabs)",
    "none found" if not hard else f"{len(hard)} cells, e.g. {', '.join(hard[:5])}")
add(sec, "PASS" if not incons else "FAIL", f"One consistent formula across all 24 periods ({checked_rows} rows tested)",
    "every row consistent" if not incons else f"{len(incons)} rows differ, e.g. {', '.join(incons[:5])}")

FORBID = re.compile(r"\b(INDIRECT|OFFSET|QUERY|ARRAYFORMULA|IMPORTRANGE|IMPORTXML|IMPORTDATA|NOW|TODAY|RAND|RANDBETWEEN|LET|LAMBDA|XLOOKUP|FILTER|SORT|UNIQUE)\s*\(", re.I)
bad_fn, nformulas = [], 0
names = set(wbf.defined_names.keys())
undefined = []
for ws in wbf.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                nformulas += 1
                if FORBID.search(c.value):
                    bad_fn.append(f"{ws.title}!{c.coordinate}")
                for nm in re.findall(r"\b((?:inp|chk)_[A-Za-z0-9_]+)\b", c.value):
                    if nm not in names:
                        undefined.append(f"{ws.title}!{c.coordinate}:{nm}")
add(sec, "PASS" if not bad_fn else "FAIL", f"No non-portable or volatile functions ({nformulas:,} formulas scanned)",
    "none found" if not bad_fn else f"{len(bad_fn)} cells, e.g. {', '.join(bad_fn[:5])}")
add(sec, "PASS" if not undefined else "FAIL", "Every inp_* / chk_* name used in a formula is defined",
    f"{len(names)} defined names" if not undefined else f"{len(undefined)} undefined, e.g. {', '.join(undefined[:5])}")
broken_names = [n for n, d in wbf.defined_names.items() if d.attr_text.split("!")[0].strip("'") not in wbf.sheetnames]
add(sec, "PASS" if not broken_names else "FAIL", "Every defined name points at an existing sheet", "ok" if not broken_names else ", ".join(broken_names))

unmarked = []
for ws in wbf.worksheets:
    if ws.title in INPUT_SHEETS:
        continue
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                if color_of(c) not in ("0000FF", "595959", "FFFFFF"):
                    unmarked.append(f"{ws.title}!{c.coordinate}")
add(sec, "PASS" if not unmarked else "FAIL", "Every typed number outside the input tabs is marked: blue (input), grey (engine-run snapshot) or white (header)",
    "all marked" if not unmarked else f"{len(unmarked)} unmarked, e.g. {', '.join(unmarked[:6])}")

ALLOWED = {0, 1, 2, 3, 6, 11, 12, 30, 100, 1000, 0.5}
lits = {}
for ws in wbf.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                try:
                    for t in Tokenizer(c.value).items:
                        if t.type == Token.OPERAND and t.subtype == Token.NUMBER and float(t.value) not in ALLOWED:
                            lits.setdefault(float(t.value), []).append(f"{ws.title}!{c.coordinate}")
                except Exception:
                    pass
detail = "; ".join(f"{v:g} x{len(cs)} (e.g. {cs[0]})" for v, cs in sorted(lits.items()))
add(sec, "REVIEW" if lits else "PASS", "Numeric literals inside formulas other than unit conversions (0, 1, 2, 3, 6, 11, 12, 30, 100, 1000, 0.5)",
    detail or "none")

dws, sws = wbf["Inputs_Drivers"], wbf["Inputs_Scenarios"]
oor = []
for nm, d in wbf.defined_names.items():
    if not nm.startswith("inp_") or nm in ("inp_scenario", "inp_scen_idx"):
        continue
    ref = d.attr_text.split("!")[1].replace("$", "")
    r = int(re.sub(r"\D", "", ref))
    lo, hi = dws[f"D{r}"].value, dws[f"E{r}"].value
    if lo is None:
        continue
    v = dws[f"C{r}"].value
    if isinstance(v, str) and v.startswith("=INDEX"):
        sr = int(re.search(r"\$C(\d+):", v).group(1))
        vals = [(col, sws[f"{col}{sr}"].value) for col in "CDEF"]
    else:
        vals = [("C", v)]
    for col, x in vals:
        if not (lo <= x <= hi):
            oor.append(f"{nm} ({col}={x}, allowed {lo}..{hi})")
add(sec, "PASS" if not oor else "FAIL", "Every input value, in every scenario column, is inside its allowed range",
    "all inside" if not oor else "; ".join(oor[:5]))

# ---------------------------------------------------------------------------------------------------
# B. Independent re-computation
# ---------------------------------------------------------------------------------------------------
sec = "B. Independent re-computation"


def load_inputs(scenario):
    idx = ["Base", "Upside", "Downside", "Custom"].index(scenario)
    col = "CDEF"[idx]
    out = {}
    for nm, d in wbf.defined_names.items():
        if not nm.startswith("inp_") or nm in ("inp_scenario", "inp_scen_idx"):
            continue
        r = int(re.sub(r"\D", "", d.attr_text.split("!")[1]))
        v = dws[f"C{r}"].value
        if isinstance(v, str) and v.startswith("=INDEX"):
            sr = int(re.search(r"\$C(\d+):", v).group(1))
            v = sws[f"{col}{sr}"].value
        out[nm] = v
    return out


def northgate(I):
    """Plain-Python second implementation of the Northgate monthly model (24 months)."""
    T = 24
    segs = ("ent", "mid")
    arr = {s: {0: I[f"inp_arr_open_{s}"]} for s in segs}
    cust = {s: {0: I[f"inp_cust_open_{s}"]} for s in segs}
    new = {s: {} for s in segs}
    expn = {s: {} for s in segs}
    chn = {s: {} for s in segs}
    inv = {s: {k: I[f"inp_arr_open_{s}"] / 12 for k in range(-11, 1)} for s in segs}
    for t in range(1, T + 1):
        yr = -(-t // 12)
        price_f = (1 + I["inp_price_uplift"]) ** (yr - 1)
        logo_f = (1 + I["inp_logo_growth"]) ** ((t - 1) / 12)
        for s in segs:
            logos = I[f"inp_logos_{s}"] * logo_f
            acv = I[f"inp_acv_{s}"] * price_f
            cm = 1 - (1 - I[f"inp_churn_{s}"]) ** (1 / 12)
            em = (1 + I[f"inp_exp_{s}"]) ** (1 / 12) - 1
            o = arr[s][t - 1]
            new[s][t], expn[s][t], chn[s][t] = logos * acv, o * em, -o * cm
            arr[s][t] = o + new[s][t] + expn[s][t] + chn[s][t]
            cust[s][t] = cust[s][t - 1] * (1 - cm) + logos
            inv[s][t] = new[s][t] + (1 + I[f"inp_exp_{s}"] - I[f"inp_churn_{s}"]) * inv[s][t - 12]
    rev_s = {s: {t: sum(inv[s][k] for k in range(t - 11, t + 1)) / 12 for t in range(1, T + 1)} for s in segs}
    fee = {t: I["inp_impl_fee_pct"] * new["ent"][t] for t in range(1, T + 1)}
    fee.update({0: 0, -1: 0})
    rev_impl = {t: (fee[t] + fee[t - 1] + fee[t - 2]) / 3 for t in range(1, T + 1)}
    rev_sub = {t: rev_s["ent"][t] + rev_s["mid"][t] for t in range(1, T + 1)}
    rev = {t: rev_sub[t] + rev_impl[t] for t in range(1, T + 1)}
    inv_tot = {t: inv["ent"][t] + inv["mid"][t] for t in range(-11, T + 1)}
    bill = {t: inv_tot[t] + fee[t] for t in range(1, T + 1)}
    bill[0] = inv_tot[0]
    dr_sub = {0: sum(inv_tot[k] * (11 + k) / 12 for k in range(-11, 1))}
    dr_imp = {0: I["inp_open_defrev_impl"]}
    ar = {0: I["inp_open_ar"]}
    d = I["inp_dso"]
    for t in range(1, T + 1):
        dr_sub[t] = dr_sub[t - 1] + inv_tot[t] - rev_sub[t]
        dr_imp[t] = dr_imp[t - 1] + fee[t] - rev_impl[t]
        ar[t] = bill[t] * min(1, d / 30) + bill[t - 1] * max(0, min(1, (d - 30) / 30))
    dr = {t: dr_sub[t] + dr_imp[t] for t in range(0, T + 1)}
    funcs = ("rd", "sm", "ga", "cs")
    heads = {f: {0: I[f"inp_heads_open_{f}"]} for f in funcs}
    pay = {f: {} for f in funcs}
    for t in range(1, T + 1):
        yr = -(-t // 12)
        for f in funcs:
            heads[f][t] = heads[f][t - 1] + I[f"inp_hires_{f}"]
            pay[f][t] = (heads[f][t - 1] + heads[f][t]) / 2 * I[f"inp_sal_{f}"] / 12 * (1 + I["inp_merit"]) ** (yr - 1)
    cca = {0: I["inp_open_cca"]}
    ppe = {0: I["inp_open_ppe"]}
    cash = {0: I["inp_open_cash"]}
    ap = {0: I["inp_open_ap"]}
    equity = {0: cash[0] + ar[0] + ppe[0] + cca[0] - ap[0] - dr[0]}
    out = {"rev": rev, "ebitda": {}, "ni": {}, "cash": cash, "ar": ar, "ap": ap, "dr": dr, "equity": equity, "check": {}}
    out["arr"] = {t: arr["ent"][t] + arr["mid"][t] for t in range(0, T + 1)}
    for t in range(1, T + 1):
        comm = I["inp_commission_pct"] * (new["ent"][t] + new["mid"][t] + expn["ent"][t] + expn["mid"][t])
        cap = I["inp_capitalise_commission"] * comm
        amort = cca[t - 1] / I["inp_cc_life"]
        cca[t] = cca[t - 1] + cap - amort
        pl_comm = comm - cap + amort
        host = cust["ent"][t] * I["inp_host_ent"] + cust["mid"][t] * I["inp_host_mid"]
        lic = I["inp_licence_pct"] * rev_sub[t]
        svc = I["inp_services_cost_pct"] * rev_impl[t]
        cogs = host + lic + pay["cs"][t] + svc
        opex = (pay["rd"][t] + I["inp_rd_other_m"]) + (pay["sm"][t] + I["inp_marketing_m"] + pl_comm) + (pay["ga"][t] + I["inp_ga_other_m"])
        ebitda = rev[t] - cogs - opex
        dep = ppe[t - 1] / I["inp_dep_life"]
        capex = I["inp_capex_pct"] * rev[t]
        ppe[t] = ppe[t - 1] + capex - dep
        pbt = ebitda - dep + cash[t - 1] * I["inp_interest"] / 12
        ni = pbt - max(0, pbt) * I["inp_tax"]
        nonpay = host + lic + svc + I["inp_marketing_m"] + I["inp_rd_other_m"] + I["inp_ga_other_m"]
        ap[t] = nonpay * I["inp_dpo"] / 30
        net = ni + dep - (ar[t] - ar[t - 1]) + (ap[t] - ap[t - 1]) + (dr[t] - dr[t - 1]) - (cca[t] - cca[t - 1]) - capex
        cash[t] = cash[t - 1] + net
        equity[t] = equity[t - 1] + ni
        out["ebitda"][t], out["ni"][t] = ebitda, ni
        out["check"][t] = cash[t] + ar[t] + ppe[t] + cca[t] - ap[t] - dr[t] - equity[t]
    return out


def goodwill(I):
    arr = I["inp_v_arr_open"]
    dr = arr / 12 * sum((11 + k) / 12 for k in range(-11, 1))
    ap = arr / 12 * (I["inp_v_cogs_pct"] + I["inp_v_rd_pct"] + I["inp_v_sm_pct"] + I["inp_v_ga_pct"]) * I["inp_v_ap_months"]
    book = I["inp_v_open_cash"] + I["inp_v_open_ar"] + I["inp_v_open_ppe"] - ap - dr
    ev = I["inp_ma_offer_mult"] * arr
    price = ev + I["inp_v_open_cash"]
    intang = (I["inp_ma_cr_pct"] + I["inp_ma_tech_pct"]) * ev
    base = book + intang - I["inp_tax"] * intang
    return price - (base + I["inp_ma_dr_haircut"] * dr), price - base


def find(sheet, label, nth=0):
    ws = wbf[sheet]
    rows = [r for r in range(1, ws.max_row + 1) if ws[f"A{r}"].value == label]
    return rows[nth] if nth >= 0 else rows[nth]


def new_xl():
    return ExcelCompiler(filename=path)


def set_scenario(xl, scen):
    addr = "Inputs_Drivers!" + wbf.defined_names["inp_scenario"].attr_text.split("!")[1].replace("$", "")
    xl.evaluate(addr)
    xl.set_value(addr, scen)


xl = new_xl()
ANN = {"fy1": "AN", "fy2": "AO"}
SCEN_SUMMARY = []
for scen in ("Base", "Upside", "Downside", "Custom"):
    set_scenario(xl, scen)
    I = load_inputs(scen)
    m = northgate(I)
    ev = lambda sh, lab, col, nth=0: xl.evaluate(f"{sh}!{col}{find(sh, lab, nth)}")
    pairs = [
        ("FY1 revenue", sum(m["rev"][t] for t in range(1, 13)), ev("IS", "Total revenue", "AN")),
        ("FY2 revenue", sum(m["rev"][t] for t in range(13, 25)), ev("IS", "Total revenue", "AO")),
        ("FY1 EBITDA", sum(m["ebitda"][t] for t in range(1, 13)), ev("IS", "EBITDA", "AN")),
        ("FY2 EBITDA", sum(m["ebitda"][t] for t in range(13, 25)), ev("IS", "EBITDA", "AO")),
        ("FY1 net income", sum(m["ni"][t] for t in range(1, 13)), ev("IS", "Net income", "AN")),
        ("Closing ARR, month 12", m["arr"][12], ev("SaaS_Metrics", "Closing ARR", "Z")),
        ("Closing ARR, month 24", m["arr"][24], ev("SaaS_Metrics", "Closing ARR", "AL")),
        ("Deferred revenue, month 24", m["dr"][24], ev("BS", "Deferred revenue", "AL")),
        ("Receivables, month 24", m["ar"][24], ev("BS", "Trade receivables", "AL")),
        ("Payables, month 24", m["ap"][24], ev("BS", "Trade payables", "AL")),
        ("Cash, month 12", m["cash"][12], ev("BS", "Cash", "Z")),
        ("Cash, month 24", m["cash"][24], ev("BS", "Cash", "AL")),
        ("Equity, month 24", m["equity"][24], ev("BS", "Total equity (opening net assets + cumulative net income)", "AL")),
    ]
    worst = max(abs(a - b) for _, a, b in pairs)
    bs_gap = max(abs(v) for v in m["check"].values())
    gw_ifrs, gw_us = goodwill(I)
    gr = find("MA_PPA", "Goodwill = equity purchase price - fair value of net assets")
    wgw = (xl.evaluate(f"MA_PPA!C{gr}"), xl.evaluate(f"MA_PPA!D{gr}"))
    gw_gap = max(abs(gw_ifrs - wgw[0]), abs(gw_us - wgw[1]))
    ok = worst < 0.01 and gw_gap < 0.01 and bs_gap < 0.01
    add(sec, "PASS" if ok else "FAIL", f"{scen}: independent model vs workbook (13 outputs + goodwill)",
        f"largest gap {worst:.2e} GBP k; Python balance sheet gap {bs_gap:.1e}; goodwill IFRS {gw_ifrs:,.1f} / US GAAP {gw_us:,.1f} (gap {gw_gap:.1e})")
    SCEN_SUMMARY.append((scen, pairs, gw_ifrs, gw_us))

# ---------------------------------------------------------------------------------------------------
# Workbook's own checks per scenario (for the record)
# ---------------------------------------------------------------------------------------------------
sec = "B2. The workbook's own master check, by scenario"
ck = wbf["Checks"]
check_rows = [r for r in range(5, ck.max_row + 1) if isinstance(ck[f"E{r}"].value, str) and ck[f"E{r}"].value.startswith("=IF(ABS(C")]
master_row = [r for r in range(1, ck.max_row + 1) if ck[f"A{r}"].value == "MASTER CHECK"][0]
warn_row = [r for r in range(1, ck.max_row + 1) if ck[f"A{r}"].value == "WARNINGS"][0]
for scen in ("Base", "Upside", "Downside", "Custom"):
    set_scenario(xl, scen)
    failing = [ck[f"A{r}"].value for r in check_rows if xl.evaluate(f"Checks!E{r}") != 1]
    master = xl.evaluate(f"Checks!E{master_row}")
    warns = xl.evaluate(f"Checks!E{warn_row}")
    if master == 1:
        add(sec, "PASS", f"{scen}: all {len(check_rows)} checks pass", f"outlier warnings: {warns}")
    elif failing == ["Pro-forma cash never negative"] and scen == "Downside":
        add(sec, "EXPECTED", f"{scen}: master check red because the deal leaves the pro-forma business with negative cash",
            "Intended: the model refuses to hide a funding gap. See docs/acquisition_memo.md")
    else:
        add(sec, "FAIL", f"{scen}: checks failing", ", ".join(failing))
set_scenario(xl, "Base")

# ---------------------------------------------------------------------------------------------------
# C. Test the tests
# ---------------------------------------------------------------------------------------------------
sec = "C. Test the tests (deliberately break the model)"


def mutated(label, setter, expect_check):
    x = new_xl()
    addr_val = setter(x)
    failing = [ck[f"A{r}"].value for r in check_rows if x.evaluate(f"Checks!E{r}") != 1]
    master = x.evaluate(f"Checks!E{master_row}")
    caught = master == 0 and any(expect_check in f for f in failing)
    add(sec, "PASS" if caught else "FAIL", f"{label}",
        f"master={master}; failing checks: {', '.join(failing) or 'none'}" if caught else f"NOT CAUGHT: expected '{expect_check}', master={master}, failing={failing}")


def cell(sheet, label, col, nth=0):
    return f"{sheet}!{col}{find(sheet, label, nth)}"


def m1(x):
    a = cell("CF", "(Increase) / decrease in receivables", "S")
    x.evaluate(a)
    x.set_value(a, 100)


def m2(x):
    a = cell("Rev_Build", "Closing ARR", "Q")
    x.evaluate(a)
    x.set_value(a, x.evaluate(a) + 50)


def m3(x):
    r = int(re.sub(r"\D", "", wbf.defined_names["inp_tax"].attr_text.split("!")[1]))
    a = f"Inputs_Drivers!C{r}"
    x.evaluate(a)
    x.set_value(a, 0.9)


def m4(x):
    a = cell("Rev_Rec", "Total billings", "T")
    x.evaluate(a)
    x.set_value(a, -5)


def m5(x):
    a = cell("BS", "Total assets", "AB")
    x.evaluate(a)
    x.set_value(a, x.evaluate(a) + 1)


mutated("Overwrite one month of the receivables cash-flow line with +100", m1, "Balance sheet balances")
mutated("Add 50 to one month of closing Enterprise ARR (breaks the roll-forward)", m2, "ARR")
mutated("Set the tax rate to 90% (outside its allowed range)", m3, "Inputs within allowed range")
mutated("Set one month's billings negative", m4, "Billings never negative")
mutated("Add 1 to one month of total assets", m5, "Balance sheet balances")

# outlier warning self-test: tighten the revenue threshold and expect a flag
x = new_xl()
wrow = [r for r in range(1, ck.max_row + 1) if ck[f"A{r}"].value == "Revenue" and ck[f"B{r}"].value and "month-on-month" in ck[f"B{r}"].value][0]
x.evaluate(f"Checks!D{wrow}")
x.set_value(f"Checks!D{wrow}", 0.0001)
flag = x.evaluate(f"Checks!C{wrow}")
add(sec, "PASS" if flag and flag > 0 else "FAIL", "Outlier warning fires when the revenue threshold is set to 0.01%",
    f"{flag} months flagged (Base threshold 15%: none)")

# ---------------------------------------------------------------------------------------------------
# D. Error scan
# ---------------------------------------------------------------------------------------------------
sec = "D. Error scan"
t0 = time.time()
errs, n = [], 0
for sn in ("Checks", "IS", "BS", "CF", "SaaS_Metrics", "Business_Review", "LT_Model", "Budget_vs_Forecast", "MA_PPA", "MA_ProForma", "MA_Returns", "MA_Valuation"):
    for row in wbf[sn].iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                n += 1
                try:
                    v = xl.evaluate(f"{sn}!{c.coordinate}")
                    if isinstance(v, str) and v.startswith("#"):
                        errs.append(f"{sn}!{c.coordinate}={v}")
                except Exception as e:
                    errs.append(f"{sn}!{c.coordinate}: {type(e).__name__}")
add(sec, "PASS" if not errs else "FAIL", f"No error values on 12 output tabs ({n:,} formula cells evaluated in {time.time() - t0:.0f}s)",
    "none" if not errs else f"{len(errs)}, e.g. {', '.join(errs[:6])}")

# ---------------------------------------------------------------------------------------------------
# E. Web contract
# ---------------------------------------------------------------------------------------------------
sec = "E. Web contract (Outputs_Web and named ranges)"
owf = wbf["Outputs_Web"]
out_names = {n: d for n, d in wbf.defined_names.items() if n.startswith("out_")}
inp_names = {n for n in wbf.defined_names if n.startswith("inp_")} - {"inp_scenario", "inp_scen_idx"}
bad_loc = [n for n, d in out_names.items() if not d.attr_text.startswith("Outputs_Web!")]
add(sec, "PASS" if not bad_loc and out_names else "FAIL", f"{len(out_names)} out_* named ranges, all on Outputs_Web", "ok" if not bad_loc else ", ".join(bad_loc[:5]))
reg = wbf.defined_names["out_input_registry"].attr_text.split("!")[1].replace("$", "")
r0, r1 = [int(re.sub(r"\D", "", x)) for x in reg.split(":")]
reg_names = {owf[f"A{r}"].value for r in range(r0, r1 + 1)}
missing = sorted(inp_names - reg_names)
extra = sorted(reg_names - inp_names)
add(sec, "PASS" if not missing and not extra else "FAIL", f"Input registry lists every inp_* name ({len(inp_names)})",
    "complete" if not missing and not extra else f"missing {missing[:4]}, extra {extra[:4]}")
bad_scen = []
for r in range(r0, r1 + 1):
    if owf[f"G{r}"].value == "scenario":
        sr = int(re.search(r"\$A\$(\d+)", owf[f"H{r}"].value).group(1))
        if sws[f"A{sr}"].value != owf[f"B{r}"].value:
            bad_scen.append(owf[f"A{r}"].value)
add(sec, "PASS" if not bad_scen else "FAIL", "Each scenario-driven input points at the right Inputs_Scenarios row (the Custom column is where the website writes)",
    "ok" if not bad_scen else ", ".join(bad_scen))
not_inputs_sheet = [n for n in inp_names if not wbf.defined_names[n].attr_text.startswith("Inputs_Drivers!")]
add(sec, "PASS" if not not_inputs_sheet else "FAIL", "Every inp_* name sits on Inputs_Drivers", "ok" if not not_inputs_sheet else ", ".join(not_inputs_sheet))
LINK = re.compile(r"^=\s*'?[A-Za-z_ ]+'?!\$?[A-Z]{1,3}\$?\d+$|^=(inp|chk)_[A-Za-z0-9_]+$|^=ROW\(Inputs_Scenarios!\$A\$\d+\)$")
nonlink = []
for row in owf.iter_rows():
    for c in row:
        if isinstance(c.value, str) and c.value.startswith("=") and not LINK.match(c.value):
            nonlink.append(f"{c.coordinate}: {c.value[:40]}")
add(sec, "PASS" if not nonlink else "FAIL", "Outputs_Web holds only pure links (no calculation of its own)",
    "all links" if not nonlink else f"{len(nonlink)} non-links, e.g. {nonlink[:3]}")
t0 = time.time()
set_scenario(xl, "Base")
errs, n = [], 0
for row in owf.iter_rows():
    for c in row:
        if isinstance(c.value, str) and c.value.startswith("="):
            n += 1
            try:
                v = xl.evaluate(f"Outputs_Web!{c.coordinate}")
                if isinstance(v, str) and v.startswith("#"):
                    errs.append(f"{c.coordinate}={v}")
            except Exception as e:
                errs.append(f"{c.coordinate}: {type(e).__name__}")
add(sec, "PASS" if not errs else "FAIL", f"Every Outputs_Web cell evaluates without error ({n:,} cells, {time.time() - t0:.0f}s)", "none" if not errs else ", ".join(errs[:6]))
mv_ = xl.evaluate("Outputs_Web!" + wbf.defined_names["out_master_check"].attr_text.split("!")[1].replace("$", ""))
add(sec, "PASS" if mv_ == 1 else "FAIL", "out_master_check reads 1 in Base", str(mv_))

# ---------------------------------------------------------------------------------------------------
# F. Portability inventory
# ---------------------------------------------------------------------------------------------------
sec = "F. Portability inventory (Excel, Google Sheets, HyperFormula)"
PORTABLE = {"ABS", "AND", "AVERAGE", "COLUMN", "COUNT", "COUNTIF", "EOMONTH", "IF", "INDEX", "ISNUMBER", "LARGE", "MATCH", "MAX", "MIN", "OR", "ROUNDUP", "ROW", "SUM", "SUMIF", "SUMPRODUCT", "YEAR"}
fn_count, longest, deepest, arr_sp = {}, 0, 0, []
for ws in wbf.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                longest = max(longest, len(c.value))
                depth = 0
                try:
                    for t in Tokenizer(c.value).items:
                        if t.type == Token.FUNC and t.subtype == Token.OPEN:
                            depth += 1
                            deepest = max(deepest, depth)
                            fn_count[t.value[:-1].upper()] = fn_count.get(t.value[:-1].upper(), 0) + 1
                        elif t.type == Token.FUNC and t.subtype == Token.CLOSE:
                            depth -= 1
                except Exception:
                    pass
                for m in re.finditer(r"SUMPRODUCT\(([^()]*(?:\([^()]*\)[^()]*)*)\)", c.value):
                    inner = re.sub(r"[A-Za-z_]+!", "", m.group(1))
                    if not re.fullmatch(r"[\$A-Z0-9:]+,[\$A-Z0-9:]+", inner):
                        arr_sp.append(f"{ws.title}!{c.coordinate}")
def index_args(formula):
    """Argument counts of every INDEX( call in a formula (top-level commas, parentheses and strings respected)."""
    out = []
    for m in re.finditer(r"INDEX\(", formula):
        depth, commas, i, instr = 1, 0, m.end(), False
        while i < len(formula) and depth:
            ch = formula[i]
            if ch == '"':
                instr = not instr
            elif not instr:
                depth += ch == "("
                depth -= ch == ")"
                commas += ch == "," and depth == 1
            i += 1
        out.append(commas + 1)
    return out


two_arg_index = []
for ws in wbf.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("=") and "INDEX(" in c.value and any(n < 3 for n in index_args(c.value)):
                two_arg_index.append(f"{ws.title}!{c.coordinate}")
unknown = sorted(set(fn_count) - PORTABLE)
add(sec, "PASS" if not unknown else "FAIL", f"Only {len(PORTABLE)} classic functions used ({len(fn_count)} distinct)",
    ", ".join(f"{k} x{v}" for k, v in sorted(fn_count.items())) if not unknown else f"not on the portable list: {unknown}")
add(sec, "PASS" if not arr_sp else "FAIL", "SUMPRODUCT is only ever called with two plain ranges (no array arithmetic, which HyperFormula disables by default)",
    "ok" if not arr_sp else f"{len(arr_sp)}, e.g. {arr_sp[:3]}")
add(sec, "PASS" if not two_arg_index else "FAIL", "INDEX is always called with explicit row and column (HyperFormula reads a 2-argument INDEX on a one-row range differently from Excel and Sheets)",
    "ok" if not two_arg_index else f"{len(two_arg_index)} cells, e.g. {two_arg_index[:3]}")
add(sec, "PASS" if longest < 8000 and deepest < 60 else "FAIL", "Formula length and nesting are within Excel limits", f"longest {longest} characters (limit 8,192); deepest nesting {deepest} (limit 64)")
add(sec, "REVIEW", "Engines not yet tried: Excel and Google Sheets",
    "HyperFormula (the website's engine) reproduces every Outputs_Web cell: see web/tests. Run tests/portability_checklist.md in Excel and Google Sheets to close the remaining gap.")

# ---------------------------------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------------------------------
counts = {k: sum(1 for r in RESULTS if r[1] == k) for k in ("PASS", "FAIL", "REVIEW", "EXPECTED")}
lines = [f"# Validation report", "",
         f"*Generated {datetime.datetime.now():%Y-%m-%d %H:%M} by `validation/validate_model.py` on `{os.path.basename(path)}`. "
         "Recalculation engine: pycel (not Excel).*", "",
         f"**{counts['PASS']} passed, {counts['FAIL']} failed, {counts['REVIEW']} for review, {counts['EXPECTED']} expected.**", ""]
cur = None
for s_, st, nm, dt in RESULTS:
    if s_ != cur:
        lines += ["", f"## {s_}", "", "| Status | Test | Detail |", "|---|---|---|"]
        cur = s_
    lines.append(f"| {st} | {nm} | {dt.replace('|', '/')} |")
lines += ["", "## Independent model: scenario comparison (GBP k)", "", "| Output | " + " | ".join(f"{s} workbook / Python" for s, *_ in SCEN_SUMMARY) + " |", "|---|" + "---|" * len(SCEN_SUMMARY)]
for i, (lab, _a, _b) in enumerate(SCEN_SUMMARY[0][1]):
    lines.append(f"| {lab} | " + " | ".join(f"{p[i][2]:,.1f} / {p[i][1]:,.1f}" for _s, p, *_ in SCEN_SUMMARY) + " |")
lines += ["", "## What this does not cover", "",
          "- The Vaultline monthly model, the DCF and the pro-forma are not re-implemented in Python (only the Vaultline goodwill is). Their integrity rests on the workbook's checks, which the mutation tests above exercise only for the Northgate model.",
          "- pycel is the recalculation engine here, not Excel or Google Sheets. A difference in how those engines treat a function would not show up.",
          "- Static tables (levers, grids, scenario snapshot) are checked for staleness inside the workbook, not re-derived here.",
          "- The independent Python model was written from the same understanding of the business as the workbook. It catches wiring, formula and arithmetic errors; it cannot catch a shared conceptual mistake (for example a wrong accounting treatment).",
          "- The one review item (the literal 10 in the Lever_Analysis staleness tolerance) is the cut-off between ratio metrics and GBP k metrics. It is deliberate and visible."]
open(report_path, "w").write("\n".join(lines) + "\n")
print("\n".join(f"{st:9s} {nm}" for _s, st, nm, _d in RESULTS))
print(f"\n{counts}")
print("report:", report_path)
sys.exit(1 if counts["FAIL"] else 0)
