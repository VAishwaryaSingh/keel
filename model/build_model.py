"""Builds keel_saas_model.xlsx (Phase 1 MVP) with openpyxl.

Every calculation is written as an Excel formula; Python only lays the formulas out.
Units: GBP thousands unless stated. Time axis (identical on every calc tab):
  A label | B unit | C:M history months -11..-1 (grouped) | N = opening (t=0)
  O:AL = months 1..24 | AN:AO = FY1, FY2
"""
import re
import sys
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "keel_saas_model.xlsx"
OPEN, F0, F1, AN1, AN2 = 14, 15, 38, 40, 41   # N, O, AL, AN, AO

BLUE, BLACK, GREEN = "0000FF", "000000", "007A33"
FONT = "Arial"
FMT = {
    "num": '#,##0;(#,##0);"-"',
    "num1": '#,##0.0;(#,##0.0);"-"',
    "num2": '#,##0.00;(#,##0.00);"-"',
    "pct": '0.0%;(0.0%);"-"',
    "x": '0.0"x"',
    "x2": '0.000',
    "mo": '0.0',
    "date": 'mmm-yy',
    "int": '0',
}
HDR_FILL = PatternFill("solid", fgColor="1F3A5F")
SEC_FILL = PatternFill("solid", fgColor="DCE6F1")
INP_FILL = PatternFill("solid", fgColor="FFF9E0")
THIN = Side(style="thin", color="808080")

wb = Workbook()
wb.remove(wb.active)
SH = {}
PURE_LINK = re.compile(r"^='?[A-Za-z_ ]+'?!\$?[A-Z]{1,3}\$?\d+$")


def put(ws, ref, value, fmt=None, bold=False, color=None, fill=None, italic=False, align=None):
    cell = ws[ref]
    cell.value = value
    if color is None:
        if isinstance(value, str) and value.startswith("="):
            color = GREEN if PURE_LINK.match(value) else BLACK
        else:
            color = BLACK
    cell.font = Font(name=FONT, size=10, bold=bold, color=color, italic=italic)
    if fmt:
        cell.number_format = FMT.get(fmt, fmt)
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = Alignment(horizontal=align)
    return cell


class Sheet:
    """Row-based time-series sheet. Formulas are lambdas (c, p, ci) -> str for one period column."""

    def __init__(self, name, title, hide_hist=True):
        self.name, self.title, self.hide_hist = name, title, hide_hist
        self.ws = wb.create_sheet(name)
        self.reg, self.items, self.next = {}, [], 8
        SH[name] = self

    def sec(self, label):
        self.items.append(("sec", self.next, label))
        self.next += 1

    def gap(self):
        self.next += 1

    def add(self, key, label, unit, fmt, f=None, opening=None, hist=None, ann="sum", bold=False, top=False):
        row = self.next
        self.next += 1
        self.reg[key] = row
        self.items.append(("row", row, dict(label=label, unit=unit, fmt=fmt, f=f, opening=opening,
                                            hist=hist, ann=ann, bold=bold, top=top)))
        return row

    def write(self):
        ws = self.ws
        ws.sheet_view.showGridLines = False
        put(ws, "A1", self.title, bold=True)
        ws["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3A5F")
        put(ws, "A2", "GBP thousands unless stated. Blue = input, black = formula, green = link to another tab.",
            italic=True)
        for r, lab in ((3, "Period #"), (4, "Period ending"), (5, "Year"), (6, "Quarter")):
            put(ws, f"A{r}", lab, bold=True, color="FFFFFF", fill=HDR_FILL)
            put(ws, f"B{r}", None, fill=HDR_FILL)
        for ci in range(3, F1 + 1):
            c = L(ci)
            put(ws, f"{c}3", "=COLUMN()-COLUMN($N$1)", "int", True, "FFFFFF", HDR_FILL, align="right")
            put(ws, f"{c}4", f"=EOMONTH(inp_model_start,{c}3-1)", "date", True, "FFFFFF", HDR_FILL, align="right")
            put(ws, f"{c}5", f"=IF({c}3<1,0,ROUNDUP({c}3/12,0))", "int", True, "FFFFFF", HDR_FILL, align="right")
            put(ws, f"{c}6", f"=IF({c}3<1,0,ROUNDUP({c}3/3,0))", "int", True, "FFFFFF", HDR_FILL, align="right")
        for ci in (AN1, AN2):
            c = L(ci)
            put(ws, f"{c}3", "Annual", None, True, "FFFFFF", HDR_FILL, align="right")
            put(ws, f"{c}4", f'="FY"&(YEAR(inp_model_start)+{c}5-1)', None, True, "FFFFFF", HDR_FILL, align="right")
            put(ws, f"{c}5", "=COLUMN()-COLUMN($AM$1)", "int", True, "FFFFFF", HDR_FILL, align="right")
            put(ws, f"{c}6", None, None, True, "FFFFFF", HDR_FILL)
        for kind, row, d in self.items:
            if kind == "sec":
                for ci in range(1, AN2 + 1):
                    ws.cell(row=row, column=ci).fill = SEC_FILL
                put(ws, f"A{row}", d, bold=True, fill=SEC_FILL)
                continue
            put(ws, f"A{row}", d["label"], bold=d["bold"])
            put(ws, f"B{row}", d["unit"], italic=True)
            fmt = d["fmt"]
            if d["hist"]:
                for ci in range(3, OPEN + 1):
                    put(ws, f"{L(ci)}{row}", d["hist"](L(ci), ci), fmt, d["bold"])
            if d["opening"] is not None:
                o = d["opening"]
                put(ws, f"{L(OPEN)}{row}", o(L(OPEN), L(OPEN - 1), OPEN) if callable(o) else o, fmt, d["bold"])
            if d["f"]:
                for ci in range(F0, F1 + 1):
                    put(ws, f"{L(ci)}{row}", d["f"](L(ci), L(ci - 1), ci), fmt, d["bold"])
            if d["ann"] and d["f"]:
                for ci in (AN1, AN2):
                    c = L(ci)
                    a = d["ann"]
                    if a == "sum":
                        fx = f"=SUMIF($O$5:$AL$5,{c}$5,$O{row}:$AL{row})"
                    elif a == "end":
                        fx = f"=INDEX($O{row}:$AL{row},1,{c}$5*12)"
                    elif a == "open":
                        fx = f"=INDEX($O{row}:$AL{row},1,({c}$5-1)*12+1)"   # opening row: value in the first month of the year
                    else:
                        fx = a(c)
                    put(ws, f"{c}{row}", fx, fmt, d["bold"])
            if d["top"]:
                for ci in range(1, AN2 + 1):
                    ws.cell(row=row, column=ci).border = Border(top=THIN)
        ws.column_dimensions["A"].width = 44
        ws.column_dimensions["B"].width = 9
        for ci in range(3, AN2 + 1):
            ws.column_dimensions[L(ci)].width = 10.5
        ws.column_dimensions["AM"].width = 3
        if self.hide_hist:
            ws.column_dimensions.group("C", "M", hidden=True)
        ws.freeze_panes = "C7"


def row_of(sh, key):
    return SH[sh].reg[key]


def X(sh, key, c):
    return f"{sh}!{c}{row_of(sh, key)}"


def cl(sh):
    """Local reference helper: cl('IS')('rev_tot', 'O') -> 'O12'."""
    return lambda key, c: f"{c}{row_of(sh, key)}"


# ---------------------------------------------------------------------------
# Inputs
# ---------------------------------------------------------------------------
# kind, name, label, unit, fmt, base, upside, downside, min, max, tip
S = "scen"
T = "static"
INPUTS = [
    ("sec", "Growth"),
    (S, "inp_logos_ent", "New Enterprise logos per month", "#/mo", "num1", 3.0, 3.6, 2.2, 0, 15, "New-logo ARR = logos x ACV. Biggest driver of ARR growth."),
    (S, "inp_logos_mid", "New Mid-market logos per month", "#/mo", "num1", 14, 16, 10, 0, 60, "Volume engine for the Mid-market segment."),
    (S, "inp_logo_growth", "Annual growth in new-logo volume", "% p.a.", "pct", 0.05, 0.10, 0.0, -0.2, 0.5, "Compounds monthly. Sales capacity proxy in the MVP."),
    ("sec", "Pricing"),
    (S, "inp_acv_ent", "Enterprise ACV (new logos)", "GBP k", "num1", 210, 220, 195, 50, 1000, "Average annual contract value of a new Enterprise logo."),
    (S, "inp_acv_mid", "Mid-market ACV (new logos)", "GBP k", "num1", 26, 27, 23, 5, 100, "Average annual contract value of a new Mid-market logo."),
    (S, "inp_price_uplift", "Annual price uplift on new-logo ACV", "% p.a.", "pct", 0.03, 0.04, 0.01, 0, 0.15, "Applied from month 13, 25, ... to new-logo ACV."),
    ("sec", "Retention"),
    (S, "inp_exp_ent", "Enterprise annual expansion", "% of ARR", "pct", 0.14, 0.17, 0.09, 0, 0.5, "Upsell / cross-sell on the existing Enterprise base."),
    (S, "inp_exp_mid", "Mid-market annual expansion", "% of ARR", "pct", 0.07, 0.09, 0.04, 0, 0.5, "Upsell on the existing Mid-market base."),
    (S, "inp_churn_ent", "Enterprise annual gross churn", "% of ARR", "pct", 0.06, 0.05, 0.10, 0, 0.5, "Lost ARR (incl. contraction) from existing Enterprise customers."),
    (S, "inp_churn_mid", "Mid-market annual gross churn", "% of ARR", "pct", 0.12, 0.10, 0.18, 0, 0.5, "Lost ARR (incl. contraction) from existing Mid-market customers."),
    ("sec", "People (hiring plan)"),
    (S, "inp_hires_rd", "R&D net hires per month", "FTE/mo", "num1", 2.5, 3.0, 1.0, 0, 10, "Drives R&D payroll."),
    (S, "inp_hires_sm", "Sales & Marketing net hires per month", "FTE/mo", "num1", 3.0, 4.0, 1.0, 0, 10, "Drives S&M payroll and CAC."),
    (S, "inp_hires_ga", "G&A net hires per month", "FTE/mo", "num1", 0.8, 1.0, 0.3, 0, 5, "Drives G&A payroll."),
    (S, "inp_hires_cs", "Customer Success net hires per month", "FTE/mo", "num1", 1.5, 2.0, 0.8, 0, 5, "Drives support payroll in COGS."),
    ("sec", "Cash"),
    (S, "inp_dso", "Customer payment terms (DSO)", "days", "num", 45, 40, 60, 0, 90, "Days to collect an invoice. Drives receivables and cash timing."),
    ("sec", "Timing and opening position (31 Dec before model start)"),
    (T, "inp_model_start", "Model start date (first forecast month)", "date", "date", dt.date(2027, 1, 1), None, None, None, None, "Month 1 of the 24-month forecast."),
    (T, "inp_arr_open_ent", "Opening ARR - Enterprise", "GBP k", "num", 26000, None, None, 0, 200000, "Contracted ARR at the opening date."),
    (T, "inp_arr_open_mid", "Opening ARR - Mid-market", "GBP k", "num", 14000, None, None, 0, 200000, "Contracted ARR at the opening date."),
    (T, "inp_cust_open_ent", "Opening customers - Enterprise", "#", "num", 130, None, None, 0, 5000, "Logo count at the opening date."),
    (T, "inp_cust_open_mid", "Opening customers - Mid-market", "#", "num", 560, None, None, 0, 20000, "Logo count at the opening date."),
    (T, "inp_open_cash", "Opening cash", "GBP k", "num", 24000, None, None, 0, 500000, "Cash at the opening date."),
    (T, "inp_open_ar", "Opening trade receivables", "GBP k", "num", 5500, None, None, 0, 100000, "About 1.5 months of billings."),
    (T, "inp_open_ppe", "Opening PP&E (net)", "GBP k", "num", 2500, None, None, 0, 100000, "Net book value at the opening date."),
    (T, "inp_open_ap", "Opening trade payables", "GBP k", "num", 1100, None, None, 0, 100000, "About one month of non-payroll spend."),
    (T, "inp_open_defrev_impl", "Opening deferred implementation revenue", "GBP k", "num", 0, None, None, 0, 10000, "Subscription deferred revenue is derived from the billing history."),
    ("sec", "People (static)"),
    (T, "inp_heads_open_rd", "Opening headcount - R&D", "FTE", "num", 150, None, None, 0, 2000, ""),
    (T, "inp_heads_open_sm", "Opening headcount - Sales & Marketing", "FTE", "num", 130, None, None, 0, 2000, ""),
    (T, "inp_heads_open_ga", "Opening headcount - G&A", "FTE", "num", 50, None, None, 0, 2000, ""),
    (T, "inp_heads_open_cs", "Opening headcount - Customer Success", "FTE", "num", 70, None, None, 0, 2000, ""),
    (T, "inp_sal_rd", "Fully loaded cost per FTE - R&D", "GBP k p.a.", "num1", 90, None, None, 30, 300, "Salary + employer NI + pension + benefits."),
    (T, "inp_sal_sm", "Fully loaded cost per FTE - S&M", "GBP k p.a.", "num1", 88, None, None, 30, 300, "Excludes commission (modelled separately)."),
    (T, "inp_sal_ga", "Fully loaded cost per FTE - G&A", "GBP k p.a.", "num1", 75, None, None, 30, 300, ""),
    (T, "inp_sal_cs", "Fully loaded cost per FTE - Customer Success", "GBP k p.a.", "num1", 65, None, None, 30, 300, ""),
    (T, "inp_merit", "Annual pay rise (applied from month 13)", "% p.a.", "pct", 0.035, None, None, 0, 0.15, ""),
    ("sec", "Costs"),
    (T, "inp_marketing_m", "Marketing programmes", "GBP k/mo", "num", 480, None, None, 0, 5000, "Non-headcount S&M."),
    (T, "inp_rd_other_m", "R&D non-headcount (tools, contractors)", "GBP k/mo", "num", 120, None, None, 0, 5000, ""),
    (T, "inp_ga_other_m", "G&A non-headcount (rent, fees, insurance)", "GBP k/mo", "num", 320, None, None, 0, 5000, ""),
    (T, "inp_commission_pct", "Sales commission", "% of gross new ARR", "pct", 0.10, None, None, 0, 0.3, "Expensed as incurred in the MVP. Phase 2 capitalises it (IFRS 15 costs to obtain a contract)."),
    (T, "inp_capitalise_commission", "Capitalise sales commission (1 = yes, 0 = expense)", "1/0", "int", 1, None, None, 0, 1, "IFRS 15 incremental costs of obtaining a contract: capitalised and amortised instead of expensed."),
    (T, "inp_cc_life", "Contract-cost amortisation (declining balance on opening asset)", "months", "num", 48, None, None, 6, 120, "Amortisation = opening asset / this life. Approximates amortising over the expected customer life."),
    (T, "inp_open_cca", "Opening capitalised contract costs", "GBP k", "num", 3000, None, None, 0, 50000, "Asset balance at the opening date."),
    (T, "inp_host_ent", "Hosting cost per Enterprise customer", "GBP k/mo", "num2", 1.4, None, None, 0, 10, ""),
    (T, "inp_host_mid", "Hosting cost per Mid-market customer", "GBP k/mo", "num2", 0.3, None, None, 0, 5, ""),
    (T, "inp_licence_pct", "Third-party licences", "% of subscription revenue", "pct", 0.02, None, None, 0, 0.2, ""),
    (T, "inp_impl_fee_pct", "Implementation fee", "% of new Enterprise ACV", "pct", 0.15, None, None, 0, 0.5, "Separate performance obligation, billed at signing, recognised over 3 months."),
    (T, "inp_services_cost_pct", "Cost of delivering implementation", "% of implementation revenue", "pct", 0.70, None, None, 0, 1.5, ""),
    ("sec", "Acquisition: Vaultline scenario drivers"),
    (S, "inp_v_logos", "Vaultline new logos per month", "#/mo", "num1", 1.6, 2.0, 1.0, 0, 10, "Vaultline's own growth engine, standalone."),
    (S, "inp_v_exp", "Vaultline annual expansion", "% of ARR", "pct", 0.08, 0.10, 0.05, 0, 0.5, ""),
    (S, "inp_v_churn", "Vaultline annual gross churn", "% of ARR", "pct", 0.06, 0.05, 0.12, 0, 0.5, "Integration risk: customers leave after a change of control."),
    (S, "inp_ma_syn_achieved", "Synergies achieved", "% of plan", "pct", 1.0, 1.2, 0.4, 0, 2, "Scales both revenue (cross-sell) and cost synergies. Integration costs are not scaled."),
    ("sec", "Acquisition: deal terms"),
    (T, "inp_ma_include", "Include acquisition in the pro-forma (1 = yes, 0 = no)", "1/0", "int", 1, None, None, 0, 1, "0 makes MA_ProForma equal Northgate standalone."),
    (T, "inp_ma_offer_mult", "Offer: enterprise value / Vaultline opening ARR", "x ARR", "num1", 4.5, None, None, 1, 12, "Sets the offer price (EV). Equity price = EV + cash acquired (Vaultline has no debt)."),
    (T, "inp_ma_debt", "Acquisition term loan", "GBP k", "num", 15000, None, None, 0, 100000, "Bullet loan; balance of the equity price is paid from Northgate cash."),
    (T, "inp_ma_loan_rate", "Acquisition loan interest rate", "% p.a.", "pct", 0.08, None, None, 0, 0.25, "Charged on the prior-period balance."),
    (T, "inp_ma_deal_cost_pct", "Transaction costs", "% of EV", "pct", 0.02, None, None, 0, 0.1, "Expensed in month 1 (IFRS 3.53)."),
    (T, "inp_ma_integration_total", "Integration costs, total", "GBP k", "num", 3000, None, None, 0, 20000, "One-off, spread evenly."),
    (T, "inp_ma_integration_months", "Integration period", "months", "num", 12, None, None, 1, 24, "Costs spread evenly over this many months from month 1."),
    (T, "inp_ma_xsell_logos_m", "Cross-sell: Vaultline product sold to Northgate customers per month (at full ramp)", "#/mo", "num1", 0.8, None, None, 0, 10, "Revenue synergy volume."),
    (T, "inp_ma_xsell_acv", "Cross-sell ACV", "GBP k", "num1", 45, None, None, 0, 500, ""),
    (T, "inp_ma_syn_start", "Synergies start in month", "month", "num", 7, None, None, 1, 24, ""),
    (T, "inp_ma_syn_ramp", "Synergy ramp-up", "months", "num", 6, None, None, 1, 24, "Linear ramp from start to 100%."),
    (T, "inp_ma_syn_margin", "Contribution margin on cross-sell revenue", "%", "pct", 0.65, None, None, 0, 1, "Revenue synergies are not free: they carry delivery and selling cost."),
    (T, "inp_ma_cost_syn", "Cost synergies, annual run-rate", "GBP k p.a.", "num", 1200, None, None, 0, 20000, "Duplicate G&A and hosting."),
    (T, "inp_ma_ic_annual", "Intercompany contract: Northgate software sold to Vaultline", "GBP k p.a.", "num", 240, None, None, 0, 5000, "Eliminated on consolidation (revenue and cost)."),
    ("sec", "Acquisition: purchase price allocation (IFRS 3)"),
    (T, "inp_ma_dr_haircut", "Deferred revenue fair-value haircut", "% of book", "pct", 0.30, None, None, 0, 0.8, "IFRS 3 measures acquired deferred revenue at fair value; US GAAP (ASU 2021-08) carries it over at book."),
    (T, "inp_ma_cr_pct", "Customer relationships", "% of EV", "pct", 0.30, None, None, 0, 0.7, "Illustrative; in practice from a valuation specialist (multi-period excess earnings)."),
    (T, "inp_ma_tech_pct", "Acquired technology", "% of EV", "pct", 0.12, None, None, 0, 0.5, "Illustrative (relief from royalty)."),
    (T, "inp_ma_cr_life", "Customer relationships: useful life", "months", "num", 96, None, None, 24, 240, ""),
    (T, "inp_ma_tech_life", "Technology: useful life", "months", "num", 60, None, None, 24, 240, ""),
    ("sec", "Acquisition: valuation"),
    (T, "inp_ma_wacc", "Discount rate (WACC)", "%", "pct", 0.115, None, None, 0.05, 0.25, "Applied to Vaultline cash flows and synergies."),
    (T, "inp_ma_tg", "Terminal growth", "%", "pct", 0.03, None, None, 0, 0.06, "Must be below WACC."),
    (T, "inp_ma_wacc_band", "DCF sensitivity band on WACC (low / high cases)", "pts", "pct", 0.01, None, None, 0, 0.05, "Low case = WACC + band; high case = WACC - band."),
    (T, "inp_ma_tg_band", "DCF sensitivity band on terminal growth", "pts", "pct", 0.005, None, None, 0, 0.02, "Low case = growth - band; high case = growth + band."),
    (T, "inp_ma_fade", "Vaultline ARR growth fade, years 3-5", "pts p.a.", "pct", 0.03, None, None, 0, 0.2, "Growth falls by this much each year after year 2."),
    ("sec", "Acquisition: Vaultline standalone (opening position and cost structure)"),
    (T, "inp_v_arr_open", "Vaultline opening ARR", "GBP k", "num", 6000, None, None, 0, 100000, ""),
    (T, "inp_v_cust_open", "Vaultline opening customers", "#", "num", 150, None, None, 0, 5000, ""),
    (T, "inp_v_acv", "Vaultline ACV of new customers", "GBP k", "num1", 42, None, None, 1, 500, ""),
    (T, "inp_v_cogs_pct", "Vaultline cost of revenue", "% of revenue", "pct", 0.20, None, None, 0, 1, "Includes the intercompany software cost."),
    (T, "inp_v_rd_pct", "Vaultline R&D", "% of revenue", "pct", 0.17, None, None, 0, 1, ""),
    (T, "inp_v_sm_pct", "Vaultline sales & marketing", "% of revenue", "pct", 0.26, None, None, 0, 1, ""),
    (T, "inp_v_ga_pct", "Vaultline G&A", "% of revenue", "pct", 0.12, None, None, 0, 1, ""),
    (T, "inp_v_dso", "Vaultline payment terms (DSO)", "days", "num", 40, None, None, 0, 90, ""),
    (T, "inp_v_ap_months", "Vaultline payables", "months of costs", "num1", 0.4, None, None, 0, 1, ""),
    (T, "inp_v_capex_pct", "Vaultline capex", "% of revenue", "pct", 0.01, None, None, 0, 0.2, ""),
    (T, "inp_v_dep_life", "Vaultline depreciation life", "months", "num", 36, None, None, 6, 120, ""),
    (T, "inp_v_open_cash", "Vaultline opening cash", "GBP k", "num", 3500, None, None, 0, 100000, ""),
    (T, "inp_v_open_ar", "Vaultline opening receivables", "GBP k", "num", 1000, None, None, 0, 50000, ""),
    (T, "inp_v_open_ppe", "Vaultline opening PP&E", "GBP k", "num", 400, None, None, 0, 50000, ""),
    ("sec", "Long-term extension (years 3-5)"),
    (T, "inp_lt_newarr_growth", "Annual growth in new-logo ARR", "% p.a.", "pct", 0.12, None, None, -0.2, 0.6, "New-logo ARR in year N = year N-1 x (1 + this)."),
    (T, "inp_lt_nrr_step", "Annual change in expansion rate", "pts p.a.", "pct", 0.005, None, None, -0.05, 0.05, "Added to the prior year's expansion rate. Churn rate is held at the prior year."),
    (T, "inp_lt_gm_step", "Annual change in gross margin", "pts p.a.", "pct", 0.005, None, None, -0.05, 0.05, ""),
    (T, "inp_lt_rd_step", "Annual reduction in R&D as % of revenue", "pts p.a.", "pct", 0.015, None, None, -0.05, 0.1, "Operating leverage."),
    (T, "inp_lt_sm_step", "Annual reduction in S&M as % of revenue", "pts p.a.", "pct", 0.020, None, None, -0.05, 0.1, "Operating leverage."),
    (T, "inp_lt_ga_step", "Annual reduction in G&A as % of revenue", "pts p.a.", "pct", 0.0075, None, None, -0.05, 0.1, "Operating leverage."),
    ("sec", "Accounting and tax"),
    (T, "inp_capex_pct", "Capex", "% of revenue", "pct", 0.015, None, None, 0, 0.2, ""),
    (T, "inp_dep_life", "Depreciation life (declining balance on opening NBV)", "months", "num", 48, None, None, 6, 120, "Depreciation = opening NBV / life."),
    (T, "inp_interest", "Interest rate earned on cash", "% p.a.", "pct", 0.035, None, None, 0, 0.15, "Applied to PRIOR-period cash: avoids circularity."),
    (T, "inp_tax", "Corporation tax rate", "%", "pct", 0.25, None, None, 0, 0.5, "Charged on positive profit before tax only. No loss carry-forward in the MVP."),
    (T, "inp_dpo", "Supplier payment terms (DPO)", "days", "num", 30, None, None, 0, 30, "Payables = non-payroll spend x DPO / 30 (valid 0-30)."),
]

SCEN_ROW, DRV_CELL, DRV_LAST = {}, {}, [0]
ws_scen = wb.create_sheet("Inputs_Scenarios")
ws_drv = wb.create_sheet("Inputs_Drivers")


def build_inputs():
    # ----- Inputs_Scenarios
    w = ws_scen
    w.sheet_view.showGridLines = False
    put(w, "A1", "Inputs_Scenarios - scenario values for every scenario-driven input", bold=True)
    w["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3A5F")
    put(w, "A2", "Blue cells are inputs. The live column on Inputs_Drivers picks one of these columns via the selector.", italic=True)
    for ci, h in enumerate(["Driver", "Unit", "Base", "Upside", "Downside", "Custom"], start=1):
        put(w, f"{L(ci)}4", h, bold=True, color="FFFFFF", fill=HDR_FILL, align="left" if ci < 3 else "right")
    r = 5
    scen_row = {}
    for item in INPUTS:
        if item[0] == "sec":
            if INPUTS[INPUTS.index(item) + 1][0] == S:
                for ci in range(1, 7):
                    w.cell(row=r, column=ci).fill = SEC_FILL
                put(w, f"A{r}", item[1], bold=True, fill=SEC_FILL)
                r += 1
            continue
        kind, name, label, unit, fmt, base, up, down = item[:8]
        if kind != S:
            continue
        scen_row[name] = r
        put(w, f"A{r}", label)
        put(w, f"B{r}", unit, italic=True)
        for ci, v in zip((3, 4, 5, 6), (base, up, down, base)):   # Custom starts as a copy of Base
            put(w, f"{L(ci)}{r}", v, fmt, color=BLUE, fill=INP_FILL)
        r += 1
    w.column_dimensions["A"].width = 44
    w.column_dimensions["B"].width = 11
    for c in "CDEF":
        w.column_dimensions[c].width = 12
    w.freeze_panes = "C5"

    # ----- Inputs_Drivers
    d = ws_drv
    d.sheet_view.showGridLines = False
    put(d, "A1", "Inputs_Drivers - every named input the model (and later the website) reads", bold=True)
    d["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3A5F")
    put(d, "A2", "Blue = typed input. Green = live value pulled from Inputs_Scenarios for the selected scenario.", italic=True)
    put(d, "A4", "Scenario selector", bold=True)
    put(d, "C4", "Base", color=BLUE, fill=INP_FILL, bold=True)
    dv = DataValidation(type="list", formula1='"Base,Upside,Downside,Custom"', allow_blank=False)
    d.add_data_validation(dv)
    dv.add("C4")
    put(d, "F4", "Pick Base, Upside, Downside or Custom.", italic=True)
    put(d, "A5", "Scenario index")
    put(d, "C5", "=MATCH(inp_scenario,Inputs_Scenarios!$C$4:$F$4,0)", "int")
    put(d, "F5", "Position of the selected scenario in Inputs_Scenarios (1-4).", italic=True)
    for ci, h in enumerate(["Driver", "Unit", "Value", "Min", "Max", "What it drives"], start=1):
        put(d, f"{L(ci)}7", h, bold=True, color="FFFFFF", fill=HDR_FILL, align="left" if ci in (1, 2, 6) else "right")
    names = {"inp_scenario": "C4", "inp_scen_idx": "C5"}
    r = 8
    for item in INPUTS:
        if item[0] == "sec":
            for ci in range(1, 7):
                d.cell(row=r, column=ci).fill = SEC_FILL
            put(d, f"A{r}", item[1], bold=True, fill=SEC_FILL)
            r += 1
            continue
        kind, name, label, unit, fmt, base, up, down, lo, hi, tip = item
        put(d, f"A{r}", label)
        put(d, f"B{r}", unit, italic=True)
        if kind == S:
            sr = scen_row[name]
            put(d, f"C{r}", f"=INDEX(Inputs_Scenarios!$C{sr}:$F{sr},1,inp_scen_idx)", fmt)
            d[f"C{r}"].font = Font(name=FONT, size=10, color=GREEN)
        else:
            put(d, f"C{r}", base, fmt, color=BLUE, fill=INP_FILL)
        if lo is not None:
            put(d, f"D{r}", lo, fmt, italic=True)
            put(d, f"E{r}", hi, fmt, italic=True)
            if kind == T:
                v = DataValidation(type="decimal", operator="between", formula1=str(lo), formula2=str(hi),
                                   showErrorMessage=True, errorTitle="Out of range",
                                   error=f"Enter a value between {lo} and {hi}.")
                d.add_data_validation(v)
                v.add(f"C{r}")
        put(d, f"F{r}", tip)
        if lo is not None:
            put(d, f"G{r}", f"=IF(OR(C{r}<D{r},C{r}>E{r}),1,0)", "int")
        names[name] = f"C{r}"
        r += 1
    d.column_dimensions["A"].width = 50
    d.column_dimensions["B"].width = 24
    for c in "CDE":
        d.column_dimensions[c].width = 12
    d.column_dimensions["F"].width = 80
    d.column_dimensions["G"].width = 16
    put(d, "G7", "Out of range", bold=True, color="FFFFFF", fill=HDR_FILL, align="right")
    DRV_LAST[0] = r - 1
    d.freeze_panes = "C8"
    SCEN_ROW.update(scen_row)
    DRV_CELL.update(names)
    for n, ref in names.items():
        wb.defined_names[n] = DefinedName(n, attr_text=f"Inputs_Drivers!${ref[0]}${ref[1:]}")


build_inputs()

# ---------------------------------------------------------------------------
# Rev_Build
# ---------------------------------------------------------------------------
RB = Sheet("Rev_Build", "Rev_Build - ARR waterfall and customers by segment")
rb = cl("Rev_Build")
SEG = (("ent", "Enterprise"), ("mid", "Mid-market"))

RB.sec("Volume and price drivers")
RB.add("price_f", "ACV price factor", "x", "x2", lambda c, p, ci: f"=(1+inp_price_uplift)^({c}$5-1)", ann="end")
RB.add("logo_f", "New-logo volume factor", "x", "x2", lambda c, p, ci: f"=(1+inp_logo_growth)^(({c}$3-1)/12)", ann="end")
for s, nm in SEG:
    RB.add(f"logos_{s}", f"New logos - {nm}", "#", "num1", lambda c, p, ci, s=s: f"=inp_logos_{s}*{rb('logo_f', c)}")
    RB.add(f"acv_{s}", f"ACV of new logos - {nm}", "GBP k", "num1", lambda c, p, ci, s=s: f"=inp_acv_{s}*{rb('price_f', c)}", ann="end")
    RB.add(f"churn_m_{s}", f"Monthly gross churn rate - {nm}", "%/mo", "pct", lambda c, p, ci, s=s: f"=1-(1-inp_churn_{s})^(1/12)", ann="end")
    RB.add(f"exp_m_{s}", f"Monthly expansion rate - {nm}", "%/mo", "pct", lambda c, p, ci, s=s: f"=(1+inp_exp_{s})^(1/12)-1", ann="end")
for s, nm in SEG:
    RB.gap()
    RB.sec(f"ARR waterfall - {nm}")
    RB.add(f"arr_open_{s}", "Opening ARR", "GBP k", "num", lambda c, p, ci, s=s: f"={p}{row_of('Rev_Build', f'arr_close_{s}')}", ann="open")
    RB.add(f"arr_new_{s}", "+ New logos", "GBP k", "num", lambda c, p, ci, s=s: f"={rb(f'logos_{s}', c)}*{rb(f'acv_{s}', c)}")
    RB.add(f"arr_exp_{s}", "+ Expansion", "GBP k", "num", lambda c, p, ci, s=s: f"={rb(f'arr_open_{s}', c)}*{rb(f'exp_m_{s}', c)}")
    RB.add(f"arr_churn_{s}", "- Churn", "GBP k", "num", lambda c, p, ci, s=s: f"=-{rb(f'arr_open_{s}', c)}*{rb(f'churn_m_{s}', c)}")
    RB.add(f"arr_close_{s}", "Closing ARR", "GBP k", "num",
           lambda c, p, ci, s=s: f"=SUM({rb(f'arr_open_{s}', c)}:{rb(f'arr_churn_{s}', c)})",
           opening=lambda c, p, ci, s=s: f"=inp_arr_open_{s}", ann="end", bold=True, top=True)
    RB.add(f"cust_open_{s}", "Opening customers", "#", "num1", lambda c, p, ci, s=s: f"={p}{row_of('Rev_Build', f'cust_close_{s}')}", ann="open")
    RB.add(f"cust_new_{s}", "+ New customers", "#", "num1", lambda c, p, ci, s=s: f"={rb(f'logos_{s}', c)}")
    RB.add(f"cust_churn_{s}", "- Churned customers", "#", "num1", lambda c, p, ci, s=s: f"=-{rb(f'cust_open_{s}', c)}*{rb(f'churn_m_{s}', c)}")
    RB.add(f"cust_close_{s}", "Closing customers", "#", "num1",
           lambda c, p, ci, s=s: f"=SUM({rb(f'cust_open_{s}', c)}:{rb(f'cust_churn_{s}', c)})",
           opening=lambda c, p, ci, s=s: f"=inp_cust_open_{s}", ann="end", bold=True, top=True)
RB.gap()
RB.sec("Total company")
for key, lab, fmt, ann in (("arr_open", "Opening ARR", "num", "open"), ("arr_new", "+ New logos", "num", "sum"),
                           ("arr_exp", "+ Expansion", "num", "sum"), ("arr_churn", "- Churn", "num", "sum")):
    RB.add(key, lab, "GBP k", fmt, lambda c, p, ci, key=key: f"={rb(key + '_ent', c)}+{rb(key + '_mid', c)}", ann=ann)
RB.add("arr_close", "Closing ARR", "GBP k", "num", lambda c, p, ci: f"=SUM({rb('arr_open', c)}:{rb('arr_churn', c)})",
       opening=lambda c, p, ci: f"={rb('arr_close_ent', c)}+{rb('arr_close_mid', c)}", ann="end", bold=True, top=True)
for key, lab, ann in (("cust_open", "Opening customers", "open"), ("cust_new", "+ New customers", "sum"),
                      ("cust_churn", "- Churned customers", "sum")):
    RB.add(key, lab, "#", "num1", lambda c, p, ci, key=key: f"={rb(key + '_ent', c)}+{rb(key + '_mid', c)}", ann=ann)
RB.add("cust_close", "Closing customers", "#", "num1", lambda c, p, ci: f"=SUM({rb('cust_open', c)}:{rb('cust_churn', c)})",
       opening=lambda c, p, ci: f"={rb('cust_close_ent', c)}+{rb('cust_close_mid', c)}", ann="end", bold=True, top=True)

# ---------------------------------------------------------------------------
# Rev_Rec
# ---------------------------------------------------------------------------
RR = Sheet("Rev_Rec", "Rev_Rec - billings, ratable revenue (IFRS 15) and deferred revenue", hide_hist=False)
rr = cl("Rev_Rec")
RR.sec("Subscription invoices raised (annual upfront). Columns C:N hold the opening-year billing history")
for s, nm in SEG:
    RR.add(f"inv_{s}", f"Invoices raised - {nm}", "GBP k", "num",
           lambda c, p, ci, s=s: f"={X('Rev_Build', f'arr_new_{s}', c)}+(1+inp_exp_{s}-inp_churn_{s})*{L(ci - 12)}{row_of('Rev_Rec', f'inv_{s}')}",
           hist=lambda c, ci, s=s: f"=inp_arr_open_{s}/12")
RR.add("inv_tot", "Total subscription invoices", "GBP k", "num",
       lambda c, p, ci: f"={rr('inv_ent', c)}+{rr('inv_mid', c)}",
       hist=lambda c, ci: f"={rr('inv_ent', c)}+{rr('inv_mid', c)}", bold=True, top=True)
RR.add("unexp", "Unexpired share of each historic invoice at opening date", "x", "x2", None,
       hist=lambda c, ci: f"=({c}$3+11)/12")
RR.gap()
RR.sec("Implementation fees (separate performance obligation, recognised over 3 months)")
RR.add("fee_impl", "Implementation fees billed", "GBP k", "num",
       lambda c, p, ci: f"=inp_impl_fee_pct*{X('Rev_Build', 'arr_new_ent', c)}")
RR.gap()
RR.sec("Revenue recognised")
for s, nm in SEG:
    RR.add(f"rev_{s}", f"Subscription revenue - {nm} (each invoice over 12 months)", "GBP k", "num",
           lambda c, p, ci, s=s: f"=SUM({L(ci - 11)}{row_of('Rev_Rec', f'inv_{s}')}:{c}{row_of('Rev_Rec', f'inv_{s}')})/12")
RR.add("rev_sub", "Total subscription revenue", "GBP k", "num", lambda c, p, ci: f"={rr('rev_ent', c)}+{rr('rev_mid', c)}", bold=True, top=True)
RR.add("rev_impl", "Implementation revenue", "GBP k", "num",
       lambda c, p, ci: f"=SUM({L(ci - 2)}{row_of('Rev_Rec', 'fee_impl')}:{c}{row_of('Rev_Rec', 'fee_impl')})/3")
RR.add("rev_tot", "Total revenue", "GBP k", "num", lambda c, p, ci: f"={rr('rev_sub', c)}+{rr('rev_impl', c)}", bold=True, top=True)
RR.add("bill_tot", "Total billings", "GBP k", "num", lambda c, p, ci: f"={rr('inv_tot', c)}+{rr('fee_impl', c)}",
       opening=lambda c, p, ci: f"={rr('inv_tot', c)}", bold=True)
RR.gap()
RR.sec("Deferred revenue roll-forward")
RR.add("def_sub_open", "Subscription: opening", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('Rev_Rec', 'def_sub_close')}", ann="open")
RR.add("def_sub_bill", "+ Invoices raised", "GBP k", "num", lambda c, p, ci: f"={rr('inv_tot', c)}")
RR.add("def_sub_rev", "- Revenue recognised", "GBP k", "num", lambda c, p, ci: f"=-{rr('rev_sub', c)}")
RR.add("def_sub_close", "Subscription: closing", "GBP k", "num",
       lambda c, p, ci: f"=SUM({rr('def_sub_open', c)}:{rr('def_sub_rev', c)})",
       opening=lambda c, p, ci: f"=SUMPRODUCT(C{row_of('Rev_Rec', 'inv_tot')}:N{row_of('Rev_Rec', 'inv_tot')},C{row_of('Rev_Rec', 'unexp')}:N{row_of('Rev_Rec', 'unexp')})",
       ann="end", bold=True, top=True)
RR.add("def_imp_open", "Implementation: opening", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('Rev_Rec', 'def_imp_close')}", ann="open")
RR.add("def_imp_bill", "+ Fees billed", "GBP k", "num", lambda c, p, ci: f"={rr('fee_impl', c)}")
RR.add("def_imp_rev", "- Revenue recognised", "GBP k", "num", lambda c, p, ci: f"=-{rr('rev_impl', c)}")
RR.add("def_imp_close", "Implementation: closing", "GBP k", "num",
       lambda c, p, ci: f"=SUM({rr('def_imp_open', c)}:{rr('def_imp_rev', c)})",
       opening=lambda c, p, ci: "=inp_open_defrev_impl", ann="end", bold=True, top=True)
RR.add("def_tot", "Total deferred revenue", "GBP k", "num", lambda c, p, ci: f"={rr('def_sub_close', c)}+{rr('def_imp_close', c)}",
       opening=lambda c, p, ci: f"={rr('def_sub_close', c)}+{rr('def_imp_close', c)}", ann="end", bold=True)
RR.gap()
RR.sec("Trade receivables (payment terms)")
RR.add("ar", "Trade receivables", "GBP k", "num",
       lambda c, p, ci: f"={rr('bill_tot', c)}*MIN(1,inp_dso/30)+{rr('bill_tot', p)}*MAX(0,MIN(1,(inp_dso-30)/30))",
       opening=lambda c, p, ci: "=inp_open_ar", ann="end", bold=True)
RR.gap()
RR.sec("Memo")
RR.add("rev_vs_arr", "Subscription revenue / (closing ARR / 12)", "%", "pct",
       lambda c, p, ci: f"={rr('rev_sub', c)}/({X('Rev_Build', 'arr_close', c)}/12)", ann=None)

# ---------------------------------------------------------------------------
# Headcount_Opex
# ---------------------------------------------------------------------------
HC = Sheet("Headcount_Opex", "Headcount_Opex - headcount by function, payroll and non-headcount opex")
hc = cl("Headcount_Opex")
FUNCS = (("rd", "R&D"), ("sm", "Sales & Marketing"), ("ga", "G&A"), ("cs", "Customer Success"))
for f_, nm in FUNCS:
    HC.sec(f"Headcount - {nm}")
    HC.add(f"h_open_{f_}", "Opening FTE", "FTE", "num1", lambda c, p, ci, f_=f_: f"={p}{row_of('Headcount_Opex', f'h_close_{f_}')}", ann="open")
    HC.add(f"h_hire_{f_}", "+ Net hires", "FTE", "num1", lambda c, p, ci, f_=f_: f"=inp_hires_{f_}")
    HC.add(f"h_close_{f_}", "Closing FTE", "FTE", "num1", lambda c, p, ci, f_=f_: f"={hc(f'h_open_{f_}', c)}+{hc(f'h_hire_{f_}', c)}",
           opening=lambda c, p, ci, f_=f_: f"=inp_heads_open_{f_}", ann="end", bold=True, top=True)
    HC.add(f"pay_{f_}", "Payroll cost (average FTE x loaded cost / 12 x pay-rise factor)", "GBP k", "num",
           lambda c, p, ci, f_=f_: f"=({hc(f'h_open_{f_}', c)}+{hc(f'h_close_{f_}', c)})/2*inp_sal_{f_}/12*(1+inp_merit)^({c}$5-1)")
    HC.gap()
HC.sec("Totals")
HC.add("fte_close", "Total closing FTE", "FTE", "num1",
       lambda c, p, ci: "=" + "+".join(hc(f"h_close_{f_}", c) for f_, _ in FUNCS),
       opening=lambda c, p, ci: "=" + "+".join(hc(f"h_close_{f_}", c) for f_, _ in FUNCS), ann="end", bold=True)
HC.add("pay_tot", "Total payroll", "GBP k", "num", lambda c, p, ci: "=" + "+".join(hc(f"pay_{f_}", c) for f_, _ in FUNCS), bold=True)
HC.gap()
HC.sec("Non-headcount opex")
HC.add("marketing", "Marketing programmes", "GBP k", "num", lambda c, p, ci: "=inp_marketing_m")
HC.add("rd_other", "R&D non-headcount", "GBP k", "num", lambda c, p, ci: "=inp_rd_other_m")
HC.add("ga_other", "G&A non-headcount", "GBP k", "num", lambda c, p, ci: "=inp_ga_other_m")
HC.add("commission", "Sales commission earned (cash; accounting treatment on Contract_Costs)", "GBP k", "num",
       lambda c, p, ci: f"=inp_commission_pct*({X('Rev_Build', 'arr_new', c)}+{X('Rev_Build', 'arr_exp', c)})")
HC.gap()
HC.sec("Operating expense by function (Customer Success payroll sits in COGS)")
HC.add("opex_rd", "R&D", "GBP k", "num", lambda c, p, ci: f"={hc('pay_rd', c)}+{hc('rd_other', c)}")
HC.add("opex_sm", "Sales & Marketing", "GBP k", "num", lambda c, p, ci: f"={hc('pay_sm', c)}+{hc('marketing', c)}+{X('Contract_Costs', 'pl_charge', c)}")
HC.add("opex_ga", "G&A", "GBP k", "num", lambda c, p, ci: f"={hc('pay_ga', c)}+{hc('ga_other', c)}")
HC.add("opex_tot", "Total operating expense", "GBP k", "num", lambda c, p, ci: f"=SUM({hc('opex_rd', c)}:{hc('opex_ga', c)})", bold=True, top=True)

# ---------------------------------------------------------------------------
# Contract_Costs (IFRS 15 incremental costs of obtaining a contract)
# ---------------------------------------------------------------------------
CC = Sheet("Contract_Costs", "Contract_Costs - capitalised sales commission (IFRS 15.91)")
cc = cl("Contract_Costs")
CC.sec("Commission earned and treatment")
CC.add("comm", "Commission earned (cash cost)", "GBP k", "num", lambda c, p, ci: f"={X('Headcount_Opex', 'commission', c)}")
CC.add("cap", "Capitalised", "GBP k", "num", lambda c, p, ci: f"=inp_capitalise_commission*{cc('comm', c)}")
CC.add("exp_now", "Expensed immediately", "GBP k", "num", lambda c, p, ci: f"={cc('comm', c)}-{cc('cap', c)}")
CC.gap()
CC.sec("Asset roll-forward")
CC.add("a_open", "Opening asset", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('Contract_Costs', 'a_close')}", ann="open")
CC.add("a_cap", "+ Capitalised", "GBP k", "num", lambda c, p, ci: f"={cc('cap', c)}")
CC.add("a_amort", "- Amortisation", "GBP k", "num", lambda c, p, ci: f"=-{cc('a_open', c)}/inp_cc_life")
CC.add("a_close", "Closing asset", "GBP k", "num", lambda c, p, ci: f"=SUM({cc('a_open', c)}:{cc('a_amort', c)})",
       opening=lambda c, p, ci: "=inp_open_cca", ann="end", bold=True, top=True)
CC.gap()
CC.sec("Charge to the income statement (S&M)")
CC.add("pl_charge", "Commission expense = expensed now + amortisation", "GBP k", "num", lambda c, p, ci: f"={cc('exp_now', c)}-{cc('a_amort', c)}", bold=True)

# ---------------------------------------------------------------------------
# COGS
# ---------------------------------------------------------------------------
CG = Sheet("COGS", "COGS - hosting, licences, support and services; gross margin by stream")
cg = cl("COGS")
CG.sec("Cost of revenue")
CG.add("host", "Hosting (closing customers x cost per customer)", "GBP k", "num",
       lambda c, p, ci: f"={X('Rev_Build', 'cust_close_ent', c)}*inp_host_ent+{X('Rev_Build', 'cust_close_mid', c)}*inp_host_mid")
CG.add("lic", "Third-party licences", "GBP k", "num", lambda c, p, ci: f"=inp_licence_pct*{X('Rev_Rec', 'rev_sub', c)}")
CG.add("support", "Customer Success payroll", "GBP k", "num", lambda c, p, ci: f"={X('Headcount_Opex', 'pay_cs', c)}")
CG.add("svc", "Implementation delivery cost", "GBP k", "num", lambda c, p, ci: f"=inp_services_cost_pct*{X('Rev_Rec', 'rev_impl', c)}")
CG.add("cogs_tot", "Total COGS", "GBP k", "num", lambda c, p, ci: f"=SUM({cg('host', c)}:{cg('svc', c)})", bold=True, top=True)
CG.gap()
CG.sec("Gross margin by stream")
CG.add("rev_sub", "Subscription revenue", "GBP k", "num", lambda c, p, ci: f"={X('Rev_Rec', 'rev_sub', c)}")
CG.add("gp_sub", "Subscription gross profit", "GBP k", "num", lambda c, p, ci: f"={cg('rev_sub', c)}-{cg('host', c)}-{cg('lic', c)}-{cg('support', c)}")
CG.add("gm_sub", "Subscription gross margin", "%", "pct", lambda c, p, ci: f"={cg('gp_sub', c)}/{cg('rev_sub', c)}", ann="end")
CG.add("rev_impl", "Implementation revenue", "GBP k", "num", lambda c, p, ci: f"={X('Rev_Rec', 'rev_impl', c)}")
CG.add("gp_svc", "Implementation gross profit", "GBP k", "num", lambda c, p, ci: f"={cg('rev_impl', c)}-{cg('svc', c)}")
CG.add("gm_svc", "Implementation gross margin", "%", "pct", lambda c, p, ci: f"=IF({cg('rev_impl', c)}=0,0,{cg('gp_svc', c)}/{cg('rev_impl', c)})", ann="end")
CG.add("gp_tot", "Total gross profit", "GBP k", "num", lambda c, p, ci: f"={cg('gp_sub', c)}+{cg('gp_svc', c)}", bold=True, top=True)
CG.add("gm_tot", "Blended gross margin", "%", "pct", lambda c, p, ci: f"={cg('gp_tot', c)}/({cg('rev_sub', c)}+{cg('rev_impl', c)})", ann="end", bold=True)

# ---------------------------------------------------------------------------
# IS / BS / CF
# ---------------------------------------------------------------------------
IS = Sheet("IS", "IS - income statement")
BS = Sheet("BS", "BS - balance sheet")
CF = Sheet("CF", "CF - cash flow statement (indirect)")
is_, bs, cf = cl("IS"), cl("BS"), cl("CF")

IS.sec("Revenue")
IS.add("rev_ent", "Subscription - Enterprise", "GBP k", "num", lambda c, p, ci: f"={X('Rev_Rec', 'rev_ent', c)}")
IS.add("rev_mid", "Subscription - Mid-market", "GBP k", "num", lambda c, p, ci: f"={X('Rev_Rec', 'rev_mid', c)}")
IS.add("rev_impl", "Implementation", "GBP k", "num", lambda c, p, ci: f"={X('Rev_Rec', 'rev_impl', c)}")
IS.add("rev_tot", "Total revenue", "GBP k", "num", lambda c, p, ci: f"=SUM({is_('rev_ent', c)}:{is_('rev_impl', c)})", bold=True, top=True)
IS.gap()
IS.sec("Cost of revenue")
for k, lab in (("host", "Hosting"), ("lic", "Third-party licences"), ("support", "Customer Success"), ("svc", "Implementation delivery")):
    IS.add(f"c_{k}", lab, "GBP k", "num", lambda c, p, ci, k=k: f"={X('COGS', k, c)}")
IS.add("cogs_tot", "Total cost of revenue", "GBP k", "num", lambda c, p, ci: f"=SUM({is_('c_host', c)}:{is_('c_svc', c)})", bold=True, top=True)
IS.add("gp", "Gross profit", "GBP k", "num", lambda c, p, ci: f"={is_('rev_tot', c)}-{is_('cogs_tot', c)}", bold=True)
IS.add("gm", "Gross margin", "%", "pct", lambda c, p, ci: f"={is_('gp', c)}/{is_('rev_tot', c)}", ann=lambda c: f"={c}{row_of('IS', 'gp')}/{c}{row_of('IS', 'rev_tot')}")
IS.gap()
IS.sec("Operating expense")
for k, lab in (("rd", "Research & development"), ("sm", "Sales & marketing"), ("ga", "General & administrative")):
    IS.add(f"o_{k}", lab, "GBP k", "num", lambda c, p, ci, k=k: f"={X('Headcount_Opex', f'opex_{k}', c)}")
IS.add("opex_tot", "Total operating expense", "GBP k", "num", lambda c, p, ci: f"=SUM({is_('o_rd', c)}:{is_('o_ga', c)})", bold=True, top=True)
IS.gap()
IS.add("ebitda", "EBITDA", "GBP k", "num", lambda c, p, ci: f"={is_('gp', c)}-{is_('opex_tot', c)}", bold=True)
IS.add("ebitda_m", "EBITDA margin", "%", "pct", lambda c, p, ci: f"={is_('ebitda', c)}/{is_('rev_tot', c)}",
       ann=lambda c: f"={c}{row_of('IS', 'ebitda')}/{c}{row_of('IS', 'rev_tot')}")
IS.add("da", "Depreciation", "GBP k", "num", lambda c, p, ci: f"={X('BS', 'dep', c)}")
IS.add("ebit", "Operating profit (EBIT)", "GBP k", "num", lambda c, p, ci: f"={is_('ebitda', c)}-{is_('da', c)}", bold=True, top=True)
IS.add("interest", "Interest income (on prior-period cash)", "GBP k", "num", lambda c, p, ci: f"={X('BS', 'cash', p)}*inp_interest/12")
IS.add("pbt", "Profit before tax", "GBP k", "num", lambda c, p, ci: f"={is_('ebit', c)}+{is_('interest', c)}", bold=True, top=True)
IS.add("tax", "Tax", "GBP k", "num", lambda c, p, ci: f"=MAX(0,{is_('pbt', c)})*inp_tax")
IS.add("ni", "Net income", "GBP k", "num", lambda c, p, ci: f"={is_('pbt', c)}-{is_('tax', c)}", bold=True, top=True)

BS.sec("Assets")
BS.add("cash", "Cash", "GBP k", "num", lambda c, p, ci: f"={X('CF', 'cash_close', c)}", opening=lambda c, p, ci: f"={X('CF', 'cash_close', c)}", ann="end")
BS.add("ar", "Trade receivables", "GBP k", "num", lambda c, p, ci: f"={X('Rev_Rec', 'ar', c)}", opening=lambda c, p, ci: f"={X('Rev_Rec', 'ar', c)}", ann="end")
BS.add("ppe", "PP&E (net)", "GBP k", "num", lambda c, p, ci: f"={bs('ppe_close', c)}", opening=lambda c, p, ci: f"={bs('ppe_close', c)}", ann="end")
BS.add("cca", "Capitalised contract costs", "GBP k", "num", lambda c, p, ci: f"={X('Contract_Costs', 'a_close', c)}", opening=lambda c, p, ci: f"={X('Contract_Costs', 'a_close', c)}", ann="end")
BS.add("ta", "Total assets", "GBP k", "num", lambda c, p, ci: f"=SUM({bs('cash', c)}:{bs('cca', c)})",
       opening=lambda c, p, ci: f"=SUM({bs('cash', c)}:{bs('cca', c)})", ann="end", bold=True, top=True)
BS.gap()
BS.sec("Liabilities")
BS.add("ap", "Trade payables", "GBP k", "num", lambda c, p, ci: f"={bs('nonpay', c)}*inp_dpo/30", opening=lambda c, p, ci: "=inp_open_ap", ann="end")
BS.add("defrev", "Deferred revenue", "GBP k", "num", lambda c, p, ci: f"={X('Rev_Rec', 'def_tot', c)}", opening=lambda c, p, ci: f"={X('Rev_Rec', 'def_tot', c)}", ann="end")
BS.add("tl", "Total liabilities", "GBP k", "num", lambda c, p, ci: f"={bs('ap', c)}+{bs('defrev', c)}",
       opening=lambda c, p, ci: f"={bs('ap', c)}+{bs('defrev', c)}", ann="end", bold=True, top=True)
BS.gap()
BS.sec("Equity")
BS.add("equity", "Total equity (opening net assets + cumulative net income)", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('BS', 'equity')}+{X('IS', 'ni', c)}",
       opening=lambda c, p, ci: f"={bs('ta', c)}-{bs('tl', c)}", ann="end", bold=True)
BS.gap()
BS.add("check", "Balance check (assets - liabilities - equity)", "GBP k", "num2", lambda c, p, ci: f"={bs('ta', c)}-{bs('tl', c)}-{bs('equity', c)}",
       opening=lambda c, p, ci: f"={bs('ta', c)}-{bs('tl', c)}-{bs('equity', c)}", ann="end", bold=True)
BS.gap()
BS.sec("Supporting schedule: PP&E")
BS.add("ppe_open", "Opening net book value", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('BS', 'ppe_close')}", ann="open")
BS.add("capex", "+ Capex", "GBP k", "num", lambda c, p, ci: f"=inp_capex_pct*{X('IS', 'rev_tot', c)}")
BS.add("dep", "- Depreciation", "GBP k", "num", lambda c, p, ci: f"={bs('ppe_open', c)}/inp_dep_life")
BS.add("ppe_close", "Closing net book value", "GBP k", "num", lambda c, p, ci: f"={bs('ppe_open', c)}+{bs('capex', c)}-{bs('dep', c)}",
       opening=lambda c, p, ci: "=inp_open_ppe", ann="end", bold=True, top=True)
BS.gap()
BS.sec("Supporting schedule: non-payroll spend (drives payables)")
BS.add("nonpay", "Hosting, licences, delivery, marketing, other opex", "GBP k", "num",
       lambda c, p, ci: f"={X('COGS', 'host', c)}+{X('COGS', 'lic', c)}+{X('COGS', 'svc', c)}+{X('Headcount_Opex', 'marketing', c)}+{X('Headcount_Opex', 'rd_other', c)}+{X('Headcount_Opex', 'ga_other', c)}")

CF.sec("Operating activities")
CF.add("ni", "Net income", "GBP k", "num", lambda c, p, ci: f"={X('IS', 'ni', c)}")
CF.add("da", "+ Depreciation", "GBP k", "num", lambda c, p, ci: f"={X('IS', 'da', c)}")
CF.add("d_ar", "(Increase) / decrease in receivables", "GBP k", "num", lambda c, p, ci: f"=-({X('BS', 'ar', c)}-{X('BS', 'ar', p)})")
CF.add("d_ap", "Increase / (decrease) in payables", "GBP k", "num", lambda c, p, ci: f"={X('BS', 'ap', c)}-{X('BS', 'ap', p)}")
CF.add("d_def", "Increase / (decrease) in deferred revenue", "GBP k", "num", lambda c, p, ci: f"={X('BS', 'defrev', c)}-{X('BS', 'defrev', p)}")
CF.add("d_cca", "(Increase) / decrease in capitalised contract costs", "GBP k", "num", lambda c, p, ci: f"=-({X('BS', 'cca', c)}-{X('BS', 'cca', p)})")
CF.add("cfo", "Cash flow from operations", "GBP k", "num", lambda c, p, ci: f"=SUM({cf('ni', c)}:{cf('d_cca', c)})", bold=True, top=True)
CF.gap()
CF.sec("Investing activities")
CF.add("capex", "Capex", "GBP k", "num", lambda c, p, ci: f"=-{X('BS', 'capex', c)}")
CF.add("cfi", "Cash flow from investing", "GBP k", "num", lambda c, p, ci: f"={cf('capex', c)}", bold=True, top=True)
CF.gap()
CF.sec("Net cash flow and cash balance (no financing flows in the MVP)")
CF.add("net", "Net change in cash (= free cash flow)", "GBP k", "num", lambda c, p, ci: f"={cf('cfo', c)}+{cf('cfi', c)}", bold=True)
CF.add("cash_open", "Opening cash", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('CF', 'cash_close')}", ann="open")
CF.add("cash_close", "Closing cash", "GBP k", "num", lambda c, p, ci: f"={cf('cash_open', c)}+{cf('net', c)}",
       opening=lambda c, p, ci: "=inp_open_cash", ann="end", bold=True, top=True)

# ---------------------------------------------------------------------------
# SaaS_Metrics
# ---------------------------------------------------------------------------
SM = Sheet("SaaS_Metrics", "SaaS_Metrics - ARR bridge and unit economics")
sm = cl("SaaS_Metrics")
SM.sec("ARR bridge")
SM.add("arr_open", "Opening ARR", "GBP k", "num", lambda c, p, ci: f"={X('Rev_Build', 'arr_open', c)}", ann="open")
SM.add("arr_new", "+ New logos", "GBP k", "num", lambda c, p, ci: f"={X('Rev_Build', 'arr_new', c)}")
SM.add("arr_exp", "+ Expansion", "GBP k", "num", lambda c, p, ci: f"={X('Rev_Build', 'arr_exp', c)}")
SM.add("arr_churn", "- Churn", "GBP k", "num", lambda c, p, ci: f"={X('Rev_Build', 'arr_churn', c)}")
SM.add("arr_close", "Closing ARR", "GBP k", "num", lambda c, p, ci: f"=SUM({sm('arr_open', c)}:{sm('arr_churn', c)})",
       opening=lambda c, p, ci: f"={X('Rev_Build', 'arr_close', c)}", ann="end", bold=True, top=True)
SM.add("arr_gap", "Check: segment ARR - total ARR", "GBP k", "num2",
       lambda c, p, ci: f"={X('Rev_Build', 'arr_close_ent', c)}+{X('Rev_Build', 'arr_close_mid', c)}-{sm('arr_close', c)}",
       opening=lambda c, p, ci: f"={X('Rev_Build', 'arr_close_ent', c)}+{X('Rev_Build', 'arr_close_mid', c)}-{sm('arr_close', c)}", ann=None)
SM.add("netnew", "Net new ARR", "GBP k", "num", lambda c, p, ci: f"={sm('arr_new', c)}+{sm('arr_exp', c)}+{sm('arr_churn', c)}")
SM.add("growth", "ARR growth, year on year", "%", "pct",
       lambda c, p, ci: f'=IF({c}$3<12,"",{sm("arr_close", c)}/INDEX($N{row_of("SaaS_Metrics", "arr_close")}:$AL{row_of("SaaS_Metrics", "arr_close")},1,{c}$3-12+1)-1)',
       ann="end")
SM.gap()
SM.sec("Retention (annualised from each month's flows)")
SM.add("nrr", "Net revenue retention (NRR)", "%", "pct", lambda c, p, ci: f"=(1+({sm('arr_exp', c)}+{sm('arr_churn', c)})/{sm('arr_open', c)})^12", ann="end", bold=True)
SM.add("grr", "Gross revenue retention (GRR)", "%", "pct", lambda c, p, ci: f"=(1+{sm('arr_churn', c)}/{sm('arr_open', c)})^12", ann="end", bold=True)
SM.add("logo_churn", "Logo churn (annualised)", "%", "pct",
       lambda c, p, ci: f"=1-(1+{X('Rev_Build', 'cust_churn', c)}/{X('Rev_Build', 'cust_open', c)})^12", ann="end")
SM.gap()
SM.sec("Customer acquisition and unit economics")
SM.add("sm_spend", "Sales & marketing spend (cash basis: P&L + capitalised - amortised commission)", "GBP k", "num",
       lambda c, p, ci: f"={X('IS', 'o_sm', c)}+{X('Contract_Costs', 'cap', c)}+{X('Contract_Costs', 'a_amort', c)}")
SM.add("new_logos", "New logos", "#", "num1", lambda c, p, ci: f"={X('Rev_Build', 'cust_new', c)}")
SM.add("cac", "CAC per new logo (blended)", "GBP k", "num1", lambda c, p, ci: f"={sm('sm_spend', c)}/{sm('new_logos', c)}",
       ann=lambda c: f"={c}{row_of('SaaS_Metrics', 'sm_spend')}/{c}{row_of('SaaS_Metrics', 'new_logos')}")
SM.add("gm", "Gross margin", "%", "pct", lambda c, p, ci: f"={X('IS', 'gm', c)}", ann="end")
SM.add("payback", "CAC payback (months of gross-margin on new + expansion ARR)", "months", "mo",
       lambda c, p, ci: f"={sm('sm_spend', c)}/(({sm('arr_new', c)}+{sm('arr_exp', c)})*{sm('gm', c)}/12)", ann="end", bold=True)
SM.add("arpa", "ARR per customer", "GBP k", "num1", lambda c, p, ci: f"={sm('arr_close', c)}/{X('Rev_Build', 'cust_close', c)}", ann="end")
SM.add("ltv", "Customer lifetime value (ARPA x GM / logo churn)", "GBP k", "num", lambda c, p, ci: f"={sm('arpa', c)}*{sm('gm', c)}/{sm('logo_churn', c)}", ann="end")
SM.add("ltv_cac", "LTV / CAC", "x", "x", lambda c, p, ci: f"={sm('ltv', c)}/{sm('cac', c)}", ann="end", bold=True)
SM.add("magic", "Magic number (3m net new ARR / prior 3m S&M)", "x", "x2",
       lambda c, p, ci: f'=IF({c}$3<6,"",SUM({L(ci - 2)}{row_of("SaaS_Metrics", "netnew")}:{c}{row_of("SaaS_Metrics", "netnew")})/SUM({L(ci - 5)}{row_of("SaaS_Metrics", "sm_spend")}:{L(ci - 3)}{row_of("SaaS_Metrics", "sm_spend")}))',
       ann="end")
SM.gap()
SM.sec("Cash efficiency")
SM.add("netburn", "Net burn (positive = cash consumed)", "GBP k", "num", lambda c, p, ci: f"=-{X('CF', 'net', c)}")
SM.add("burn_mult", "Burn multiple (net burn / net new ARR)", "x", "x",
       lambda c, p, ci: f'=IF({sm("netnew", c)}<=0,"n/a",MAX(0,{sm("netburn", c)})/{sm("netnew", c)})', ann="end", bold=True)
SM.add("cash", "Closing cash", "GBP k", "num", lambda c, p, ci: f"={X('BS', 'cash', c)}", ann="end")
SM.add("runway", "Cash runway (months at trailing 3m average burn)", "months", "mo",
       lambda c, p, ci: f'=IF(AVERAGE({L(ci - 2)}{row_of("SaaS_Metrics", "netburn")}:{c}{row_of("SaaS_Metrics", "netburn")})<=0,"n/a",{sm("cash", c)}/AVERAGE({L(ci - 2)}{row_of("SaaS_Metrics", "netburn")}:{c}{row_of("SaaS_Metrics", "netburn")}))',
       ann="end", bold=True)
SM.add("rule40", "Rule of 40 (ARR growth + EBITDA margin)", "%", "pct",
       lambda c, p, ci: f'=IF({c}$3<12,"",{sm("growth", c)}+{X("IS", "ebitda_m", c)})', ann="end", bold=True)
SM.add("arr_fte", "ARR per FTE", "GBP k", "num", lambda c, p, ci: f"={sm('arr_close', c)}/{X('Headcount_Opex', 'fte_close', c)}", ann="end")


# ---------------------------------------------------------------------------
# Phase 2 sheets
# ---------------------------------------------------------------------------
import json
import os

_HERE = os.path.dirname(os.path.abspath(OUT))
SPEC_PATH = os.path.join(_HERE, "snap_spec.json")
SNAP_PATH = os.path.join(_HERE, "snapshots.json")
SNAP = json.load(open(SNAP_PATH)) if os.path.exists(SNAP_PATH) else None
GREY = "595959"


def rng(sh, key, a="O", b="AL"):
    r_ = row_of(sh, key)
    return f"{sh}!${a}${r_}:${b}${r_}"

EXTRA_CHECKS = []


def sn(*path):
    """Look up a value in the engine-run snapshot (None on the first build pass)."""
    d = SNAP
    for k in path:
        if d is None:
            return None
        d = d.get(k) if isinstance(d, dict) else None
    return d


def new_sheet(name, title, sub, widths):
    w = wb.create_sheet(name)
    w.sheet_view.showGridLines = False
    put(w, "A1", title, bold=True)
    w["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3A5F")
    put(w, "A2", sub, italic=True)
    for col, wd in widths.items():
        w.column_dimensions[col].width = wd
    return w


def hdr(w, row, labels, start=1):
    for i, lab in enumerate(labels):
        put(w, f"{L(start + i)}{row}", lab, bold=True, color="FFFFFF", fill=HDR_FILL, align="left" if (start + i) == 1 else "right")


def secrow(w, row, label, ncol):
    for ci in range(1, ncol + 1):
        w.cell(row=row, column=ci).fill = SEC_FILL
    put(w, f"A{row}", label, bold=True, fill=SEC_FILL)


def static(w, ref, value, fmt=None):
    """A pasted engine-run value: grey italic so it is never mistaken for a live formula."""
    put(w, ref, value, fmt, color=GREY, italic=True)


# ---------------------------------------------------------------------------
# Phase 3: Vaultline acquisition module
# ---------------------------------------------------------------------------
# ----- MA_Target: Vaultline standalone (simplified), same time axis as the Northgate tabs ------
MT = Sheet("MA_Target", "MA_Target - Vaultline Archive Ltd, standalone (simplified three statements)")
mt = cl("MA_Target")
MT.sec("Drivers")
MT.add("churn_m", "Monthly gross churn rate", "%/mo", "pct", lambda c, p, ci: "=1-(1-inp_v_churn)^(1/12)", ann="end")
MT.add("exp_m", "Monthly expansion rate", "%/mo", "pct", lambda c, p, ci: "=(1+inp_v_exp)^(1/12)-1", ann="end")
MT.gap()
MT.sec("ARR and customers")
MT.add("arr_open", "Opening ARR", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('MA_Target', 'arr_close')}", ann="open")
MT.add("arr_new", "+ New logos", "GBP k", "num", lambda c, p, ci: "=inp_v_logos*inp_v_acv")
MT.add("arr_exp", "+ Expansion", "GBP k", "num", lambda c, p, ci: f"={mt('arr_open', c)}*{mt('exp_m', c)}")
MT.add("arr_churn", "- Churn", "GBP k", "num", lambda c, p, ci: f"=-{mt('arr_open', c)}*{mt('churn_m', c)}")
MT.add("arr_close", "Closing ARR", "GBP k", "num", lambda c, p, ci: f"=SUM({mt('arr_open', c)}:{mt('arr_churn', c)})",
       opening=lambda c, p, ci: "=inp_v_arr_open", ann="end", bold=True, top=True)
MT.add("cust_close", "Closing customers", "#", "num1",
       lambda c, p, ci: f"={p}{row_of('MA_Target', 'cust_close')}*(1-{mt('churn_m', c)})+inp_v_logos",
       opening=lambda c, p, ci: "=inp_v_cust_open", ann="end")
MT.gap()
MT.sec("Billings and revenue (annual upfront, recognised ratably; columns C:N hold opening-year history)")
MT.add("inv", "Invoices raised", "GBP k", "num",
       lambda c, p, ci: f"={mt('arr_new', c)}+(1+inp_v_exp-inp_v_churn)*{L(ci - 12)}{row_of('MA_Target', 'inv')}",
       hist=lambda c, ci: "=inp_v_arr_open/12")
MT.add("unexp", "Unexpired share of each historic invoice at opening date", "x", "x2", None, hist=lambda c, ci: f"=({c}$3+11)/12")
MT.add("rev", "Revenue (each invoice over 12 months)", "GBP k", "num",
       lambda c, p, ci: f"=SUM({L(ci - 11)}{row_of('MA_Target', 'inv')}:{c}{row_of('MA_Target', 'inv')})/12", bold=True)
MT.add("dr_open", "Deferred revenue: opening", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('MA_Target', 'dr_close')}", ann="open")
MT.add("dr_bill", "+ Invoices raised", "GBP k", "num", lambda c, p, ci: f"={mt('inv', c)}")
MT.add("dr_rev", "- Revenue recognised", "GBP k", "num", lambda c, p, ci: f"=-{mt('rev', c)}")
MT.add("dr_close", "Deferred revenue: closing", "GBP k", "num", lambda c, p, ci: f"=SUM({mt('dr_open', c)}:{mt('dr_rev', c)})",
       opening=lambda c, p, ci: f"=SUMPRODUCT(C{row_of('MA_Target', 'inv')}:N{row_of('MA_Target', 'inv')},C{row_of('MA_Target', 'unexp')}:N{row_of('MA_Target', 'unexp')})",
       ann="end", bold=True, top=True)
MT.add("ar", "Trade receivables", "GBP k", "num",
       lambda c, p, ci: f"={mt('inv', c)}*MIN(1,inp_v_dso/30)+{mt('inv', p)}*MAX(0,MIN(1,(inp_v_dso-30)/30))",
       opening=lambda c, p, ci: "=inp_v_open_ar", ann="end")
MT.gap()
MT.sec("Income statement")
MT.add("cogs", "Cost of revenue", "GBP k", "num", lambda c, p, ci: f"={mt('rev', c)}*inp_v_cogs_pct")
MT.add("gp", "Gross profit", "GBP k", "num", lambda c, p, ci: f"={mt('rev', c)}-{mt('cogs', c)}", bold=True)
MT.add("rd", "R&D", "GBP k", "num", lambda c, p, ci: f"={mt('rev', c)}*inp_v_rd_pct")
MT.add("sm", "Sales & marketing", "GBP k", "num", lambda c, p, ci: f"={mt('rev', c)}*inp_v_sm_pct")
MT.add("ga", "G&A", "GBP k", "num", lambda c, p, ci: f"={mt('rev', c)}*inp_v_ga_pct")
MT.add("ebitda", "EBITDA", "GBP k", "num", lambda c, p, ci: f"={mt('gp', c)}-{mt('rd', c)}-{mt('sm', c)}-{mt('ga', c)}", bold=True, top=True)
MT.add("ebitda_m", "EBITDA margin", "%", "pct", lambda c, p, ci: f"={mt('ebitda', c)}/{mt('rev', c)}",
       ann=lambda c: f"={c}{row_of('MA_Target', 'ebitda')}/{c}{row_of('MA_Target', 'rev')}")
MT.add("dep", "Depreciation", "GBP k", "num", lambda c, p, ci: f"={mt('ppe_open', c)}/inp_v_dep_life")
MT.add("interest", "Interest income (prior-period cash)", "GBP k", "num", lambda c, p, ci: f"={mt('cash', p)}*inp_interest/12")
MT.add("pbt", "Profit before tax", "GBP k", "num", lambda c, p, ci: f"={mt('ebitda', c)}-{mt('dep', c)}+{mt('interest', c)}", top=True)
MT.add("tax", "Tax", "GBP k", "num", lambda c, p, ci: f"=MAX(0,{mt('pbt', c)})*inp_tax")
MT.add("ni", "Net income", "GBP k", "num", lambda c, p, ci: f"={mt('pbt', c)}-{mt('tax', c)}", bold=True, top=True)
MT.gap()
MT.sec("PP&E schedule")
MT.add("ppe_open", "Opening NBV", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('MA_Target', 'ppe_close')}", ann="open")
MT.add("capex", "+ Capex", "GBP k", "num", lambda c, p, ci: f"={mt('rev', c)}*inp_v_capex_pct")
MT.add("ppe_dep", "- Depreciation", "GBP k", "num", lambda c, p, ci: f"=-{mt('dep', c)}")
MT.add("ppe_close", "Closing NBV", "GBP k", "num", lambda c, p, ci: f"=SUM({mt('ppe_open', c)}:{mt('ppe_dep', c)})",
       opening=lambda c, p, ci: "=inp_v_open_ppe", ann="end", bold=True, top=True)
MT.gap()
MT.sec("Balance sheet")
MT.add("cash", "Cash", "GBP k", "num", lambda c, p, ci: f"={mt('cash_close', c)}", opening=lambda c, p, ci: f"={mt('cash_close', c)}", ann="end")
MT.add("bs_ar", "Trade receivables", "GBP k", "num", lambda c, p, ci: f"={mt('ar', c)}", opening=lambda c, p, ci: f"={mt('ar', c)}", ann="end")
MT.add("bs_ppe", "PP&E (net)", "GBP k", "num", lambda c, p, ci: f"={mt('ppe_close', c)}", opening=lambda c, p, ci: f"={mt('ppe_close', c)}", ann="end")
MT.add("ta", "Total assets", "GBP k", "num", lambda c, p, ci: f"=SUM({mt('cash', c)}:{mt('bs_ppe', c)})",
       opening=lambda c, p, ci: f"=SUM({mt('cash', c)}:{mt('bs_ppe', c)})", ann="end", bold=True, top=True)
MT.add("ap", "Trade payables", "GBP k", "num", lambda c, p, ci: f"=({mt('cogs', c)}+{mt('rd', c)}+{mt('sm', c)}+{mt('ga', c)})*inp_v_ap_months",
       opening=lambda c, p, ci: "=inp_v_arr_open/12*(inp_v_cogs_pct+inp_v_rd_pct+inp_v_sm_pct+inp_v_ga_pct)*inp_v_ap_months", ann="end")
MT.add("bs_dr", "Deferred revenue", "GBP k", "num", lambda c, p, ci: f"={mt('dr_close', c)}", opening=lambda c, p, ci: f"={mt('dr_close', c)}", ann="end")
MT.add("tl", "Total liabilities", "GBP k", "num", lambda c, p, ci: f"={mt('ap', c)}+{mt('bs_dr', c)}",
       opening=lambda c, p, ci: f"={mt('ap', c)}+{mt('bs_dr', c)}", ann="end", bold=True, top=True)
MT.add("equity", "Equity (opening net assets + cumulative net income)", "GBP k", "num",
       lambda c, p, ci: f"={p}{row_of('MA_Target', 'equity')}+{mt('ni', c)}", opening=lambda c, p, ci: f"={mt('ta', c)}-{mt('tl', c)}", ann="end", bold=True)
MT.add("check", "Balance check", "GBP k", "num2", lambda c, p, ci: f"={mt('ta', c)}-{mt('tl', c)}-{mt('equity', c)}",
       opening=lambda c, p, ci: f"={mt('ta', c)}-{mt('tl', c)}-{mt('equity', c)}", ann="end")
MT.gap()
MT.sec("Cash flow")
MT.add("cf_dep", "Depreciation", "GBP k", "num", lambda c, p, ci: f"={mt('dep', c)}")
MT.add("cf_ar", "(Increase) / decrease in receivables", "GBP k", "num", lambda c, p, ci: f"=-({mt('bs_ar', c)}-{mt('bs_ar', p)})")
MT.add("cf_ap", "Increase / (decrease) in payables", "GBP k", "num", lambda c, p, ci: f"={mt('ap', c)}-{mt('ap', p)}")
MT.add("cf_dr", "Increase / (decrease) in deferred revenue", "GBP k", "num", lambda c, p, ci: f"={mt('bs_dr', c)}-{mt('bs_dr', p)}")
MT.add("cf_net", "Net cash flow (net income + non-cash + working capital - capex)", "GBP k", "num",
       lambda c, p, ci: f"={mt('ni', c)}+{mt('cf_dep', c)}+{mt('cf_ar', c)}+{mt('cf_ap', c)}+{mt('cf_dr', c)}-{mt('capex', c)}", bold=True, top=True)
MT.add("cash_open", "Opening cash", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('MA_Target', 'cash_close')}", ann="open")
MT.add("cash_close", "Closing cash", "GBP k", "num", lambda c, p, ci: f"={mt('cash_open', c)}+{mt('cf_net', c)}",
       opening=lambda c, p, ci: "=inp_v_open_cash", ann="end", bold=True, top=True)
MT.add("ufcf", "Unlevered free cash flow (EBITDA - tax - capex + working capital)", "GBP k", "num",
       lambda c, p, ci: f"={mt('ebitda', c)}-{mt('tax', c)}-{mt('capex', c)}+{mt('cf_ar', c)}+{mt('cf_ap', c)}+{mt('cf_dr', c)}", bold=True)

# ----- MA_PPA: IFRS 3 purchase price allocation --------------------------------------------------
pa = new_sheet("MA_PPA", "MA_PPA - IFRS 3 purchase price allocation (deal closes at the model's opening date)",
               "GBP thousands. Column C = IFRS 3 (acquired deferred revenue at fair value). Column D = US GAAP after ASU 2021-08 (deferred revenue carried over per ASC 606), shown for comparison.",
               {"A": 58, "B": 3, "C": 16, "D": 16, "E": 80})
hdr(pa, 4, ["", "", "IFRS 3", "US GAAP", "Note"])
PPA = {}
_pr = [5]


def pline(key, label, c_fx, d_fx=None, fmt="num", bold=False, note="", top=False):
    r = _pr[0]
    _pr[0] += 1
    PPA[key] = r
    put(pa, f"A{r}", label, bold=bold)
    put(pa, f"C{r}", c_fx, fmt, bold)
    put(pa, f"D{r}", c_fx if d_fx is None else d_fx, fmt, bold)
    if note:
        put(pa, f"E{r}", note, italic=True)
    if top:
        for col in "ACD":
            pa[f"{col}{r}"].border = Border(top=THIN)


def psec(label):
    secrow(pa, _pr[0], label, 5)
    _pr[0] += 1


def pc(key, col="C"):
    return f"{col}{PPA[key]}"


def PA(key, col="C"):
    return f"MA_PPA!${col}${PPA[key]}"


psec("Consideration")
pline("arr", "Vaultline opening ARR", f"={X('MA_Target', 'arr_close', 'N')}")
pline("mult", "Offer multiple (EV / ARR)", "=inp_ma_offer_mult", fmt="x")
pline("ev", "Enterprise value offered", f"={pc('arr')}*{pc('mult')}", bold=True, top=True)
pline("cash_acq", "+ Cash acquired (Vaultline has no debt)", f"={X('MA_Target', 'cash', 'N')}")
pline("price", "Equity purchase price", f"={pc('ev')}+{pc('cash_acq')}", bold=True, top=True)
pline("loan", "Funded by: acquisition term loan", "=inp_ma_debt")
pline("cash_paid", "Funded by: Northgate cash", f"={pc('price')}-{pc('loan')}", note="Check the Checks tab: pro-forma cash must stay positive")
pline("txn", "Transaction costs (expensed in month 1, not part of consideration)", "=inp_ma_deal_cost_pct*" + pc("ev"), note="IFRS 3.53: acquisition-related costs are expensed")
_pr[0] += 1
psec("Net assets acquired at book value")
pline("b_cash", "Cash", f"={X('MA_Target', 'cash', 'N')}")
pline("b_ar", "Trade receivables", f"={X('MA_Target', 'bs_ar', 'N')}")
pline("b_ppe", "PP&E", f"={X('MA_Target', 'bs_ppe', 'N')}")
pline("b_ap", "Trade payables", f"=-{X('MA_Target', 'ap', 'N')}")
pline("b_dr", "Deferred revenue (book)", f"=-{X('MA_Target', 'bs_dr', 'N')}")
pline("book_na", "Net assets at book value", f"=SUM({pc('b_cash')}:{pc('b_dr')})", bold=True, top=True, note="Equals Vaultline's opening equity")
_pr[0] += 1
psec("Fair value adjustments")
pline("dr_adj", "Deferred revenue: fair-value haircut (reduces the liability)", f"=inp_ma_dr_haircut*-{pc('b_dr')}", "=0",
      note="IFRS 3 measures it at fair value: less than book. US GAAP (ASU 2021-08) carries it over, so no haircut")
pline("cr", "Customer relationships (intangible)", f"=inp_ma_cr_pct*{pc('ev')}", note="Amortised straight-line over inp_ma_cr_life")
pline("tech", "Acquired technology (intangible)", f"=inp_ma_tech_pct*{pc('ev')}", note="Amortised straight-line over inp_ma_tech_life")
pline("dtl", "Deferred tax liability on intangibles", f"=-inp_tax*({pc('cr')}+{pc('tech')})", note="Intangibles have no tax base in a share deal, so a DTL is recognised")
pline("fv_na", "Fair value of identifiable net assets", f"={pc('book_na')}+{pc('dr_adj')}+{pc('cr')}+{pc('tech')}+{pc('dtl')}",
      f"={pc('book_na', 'D')}+{pc('dr_adj', 'D')}+{pc('cr', 'D')}+{pc('tech', 'D')}+{pc('dtl', 'D')}", bold=True, top=True)
_pr[0] += 1
psec("Goodwill")
pline("gw", "Goodwill = equity purchase price - fair value of net assets", f"={pc('price')}-{pc('fv_na')}", f"={pc('price', 'D')}-{pc('fv_na', 'D')}", bold=True,
      note="Not amortised under IFRS; tested annually for impairment. Positive = goodwill; negative would be a bargain purchase")
pline("gw_pct", "Goodwill as % of equity purchase price", f"={pc('gw')}/{pc('price')}", f"={pc('gw', 'D')}/{pc('price', 'D')}", fmt="pct")
_pr[0] += 1
psec("What the choice of standard changes after closing")
pline("hc_loss_y1", "Year 1 revenue lost because acquired deferred revenue is written down", f"={pc('dr_adj')}", "=0",
      note="Unwinds over 12 months. The cash is still collected; reported revenue and EBITDA are lower under IFRS 3")
pline("amort_m", "Monthly amortisation of acquired intangibles", f"={pc('cr')}/inp_ma_cr_life+{pc('tech')}/inp_ma_tech_life")
pline("amort_y", "Annual amortisation of acquired intangibles", f"={pc('amort_m')}*12", note="Below EBITDA but reduces operating profit and net income")
for col in "CD":
    pass
pa.freeze_panes = "A5"
EXTRA_CHECKS.append(("PPA reconciles", "Consideration = fair value of net assets + goodwill (both bases)",
                     f"=ABS(MA_PPA!C{PPA['price']}-MA_PPA!C{PPA['fv_na']}-MA_PPA!C{PPA['gw']})+ABS(MA_PPA!D{PPA['price']}-MA_PPA!D{PPA['fv_na']}-MA_PPA!D{PPA['gw']})", 0.01))
EXTRA_CHECKS.append(("Goodwill is positive", "A negative figure would be a bargain purchase needing re-assessment (shortfall shown)", f"=MIN(0,MA_PPA!C{PPA['gw']})", 0))
EXTRA_CHECKS.append(("Vaultline book net assets = equity", "PPA net assets at book = Vaultline opening equity", f"=MA_PPA!C{PPA['book_na']}-{X('MA_Target', 'equity', 'N')}", 0.01))
EXTRA_CHECKS.append(("Vaultline balance sheet balances", "Largest absolute balance-sheet gap over 24 months", f"=MAX(MAX({rng('MA_Target', 'check', 'N')}),-MIN({rng('MA_Target', 'check', 'N')}))", 0.01))
EXTRA_CHECKS.append(("Vaultline ARR roll-forward", "Closing ARR = opening + new + expansion - churn over 24 months",
                     f"={X('MA_Target', 'arr_close', 'AL')}-({X('MA_Target', 'arr_close', 'N')}+SUM({rng('MA_Target', 'arr_new')})+SUM({rng('MA_Target', 'arr_exp')})+SUM({rng('MA_Target', 'arr_churn')}))", 0.01))
EXTRA_CHECKS.append(("Vaultline deferred revenue roll-forward", "Closing = opening + invoices - revenue",
                     f"={X('MA_Target', 'dr_close', 'AL')}-({X('MA_Target', 'dr_close', 'N')}+SUM({rng('MA_Target', 'inv')})-SUM({rng('MA_Target', 'rev')}))", 0.01))
EXTRA_CHECKS.append(("Valuation: WACC above terminal growth", "Gordon growth needs WACC > g (shortfall shown)", "=MIN(0,inp_ma_wacc-inp_ma_wacc_band-inp_ma_tg-inp_ma_tg_band)", 0))

# ----- MA_ProForma: combined entity --------------------------------------------------------------
MP = Sheet("MA_ProForma", "MA_ProForma - Northgate + Vaultline: eliminations, synergies, integration, PPA effects (deal closes at t=0)")
mp = cl("MA_ProForma")
INC = "inp_ma_include"
MP.sec("Northgate standalone (links)")
for key, lab, sh, k, ann in (("n_arr_open", "Opening ARR", "SaaS_Metrics", "arr_open", "open"), ("n_arr_new", "+ New logos", "SaaS_Metrics", "arr_new", "sum"),
                             ("n_arr_exp", "+ Expansion", "SaaS_Metrics", "arr_exp", "sum"), ("n_arr_churn", "- Churn", "SaaS_Metrics", "arr_churn", "sum")):
    MP.add(key, "Northgate ARR: " + lab, "GBP k", "num", lambda c, p, ci, sh=sh, k=k: f"={X(sh, k, c)}", ann=ann)
MP.add("n_arr", "Northgate closing ARR", "GBP k", "num", lambda c, p, ci: f"={X('SaaS_Metrics', 'arr_close', c)}", opening=lambda c, p, ci: f"={X('SaaS_Metrics', 'arr_close', c)}", ann="end")
MP.add("n_rev", "Northgate revenue", "GBP k", "num", lambda c, p, ci: f"={X('IS', 'rev_tot', c)}")
MP.add("n_ebitda", "Northgate EBITDA", "GBP k", "num", lambda c, p, ci: f"={X('IS', 'ebitda', c)}")
MP.add("n_ni", "Northgate net income", "GBP k", "num", lambda c, p, ci: f"={X('IS', 'ni', c)}")
MP.add("n_fcf", "Northgate free cash flow", "GBP k", "num", lambda c, p, ci: f"={X('CF', 'net', c)}")
for key, lab, k in (("n_cash", "cash", "cash"), ("n_ar", "receivables", "ar"), ("n_ppe", "PP&E", "ppe"), ("n_cca", "contract-cost asset", "cca"),
                    ("n_ap", "payables", "ap"), ("n_dr", "deferred revenue", "defrev"), ("n_eq", "equity", "equity")):
    MP.add(key, f"Northgate {lab}", "GBP k", "num", lambda c, p, ci, k=k: f"={X('BS', k, c)}", opening=lambda c, p, ci, k=k: f"={X('BS', k, c)}", ann="end")
MP.gap()
MP.sec("Vaultline (multiplied by the include switch)")
for key, lab, k, ann in (("v_arr_open", "Opening ARR", "arr_open", "open"), ("v_arr_new", "+ New logos", "arr_new", "sum"),
                         ("v_arr_exp", "+ Expansion", "arr_exp", "sum"), ("v_arr_churn", "- Churn", "arr_churn", "sum")):
    MP.add(key, "Vaultline ARR: " + lab, "GBP k", "num", lambda c, p, ci, k=k: f"={INC}*{X('MA_Target', k, c)}", ann=ann)
MP.add("v_arr", "Vaultline closing ARR", "GBP k", "num", lambda c, p, ci: f"={INC}*{X('MA_Target', 'arr_close', c)}", opening=lambda c, p, ci: f"={INC}*{X('MA_Target', 'arr_close', c)}", ann="end")
MP.add("v_rev", "Vaultline revenue", "GBP k", "num", lambda c, p, ci: f"={INC}*{X('MA_Target', 'rev', c)}")
MP.add("v_ebitda", "Vaultline EBITDA", "GBP k", "num", lambda c, p, ci: f"={INC}*{X('MA_Target', 'ebitda', c)}")
MP.add("v_ni", "Vaultline net income", "GBP k", "num", lambda c, p, ci: f"={INC}*{X('MA_Target', 'ni', c)}")
MP.add("v_fcf", "Vaultline free cash flow", "GBP k", "num", lambda c, p, ci: f"={INC}*{X('MA_Target', 'cf_net', c)}")
for key, lab, k in (("v_cash", "cash", "cash"), ("v_ar", "receivables", "bs_ar"), ("v_ppe", "PP&E", "bs_ppe"), ("v_ap", "payables", "ap"), ("v_dr", "deferred revenue (book)", "bs_dr")):
    MP.add(key, f"Vaultline {lab}", "GBP k", "num", lambda c, p, ci, k=k: f"={INC}*{X('MA_Target', k, c)}", opening=lambda c, p, ci, k=k: f"={INC}*{X('MA_Target', k, c)}", ann="end")
MP.gap()
MP.sec("Intercompany elimination (Northgate software contract with Vaultline)")
MP.add("ic_rev", "Eliminate intercompany revenue", "GBP k", "num", lambda c, p, ci: f"=-{INC}*inp_ma_ic_annual/12")
MP.add("ic_cost", "Eliminate intercompany cost (EBITDA effect nil)", "GBP k", "num", lambda c, p, ci: f"={INC}*inp_ma_ic_annual/12")
MP.gap()
MP.sec("Synergies at 100% of plan (before include switch and achievement %)")
MP.add("ramp", "Ramp factor", "x", "x2", lambda c, p, ci: f"=MIN(1,MAX(0,({c}$3-inp_ma_syn_start+1)/inp_ma_syn_ramp))", ann="end")
MP.add("xs_new100", "Cross-sell new ARR", "GBP k", "num", lambda c, p, ci: f"=inp_ma_xsell_logos_m*inp_ma_xsell_acv*{mp('ramp', c)}")
MP.add("xs_arr100", "Cross-sell ARR (closing)", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('MA_ProForma', 'xs_arr100')}+{mp('xs_new100', c)}",
       opening=lambda c, p, ci: "=0", ann="end")
MP.add("xs_rev100", "Cross-sell revenue", "GBP k", "num", lambda c, p, ci: f"=({p}{row_of('MA_ProForma', 'xs_arr100')}+{mp('xs_arr100', c)})/2/12")
MP.add("xs_ebitda100", "Cross-sell contribution", "GBP k", "num", lambda c, p, ci: f"={mp('xs_rev100', c)}*inp_ma_syn_margin")
MP.add("cs100", "Cost synergies", "GBP k", "num", lambda c, p, ci: f"=inp_ma_cost_syn/12*{mp('ramp', c)}")
MP.add("integ_raw", "Integration cost", "GBP k", "num", lambda c, p, ci: f"=IF({c}$3<=inp_ma_integration_months,inp_ma_integration_total/inp_ma_integration_months,0)")
MP.gap()
MP.sec("Adjustments in the pro-forma (include switch and achievement % applied)")
MP.add("syn_arr", "Synergy ARR (closing)", "GBP k", "num", lambda c, p, ci: f"={INC}*inp_ma_syn_achieved*{mp('xs_arr100', c)}",
       opening=lambda c, p, ci: "=0", ann="end")
MP.add("syn_rev", "Revenue synergies", "GBP k", "num", lambda c, p, ci: f"={INC}*inp_ma_syn_achieved*{mp('xs_rev100', c)}")
MP.add("syn_ebitda", "EBITDA from synergies (cross-sell contribution + cost savings)", "GBP k", "num",
       lambda c, p, ci: f"={INC}*inp_ma_syn_achieved*({mp('xs_ebitda100', c)}+{mp('cs100', c)})")
MP.add("integ", "Integration costs", "GBP k", "num", lambda c, p, ci: f"=-{INC}*{mp('integ_raw', c)}")
MP.add("txn", "Transaction costs", "GBP k", "num", lambda c, p, ci: f"=-{INC}*IF({c}$3=1,{PA('txn')},0)")
MP.add("hc_loss", "Revenue lost: acquired deferred revenue at fair value (IFRS 3)", "GBP k", "num",
       lambda c, p, ci: f"=-{INC}*IF({c}$3<=12,{PA('dr_adj')}/12,0)")
MP.add("amort", "Amortisation of acquired intangibles", "GBP k", "num",
       lambda c, p, ci: f"=-{INC}*({PA('cr')}/inp_ma_cr_life*IF({c}$3<=inp_ma_cr_life,1,0)+{PA('tech')}/inp_ma_tech_life*IF({c}$3<=inp_ma_tech_life,1,0))")
MP.add("dtl_credit", "Deferred tax credit on amortisation (non-cash)", "GBP k", "num", lambda c, p, ci: f"=-{mp('amort', c)}*inp_tax")
MP.add("loan_bal", "Acquisition loan balance (bullet)", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('MA_ProForma', 'loan_bal')}",
       opening=lambda c, p, ci: f"={INC}*{PA('loan')}", ann="end")
MP.add("int_loan", "Interest on acquisition loan (prior-period balance)", "GBP k", "num", lambda c, p, ci: f"=-{mp('loan_bal', p)}*inp_ma_loan_rate/12")
MP.add("int_cash", "Interest income lost on cash paid out (prior-period)", "GBP k", "num",
       lambda c, p, ci: f"=({mp('pf_cash', p)}-{mp('n_cash', p)}-{mp('v_cash', p)})*inp_interest/12")
MP.gap()
MP.sec("Pro-forma income and cash")
MP.add("pf_rev", "Pro-forma revenue", "GBP k", "num", lambda c, p, ci: f"={mp('n_rev', c)}+{mp('v_rev', c)}+{mp('syn_rev', c)}+{mp('ic_rev', c)}+{mp('hc_loss', c)}", bold=True)
MP.add("pf_ebitda", "Pro-forma EBITDA", "GBP k", "num",
       lambda c, p, ci: f"={mp('n_ebitda', c)}+{mp('v_ebitda', c)}+{mp('syn_ebitda', c)}+{mp('integ', c)}+{mp('txn', c)}+{mp('hc_loss', c)}", bold=True)
MP.add("pf_ni", "Pro-forma net income", "GBP k", "num",
       lambda c, p, ci: f"={mp('n_ni', c)}+{mp('v_ni', c)}+{mp('syn_ebitda', c)}+{mp('integ', c)}+{mp('txn', c)}+{mp('hc_loss', c)}+{mp('amort', c)}+{mp('dtl_credit', c)}+{mp('int_loan', c)}+{mp('int_cash', c)}", bold=True)
MP.add("adj_cash", "Cash effect of adjustments", "GBP k", "num", lambda c, p, ci: f"={mp('syn_ebitda', c)}+{mp('integ', c)}+{mp('txn', c)}+{mp('int_loan', c)}+{mp('int_cash', c)}")
MP.add("pf_fcf", "Pro-forma free cash flow", "GBP k", "num", lambda c, p, ci: f"={mp('n_fcf', c)}+{mp('v_fcf', c)}+{mp('adj_cash', c)}", bold=True)
MP.add("pf_cash", "Pro-forma closing cash", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('MA_ProForma', 'pf_cash')}+{mp('pf_fcf', c)}",
       opening=lambda c, p, ci: f"={mp('n_cash', c)}+{mp('v_cash', c)}-{INC}*{PA('price')}+{INC}*{PA('loan')}", ann="end", bold=True, top=True)
MP.gap()
MP.sec("Pro-forma balance sheet")
MP.add("pb_cash", "Cash", "GBP k", "num", lambda c, p, ci: f"={mp('pf_cash', c)}", opening=lambda c, p, ci: f"={mp('pf_cash', c)}", ann="end")
MP.add("pb_ar", "Trade receivables", "GBP k", "num", lambda c, p, ci: f"={mp('n_ar', c)}+{mp('v_ar', c)}", opening=lambda c, p, ci: f"={mp('n_ar', c)}+{mp('v_ar', c)}", ann="end")
MP.add("pb_ppe", "PP&E", "GBP k", "num", lambda c, p, ci: f"={mp('n_ppe', c)}+{mp('v_ppe', c)}", opening=lambda c, p, ci: f"={mp('n_ppe', c)}+{mp('v_ppe', c)}", ann="end")
MP.add("pb_cca", "Capitalised contract costs", "GBP k", "num", lambda c, p, ci: f"={mp('n_cca', c)}", opening=lambda c, p, ci: f"={mp('n_cca', c)}", ann="end")
MP.add("pb_int", "Acquired intangibles (net)", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('MA_ProForma', 'pb_int')}+{mp('amort', c)}",
       opening=lambda c, p, ci: f"={INC}*({PA('cr')}+{PA('tech')})", ann="end")
MP.add("pb_gw", "Goodwill", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('MA_ProForma', 'pb_gw')}", opening=lambda c, p, ci: f"={INC}*{PA('gw')}", ann="end")
MP.add("pb_ta", "Total assets", "GBP k", "num", lambda c, p, ci: f"=SUM({mp('pb_cash', c)}:{mp('pb_gw', c)})",
       opening=lambda c, p, ci: f"=SUM({mp('pb_cash', c)}:{mp('pb_gw', c)})", ann="end", bold=True, top=True)
MP.add("pb_ap", "Trade payables", "GBP k", "num", lambda c, p, ci: f"={mp('n_ap', c)}+{mp('v_ap', c)}", opening=lambda c, p, ci: f"={mp('n_ap', c)}+{mp('v_ap', c)}", ann="end")
MP.add("pb_dr_rem", "Memo: fair-value write-down of acquired deferred revenue still to unwind", "GBP k", "num",
       lambda c, p, ci: f"={p}{row_of('MA_ProForma', 'pb_dr_rem')}+{mp('hc_loss', c)}", opening=lambda c, p, ci: f"={INC}*{PA('dr_adj')}", ann="end")
MP.add("pb_dr", "Deferred revenue (net of write-down)", "GBP k", "num", lambda c, p, ci: f"={mp('n_dr', c)}+{mp('v_dr', c)}-{mp('pb_dr_rem', c)}",
       opening=lambda c, p, ci: f"={mp('n_dr', c)}+{mp('v_dr', c)}-{mp('pb_dr_rem', c)}", ann="end")
MP.add("pb_loan", "Acquisition loan", "GBP k", "num", lambda c, p, ci: f"={mp('loan_bal', c)}", opening=lambda c, p, ci: f"={mp('loan_bal', c)}", ann="end")
MP.add("pb_dtl", "Deferred tax liability", "GBP k", "num", lambda c, p, ci: f"={p}{row_of('MA_ProForma', 'pb_dtl')}-{mp('dtl_credit', c)}",
       opening=lambda c, p, ci: f"=-{INC}*{PA('dtl')}", ann="end")
MP.add("pb_tl", "Total liabilities", "GBP k", "num", lambda c, p, ci: f"={mp('pb_ap', c)}+{mp('pb_dr', c)}+{mp('pb_loan', c)}+{mp('pb_dtl', c)}",
       opening=lambda c, p, ci: f"={mp('pb_ap', c)}+{mp('pb_dr', c)}+{mp('pb_loan', c)}+{mp('pb_dtl', c)}", ann="end", bold=True, top=True)
MP.add("pb_eq", "Equity (Northgate opening equity + pro-forma net income)", "GBP k", "num",
       lambda c, p, ci: f"={p}{row_of('MA_ProForma', 'pb_eq')}+{mp('pf_ni', c)}", opening=lambda c, p, ci: f"={mp('n_eq', c)}", ann="end", bold=True)
MP.add("pb_check", "Balance check", "GBP k", "num2", lambda c, p, ci: f"={mp('pb_ta', c)}-{mp('pb_tl', c)}-{mp('pb_eq', c)}",
       opening=lambda c, p, ci: f"={mp('pb_ta', c)}-{mp('pb_tl', c)}-{mp('pb_eq', c)}", ann="end", bold=True)
MP.gap()
MP.sec("Pro-forma SaaS metrics")
MP.add("pf_arr", "Pro-forma closing ARR", "GBP k", "num", lambda c, p, ci: f"={mp('n_arr', c)}+{mp('v_arr', c)}+{mp('syn_arr', c)}",
       opening=lambda c, p, ci: f"={mp('n_arr', c)}+{mp('v_arr', c)}+{mp('syn_arr', c)}", ann="end", bold=True)
MP.add("pf_netnew", "Net new ARR", "GBP k", "num", lambda c, p, ci: f"={mp('pf_arr', c)}-{mp('pf_arr', p)}")
MP.add("pf_growth", "ARR growth, year on year", "%", "pct",
       lambda c, p, ci: f'=IF({c}$3<12,"",{mp("pf_arr", c)}/INDEX($N{row_of("MA_ProForma", "pf_arr")}:$AL{row_of("MA_ProForma", "pf_arr")},1,{c}$3-12+1)-1)', ann="end")
MP.add("pf_nrr", "Net revenue retention (annualised)", "%", "pct",
       lambda c, p, ci: f"=(1+({mp('n_arr_exp', c)}+{mp('n_arr_churn', c)}+{mp('v_arr_exp', c)}+{mp('v_arr_churn', c)})/({mp('n_arr_open', c)}+{mp('v_arr_open', c)}))^12", ann="end")
MP.add("pf_ebitda_m", "EBITDA margin", "%", "pct", lambda c, p, ci: f"={mp('pf_ebitda', c)}/{mp('pf_rev', c)}",
       ann=lambda c: f"={c}{row_of('MA_ProForma', 'pf_ebitda')}/{c}{row_of('MA_ProForma', 'pf_rev')}")
MP.add("pf_rule40", "Rule of 40", "%", "pct", lambda c, p, ci: f'=IF({c}$3<12,"",{mp("pf_growth", c)}+{mp("pf_ebitda_m", c)})', ann="end")
MP.add("pf_burn", "Net burn (positive = cash consumed)", "GBP k", "num", lambda c, p, ci: f"=-{mp('pf_fcf', c)}")
MP.add("pf_burn_mult", "Burn multiple", "x", "x", lambda c, p, ci: f'=IF({mp("pf_netnew", c)}<=0,"n/a",MAX(0,{mp("pf_burn", c)})/{mp("pf_netnew", c)})', ann="end")
MP.add("pf_runway", "Cash runway (months at trailing 3m average burn)", "months", "mo",
       lambda c, p, ci: f'=IF(AVERAGE({L(ci - 2)}{row_of("MA_ProForma", "pf_burn")}:{c}{row_of("MA_ProForma", "pf_burn")})<=0,"n/a",{mp("pf_cash", c)}/AVERAGE({L(ci - 2)}{row_of("MA_ProForma", "pf_burn")}:{c}{row_of("MA_ProForma", "pf_burn")}))', ann="end")
EXTRA_CHECKS.append(("Pro-forma balance sheet balances", "Largest absolute gap over 24 months (assets - liabilities - equity)",
                     f"=MAX(MAX({rng('MA_ProForma', 'pb_check', 'N')}),-MIN({rng('MA_ProForma', 'pb_check', 'N')}))", 0.01))
EXTRA_CHECKS.append(("Pro-forma cash never negative", "Lowest pro-forma cash balance is >= 0 (shortfall shown). Fails when the deal funding is too thin",
                     f"=MIN(0,MIN({rng('MA_ProForma', 'pf_cash', 'N')}))", 0))
EXTRA_CHECKS.append(("Intercompany elimination is EBITDA-neutral", "Eliminated revenue and cost net to zero over 24 months",
                     f"=SUM({rng('MA_ProForma', 'ic_rev')})+SUM({rng('MA_ProForma', 'ic_cost')})", 0.01))

# ----- Inputs_Comps: synthetic trading comparables -------------------------------------------------
ic = new_sheet("Inputs_Comps", "Inputs_Comps - illustrative trading comparables (synthetic, fictitious companies)",
               "Typed inputs. Not real market data: for demonstrating the method only.", {"A": 30, "B": 14, "C": 16, "D": 14})
hdr(ic, 4, ["Peer", "EV / ARR", "EV / revenue (NTM)", "ARR growth"])
PEERS = [("ArchivePoint plc", 5.8, 5.4, 0.18), ("ComplyStream Inc", 7.2, 6.8, 0.24), ("RecordKeep Software", 3.9, 3.6, 0.10),
         ("SurveilCore Ltd", 6.5, 6.0, 0.21), ("Datavault Group", 4.4, 4.1, 0.12), ("Regulyx", 5.1, 4.8, 0.15)]
for i, (nm, a, b, g) in enumerate(PEERS):
    r = 5 + i
    put(ic, f"A{r}", nm)
    put(ic, f"B{r}", a, "x", color=BLUE, fill=INP_FILL)
    put(ic, f"C{r}", b, "x", color=BLUE, fill=INP_FILL)
    put(ic, f"D{r}", g, "pct", color=BLUE, fill=INP_FILL)
CP0, CP1 = 5, 4 + len(PEERS)
for j, (lab, fn) in enumerate((("Minimum", "MIN"), ("Average", "AVERAGE"), ("Maximum", "MAX"))):
    r = CP1 + 2 + j
    put(ic, f"A{r}", lab, bold=True)
    for col, fm in (("B", "x"), ("C", "x"), ("D", "pct")):
        put(ic, f"{col}{r}", f"={fn}({col}{CP0}:{col}{CP1})", fm, True)
COMP = {"min": CP1 + 2, "med": CP1 + 3, "max": CP1 + 4}
put(ic, f"A{CP1 + 6}", "Vaultline's own growth is lower than most peers, which argues for the lower half of the range.", italic=True)

# ----- MA_Valuation: DCF, comps, offer, football field ---------------------------------------------
mv = new_sheet("MA_Valuation", "MA_Valuation - Vaultline: DCF, trading comps, offer range and football field",
               "GBP thousands unless stated. Years 1-2 from MA_Target; years 3-5 fade growth by inp_ma_fade with free cash flow margin held at year 2.",
               {"A": 60, "B": 9, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13})
hdr(mv, 4, ["Vaultline standalone DCF", "Unit", "", "", "", "", ""])
for c in "CDEFG":
    put(mv, f"{c}4", f'="FY"&(YEAR(inp_model_start)+{c}5-1)', None, True, "FFFFFF", HDR_FILL, align="right")
    put(mv, f"{c}5", "=COLUMN()-COLUMN($B$1)", "int", italic=True)
put(mv, "A5", "Year index", italic=True)
MV = {}
MVSPEC = []
_mr = [6]


def mvrow(key, label, unit, fmt, y12, y35, bold=False, top=False):
    r = _mr[0]
    _mr[0] += 1
    MV[key] = r
    MVSPEC.append((r, label, unit, fmt, y12, y35, bold, top))


def mvr(k, c):
    return f"{c}{MV[k]}"


mvrow("arr_open", "Opening ARR", "GBP k", "num", lambda c, a: f"={X('MA_Target', 'arr_open', a)}", lambda c, p: f"={mvr('arr_close', p)}")
mvrow("growth", "ARR growth", "%", "pct", lambda c, a: f"={mvr('arr_close', c)}/{mvr('arr_open', c)}-1", lambda c, p: f"={mvr('growth', p)}-inp_ma_fade")
mvrow("arr_close", "Closing ARR", "GBP k", "num", lambda c, a: f"={X('MA_Target', 'arr_close', a)}", lambda c, p: f"={mvr('arr_open', c)}*(1+{mvr('growth', c)})", True)
mvrow("rev", "Revenue", "GBP k", "num", lambda c, a: f"={X('MA_Target', 'rev', a)}", lambda c, p: f"={mvr('rev', p)}*{mvr('arr_close', c)}/{mvr('arr_close', p)}")
mvrow("ebitda", "EBITDA", "GBP k", "num", lambda c, a: f"={X('MA_Target', 'ebitda', a)}", lambda c, p: f"={mvr('rev', c)}*{mvr('ebitda', p)}/{mvr('rev', p)}")
mvrow("margin", "Unlevered FCF margin (held at year 2 from year 3)", "%", "pct", lambda c, a: f"={mvr('ufcf', c)}/{mvr('rev', c)}", lambda c, p: f"={mvr('margin', p)}")
mvrow("ufcf", "Unlevered free cash flow", "GBP k", "num", lambda c, a: f"={X('MA_Target', 'ufcf', a)}", lambda c, p: f"={mvr('rev', c)}*{mvr('margin', c)}", True, True)
for r, label, unit, fmt, y12, y35, bold, top in MVSPEC:
    put(mv, f"A{r}", label, bold=bold)
    put(mv, f"B{r}", unit, italic=True)
    for i, c in enumerate("CDEFG"):
        fx = y12(c, ("AN", "AO")[i]) if i < 2 else y35(c, "CDEFG"[i - 1])
        put(mv, f"{c}{r}", fx, fmt, bold)
        if top:
            mv[f"{c}{r}"].border = Border(top=THIN)
r0 = _mr[0] + 1
secrow(mv, r0, "DCF cases (low = WACC +1pt and growth -0.5pt; high = WACC -1pt and growth +0.5pt)", 7)
hdr(mv, r0 + 1, ["", "Unit", "Low", "Base", "High"])
MV["wacc"], MV["tg"] = r0 + 2, r0 + 3
put(mv, f"A{r0 + 2}", "Discount rate (WACC)")
put(mv, f"A{r0 + 3}", "Terminal growth")
for col, dw, dg in (("C", "+inp_ma_wacc_band", "-inp_ma_tg_band"), ("D", "", ""), ("E", "-inp_ma_wacc_band", "+inp_ma_tg_band")):
    put(mv, f"{col}{r0 + 2}", f"=inp_ma_wacc{dw}", "pct")
    put(mv, f"{col}{r0 + 3}", f"=inp_ma_tg{dg}", "pct")
for k, (lab, col) in enumerate((("Low", "C"), ("Base", "D"), ("High", "E"))):
    r = r0 + 4 + k
    MV[f"df_{lab.lower()}"] = r
    put(mv, f"A{r}", f"Discount factor, {lab.lower()} case")
    for c in "CDEFG":
        put(mv, f"{c}{r}", f"=1/(1+${col}${r0 + 2})^{c}$5", "x2")
for lab, k, fn in (("Present value of years 1-5 cash flows", "pv", None), ("Terminal value (Gordon growth on year 5)", "tv", None),
                   ("Present value of terminal value", "pvtv", None), ("Enterprise value (DCF)", "ev", None)):
    r = r0 + 7 + ["pv", "tv", "pvtv", "ev"].index(k)
    MV["c_" + k] = r
    put(mv, f"A{r}", lab, bold=(k == "ev"))
    for k2, (name, col) in enumerate((("low", "C"), ("base", "D"), ("high", "E"))):
        dfr = MV[f"df_{name}"]
        if k == "pv":
            fx = f"=SUMPRODUCT($C${MV['ufcf']}:$G${MV['ufcf']},$C${dfr}:$G${dfr})"
        elif k == "tv":
            fx = f"=$G${MV['ufcf']}*(1+{col}{r0 + 3})/({col}{r0 + 2}-{col}{r0 + 3})"
        elif k == "pvtv":
            fx = f"={col}{r0 + 8}*$G${dfr}"
        else:
            fx = f"={col}{r0 + 7}+{col}{r0 + 9}"
        put(mv, f"{col}{r}", fx, "num", k == "ev")
MV["tv_pct"] = r0 + 11
put(mv, f"A{r0 + 11}", "Terminal value as % of EV (base)")
put(mv, f"D{r0 + 11}", f"=D{r0 + 9}/D{r0 + 10}", "pct")
DCF_EV = {"low": f"MA_Valuation!$C${MV['c_ev']}", "base": f"MA_Valuation!$D${MV['c_ev']}", "high": f"MA_Valuation!$E${MV['c_ev']}"}
_mr[0] = r0 + 13

# ----- MA_Returns: does the deal create value? ---------------------------------------------------
rt_ = new_sheet("MA_Returns", "MA_Returns - value creation, payback, ARR accretion, cash runway impact",
                "GBP thousands. Value created = standalone DCF + present value of net synergies - enterprise value paid - transaction costs - present value of integration costs.",
                {"A": 62, "B": 9, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14})
hdr(rt_, 4, ["Synergy and deal cash flows (full plan = 100% achieved)", "Unit", "", "", "", "", ""])
for c in "CDEFG":
    put(rt_, f"{c}4", f'="FY"&(YEAR(inp_model_start)+{c}5-1)', None, True, "FFFFFF", HDR_FILL, align="right")
    put(rt_, f"{c}5", "=COLUMN()-COLUMN($B$1)", "int", italic=True)
put(rt_, "A5", "Year index", italic=True)
MR = {}
MRSPEC = []
_rr = [6]


def mrrow(key, label, unit, fmt, y12, y35, bold=False, top=False):
    r = _rr[0]
    _rr[0] += 1
    MR[key] = r
    MRSPEC.append((r, label, unit, fmt, y12, y35, bold, top))


def mrr(k, c):
    return f"{c}{MR[k]}"


mrrow("xs_arr100", "Cross-sell ARR (closing)", "GBP k", "num", lambda c, a: f"={X('MA_ProForma', 'xs_arr100', a)}", lambda c, p: f"={mrr('xs_arr100', p)}+($D{MR['xs_arr100']}-$C{MR['xs_arr100']})")
mrrow("xs_rev100", "Cross-sell revenue", "GBP k", "num", lambda c, a: f"={X('MA_ProForma', 'xs_rev100', a)}", lambda c, p: f"=({mrr('xs_arr100', p)}+{mrr('xs_arr100', c)})/2")
mrrow("xs_e100", "Cross-sell contribution", "GBP k", "num", lambda c, a: f"={mrr('xs_rev100', c)}*inp_ma_syn_margin", lambda c, p: f"={mrr('xs_rev100', c)}*inp_ma_syn_margin")
mrrow("cs100", "Cost synergies", "GBP k", "num", lambda c, a: f"={X('MA_ProForma', 'cs100', a)}", lambda c, p: "=inp_ma_cost_syn")
mrrow("net100", "Net synergies after tax", "GBP k", "num", lambda c, a: f"=({mrr('xs_e100', c)}+{mrr('cs100', c)})*(1-inp_tax)", lambda c, p: f"=({mrr('xs_e100', c)}+{mrr('cs100', c)})*(1-inp_tax)", True, True)
mrrow("integ", "Integration costs after tax", "GBP k", "num", lambda c, a: f"={X('MA_ProForma', 'integ_raw', a)}*(1-inp_tax)", lambda c, p: "=0")
mrrow("ufcf", "Vaultline standalone free cash flow", "GBP k", "num", lambda c, a: f"=MA_Valuation!{c}{MV['ufcf']}", lambda c, p: f"=MA_Valuation!{c}{MV['ufcf']}")
mrrow("inc_cf", "Incremental cash flow to Northgate (at achieved synergies)", "GBP k", "num",
      lambda c, a: f"={mrr('ufcf', c)}+inp_ma_syn_achieved*{mrr('net100', c)}-{mrr('integ', c)}",
      lambda c, p: f"={mrr('ufcf', c)}+inp_ma_syn_achieved*{mrr('net100', c)}-{mrr('integ', c)}", True, True)
mrrow("cum", "Cumulative incremental cash flow", "GBP k", "num", lambda c, a: f"=SUM($C{MR['inc_cf']}:{c}{MR['inc_cf']})", lambda c, p: f"=SUM($C{MR['inc_cf']}:{c}{MR['inc_cf']})")
mrrow("cum_prev", "Cumulative, prior year", "GBP k", "num", lambda c, a: "=0" if c == "C" else f"={mrr('cum', 'C')}", lambda c, p: f"={mrr('cum', p)}")
mrrow("df", "Discount factor (base WACC)", "x", "x2", lambda c, a: f"=1/(1+inp_ma_wacc)^{c}$5", lambda c, p: f"=1/(1+inp_ma_wacc)^{c}$5")
for r, label, unit, fmt, y12, y35, bold, top in MRSPEC:
    put(rt_, f"A{r}", label, bold=bold)
    put(rt_, f"B{r}", unit, italic=True)
    for i, c in enumerate("CDEFG"):
        fx = y12(c, ("AN", "AO")[i]) if i < 2 else y35(c, "CDEFG"[i - 1])
        put(rt_, f"{c}{r}", fx, fmt, bold)
        if top:
            rt_[f"{c}{r}"].border = Border(top=THIN)
_rr[0] += 1
MRS = {}


def mrs(key, label, fx, fmt="num", bold=False, note=""):
    r = _rr[0]
    _rr[0] += 1
    MRS[key] = r
    put(rt_, f"A{r}", label, bold=bold)
    put(rt_, f"C{r}", fx, fmt, bold, align="right" if fmt is None else None)
    if note:
        put(rt_, f"D{r}", note, italic=True)


def msec(label):
    secrow(rt_, _rr[0], label, 7)
    _rr[0] += 1


def ms(k):
    return f"C{MRS[k]}"


msec("A. Value creation (net present value at the base WACC)")
mrs("dcf", "Vaultline standalone value (DCF enterprise value, base case)", f"={DCF_EV['base']}")
mrs("pv_syn100", "Present value of net synergies at 100% of plan (incl. terminal value)",
    f"=SUMPRODUCT(C{MR['net100']}:G{MR['net100']},C{MR['df']}:G{MR['df']})+G{MR['net100']}*(1+inp_ma_tg)/(inp_ma_wacc-inp_ma_tg)*G{MR['df']}")
mrs("pv_syn", "Present value of synergies at the achieved %", f"=inp_ma_syn_achieved*{ms('pv_syn100')}")
mrs("pv_integ", "Present value of integration costs (after tax)", f"=SUMPRODUCT(C{MR['integ']}:G{MR['integ']},C{MR['df']}:G{MR['df']})")
mrs("ev_paid", "Enterprise value paid", f"={PA('ev')}")
mrs("txn", "Transaction costs", f"={PA('txn')}")
mrs("value", "Value created (NPV)", f"={ms('dcf')}+{ms('pv_syn')}-{ms('pv_integ')}-{ms('ev_paid')}-{ms('txn')}", bold=True)
mrs("verdict", "Verdict", f'=IF({ms("value")}>0,"Creates value","Destroys value")', None, True)
mrs("premium", "Premium paid over standalone DCF value", f"={ms('ev_paid')}-{ms('dcf')}", note="The price of control: synergies must justify at least this")
mrs("breakeven", "Synergies needed to break even (% of plan)", f"=({ms('ev_paid')}+{ms('txn')}+{ms('pv_integ')}-{ms('dcf')})/{ms('pv_syn100')}", "pct", True,
    "Compare with synergies achieved (inp_ma_syn_achieved)")
_rr[0] += 1
msec("B. Payback (undiscounted, 5 years)")
mrs("hurdle", "Cash to recover: enterprise value + transaction costs", f"={ms('ev_paid')}+{ms('txn')}")
mrs("n", "First year in which cumulative cash flow reaches the hurdle", f'=COUNTIF(C{MR["cum"]}:G{MR["cum"]},"<"&{ms("hurdle")})+1', "int")
mrs("payback", "Payback (years)", f'=IF({ms("n")}>COUNT(C5:G5),">5 years",{ms("n")}-1+({ms("hurdle")}-INDEX(C{MR["cum_prev"]}:G{MR["cum_prev"]},1,{ms("n")}))/INDEX(C{MR["inc_cf"]}:G{MR["inc_cf"]},1,{ms("n")}))', "num1", True)
_rr[0] += 1
msec("C. ARR accretion (live pro-forma vs Northgate standalone)")
mrs("arr_sa", "Northgate standalone ARR, month 24", f"={X('SaaS_Metrics', 'arr_close', 'AL')}")
mrs("arr_pf", "Pro-forma ARR, month 24", f"={X('MA_ProForma', 'pf_arr', 'AL')}")
mrs("arr_acc", "ARR accretion", f"={ms('arr_pf')}/{ms('arr_sa')}-1", "pct", True)
mrs("g_sa", "Northgate standalone ARR growth, month 12", f"={X('SaaS_Metrics', 'growth', 'Z')}", "pct")
mrs("g_pf", "Pro-forma ARR growth, month 12", f"={X('MA_ProForma', 'pf_growth', 'Z')}", "pct")
mrs("nrr_sa", "Northgate standalone NRR, month 12", f"={X('SaaS_Metrics', 'nrr', 'Z')}", "pct")
mrs("nrr_pf", "Pro-forma NRR, month 12", f"={X('MA_ProForma', 'pf_nrr', 'Z')}", "pct")
mrs("r40_sa", "Northgate standalone Rule of 40, month 24", f"={X('SaaS_Metrics', 'rule40', 'AL')}", "pct")
mrs("r40_pf", "Pro-forma Rule of 40, month 24", f"={X('MA_ProForma', 'pf_rule40', 'AL')}", "pct")
_rr[0] += 1
msec("D. Cash runway impact")
mrs("cashpaid", "Northgate cash used at closing", f"={PA('cash_paid')}")
mrs("min_sa", "Lowest month-end cash, standalone", "=Sensitivity!B4")
mrs("min_pf", "Lowest month-end cash, pro-forma", f"=MIN({rng('MA_ProForma', 'pf_cash', 'N')})", bold=True)
mrs("cash_sa", "Cash, month 24, standalone", f"={X('BS', 'cash', 'AL')}")
mrs("cash_pf", "Cash, month 24, pro-forma", f"={X('MA_ProForma', 'pf_cash', 'AL')}")
mrs("run_pf", "Pro-forma cash runway at month 24 (months)", f"={X('MA_ProForma', 'pf_runway', 'AL')}", "mo")
put(rt_, f"A{_rr[0]}", "Thin-headroom threshold (lowest cash as % of cash used at closing)")
put(rt_, f"C{_rr[0]}", 0.25, "pct", color=BLUE, fill=INP_FILL)
HEADROOM = f"C{_rr[0]}"
_rr[0] += 1
mrs("cash_flag", "Funding flag", f'=IF({ms("min_pf")}<0,"FLAG: cash goes negative",IF({ms("min_pf")}<{ms("cashpaid")}*{HEADROOM},"WATCH: thin headroom","OK"))', None, True)
_rr[0] += 1
MA_METRICS = [
    ("ma_value", "Deal value created (NPV)", "num", f"MA_Returns!C{MRS['value']}"),
    ("pf_arr24", "Pro-forma ARR, month 24", "num", f"MA_ProForma!AL{row_of('MA_ProForma', 'pf_arr')}"),
    ("pf_cash24", "Pro-forma cash, month 24", "num", f"MA_ProForma!AL{row_of('MA_ProForma', 'pf_cash')}"),
    ("pf_mincash", "Pro-forma lowest month-end cash", "num", f"MA_Returns!C{MRS['min_pf']}"),
]
msec("E. Deal outcome by scenario (engine-run snapshots; the Downside case lowers Vaultline growth, raises its churn and cuts synergies)")
hdr(rt_, _rr[0], ["Metric", "", "Base", "Upside", "Downside"])
_rr[0] += 1
MR_SNAP_START = _rr[0]
for mid, lab, fmt, cell in MA_METRICS:
    r = _rr[0]
    _rr[0] += 1
    put(rt_, f"A{r}", lab)
    for j, sc in enumerate(("Base", "Upside", "Downside")):
        static(rt_, f"{L(3 + j)}{r}", sn("scen", sc, mid), fmt)
    if mid == "ma_value":
        VAL_ROW = r
put(rt_, f"A{_rr[0]}", "Verdict", bold=True)
for j in range(3):
    c = L(3 + j)
    put(rt_, f"{c}{_rr[0]}", f'=IF({c}{VAL_ROW}>0,"Creates value","Destroys value")', None, True, align="right")
rt_.freeze_panes = "A6"
EXTRA_CHECKS.append(("Cumulative cash flow row ties", "Year-5 cumulative incremental cash flow = sum of the five annual flows",
                     f"=MA_Returns!G{MR['cum']}-SUM(MA_Returns!C{MR['inc_cf']}:G{MR['inc_cf']})", 0.01))

# Valuation tail: comps, offer, football field (needs the returns cells above)
v0 = _mr[0]
secrow(mv, v0, "Trading comps applied to Vaultline (Inputs_Comps)", 7)
hdr(mv, v0 + 1, ["", "Unit", "Low (min)", "Mid (average)", "High (max)"])
put(mv, f"A{v0 + 2}", "EV / ARR multiple")
put(mv, f"A{v0 + 3}", "Implied EV from ARR")
put(mv, f"A{v0 + 4}", "EV / revenue multiple (NTM)")
put(mv, f"A{v0 + 5}", "Implied EV from FY1 revenue")
for col, key in (("C", "min"), ("D", "med"), ("E", "max")):
    put(mv, f"{col}{v0 + 2}", f"=Inputs_Comps!B{COMP[key]}", "x")
    put(mv, f"{col}{v0 + 3}", f"={col}{v0 + 2}*{PA('arr')}", "num")
    put(mv, f"{col}{v0 + 4}", f"=Inputs_Comps!C{COMP[key]}", "x")
    put(mv, f"{col}{v0 + 5}", f"={col}{v0 + 4}*C{MV['rev']}", "num")
o0 = v0 + 7
secrow(mv, o0, "Offer", 7)
hdr(mv, o0 + 1, ["", "Unit", "Low", "Offer", "High"])
put(mv, f"A{o0 + 2}", "EV / ARR (offer -0.5x / offer / +0.5x)")
put(mv, f"A{o0 + 3}", "Enterprise value")
put(mv, f"A{o0 + 4}", "Equity purchase price (EV + cash acquired)")
put(mv, f"A{o0 + 5}", "Implied EV / FY1 revenue")
put(mv, f"A{o0 + 6}", "Implied EV / FY1 EBITDA")
for col, dm in (("C", "-0.5"), ("D", ""), ("E", "+0.5")):
    put(mv, f"{col}{o0 + 2}", f"=inp_ma_offer_mult{dm}", "x")
    put(mv, f"{col}{o0 + 3}", f"={col}{o0 + 2}*{PA('arr')}", "num")
    put(mv, f"{col}{o0 + 4}", f"={col}{o0 + 3}+{PA('cash_acq')}", "num")
    put(mv, f"{col}{o0 + 5}", f"={col}{o0 + 3}/$C${MV['rev']}", "x")
    put(mv, f"{col}{o0 + 6}", f"={col}{o0 + 3}/$C${MV['ebitda']}", "x")
f0 = o0 + 8
secrow(mv, f0, "Football field: enterprise value (GBP m)", 7)
hdr(mv, f0 + 1, ["Method", "", "Low", "Mid", "High", "Spread", ""])
FF = [
    ("DCF (WACC +/-1pt, growth +/-0.5pt)", f"={DCF_EV['low']}/1000", f"={DCF_EV['base']}/1000", f"={DCF_EV['high']}/1000"),
    ("DCF plus net synergies (50% / achieved / 100%)", f"=({DCF_EV['low']}+0.5*MA_Returns!{ms('pv_syn100')}-MA_Returns!{ms('pv_integ')})/1000",
     f"=({DCF_EV['base']}+MA_Returns!{ms('pv_syn')}-MA_Returns!{ms('pv_integ')})/1000", f"=({DCF_EV['high']}+MA_Returns!{ms('pv_syn100')}-MA_Returns!{ms('pv_integ')})/1000"),
    ("Trading comps: EV / ARR", f"=C{v0 + 3}/1000", f"=D{v0 + 3}/1000", f"=E{v0 + 3}/1000"),
    ("Trading comps: EV / FY1 revenue", f"=C{v0 + 5}/1000", f"=D{v0 + 5}/1000", f"=E{v0 + 5}/1000"),
    ("Offer range (+/-0.5x ARR)", f"=C{o0 + 3}/1000", f"=D{o0 + 3}/1000", f"=E{o0 + 3}/1000"),
]
for i, (lab, lo, mid_, hi) in enumerate(FF):
    r = f0 + 2 + i
    put(mv, f"A{r}", lab, bold=(i == len(FF) - 1))
    put(mv, f"C{r}", lo, "num1")
    put(mv, f"D{r}", mid_, "num1")
    put(mv, f"E{r}", hi, "num1")
    put(mv, f"F{r}", f"=E{r}-C{r}", "num1")
from openpyxl.chart import BarChart, Reference
ch = BarChart()
ch.type = "bar"
ch.grouping = "stacked"
ch.overlap = 100
ch.title = "Football field: Vaultline enterprise value (GBP m)"
ch.add_data(Reference(mv, min_col=3, min_row=f0 + 1, max_row=f0 + 1 + len(FF)), titles_from_data=True)
ch.add_data(Reference(mv, min_col=6, min_row=f0 + 1, max_row=f0 + 1 + len(FF)), titles_from_data=True)
ch.set_categories(Reference(mv, min_col=1, min_row=f0 + 2, max_row=f0 + 1 + len(FF)))
ch.series[0].graphicalProperties.noFill = True
ch.series[0].graphicalProperties.line.noFill = True
ch.legend = None
ch.height, ch.width = 8, 22
mv.add_chart(ch, f"A{f0 + 9}")
mv.freeze_panes = "A6"

# ----- LT_Model: years 3-5 driver-based extension --------------------------------------------
lt = new_sheet("LT_Model", "LT_Model - five-year view (years 1-2 from the monthly model, years 3-5 driver-based)",
               "GBP thousands. Years 1-2 are links to the monthly tabs; years 3-5 use the inp_lt_* drivers. Simplified cash view: no interest income, no full balance sheet.",
               {"A": 52, "B": 9, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12})
hdr(lt, 4, ["Year", "Unit", "", "", "", "", ""])
for i, c in enumerate("CDEFG", start=1):
    put(lt, f"{c}4", f'="FY"&(YEAR(inp_model_start)+{c}5-1)', None, True, "FFFFFF", HDR_FILL, align="right")
    put(lt, f"{c}5", "=COLUMN()-COLUMN($B$1)", "int", italic=True)
put(lt, "A5", "Year index", italic=True)
LT = {}
_ltr = [7]


LTSPEC = []


def ltrow(key, label, unit, fmt, y12, y35, bold=False, top=False):
    r = _ltr[0]
    _ltr[0] += 1
    LT[key] = r
    LTSPEC.append((r, label, unit, fmt, y12, y35, bold, top))


def ltwrite():
    for r, label, unit, fmt, y12, y35, bold, top in LTSPEC:
        put(lt, f"A{r}", label, bold=bold)
        put(lt, f"B{r}", unit, italic=True)
        for i, c in enumerate("CDEFG"):
            p = "CDEFG"[i - 1] if i else None
            ann = ("AN", "AO")[i] if i < 2 else None
            put(lt, f"{c}{r}", y12(c, ann) if i < 2 else y35(c, p), fmt, bold)
            if top:
                lt[f"{c}{r}"].border = Border(top=THIN)


def ltsec(label):
    secrow(lt, _ltr[0], label, 7)
    _ltr[0] += 1


def lr(k, c):
    return f"{c}{LT[k]}"


ltsec("ARR")
ltrow("arr_open", "Opening ARR", "GBP k", "num", lambda c, a: f"={X('Rev_Build', 'arr_open', a)}", lambda c, p: f"={lr('arr_close', p)}")
ltrow("arr_new", "+ New logos", "GBP k", "num", lambda c, a: f"={X('Rev_Build', 'arr_new', a)}", lambda c, p: f"={lr('arr_new', p)}*(1+inp_lt_newarr_growth)")
ltrow("arr_exp", "+ Expansion", "GBP k", "num", lambda c, a: f"={X('Rev_Build', 'arr_exp', a)}", lambda c, p: f"={lr('arr_open', c)}*{lr('exp_rate', c)}")
ltrow("arr_churn", "- Churn", "GBP k", "num", lambda c, a: f"={X('Rev_Build', 'arr_churn', a)}", lambda c, p: f"=-{lr('arr_open', c)}*{lr('churn_rate', c)}")
ltrow("arr_close", "Closing ARR", "GBP k", "num", lambda c, a: f"={X('Rev_Build', 'arr_close', a)}", lambda c, p: f"=SUM({lr('arr_open', c)}:{lr('arr_churn', c)})", True, True)
ltrow("exp_rate", "Expansion rate (% of opening ARR)", "%", "pct", lambda c, a: f"={lr('arr_exp', c)}/{lr('arr_open', c)}", lambda c, p: f"={lr('exp_rate', p)}+inp_lt_nrr_step")
ltrow("churn_rate", "Gross churn rate (% of opening ARR)", "%", "pct", lambda c, a: f"=-{lr('arr_churn', c)}/{lr('arr_open', c)}", lambda c, p: f"={lr('churn_rate', p)}")
ltrow("growth", "ARR growth", "%", "pct", lambda c, a: f"={lr('arr_close', c)}/{lr('arr_open', c)}-1", lambda c, p: f"={lr('arr_close', c)}/{lr('arr_open', c)}-1")
ltrow("nrr", "Net revenue retention", "%", "pct", lambda c, a: f"=1+({lr('arr_exp', c)}+{lr('arr_churn', c)})/{lr('arr_open', c)}", lambda c, p: f"=1+({lr('arr_exp', c)}+{lr('arr_churn', c)})/{lr('arr_open', c)}")
_ltr[0] += 1
ltsec("Income statement")
ltrow("rev", "Revenue", "GBP k", "num", lambda c, a: f"={X('IS', 'rev_tot', a)}", lambda c, p: f"=({lr('arr_open', c)}+{lr('arr_close', c)})/2*{lr('rev_conv', c)}", True)
ltrow("rev_conv", "Revenue / average ARR", "x", "x2", lambda c, a: f"={lr('rev', c)}/(({lr('arr_open', c)}+{lr('arr_close', c)})/2)", lambda c, p: f"={lr('rev_conv', p)}")
ltrow("gm", "Gross margin", "%", "pct", lambda c, a: f"={X('IS', 'gm', a)}", lambda c, p: f"={lr('gm', p)}+inp_lt_gm_step")
ltrow("gp", "Gross profit", "GBP k", "num", lambda c, a: f"={X('IS', 'gp', a)}", lambda c, p: f"={lr('rev', c)}*{lr('gm', c)}")
for k, nm, step in (("rd", "R&D", "inp_lt_rd_step"), ("sm", "S&M", "inp_lt_sm_step"), ("ga", "G&A", "inp_lt_ga_step")):
    ltrow(f"{k}_pct", f"{nm} as % of revenue", "%", "pct", lambda c, a, k=k: f"={lr(k, c)}/{lr('rev', c)}", lambda c, p, k=k, step=step: f"={lr(k + '_pct', p)}-{step}")
    ltrow(k, f"{nm} expense", "GBP k", "num", lambda c, a, k=k: f"={X('IS', 'o_' + k, a)}", lambda c, p, k=k: f"={lr('rev', c)}*{lr(k + '_pct', c)}")
ltrow("ebitda", "EBITDA", "GBP k", "num", lambda c, a: f"={lr('gp', c)}-{lr('rd', c)}-{lr('sm', c)}-{lr('ga', c)}", lambda c, p: f"={lr('gp', c)}-{lr('rd', c)}-{lr('sm', c)}-{lr('ga', c)}", True, True)
ltrow("ebitda_m", "EBITDA margin", "%", "pct", lambda c, a: f"={lr('ebitda', c)}/{lr('rev', c)}", lambda c, p: f"={lr('ebitda', c)}/{lr('rev', c)}")
ltrow("da_pct", "Depreciation as % of revenue", "%", "pct", lambda c, a: f"={lr('da', c)}/{lr('rev', c)}", lambda c, p: f"={lr('da_pct', p)}")
ltrow("da", "Depreciation", "GBP k", "num", lambda c, a: f"={X('IS', 'da', a)}", lambda c, p: f"={lr('rev', c)}*{lr('da_pct', c)}")
ltrow("tax", "Tax (on EBITDA less depreciation, if positive)", "GBP k", "num", lambda c, a: f"={X('IS', 'tax', a)}", lambda c, p: f"=MAX(0,{lr('ebitda', c)}-{lr('da', c)})*inp_tax")
_ltr[0] += 1
ltsec("Cash (simplified)")
ltrow("capex", "Capex", "GBP k", "num", lambda c, a: f"={X('BS', 'capex', a)}", lambda c, p: f"={lr('rev', c)}*inp_capex_pct")
ltrow("def_pct", "Deferred revenue as % of closing ARR", "%", "pct", lambda c, a: f"={lr('def_close', c)}/{lr('arr_close', c)}", lambda c, p: f"={lr('def_pct', p)}")
ltrow("def_close", "Deferred revenue (closing)", "GBP k", "num", lambda c, a: f"={X('BS', 'defrev', a)}", lambda c, p: f"={lr('arr_close', c)}*{lr('def_pct', c)}")
ltrow("ar_pct", "Receivables as % of closing ARR", "%", "pct", lambda c, a: f"={lr('ar_close', c)}/{lr('arr_close', c)}", lambda c, p: f"={lr('ar_pct', p)}")
ltrow("ar_close", "Receivables (closing)", "GBP k", "num", lambda c, a: f"={X('BS', 'ar', a)}", lambda c, p: f"={lr('arr_close', c)}*{lr('ar_pct', c)}")
ltrow("cca_pct", "Contract-cost asset as % of closing ARR", "%", "pct", lambda c, a: f"={lr('cca_close', c)}/{lr('arr_close', c)}", lambda c, p: f"={lr('cca_pct', p)}")
ltrow("cca_close", "Capitalised contract costs (closing)", "GBP k", "num", lambda c, a: f"={X('BS', 'cca', a)}", lambda c, p: f"={lr('arr_close', c)}*{lr('cca_pct', c)}")
_prev_open = {"C": None, "D": "C", "E": "D", "F": "E", "G": "F"}
ltrow("fcf", "Free cash flow", "GBP k", "num", lambda c, a: f"={X('CF', 'net', a)}",
      lambda c, p: f"={lr('ebitda', c)}-{lr('tax', c)}-{lr('capex', c)}+({lr('def_close', c)}-{lr('def_close', p)})-({lr('ar_close', c)}-{lr('ar_close', p)})-({lr('cca_close', c)}-{lr('cca_close', p)})", True, True)
ltrow("cash_close", "Closing cash", "GBP k", "num", lambda c, a: f"={X('CF', 'cash_close', a)}", lambda c, p: f"={lr('cash_close', p)}+{lr('fcf', c)}", True)
_ltr[0] += 1
ltsec("Headline metrics")
ltrow("burn_mult", "Burn multiple", "x", "x", lambda c, a: f'=IF(({lr("arr_close", c)}-{lr("arr_open", c)})<=0,"n/a",MAX(0,-{lr("fcf", c)})/({lr("arr_close", c)}-{lr("arr_open", c)}))',
      lambda c, p: f'=IF(({lr("arr_close", c)}-{lr("arr_open", c)})<=0,"n/a",MAX(0,-{lr("fcf", c)})/({lr("arr_close", c)}-{lr("arr_open", c)}))')
ltrow("rule40", "Rule of 40 (ARR growth + EBITDA margin)", "%", "pct", lambda c, a: f"={lr('growth', c)}+{lr('ebitda_m', c)}", lambda c, p: f"={lr('growth', c)}+{lr('ebitda_m', c)}")
ltwrite()
lt.freeze_panes = "C6"
EXTRA_CHECKS.append(("Long-term view ties to monthly model", "Year 2 closing ARR and cash on LT_Model = monthly model (sum of absolute gaps)",
                     f"=ABS(LT_Model!D{LT['arr_close']}-{X('Rev_Build', 'arr_close', 'AL')})+ABS(LT_Model!D{LT['cash_close']}-{X('BS', 'cash', 'AL')})+ABS(LT_Model!D{LT['rev']}-{X('IS', 'rev_tot', 'AO')})", 0.01))
EXTRA_CHECKS.append(("Long-term ARR roll-forward", "Year 5 closing ARR = year 1 opening + all new + expansion + churn",
                     f"=LT_Model!G{LT['arr_close']}-(LT_Model!C{LT['arr_open']}+SUM(LT_Model!C{LT['arr_new']}:G{LT['arr_new']})+SUM(LT_Model!C{LT['arr_exp']}:G{LT['arr_exp']})+SUM(LT_Model!C{LT['arr_churn']}:G{LT['arr_churn']}))", 0.01))

# ----- Inputs_Budget: a board-approved budget, deliberately in a different layout ---------------
ib = new_sheet("Inputs_Budget", "Inputs_Budget - board-approved FY1 budget, as received",
               "Board-pack layout: GBP millions, quarters across, costs shown as negatives, different line names from the model. Mapped on Budget_vs_Forecast. Totals are the only formulas.",
               {"A": 40, "B": 11, "C": 11, "D": 11, "E": 11, "F": 11})
hdr(ib, 4, ["GBP m", "Q1", "Q2", "Q3", "Q4", "FY"])
BUD = {
    5: ("Subscription revenue", [10.70, 11.55, 12.45, 13.35]),
    6: ("Professional services revenue", [0.25, 0.30, 0.30, 0.30]),
    8: ("Cost of sales", [-2.55, -2.75, -2.90, -3.05]),
    10: ("Sales & marketing", [-4.80, -4.95, -5.10, -5.25]),
    11: ("Research & development", [-3.75, -3.90, -4.05, -4.20]),
    12: ("General & administrative", [-1.85, -1.90, -1.95, -2.00]),
}
for r, (lab, vals) in BUD.items():
    put(ib, f"A{r}", lab)
    for c, v in zip("BCDE", vals):
        put(ib, f"{c}{r}", v, "#,##0.00;(#,##0.00)", color=BLUE, fill=INP_FILL)
for r, lab, fx in ((7, "Total revenue", "={c}5+{c}6"), (9, "Gross profit", "={c}7+{c}8"),
                   (13, "Total operating costs", "=SUM({c}10:{c}12)"), (14, "EBITDA", "={c}9+{c}13")):
    put(ib, f"A{r}", lab, bold=True)
    for c in "BCDE":
        put(ib, f"{c}{r}", fx.format(c=c), "#,##0.00;(#,##0.00)", True)
for r in range(5, 15):
    put(ib, f"F{r}", f"=SUM(B{r}:E{r})", "#,##0.00;(#,##0.00)", r in (7, 9, 13, 14))
secrow(ib, 16, "Operating KPIs, full year (GBP m unless stated)", 6)
KPI = {17: ("Opening ARR", 40.00), 18: ("New-customer ARR", 12.93), 19: ("Expansion ARR", 4.50), 20: ("Churned ARR", -2.90),
       22: ("New customers (number)", 226), 23: ("Closing headcount (FTE)", 485)}
for r, (lab, v) in KPI.items():
    put(ib, f"A{r}", lab)
    put(ib, f"B{r}", v, "#,##0.00;(#,##0.00)" if r < 22 else "#,##0", color=BLUE, fill=INP_FILL)
put(ib, "A21", "Closing ARR", bold=True)
put(ib, "B21", "=SUM(B17:B20)", "#,##0.00;(#,##0.00)", True)

# ----- Budget_vs_Forecast ---------------------------------------------------------------------
bv = new_sheet("Budget_vs_Forecast", "Budget_vs_Forecast - FY1 budget mapped to the model, with variance bridges",
               "GBP thousands. Budget (GBP m, costs negative) x 1,000 vs model forecast on the same sign basis. Variance = forecast - budget, so a positive number is favourable on every line.",
               {"A": 44, "B": 13, "C": 13, "D": 13, "E": 11, "F": 13, "G": 70})
hdr(bv, 4, ["FY1 P&L", "Budget", "Forecast", "Variance", "Var %", "Flag", "Commentary (owner)"])
put(bv, "A3", "Treat a variance within this % of budget as 'On budget'", bold=True)
put(bv, "B3", 0.005, "pct", color=BLUE, fill=INP_FILL)
G_ = lambda r: f"=Inputs_Budget!F{r}*1000"
FC_SUB = f"{X('IS', 'rev_ent', 'AN')}+{X('IS', 'rev_mid', 'AN')}"
PL = [
    (5, "Subscription revenue", G_(6 - 1), "=" + FC_SUB, "Budget assumed faster new-logo conversion and lower early-year churn than the forecast; see the ARR bridge below."),
    (6, "Professional services revenue", G_(6), f"={X('IS', 'rev_impl', 'AN')}", "Implementation fees follow Enterprise new-logo volume."),
    (7, "Total revenue", "=B5+B6", "=C5+C6", ""),
    (8, "Cost of sales", G_(8), f"=-{X('IS', 'cogs_tot', 'AN')}", "Hosting and Customer Success run slightly above the budgeted unit costs."),
    (9, "Gross profit", "=B7+B8", "=C7+C8", ""),
    (10, "Sales & marketing", G_(10), f"=-{X('IS', 'o_sm', 'AN')}", "Favourable largely because commission is now capitalised (IFRS 15); cash S&M is higher than the P&L shows."),
    (11, "Research & development", G_(11), f"=-{X('IS', 'o_rd', 'AN')}", "Hiring ahead of the budgeted profile."),
    (12, "General & administrative", G_(12), f"=-{X('IS', 'o_ga', 'AN')}", "Professional fees and G&A hiring above budget."),
    (13, "Total operating costs", "=SUM(B10:B12)", "=SUM(C10:C12)", ""),
    (14, "EBITDA", "=B9+B13", "=C9+C13", "Net of the above."),
]
for r, lab, b, f_, cm in PL:
    tot = r in (7, 9, 13, 14)
    put(bv, f"A{r}", lab, bold=tot)
    put(bv, f"B{r}", b, "num", tot)
    put(bv, f"C{r}", f_, "num", tot)
    put(bv, f"D{r}", f"=C{r}-B{r}", "num", tot)
    put(bv, f"E{r}", f"=IF(B{r}=0,0,D{r}/ABS(B{r}))", "pct", tot)
    put(bv, f"F{r}", f'=IF(ABS(E{r})<$B$3,"On budget",IF(D{r}>0,"Favourable","Adverse"))', None, tot, align="right")
    put(bv, f"G{r}", cm, italic=True)
BUD_EBITDA_CHECK = "Budget_vs_Forecast!B14"
secrow(bv, 16, "Variance by quarter (forecast - budget, GBP k)", 7)
hdr(bv, 17, ["Line", "Q1", "Q2", "Q3", "Q4", "FY1", ""])
QL = [(5, "Subscription revenue", 1, ["IS|rev_ent", "IS|rev_mid"], 5), (6, "Professional services revenue", 1, ["IS|rev_impl"], 6),
      (8, "Cost of sales", -1, ["IS|cogs_tot"], 8), (10, "Sales & marketing", -1, ["IS|o_sm"], 10),
      (11, "Research & development", -1, ["IS|o_rd"], 11), (12, "General & administrative", -1, ["IS|o_ga"], 12)]
qr = 18
for brow, lab, sign, srcs, ibr in QL:
    put(bv, f"A{qr}", lab)
    for qi, c in enumerate("BCDE", start=1):
        fc = "+".join(f"SUMIF({sh}!$O$6:$AL$6,COLUMN()-COLUMN($A$1),{sh}!$O${row_of(sh, k)}:$AL${row_of(sh, k)})" for sh, k in (s_.split("|") for s_ in srcs))
        put(bv, f"{c}{qr}", f"={'-' if sign < 0 else ''}({fc})-Inputs_Budget!{'BCDE'[qi - 1]}{ibr}*1000", "num")
    put(bv, f"F{qr}", f"=SUM(B{qr}:E{qr})", "num")
    qr += 1
put(bv, f"A{qr}", "EBITDA variance", bold=True)
for c in "BCDEF":
    put(bv, f"{c}{qr}", f"={c}18+{c}19+{c}20+{c}21+{c}22+{c}23", "num", True)
    bv[f"{c}{qr}"].border = Border(top=THIN)
Q_EBITDA_ROW = qr
secrow(bv, qr + 2, "FY1 ARR drivers: budget vs forecast (GBP k)", 7)
hdr(bv, qr + 3, ["Driver", "Budget", "Forecast", "Variance", "", "", ""])
k0 = qr + 4
KP = [("Opening ARR", "=Inputs_Budget!B17*1000", f"={X('Rev_Build', 'arr_open', 'AN')}"),
      ("New-customer ARR", "=Inputs_Budget!B18*1000", f"={X('Rev_Build', 'arr_new', 'AN')}"),
      ("Expansion ARR", "=Inputs_Budget!B19*1000", f"={X('Rev_Build', 'arr_exp', 'AN')}"),
      ("Churned ARR", "=Inputs_Budget!B20*1000", f"={X('Rev_Build', 'arr_churn', 'AN')}"),
      ("Closing ARR", f"=SUM(B{k0}:B{k0 + 3})", f"=SUM(C{k0}:C{k0 + 3})"),
      ("New customers (number)", "=Inputs_Budget!B22", f"={X('Rev_Build', 'cust_new', 'AN')}"),
      ("Average ACV of new customers", f"=B{k0 + 1}/B{k0 + 5}", f"=C{k0 + 1}/C{k0 + 5}"),
      ("Closing headcount (FTE)", "=Inputs_Budget!B23", f"={X('Headcount_Opex', 'fte_close', 'AN')}")]
for i, (lab, b, f_) in enumerate(KP):
    r = k0 + i
    put(bv, f"A{r}", lab, bold=(lab == "Closing ARR"))
    fm = "num1" if "ACV" in lab else "num"
    put(bv, f"B{r}", b, fm)
    put(bv, f"C{r}", f_, fm)
    put(bv, f"D{r}", f"=C{r}-B{r}", fm)
K = {"open": k0, "new": k0 + 1, "exp": k0 + 2, "churn": k0 + 3, "close": k0 + 4, "n": k0 + 5, "acv": k0 + 6}
b0 = k0 + len(KP) + 1
secrow(bv, b0, "Bridge 1: FY1 closing ARR, budget to forecast (GBP k)", 7)
hdr(bv, b0 + 1, ["Step", "GBP k", "Cumulative", "", "", "", "How it is calculated"])
ARR_B = [
    ("Budget closing ARR", f"=B{K['close']}", "Budget"),
    ("Opening ARR difference", f"=C{K['open']}-B{K['open']}", "Forecast opening ARR - budget opening ARR"),
    ("Volume: number of new customers", f"=(C{K['n']}-B{K['n']})*B{K['acv']}", "(forecast customers - budget customers) x budget ACV"),
    ("Price: ACV of new customers", f"=C{K['n']}*(C{K['acv']}-B{K['acv']})", "forecast customers x (forecast ACV - budget ACV)"),
    ("Expansion", f"=C{K['exp']}-B{K['exp']}", "Forecast expansion - budget expansion"),
    ("Churn", f"=C{K['churn']}-B{K['churn']}", "Forecast churn - budget churn (negative = more churn than budget)"),
]
for i, (lab, fx, how) in enumerate(ARR_B):
    r = b0 + 2 + i
    put(bv, f"A{r}", lab, bold=(i == 0))
    put(bv, f"B{r}", fx, "num", i == 0)
    put(bv, f"C{r}", f"=B{r}" if i == 0 else f"=C{r - 1}+B{r}", "num")
    put(bv, f"G{r}", how, italic=True)
ra = b0 + 2 + len(ARR_B)
put(bv, f"A{ra}", "Forecast closing ARR", bold=True)
put(bv, f"B{ra}", f"=C{K['close']}", "num", True)
put(bv, f"C{ra}", f"=C{ra - 1}", "num", True)
put(bv, f"G{ra}", "Bridge ends here; tie-out below must be zero", italic=True)
ARR_BR = {"vol": b0 + 4, "price": b0 + 5, "exp": b0 + 6, "churn": b0 + 7, "first": b0 + 2, "last": ra}
c0 = ra + 2
secrow(bv, c0, "Bridge 2: FY1 revenue, budget to forecast (GBP k)", 7)
hdr(bv, c0 + 1, ["Step", "GBP k", "Cumulative", "", "", "", "How it is calculated"])
REV_B = [
    ("Budget revenue", "=B7", "Budget"),
    ("Volume", f"=B{ARR_BR['vol']}*0.5", "ARR volume effect x 0.5: flows land mid-year on average, so about half reaches FY1 revenue"),
    ("Price", f"=B{ARR_BR['price']}*0.5", "ARR price effect x 0.5 (same mid-year convention)"),
    ("Expansion", f"=B{ARR_BR['exp']}*0.5", "ARR expansion effect x 0.5"),
    ("Churn", f"=B{ARR_BR['churn']}*0.5", "ARR churn effect x 0.5"),
]
for i, (lab, fx, how) in enumerate(REV_B):
    r = c0 + 2 + i
    put(bv, f"A{r}", lab, bold=(i == 0))
    put(bv, f"B{r}", fx, "num", i == 0)
    put(bv, f"C{r}", f"=B{r}" if i == 0 else f"=C{r - 1}+B{r}", "num")
    put(bv, f"G{r}", how, italic=True)
rt = c0 + 2 + len(REV_B)
put(bv, f"A{rt}", "Timing / other (residual)")
put(bv, f"B{rt}", f"=C7-C{rt - 1}", "num")
put(bv, f"C{rt}", f"=C{rt - 1}+B{rt}", "num")
put(bv, f"G{rt}", "Billing-to-revenue timing, implementation fees, intra-year phasing: whatever the ARR effects do not explain", italic=True)
put(bv, f"A{rt + 1}", "Forecast revenue", bold=True)
put(bv, f"B{rt + 1}", "=C7", "num", True)
put(bv, f"C{rt + 1}", f"=C{rt}", "num", True)
REV_BR = {"last": rt + 1}
e0 = rt + 3
secrow(bv, e0, "Bridge 3: FY1 EBITDA, budget to forecast (GBP k)", 7)
hdr(bv, e0 + 1, ["Step", "GBP k", "Cumulative", "", "", "", "How it is calculated"])
EB = [("Budget EBITDA", "=B14", "Budget"), ("Revenue", "=D7", "Total revenue variance"), ("Cost of sales", "=D8", "Positive = lower cost than budget"),
      ("Sales & marketing", "=D10", "Includes the effect of capitalising commission"), ("Research & development", "=D11", ""),
      ("General & administrative", "=D12", "")]
for i, (lab, fx, how) in enumerate(EB):
    r = e0 + 2 + i
    put(bv, f"A{r}", lab, bold=(i == 0))
    put(bv, f"B{r}", fx, "num", i == 0)
    put(bv, f"C{r}", f"=B{r}" if i == 0 else f"=C{r - 1}+B{r}", "num")
    put(bv, f"G{r}", how, italic=True)
re_ = e0 + 2 + len(EB)
put(bv, f"A{re_}", "Forecast EBITDA", bold=True)
put(bv, f"B{re_}", "=C14", "num", True)
put(bv, f"C{re_}", f"=C{re_ - 1}", "num", True)
bv.freeze_panes = "B5"
EXTRA_CHECKS.append(("Bridge 1 ties (ARR)", "Budget ARR + bridge steps = forecast closing ARR",
                     f"=Budget_vs_Forecast!C{ARR_BR['last'] - 1}-Budget_vs_Forecast!B{ARR_BR['last']}", 0.01))
EXTRA_CHECKS.append(("Bridge 3 ties (EBITDA)", "Budget EBITDA + bridge steps = forecast EBITDA",
                     f"=Budget_vs_Forecast!C{re_ - 1}-Budget_vs_Forecast!B{re_}", 0.01))
EXTRA_CHECKS.append(("Quarterly variances sum to FY1", "Sum of the quarterly EBITDA variances = FY1 EBITDA variance",
                     f"=Budget_vs_Forecast!F{Q_EBITDA_ROW}-Budget_vs_Forecast!D14", 0.01))

# ----- Metrics catalogue used by the snapshot tables ---------------------------------------------
sens = new_sheet("Sensitivity", "Sensitivity - two-way grids (static, engine-run)",
                 "Sheets and Excel can't re-run the whole model per cell without data tables, so grid values below were produced by recalculating the full workbook "
                 "(validation/snapshots.py) and pasted as grey values. A staleness check on the Checks tab flags them if the model changes.",
                 {"A": 44, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13})
put(sens, "A4", "Lowest month-end cash over the 24 months (live)", bold=True)
put(sens, "B4", f"=MIN({rng('BS', 'cash')})", "num", True)
METRICS = [
    ("rev_fy1", "FY1 revenue", "num", f"IS!AN{row_of('IS', 'rev_tot')}"),
    ("rev_fy2", "FY2 revenue", "num", f"IS!AO{row_of('IS', 'rev_tot')}"),
    ("ebitda_fy1", "FY1 EBITDA", "num", f"IS!AN{row_of('IS', 'ebitda')}"),
    ("ebitda_fy2", "FY2 EBITDA", "num", f"IS!AO{row_of('IS', 'ebitda')}"),
    ("arr12", "ARR, month 12", "num", f"SaaS_Metrics!Z{row_of('SaaS_Metrics', 'arr_close')}"),
    ("arr24", "ARR, month 24", "num", f"SaaS_Metrics!AL{row_of('SaaS_Metrics', 'arr_close')}"),
    ("growth12", "ARR growth, month 12", "pct", f"SaaS_Metrics!Z{row_of('SaaS_Metrics', 'growth')}"),
    ("nrr12", "NRR (annualised), month 12", "pct", f"SaaS_Metrics!Z{row_of('SaaS_Metrics', 'nrr')}"),
    ("grr12", "GRR (annualised), month 12", "pct", f"SaaS_Metrics!Z{row_of('SaaS_Metrics', 'grr')}"),
    ("payback12", "CAC payback (months), month 12", "mo", f"SaaS_Metrics!Z{row_of('SaaS_Metrics', 'payback')}"),
    ("ltvcac12", "LTV / CAC, month 12", "x", f"SaaS_Metrics!Z{row_of('SaaS_Metrics', 'ltv_cac')}"),
    ("cash24", "Cash, month 24", "num", f"BS!AL{row_of('BS', 'cash')}"),
    ("mincash", "Lowest month-end cash", "num", "Sensitivity!B4"),
    ("fte24", "Headcount (FTE), month 24", "num", f"Headcount_Opex!AL{row_of('Headcount_Opex', 'fte_close')}"),
    ("rule40_24", "Rule of 40, month 24", "pct", f"SaaS_Metrics!AL{row_of('SaaS_Metrics', 'rule40')}"),
] + MA_METRICS
MET = {m[0]: m for m in METRICS}

# Sensitivity grids
CHURN_SHIFTS = [-0.04, -0.02, 0.0, 0.02, 0.04]
SCALES = [0.8, 0.9, 1.0, 1.1, 1.2]
secrow(sens, 6, "Grid 1: lowest month-end cash (GBP k) - churn shift (rows, pts added to both segments) x new-logo volume (columns)", 7)
hdr(sens, 7, ["Churn shift  \\  new-logo volume", "80%", "90%", "100%", "110%", "120%", ""])
for j, sc in enumerate(SCALES):
    put(sens, f"{L(2 + j)}7", sc, '0%', True, "FFFFFF", HDR_FILL, align="right")
for i, cs in enumerate(CHURN_SHIFTS):
    r = 8 + i
    put(sens, f"A{r}", cs, '+0.0%;-0.0%;0.0%', True, BLUE, align="left")
    for j in range(5):
        static(sens, f"{L(2 + j)}{r}", sn("grid_mincash", i, j), "num")
sens["D10"].fill = INP_FILL
put(sens, "A13", "Highlighted cell = Base case. Grey italic = pasted engine output.", italic=True)
secrow(sens, 15, "Grid 2: enterprise value (GBP m) - ARR multiple (rows) x new-logo volume (columns), on month-24 ARR", 7)
hdr(sens, 16, ["ARR multiple  \\  new-logo volume", "80%", "90%", "100%", "110%", "120%", ""])
for j, sc in enumerate(SCALES):
    put(sens, f"{L(2 + j)}16", sc, '0%', True, "FFFFFF", HDR_FILL, align="right")
put(sens, "A17", "Month-24 ARR at that volume (GBP k, engine-run)", italic=True)
for j in range(5):
    static(sens, f"{L(2 + j)}17", sn("arr24_by_scale", j), "num")
for i, mult in enumerate([4, 5, 6, 7, 8]):
    r = 18 + i
    put(sens, f"A{r}", mult, '0.0"x"', color=BLUE, bold=True, align="left", fill=INP_FILL)
    for j in range(5):
        put(sens, f"{L(2 + j)}{r}", f"=$A{r}*{L(2 + j)}$17/1000", "num1")
sens["D20"].fill = INP_FILL
put(sens, "A24", "EV = multiple (blue, editable) x snapshot ARR. Multiples are illustrative, synthetic.", italic=True)

# Lever_Analysis
LEVERS = [
    ("inp_churn_ent", "Enterprise gross churn", 0.01, "pct"), ("inp_churn_mid", "Mid-market gross churn", 0.01, "pct"),
    ("inp_exp_ent", "Enterprise expansion", 0.01, "pct"), ("inp_exp_mid", "Mid-market expansion", 0.01, "pct"),
    ("inp_logos_ent", "Enterprise new logos / month", 0.5, "num1"), ("inp_logos_mid", "Mid-market new logos / month", 2, "num1"),
    ("inp_acv_ent", "Enterprise ACV (GBP k)", 10, "num1"), ("inp_acv_mid", "Mid-market ACV (GBP k)", 2, "num1"),
    ("inp_price_uplift", "Annual price uplift", 0.01, "pct"), ("inp_hires_sm", "S&M net hires / month", 0.5, "num1"),
    ("inp_hires_rd", "R&D net hires / month", 0.5, "num1"), ("inp_dso", "Customer payment terms (days)", 5, "num"),
]
LEV_METRICS = ["arr24", "nrr12", "payback12", "cash24", "ebitda_fy2"]
lv = new_sheet("Lever_Analysis", "Lever_Analysis - which drivers move each metric, and by how much",
               "Static, engine-run: each driver is moved one step down and up with everything else at Base, and the whole workbook is recalculated "
               "(validation/snapshots.py). Grey italic = pasted engine output. Section A shows whether the snapshot still matches the live model.",
               {"A": 40, "B": 10})
secrow(lv, 4, "A. Snapshot check: live model vs Base snapshot", 8)
hdr(lv, 5, ["Metric", "Live", "Snapshot", "Gap", "Status"])
a0 = 6
for i, (mid, lab, fmt, cell) in enumerate(METRICS):
    r = a0 + i
    put(lv, f"A{r}", lab)
    put(lv, f"B{r}", f"={cell}", fmt)
    static(lv, f"C{r}", sn("base", mid), fmt)
    put(lv, f"D{r}", f'=IF(ISNUMBER(B{r}),B{r}-C{r},0)', fmt)
    put(lv, f"E{r}", f'=IF(inp_scenario<>"Base","switch to Base to check",IF(ABS(D{r})<=IF(ABS(C{r})<10,chk_tol/100,chk_tol),"OK","STALE"))', None, align="right")
    lv.column_dimensions["E"].width = 22
STALE_LAST = a0 + len(METRICS) - 1
STALE_ROWS = {m: a0 + i for i, (m, *_r) in enumerate(METRICS)}
lv_b = STALE_LAST + 3
secrow(lv, lv_b, "B. Impact matrix: metric value with the driver one step DOWN / UP; impact = (up - down) / 2 per step", 8)
hdr(lv, lv_b + 1, ["Driver", "Step"])
col = 3
LEV_COL = {}
for mid in LEV_METRICS:
    put(lv, f"{L(col)}{lv_b}", MET[mid][1], bold=True, fill=SEC_FILL)
    for k, lab in enumerate(("down", "up", "impact")):
        put(lv, f"{L(col + k)}{lv_b + 1}", lab, bold=True, color="FFFFFF", fill=HDR_FILL, align="right")
        lv.column_dimensions[L(col + k)].width = 11
    LEV_COL[mid] = col
    col += 3
abs_col0 = col + 1
for j, mid in enumerate(LEV_METRICS):
    put(lv, f"{L(abs_col0 + j)}{lv_b + 1}", "|" + MET[mid][1][:14] + "|", bold=True, color="FFFFFF", fill=HDR_FILL, align="right")
    lv.column_dimensions[L(abs_col0 + j)].width = 13
m0 = lv_b + 2
for i, (drv, lab, step, sf) in enumerate(LEVERS):
    r = m0 + i
    put(lv, f"A{r}", lab)
    put(lv, f"B{r}", step, sf, color=BLUE, fill=INP_FILL)
    for j, mid in enumerate(LEV_METRICS):
        cc_ = LEV_COL[mid]
        fm = MET[mid][2]
        static(lv, f"{L(cc_)}{r}", sn("levers", drv, "dn", mid), fm)
        static(lv, f"{L(cc_ + 1)}{r}", sn("levers", drv, "up", mid), fm)
        put(lv, f"{L(cc_ + 2)}{r}", f"=({L(cc_ + 1)}{r}-{L(cc_)}{r})/2", fm)
        put(lv, f"{L(abs_col0 + j)}{r}", f"=ABS({L(cc_ + 2)}{r})", fm)
m1 = m0 + len(LEVERS) - 1
put(lv, f"A{m1 + 1}", "Step sizes (blue) are labels for the snapshot run; changing them here does not rerun it.", italic=True)
t0 = m1 + 4
secrow(lv, t0, "C. Top three levers per metric (ranked by absolute impact per step; live formulas on the matrix above)", 8)
hdr(lv, t0 + 1, ["Metric", "Rank 1", "Rank 2", "Rank 3"])
for i, mid in enumerate(LEV_METRICS):
    r = t0 + 2 + i
    put(lv, f"A{r}", MET[mid][1])
    ac = L(abs_col0 + i)
    for k in range(3):
        put(lv, f"{L(2 + k)}{r}", f"=INDEX($A${m0}:$A${m1},MATCH(LARGE({ac}${m0}:{ac}${m1},{k + 1}),{ac}${m0}:{ac}${m1},0),1)", None, align="right")
for _c in "BCD":
    lv.column_dimensions[_c].width = 24
EXTRA_CHECKS.append(("Static tables current (Base only)", "Live Base outputs = snapshot behind Lever_Analysis / Sensitivity / Business_Review (sum of absolute gaps; 0 when another scenario is selected)",
                     "=IF(inp_scenario=\"Base\"," + "+".join(f"ABS(Lever_Analysis!D{STALE_ROWS[m]})" for m in ("rev_fy1", "ebitda_fy2", "arr24", "cash24", "mincash")) + ",0)", 1))

# Business_Review
br = new_sheet("Business_Review", "Business_Review - one-page pack: KPIs, what moved, risks, downside alongside the point estimate",
               "Section 1 is live for the selected scenario. Section 3 compares engine-run Base / Upside / Downside snapshots (grey italic).",
               {"A": 46, "B": 14, "C": 14, "D": 14, "E": 14, "F": 60})
put(br, "A3", '="Scenario shown: "&inp_scenario&"   |   Master check: "&IF(chk_master=1,"ALL CHECKS PASS","CHECK FAILURE")', bold=True)
secrow(br, 5, "1. Headline KPIs (live)", 6)
hdr(br, 6, ["KPI", "FY1", "FY2"])
KP2 = [("Revenue", "IS", "rev_tot", "num"), ("Gross margin", "IS", "gm", "pct"), ("EBITDA", "IS", "ebitda", "num"), ("EBITDA margin", "IS", "ebitda_m", "pct"),
       ("Closing ARR", "SaaS_Metrics", "arr_close", "num"), ("ARR growth (year on year)", "SaaS_Metrics", "growth", "pct"),
       ("Net revenue retention", "SaaS_Metrics", "nrr", "pct"), ("Gross revenue retention", "SaaS_Metrics", "grr", "pct"),
       ("CAC payback (months)", "SaaS_Metrics", "payback", "mo"), ("LTV / CAC", "SaaS_Metrics", "ltv_cac", "x"),
       ("Burn multiple", "SaaS_Metrics", "burn_mult", "x"), ("Rule of 40", "SaaS_Metrics", "rule40", "pct"),
       ("Closing cash", "BS", "cash", "num"), ("Closing headcount (FTE)", "Headcount_Opex", "fte_close", "num")]
for i, (lab, sh, k, fm) in enumerate(KP2):
    r = 7 + i
    put(br, f"A{r}", lab)
    put(br, f"B{r}", f"={X(sh, k, 'AN')}", fm)
    put(br, f"C{r}", f"={X(sh, k, 'AO')}", fm)
r = 7 + len(KP2) + 1
secrow(br, r, "2. What moved vs budget in FY1, and why (live, from Budget_vs_Forecast)", 6)
WM = [("Closing ARR vs budget", f"=Budget_vs_Forecast!D{K['close']}", "num", "Sum of the four ARR drivers below"),
      ("   Volume (new customers)", f"=Budget_vs_Forecast!B{ARR_BR['vol']}", "num", "Fewer/more new logos than budgeted"),
      ("   Price (ACV)", f"=Budget_vs_Forecast!B{ARR_BR['price']}", "num", "Deal size vs budget"),
      ("   Expansion", f"=Budget_vs_Forecast!B{ARR_BR['exp']}", "num", "Upsell on the installed base"),
      ("   Churn", f"=Budget_vs_Forecast!B{ARR_BR['churn']}", "num", "Negative = more ARR lost than budgeted"),
      ("Revenue vs budget", "=Budget_vs_Forecast!D7", "num", ""), ("EBITDA vs budget", "=Budget_vs_Forecast!D14", "num", "")]
for i, (lab, fx, fm, why) in enumerate(WM):
    put(br, f"A{r + 1 + i}", lab)
    put(br, f"B{r + 1 + i}", fx, fm)
    put(br, f"F{r + 1 + i}", why, italic=True)
r2 = r + len(WM) + 2
secrow(br, r2, "3. Point estimate vs downside (engine-run snapshots)", 6)
hdr(br, r2 + 1, ["Metric", "Base", "Upside", "Downside", "Downside vs Base"])
for i, (mid, lab, fmt, cell) in enumerate(METRICS):
    rr_ = r2 + 2 + i
    put(br, f"A{rr_}", lab)
    for j, sc in enumerate(("Base", "Upside", "Downside")):
        static(br, f"{L(2 + j)}{rr_}", sn("scen", sc, mid), fmt)
    put(br, f"E{rr_}", f'=IF(AND(ISNUMBER(B{rr_}),ISNUMBER(D{rr_})),D{rr_}-B{rr_},"")', fmt)
r3 = r2 + 2 + len(METRICS) + 1
secrow(br, r3, "4. Risk flags (live thresholds on the selected scenario)", 6)
RF = [("Cash runway under this many months (month 24)", 12, lambda t: f'=IF(ISNUMBER({X("SaaS_Metrics", "runway", "AL")}),IF({X("SaaS_Metrics", "runway", "AL")}<{t},"FLAG","OK"),"OK (cash generative)")', "Runway = cash / trailing 3-month average burn.", "mo"),
      ("Net revenue retention below (month 12)", 1.0, lambda t: f'=IF({X("SaaS_Metrics", "nrr", "Z")}<{t},"FLAG","OK")', "The installed base is shrinking before new sales.", "pct"),
      ("CAC payback above this many months (month 12)", 24, lambda t: f'=IF({X("SaaS_Metrics", "payback", "Z")}>{t},"FLAG","OK")', "Growth is expensive to buy.", "mo"),
      ("Lowest month-end cash below (GBP k)", 0, lambda t: f'=IF(Sensitivity!B4<{t},"FLAG","OK")', "Funding gap.", "num")]
for i, (lab, thr, fxf, why, tf) in enumerate(RF):
    put(br, f"A{r3 + 1 + i}", lab)
    put(br, f"C{r3 + 1 + i}", thr, tf, color=BLUE, fill=INP_FILL)
    put(br, f"B{r3 + 1 + i}", fxf(f"$C{r3 + 1 + i}"), None, True, align="right")
    put(br, f"F{r3 + 1 + i}", why, italic=True)
BR_LAST = r3 + len(RF)
br.freeze_panes = "A5"

SPEC = {
    "selector": DRV_CELL["inp_scenario"],
    "metrics": {m[0]: m[3] for m in METRICS},
    "scen_cells": {n: f"F{r}" for n, r in SCEN_ROW.items()},
    "levers": [{"driver": d, "step": st} for d, _l, st, _f in LEVERS],
    "lever_metrics": LEV_METRICS,
    "churn_shifts": CHURN_SHIFTS,
    "scales": SCALES,
}
json.dump(SPEC, open(SPEC_PATH, "w"), indent=1)

# ----- Sensitivity grid 3: acquisition value created (closed-form, live formulas) -----------------
g0 = 27
secrow(sens, g0, "Grid 3 (live formulas): deal value created (GBP k) - offer multiple (rows) x synergies achieved (columns)", 7)
hdr(sens, g0 + 1, ["Offer EV / ARR  \\  synergies achieved", "", "", "", "", "", "Break-even synergies"])
for j, a_ in enumerate([0.0, 0.4, 0.7, 1.0, 1.3]):
    put(sens, f"{L(2 + j)}{g0 + 1}", a_, '0%', True, "FFFFFF", HDR_FILL, align="right")
for i, m_ in enumerate([3.5, 4.0, 4.5, 5.0, 5.5]):
    r = g0 + 2 + i
    put(sens, f"A{r}", m_, '0.0"x"', color=BLUE, bold=True, align="left", fill=INP_FILL)
    for j in range(5):
        c = L(2 + j)
        put(sens, f"{c}{r}", f"=MA_Returns!$C${MRS['dcf']}+{c}${g0 + 1}*MA_Returns!$C${MRS['pv_syn100']}-MA_Returns!$C${MRS['pv_integ']}-$A{r}*{PA('arr')}*(1+inp_ma_deal_cost_pct)", "num")
    put(sens, f"G{r}", f"=($A{r}*{PA('arr')}*(1+inp_ma_deal_cost_pct)+MA_Returns!$C${MRS['pv_integ']}-MA_Returns!$C${MRS['dcf']})/MA_Returns!$C${MRS['pv_syn100']}", "pct")
put(sens, f"A{g0 + 7}", "Value is linear in the synergy % and in the price, so this grid needs no model re-run. Negative = the deal destroys value.", italic=True)

for sh in (RB, RR, HC, CC, CG, IS, BS, CF, SM, MT, MP):
    sh.write()

# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------
ck = wb.create_sheet("Checks")
ck.sheet_view.showGridLines = False
put(ck, "A1", "Checks - integrity tests and master flag", bold=True)
ck["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3A5F")
put(ck, "A2", "Each check returns a difference (0 = perfect). It passes when the absolute value is within tolerance.", italic=True)
for ci, h in enumerate(["Check", "What it proves", "Value", "Tolerance", "Pass (1/0)"], start=1):
    put(ck, f"{L(ci)}4", h, bold=True, color="FFFFFF", fill=HDR_FILL)
put(ck, "A3", "Global tolerance for tie-out checks (GBP k)", bold=True)
put(ck, "C3", 0.01, "num2", color=BLUE, fill=INP_FILL)
wb.defined_names["chk_tol"] = DefinedName("chk_tol", attr_text="Checks!$C$3")


def rng(sh, key, a="O", b="AL"):
    r_ = row_of(sh, key)
    return f"{sh}!${a}${r_}:${b}${r_}"


CHECKS = [
    ("Balance sheet balances", "Assets = liabilities + equity in every period (largest absolute gap, GBP k)",
     f"=MAX(MAX({rng('BS', 'check', 'N')}),-MIN({rng('BS', 'check', 'N')}))", 0.01),
    ("Cash ties to cash flow", "Closing cash = opening cash + sum of net cash flows",
     f"=ABS({X('CF', 'cash_close', 'AL')}-({X('CF', 'cash_close', 'N')}+SUM({rng('CF', 'net')})))+ABS({X('BS', 'cash', 'AL')}-{X('CF', 'cash_close', 'AL')})", 0.01),
    ("ARR roll-forward", "Closing ARR = opening ARR + new + expansion - churn over 24 months",
     f"={X('Rev_Build', 'arr_close', 'AL')}-({X('Rev_Build', 'arr_close', 'N')}+SUM({rng('Rev_Build', 'arr_new')})+SUM({rng('Rev_Build', 'arr_exp')})+SUM({rng('Rev_Build', 'arr_churn')}))", 0.01),
    ("ARR segments sum to total", "Enterprise + Mid-market closing ARR = SaaS_Metrics closing ARR in every period",
     f"=MAX(MAX({rng('SaaS_Metrics', 'arr_gap', 'N')}),-MIN({rng('SaaS_Metrics', 'arr_gap', 'N')}))", 0.01),
    ("Customer roll-forward", "Closing customers = opening + new - churned over 24 months",
     f"={X('Rev_Build', 'cust_close', 'AL')}-({X('Rev_Build', 'cust_close', 'N')}+SUM({rng('Rev_Build', 'cust_new')})+SUM({rng('Rev_Build', 'cust_churn')}))", 0.01),
    ("Deferred revenue roll-forward", "Closing deferred revenue = opening + billings - revenue recognised",
     f"={X('Rev_Rec', 'def_tot', 'AL')}-({X('Rev_Rec', 'def_tot', 'N')}+SUM({rng('Rev_Rec', 'bill_tot')})-SUM({rng('Rev_Rec', 'rev_tot')}))", 0.01),
    ("Revenue ties across tabs", "Income statement revenue = Rev_Rec revenue (24-month total)",
     f"=SUM({rng('IS', 'rev_tot')})-SUM({rng('Rev_Rec', 'rev_tot')})", 0.01),
    ("Net income ties to equity", "Closing equity = opening equity + sum of net income",
     f"={X('BS', 'equity', 'AL')}-({X('BS', 'equity', 'N')}+SUM({rng('IS', 'ni')}))", 0.01),
    ("Deferred revenue never negative", "Lowest deferred revenue balance is >= 0 (shortfall shown)",
     f"=MIN(0,MIN({rng('Rev_Rec', 'def_tot')}))", 0),
    ("Cash never negative", "Lowest cash balance is >= 0 (shortfall shown). Fails when the plan runs out of cash",
     f"=MIN(0,MIN({rng('BS', 'cash')}))", 0),
    ("Billings never negative", "Lowest monthly billings is >= 0 (shortfall shown)",
     f"=MIN(0,MIN({rng('Rev_Rec', 'bill_tot')}))", 0),
    ("Contract-cost roll-forward", "Closing asset = opening + capitalised - amortised over 24 months",
     f"={X('Contract_Costs', 'a_close', 'AL')}-({X('Contract_Costs', 'a_close', 'N')}+SUM({rng('Contract_Costs', 'a_cap')})+SUM({rng('Contract_Costs', 'a_amort')}))", 0.01),
    ("Scenario selector valid", "Selected scenario exists in Inputs_Scenarios (0 = valid)",
     "=IF(ISNUMBER(inp_scen_idx),0,1)", 0),
]

_R = lambda sh, k, a="O", b="AL": rng(sh, k, a, b)
EXTRA_CHECKS += [
    ("Annual: FY1 opening ARR = opening balance", "Annual 'opening' column for FY1 equals the opening ARR (would have caught the v0.1 FY1 opening bug)",
     f"={X('Rev_Build', 'arr_open', 'AN')}-{X('Rev_Build', 'arr_close', 'N')}", "tol"),
    ("Annual: FY2 opening ARR = FY1 closing ARR", "Annual columns chain from year to year",
     f"={X('Rev_Build', 'arr_open', 'AO')}-{X('Rev_Build', 'arr_close', 'AN')}", "tol"),
    ("Annual: FY2 opening cash = FY1 closing cash", "Annual cash columns chain from year to year",
     f"={X('CF', 'cash_open', 'AO')}-{X('CF', 'cash_close', 'AN')}", "tol"),
    ("Annual flows = 24-month totals", "FY1 + FY2 revenue, EBITDA and free cash flow = sum of the 24 months (sum of absolute gaps)",
     f"=ABS({X('IS', 'rev_tot', 'AN')}+{X('IS', 'rev_tot', 'AO')}-SUM({_R('IS', 'rev_tot')}))+ABS({X('IS', 'ebitda', 'AN')}+{X('IS', 'ebitda', 'AO')}-SUM({_R('IS', 'ebitda')}))+ABS({X('CF', 'net', 'AN')}+{X('CF', 'net', 'AO')}-SUM({_R('CF', 'net')}))", "tol"),
    ("EBITDA recomputed from other tabs", "IS EBITDA = Rev_Rec revenue - COGS tab cost of revenue - Headcount_Opex opex (independent route, 24-month total)",
     f"=SUM({_R('IS', 'ebitda')})-(SUM({_R('Rev_Rec', 'rev_tot')})-SUM({_R('COGS', 'cogs_tot')})-SUM({_R('Headcount_Opex', 'opex_tot')}))", "tol"),
    ("Sign check: revenue, costs and capex are not negative", "Shortfall below zero across IS revenue, cost of revenue, opex and capex (0 = fine)",
     f"=MIN(0,MIN({_R('IS', 'rev_tot')}))+MIN(0,MIN({_R('IS', 'cogs_tot')}))+MIN(0,MIN({_R('IS', 'opex_tot')}))+MIN(0,MIN({_R('BS', 'capex')}))", 0),
    ("Sign check: ARR, customers and headcount are not negative", "Shortfall below zero (0 = fine)",
     f"=MIN(0,MIN({_R('Rev_Build', 'arr_close', 'N')}))+MIN(0,MIN({_R('Rev_Build', 'cust_close', 'N')}))+MIN(0,MIN({_R('Headcount_Opex', 'fte_close', 'N')}))", 0),
    ("Sign check: churn lines are not positive", "Churned ARR and churned customers must be zero or negative (excess shown)",
     f"=MAX(0,MAX({_R('Rev_Build', 'arr_churn')}))+MAX(0,MAX({_R('Rev_Build', 'cust_churn')}))", 0),
    ("Inputs within allowed range", "Count of inputs outside the min/max on Inputs_Drivers (live scenario values included)",
     f"=SUM(Inputs_Drivers!G8:G{DRV_LAST[0]})", 0),
]
CHECKS.extend(EXTRA_CHECKS)
for i, (nm, why, fx, tol) in enumerate(CHECKS):
    r = 5 + i
    put(ck, f"A{r}", nm)
    put(ck, f"B{r}", why)
    put(ck, f"C{r}", fx, "num2")
    if tol in (0.01, "tol"):
        put(ck, f"D{r}", "=chk_tol", "num2")
    else:
        put(ck, f"D{r}", tol, "num2", color=BLUE, fill=INP_FILL)
    put(ck, f"E{r}", f"=IF(ABS(C{r})<=D{r},1,0)", "int")
last = 4 + len(CHECKS)
mr = last + 2
put(ck, f"A{mr}", "MASTER CHECK", bold=True)
put(ck, f"B{mr}", "1 = every check passes")
put(ck, f"E{mr}", f"=IF(SUM(E5:E{last})=COUNT(E5:E{last}),1,0)", "int", bold=True)
put(ck, f"A{mr + 1}", "Status")
put(ck, f"E{mr + 1}", f'=IF(E{mr}=1,"ALL CHECKS PASS","CHECK FAILURE")', bold=True)
wb.defined_names["chk_master"] = DefinedName("chk_master", attr_text=f"Checks!$E${mr}")
# Outlier warnings: judgement flags that do not fail the master check
wr = mr + 4
secrow(ck, wr, "Outlier warnings (month-on-month moves larger than the threshold; these prompt a review, they do not fail the model)", 5)
for ci, h in enumerate(["Series", "Test", "Months flagged", "Threshold", "Status"], start=1):
    put(ck, f"{L(ci)}{wr + 1}", h, bold=True, color="FFFFFF", fill=HDR_FILL)
OUTL = [("Revenue", "IS", "rev_tot", 0.15, "pct"), ("Total operating expense", "IS", "opex_tot", 0.15, "pct"), ("Closing ARR", "Rev_Build", "arr_close", 0.10, "pct"),
        ("Headcount", "Headcount_Opex", "fte_close", 0.05, "pct"), ("Deferred revenue", "BS", "defrev", 0.15, "pct"), ("Cost of revenue", "IS", "cogs_tot", 0.15, "pct"),
        ("Gross margin (absolute change, pts)", "IS", "gm", 0.02, "abs"), ("EBITDA margin (absolute change, pts)", "IS", "ebitda_m", 0.05, "abs")]
put(ck, f"H{wr + 1}", "Month-on-month flags, months 2 to 24 (1 = beyond threshold)", bold=True, color="FFFFFF", fill=HDR_FILL)
for ci in range(9, 31):
    put(ck, f"{L(ci)}{wr + 1}", None, fill=HDR_FILL)
for i, (nm, sh, key, thr, kind) in enumerate(OUTL):
    r = wr + 2 + i
    rw = row_of(sh, key)
    put(ck, f"A{r}", nm)
    put(ck, f"B{r}", "Absolute month-on-month % change" if kind == "pct" else "Absolute month-on-month change")
    put(ck, f"D{r}", thr, "pct", color=BLUE, fill=INP_FILL)
    for k in range(23):
        cur, prv = f"{sh}!{L(16 + k)}{rw}", f"{sh}!{L(15 + k)}{rw}"
        test = f"ABS({cur}/{prv}-1)" if kind == "pct" else f"ABS({cur}-{prv})"
        put(ck, f"{L(8 + k)}{r}", f"=IF({test}>$D{r},1,0)", "int")
    put(ck, f"C{r}", f"=SUM(H{r}:AD{r})", "int")
    put(ck, f"E{r}", f'=IF(C{r}=0,"OK","REVIEW")', None, align="right")
wl = wr + 1 + len(OUTL)
put(ck, f"A{wl + 1}", "WARNINGS", bold=True)
put(ck, f"B{wl + 1}", "Number of series with at least one flagged month")
put(ck, f"E{wl + 1}", f'=COUNTIF(E{wr + 2}:E{wl},"REVIEW")', "int", bold=True)
wb.defined_names["chk_warnings"] = DefinedName("chk_warnings", attr_text=f"Checks!$E${wl + 1}")
ck.column_dimensions["A"].width = 34
ck.column_dimensions["B"].width = 92
for c in "CDE":
    ck.column_dimensions[c].width = 16

# ---------------------------------------------------------------------------
# SaaS_Metrics lever documentation (MVP: which drivers move each metric)
# ---------------------------------------------------------------------------
ws = SH["SaaS_Metrics"].ws
r = ws.max_row + 3
put(ws, f"A{r}", "What moves each metric (levers). Quantified lever impacts arrive in Lever_Analysis (Phase 2).", bold=True)
for ci in range(1, 8):
    ws.cell(row=r, column=ci).fill = SEC_FILL
LEVERS = [
    ("ARR growth", "inp_logos_ent / inp_logos_mid (volume), inp_acv_* and inp_price_uplift (price), inp_churn_* and inp_exp_* (existing base)"),
    ("NRR", "inp_exp_ent / inp_exp_mid (up), inp_churn_ent / inp_churn_mid (down). Mix shifts toward Enterprise raise it"),
    ("GRR", "inp_churn_ent and inp_churn_mid only"),
    ("CAC", "inp_hires_sm and inp_sal_sm (team cost), inp_marketing_m, inp_commission_pct, new-logo volume"),
    ("CAC payback", "CAC levers, plus gross margin and gross new ARR (new logos, ACV, expansion)"),
    ("LTV / CAC", "ARR per customer, gross margin, logo churn (inp_churn_*) and every CAC lever"),
    ("Burn multiple", "Net burn (opex, hiring, collections via inp_dso) over net new ARR (growth and churn levers)"),
    ("Rule of 40", "ARR growth levers plus EBITDA margin (hiring plan, hosting cost per customer, marketing spend)"),
    ("Cash runway", "Cash balance and burn: hiring plan, DSO, churn, new logos, billing terms"),
    ("Gross margin", "inp_host_ent / inp_host_mid, inp_licence_pct, inp_hires_cs, inp_sal_cs, inp_services_cost_pct"),
]
for i, (m, lev) in enumerate(LEVERS, start=1):
    put(ws, f"A{r + i}", m, bold=True)
    put(ws, f"C{r + i}", lev)

# ---------------------------------------------------------------------------
# Outputs_Web: the only tab the website reads (plus the input registry on it)
# ---------------------------------------------------------------------------
ow = new_sheet("Outputs_Web", "Outputs_Web - flat contract between the workbook and the website",
               "Everything here is a link (green) to a calculation or output tab. The website reads only the named ranges defined on this tab (out_*) and writes only to the named inputs (inp_*). "
               "Series ranges run from the opening balance (column D) through 24 months (E:AB) to FY1 and FY2 (AC:AD).",
               {"A": 30, "B": 46, "C": 10})
for ci in range(4, 31):
    ow.column_dimensions[L(ci)].width = 11
OUTNAMES = {}


def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)
    OUTNAMES[name] = ref


_o = [4]


def ow_sec(label):
    secrow(ow, _o[0], label, 30)
    _o[0] += 1


# ---- A. scalar outputs
ow_sec("A. Scalar outputs (one cell each; name in column A)")
hdr(ow, _o[0], ["Name", "Label", "Unit", "Value"])
_o[0] += 1
SCALARS = [(f"out_{m[0]}", m[1], m[2], f"={m[3]}") for m in METRICS]
SCALARS += [
    ("out_master_check", "Master check (1 = all pass)", "int", "=chk_master"),
    ("out_warnings", "Outlier warnings (series flagged)", "int", "=chk_warnings"),
    ("out_scenario", "Live scenario", None, "=inp_scenario"),
    ("out_goodwill_ifrs", "Goodwill, IFRS 3", "num", f"=MA_PPA!C{PPA['gw']}"),
    ("out_goodwill_usgaap", "Goodwill, US GAAP (ASU 2021-08)", "num", f"=MA_PPA!D{PPA['gw']}"),
    ("out_ma_ev", "Offer enterprise value", "num", f"={PA('ev')}"),
    ("out_ma_price", "Equity purchase price", "num", f"={PA('price')}"),
    ("out_ma_verdict", "Deal verdict", None, f"=MA_Returns!C{MRS['verdict']}"),
    ("out_ma_breakeven", "Synergies needed to break even (% of plan)", "pct", f"=MA_Returns!C{MRS['breakeven']}"),
    ("out_ma_payback", "Deal payback (years)", "num1", f"=MA_Returns!C{MRS['payback']}"),
    ("out_ma_premium", "Premium over standalone DCF", "num", f"=MA_Returns!C{MRS['premium']}"),
    ("out_ma_accretion", "ARR accretion, month 24", "pct", f"=MA_Returns!C{MRS['arr_acc']}"),
    ("out_ma_funding_flag", "Funding flag", None, f"=MA_Returns!C{MRS['cash_flag']}"),
    ("out_ma_dcf_ev", "Vaultline standalone DCF EV", "num", f"={DCF_EV['base']}"),
]
for nm, lab, fm, fx in SCALARS:
    r = _o[0]
    put(ow, f"A{r}", nm)
    put(ow, f"B{r}", lab)
    put(ow, f"C{r}", fm or "text", italic=True)
    put(ow, f"D{r}", fx, fm)
    defname(nm, f"Outputs_Web!$D${r}")
    _o[0] += 1
_o[0] += 1

# ---- B. monthly series
ow_sec("B. Monthly series (opening, months 1-24, FY1, FY2)")
hdr(ow, _o[0], ["Name", "Label", "Unit"])
SER_HDR = _o[0]
put(ow, f"D{SER_HDR}", "Opening", bold=True, color="FFFFFF", fill=HDR_FILL, align="right")
for i in range(24):
    put(ow, f"{L(5 + i)}{SER_HDR}", f"=Rev_Build!{L(15 + i)}4", "date", True, "FFFFFF", HDR_FILL, align="right")
put(ow, f"AC{SER_HDR}", "FY1", bold=True, color="FFFFFF", fill=HDR_FILL, align="right")
put(ow, f"AD{SER_HDR}", "FY2", bold=True, color="FFFFFF", fill=HDR_FILL, align="right")
_o[0] += 1
for nm, rr_, lab in (("out_t", 3, "Period number"), ("out_period_end", 4, "Period end date"), ("out_year", 5, "Year index"), ("out_quarter", 6, "Quarter index")):
    r = _o[0]
    put(ow, f"A{r}", nm)
    put(ow, f"B{r}", lab)
    for i in range(25):
        put(ow, f"{L(4 + i)}{r}", f"=Rev_Build!{L(14 + i)}{rr_}", "date" if rr_ == 4 else "int")
    defname(nm, f"Outputs_Web!$D${r}:$AB${r}")
    _o[0] += 1
SER = [
    ("rb", "Rev_Build", ["logos_ent", "acv_ent", "churn_m_ent", "exp_m_ent", "logos_mid", "acv_mid", "churn_m_mid", "exp_m_mid", "arr_open", "arr_new", "arr_exp", "arr_churn", "arr_close", "cust_close", "arr_close_ent", "arr_close_mid", "arr_new_ent", "arr_new_mid",
                         "arr_exp_ent", "arr_exp_mid", "arr_churn_ent", "arr_churn_mid", "cust_close_ent", "cust_close_mid"]),
    ("rr", "Rev_Rec", ["inv_tot", "fee_impl", "bill_tot", "rev_ent", "rev_mid", "rev_sub", "rev_impl", "rev_tot", "def_sub_close", "def_imp_close", "def_tot", "ar"]),
    ("hc", "Headcount_Opex", ["h_close_rd", "h_close_sm", "h_close_ga", "h_close_cs", "fte_close", "pay_tot", "commission", "opex_rd", "opex_sm", "opex_ga", "opex_tot"]),
    ("cc", "Contract_Costs", ["cap", "a_amort", "a_close", "pl_charge"]),
    ("cogs", "COGS", ["host", "lic", "support", "svc", "cogs_tot", "gm_sub", "gm_svc", "gm_tot"]),
    ("is", "IS", ["rev_ent", "rev_mid", "rev_impl", "rev_tot", "c_host", "c_lic", "c_support", "c_svc", "cogs_tot", "gp", "gm", "o_rd", "o_sm", "o_ga", "opex_tot",
                  "ebitda", "ebitda_m", "da", "ebit", "interest", "pbt", "tax", "ni"]),
    ("bs", "BS", ["cash", "ar", "ppe", "cca", "ta", "ap", "defrev", "tl", "equity", "check"]),
    ("cf", "CF", ["ni", "da", "d_ar", "d_ap", "d_def", "d_cca", "cfo", "capex", "cfi", "net", "cash_open", "cash_close"]),
    ("sm", "SaaS_Metrics", ["arr_open", "arr_new", "arr_exp", "arr_churn", "arr_close", "netnew", "growth", "nrr", "grr", "logo_churn", "sm_spend", "cac", "gm", "payback",
                            "arpa", "ltv", "ltv_cac", "magic", "netburn", "burn_mult", "runway", "rule40", "arr_fte"]),
    ("mt", "MA_Target", ["arr_close", "rev", "ebitda", "ni", "cash_close", "ufcf"]),
    ("pf", "MA_ProForma", ["pf_rev", "pf_ebitda", "pf_ni", "pf_fcf", "pf_cash", "pf_arr", "pf_growth", "pf_nrr", "pf_ebitda_m", "pf_rule40", "pf_burn_mult", "pf_runway",
                           "syn_arr", "syn_rev", "syn_ebitda", "integ", "txn", "hc_loss", "amort", "int_loan", "ic_rev", "pb_cash", "pb_ar", "pb_ppe", "pb_cca", "pb_int",
                           "pb_gw", "pb_ta", "pb_ap", "pb_dr", "pb_loan", "pb_dtl", "pb_tl", "pb_eq", "pb_check"]),
]
NSER = 0
for pref, shname, keys in SER:
    for key in keys:
        src = SH[shname]
        row = src.reg[key]
        spec = next(d for k_, r_, d in src.items if k_ == "row" and r_ == row)
        r = _o[0]
        nm = f"out_{pref}_{key}"
        put(ow, f"A{r}", nm)
        put(ow, f"B{r}", f"{shname}: {spec['label']}")
        put(ow, f"C{r}", spec["unit"], italic=True)
        fm = spec["fmt"]
        if spec["opening"] is not None or spec["hist"]:
            put(ow, f"D{r}", f"={shname}!N{row}", fm)
        for i in range(24):
            put(ow, f"{L(5 + i)}{r}", f"={shname}!{L(15 + i)}{row}", fm)
        if spec["ann"] and spec["f"]:
            put(ow, f"AC{r}", f"={shname}!AN{row}", fm)
            put(ow, f"AD{r}", f"={shname}!AO{row}", fm)
        defname(nm, f"Outputs_Web!$D${r}:$AD${r}")
        _o[0] += 1
        NSER += 1
_o[0] += 1

# ---- C. long-term view (years 1-5)
ow_sec("C. Five-year view (years 1-5 in columns D:H)")
hdr(ow, _o[0], ["Name", "Label", "Unit"])
for i, c in enumerate("DEFGH"):
    put(ow, f"{c}{_o[0]}", f"=LT_Model!{'CDEFG'[i]}4", None, True, "FFFFFF", HDR_FILL, align="right")
_o[0] += 1
for key, row in LT.items():
    r = _o[0]
    spec = next(x for x in LTSPEC if x[0] == row)
    put(ow, f"A{r}", f"out_lt_{key}")
    put(ow, f"B{r}", f"LT_Model: {spec[1]}")
    put(ow, f"C{r}", spec[2], italic=True)
    for i, c in enumerate("DEFGH"):
        put(ow, f"{c}{r}", f"=LT_Model!{'CDEFG'[i]}{row}", spec[3])
    defname(f"out_lt_{key}", f"Outputs_Web!$D${r}:$H${r}")
    _o[0] += 1
_o[0] += 1


def table(name, title, header, rows):
    """Write a 2-D linked block and name its data area."""
    ow_sec(title)
    hdr(ow, _o[0], header)
    _o[0] += 1
    first = _o[0]
    for rowvals in rows:
        for j, (fx, fm) in enumerate(rowvals):
            put(ow, f"{L(1 + j)}{_o[0]}", fx, fm)
        _o[0] += 1
    defname(name, f"Outputs_Web!$A${first}:${L(len(header))}${_o[0] - 1}")
    _o[0] += 1


bvr = lambda r_, cols, fms: [(f"=Budget_vs_Forecast!{c}{r_}", f) for c, f in zip(cols, fms)]
table("out_bvf_pl", "D. FY1 budget vs forecast (line, budget, forecast, variance, variance %, flag, commentary)", ["Line", "Budget", "Forecast", "Variance", "Var %", "Flag", "Commentary"],
      [bvr(r_, "ABCDEFG", [None, "num", "num", "num", "pct", None, None]) for r_ in range(5, 15)])
_lh = next(r_ for r_ in range(1, SH["SaaS_Metrics"].ws.max_row + 1) if str(SH["SaaS_Metrics"].ws[f"A{r_}"].value or "").startswith("What moves each metric"))
table("out_levers", "D2. What moves each metric (metric, levers)", ["Metric", "Levers"],
      [[(f"=SaaS_Metrics!A{_lh + i}", None), (f"=SaaS_Metrics!C{_lh + i}", None)] for i in range(1, len(LEVERS) + 1)])
table("out_bridge_arr", "E1. ARR bridge, budget to forecast (step, GBP k, cumulative)", ["Step", "GBP k", "Cumulative"],
      [bvr(r_, "ABC", [None, "num", "num"]) for r_ in range(ARR_BR["first"] - 0, ARR_BR["last"] + 1)])
table("out_bridge_rev", "E2. Revenue bridge, budget to forecast", ["Step", "GBP k", "Cumulative"],
      [bvr(r_, "ABC", [None, "num", "num"]) for r_ in range(c0 + 2, REV_BR["last"] + 1)])
table("out_bridge_ebitda", "E3. EBITDA bridge, budget to forecast", ["Step", "GBP k", "Cumulative"],
      [bvr(r_, "ABC", [None, "num", "num"]) for r_ in range(e0 + 2, re_ + 1)])
table("out_football", "F. Football field, enterprise value GBP m (method, low, mid, high)", ["Method", "Low", "Mid", "High"],
      [[(f"=MA_Valuation!A{f0 + 2 + i}", None)] + [(f"=MA_Valuation!{c}{f0 + 2 + i}", "num1") for c in "CDE"] for i in range(len(FF))])
table("out_ppa", "G. Purchase price allocation (line, IFRS 3, US GAAP)", ["Line", "IFRS 3", "US GAAP"],
      [[(f"=MA_PPA!A{r_}", None), (f"=MA_PPA!C{r_}", "num"), (f"=MA_PPA!D{r_}", "num")] for r_ in sorted(PPA.values())])
table("out_checks", "H. Every integrity check (name, value, tolerance, pass)", ["Check", "Value", "Tolerance", "Pass"],
      [[(f"=Checks!A{5 + i}", None), (f"=Checks!C{5 + i}", "num2"), (f"=Checks!D{5 + i}", "num2"), (f"=Checks!E{5 + i}", "int")] for i in range(len(CHECKS))])

# ---- I. input registry
ow_sec("I. Input registry: every web-facing input (the site builds its input panel from this)")
hdr(ow, _o[0], ["Name", "Label", "Unit", "Live value", "Min", "Max", "Kind", "Scenario row", "Section", "Tooltip"])
_o[0] += 1
reg_first = _o[0]
_cur_sec = ""
for item in INPUTS:
    if item[0] == "sec":
        _cur_sec = item[1]
        continue
    kind, name, label, unit, fmt, base, up, down, lo, hi, tip = item
    dr = int(re.sub(r"\D", "", DRV_CELL[name]))
    put(ow, f"I{_o[0]}", _cur_sec)
    put(ow, f"J{_o[0]}", tip)
    put(ow, f"A{_o[0]}", name)
    put(ow, f"B{_o[0]}", label)
    put(ow, f"C{_o[0]}", unit, italic=True)
    put(ow, f"D{_o[0]}", f"=Inputs_Drivers!C{dr}", fmt)
    if lo is not None:
        put(ow, f"E{_o[0]}", f"=Inputs_Drivers!D{dr}", fmt)
        put(ow, f"F{_o[0]}", f"=Inputs_Drivers!E{dr}", fmt)
    put(ow, f"G{_o[0]}", "scenario" if kind == S else "static", italic=True)
    if kind == S:
        put(ow, f"H{_o[0]}", f"=ROW(Inputs_Scenarios!$A${SCEN_ROW[name]})", "int", italic=True)
    _o[0] += 1
defname("out_input_registry", f"Outputs_Web!$A${reg_first}:$J${_o[0] - 1}")
defname("scen_table", f"Inputs_Scenarios!$A$5:$F${max(SCEN_ROW.values())}")
defname("scen_names", "Inputs_Scenarios!$C$4:$F$4")
ow.freeze_panes = "D4"


# ---------------------------------------------------------------------------
# Cover
# ---------------------------------------------------------------------------
cv = wb.create_sheet("Cover", 0)
cv.sheet_view.showGridLines = False
put(cv, "A1", "Keel - SaaS operating model for Northgate Controls Ltd", bold=True)
cv["A1"].font = Font(name=FONT, size=16, bold=True, color="1F3A5F")
put(cv, "A2", "Synthetic data. A formula-driven FP&A model: ARR build, revenue recognition, three statements, SaaS unit economics.", italic=True)
put(cv, "A4", "Master check", bold=True)
put(cv, "B4", "=Checks!E" + str(mr + 1), bold=True)
put(cv, "A7", "Outlier warnings", bold=True)
put(cv, "B7", '=chk_warnings&" series flagged for review (see Checks)"')
put(cv, "A5", "Live scenario", bold=True)
put(cv, "B5", "=inp_scenario")
put(cv, "A6", "Units", bold=True)
put(cv, "B6", "GBP thousands unless stated. 24 monthly periods from the model start date, FY1 and FY2 annual columns.")
put(cv, "A8", "Colour convention", bold=True)
put(cv, "B8", "Blue = typed input", color=BLUE)
put(cv, "B9", "Black = formula")
put(cv, "B10", "Green = link to another tab", color=GREEN)
put(cv, "A12", "Model map", bold=True)
MAP = [
    ("Inputs_Drivers", "Named inputs (inp_*). Scenario selector. Live scenario values", "Phase 1"),
    ("Inputs_Scenarios", "Base / Upside / Downside / Custom values", "Phase 1"),
    ("Rev_Build", "ARR waterfall and customers by segment", "Phase 1"),
    ("Rev_Rec", "Billings, ratable revenue, deferred revenue, receivables", "Phase 1"),
    ("Headcount_Opex", "Headcount, payroll, opex by function", "Phase 1"),
    ("COGS", "Hosting, licences, support; gross margin by stream", "Phase 1"),
    ("IS / BS / CF", "Integrated three statements (balance check on BS)", "Phase 1"),
    ("SaaS_Metrics", "ARR bridge, NRR, GRR, CAC, payback, LTV/CAC, burn multiple, runway, Rule of 40", "Phase 1"),
    ("Checks", "Integrity tests and master flag", "Phase 1"),
    ("Contract_Costs", "Capitalised sales commission (IFRS 15), amortisation, S&M charge", "Phase 2"),
    ("Inputs_Budget, Budget_vs_Forecast", "Board budget as received; mapped variances and three bridges (ARR, revenue, EBITDA)", "Phase 2"),
    ("LT_Model", "Five-year view: years 1-2 linked, years 3-5 driver-based", "Phase 2"),
    ("Lever_Analysis, Sensitivity", "Impact of each driver; two-way grids (static, engine-run, staleness-checked)", "Phase 2"),
    ("Outputs_Web", "The only tab the website reads: scalar outputs, monthly series, five-year view, bridges, football field, PPA, checks and the input registry (all named out_*)", "Phase 5"),
    ("Business_Review", "One-page pack: KPIs, what moved vs budget, risks, downside alongside Base", "Phase 2"),
    ("MA_Target, MA_PPA, MA_ProForma", "Vaultline standalone; IFRS 3 purchase price allocation; combined entity with eliminations, synergies, integration", "Phase 3"),
    ("MA_Valuation, Inputs_Comps, MA_Returns", "DCF, comps, offer range, football field; value creation, payback, accretion, runway impact", "Phase 3"),
    ("MA_Target, MA_Valuation, MA_PPA, MA_ProForma, MA_Returns", "Vaultline acquisition module", "Phase 3 (not built)"),
]
for i, (t, d_, ph) in enumerate(MAP):
    put(cv, f"A{13 + i}", t)
    put(cv, f"B{13 + i}", d_)
    put(cv, f"C{13 + i}", ph, italic=True)
vr = 13 + len(MAP) + 1
put(cv, f"A{vr}", "Version log", bold=True)
put(cv, f"A{vr + 1}", "v0.1")
put(cv, f"B{vr + 1}", "Phase 1 MVP: ARR build, rev rec, opex, COGS, three statements, SaaS metrics, scenarios (Base/Upside/Downside/Custom), checks.")
put(cv, f"A{vr + 5}", "v0.5")
put(cv, f"B{vr + 5}", "Phase 5: Outputs_Web contract (named out_* ranges, input registry), array formulas replaced by helper rows, named_ranges.json export.")
put(cv, f"A{vr + 4}", "v0.4")
put(cv, f"B{vr + 4}", "Phase 4: global tolerance, sign / annual-column / range checks, outlier warnings (validation lives in validation/).")
put(cv, f"A{vr + 3}", "v0.3")
put(cv, f"B{vr + 3}", "Phase 3: Vaultline acquisition module (MA_* tabs, Inputs_Comps), acquisition scenario drivers, deal outcome in the scenario snapshots.")
put(cv, f"A{vr + 2}", "v0.2")
put(cv, f"B{vr + 2}", "Phase 2: Contract_Costs, budget and variance bridges, LT_Model, Lever_Analysis, Sensitivity, Business_Review. Fixed FY opening-balance columns.")
cv.column_dimensions["A"].width = 38
cv.column_dimensions["B"].width = 90
cv.column_dimensions["C"].width = 20

order = ["Cover", "Inputs_Drivers", "Inputs_Scenarios", "Rev_Build", "Rev_Rec", "Headcount_Opex", "Contract_Costs", "COGS",
         "IS", "BS", "CF", "SaaS_Metrics", "LT_Model", "Inputs_Budget", "Budget_vs_Forecast", "Lever_Analysis", "Sensitivity",
         "Business_Review", "MA_Target", "MA_PPA", "MA_ProForma", "Inputs_Comps", "MA_Valuation", "MA_Returns", "Outputs_Web", "Checks"]
wb._sheets = [wb[n] for n in order]
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
# Web copy: identical workbook without the football-field chart (the website draws its own and ExcelJS cannot read openpyxl chart drawings)
mv._charts = []
wb.save(OUT.replace(".xlsx", "_web.xlsx"))
json.dump({n: d.attr_text for n, d in wb.defined_names.items()}, open(os.path.join(_HERE, "named_ranges.json"), "w"), indent=1, sort_keys=True)
print("saved", OUT)
