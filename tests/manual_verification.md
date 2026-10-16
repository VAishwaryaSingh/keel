# Manual verification (build principle 10; Phase 5 task)

**Goal:** rebuild three calculations from the inputs alone, on blank sheets, and confirm they match the model. The point is that you understand them, not that a script agrees. Do **not** read the workbook's formulas first, and do not ask a tool to write the formulas for you.

**Kit:** `tests/manual_rebuild_template.xlsx` holds the Base-scenario inputs (sheet `Inputs`), the model's answers (sheet `Compare`, column B), and three empty sheets for your work. Type into `Compare` column C a formula that points at your result; columns D and E report the difference and `MATCH` / `DIFFERENT`. Tolerance: 0.01 GBP k.

## 1. ARR bridge, months 1 to 24, by segment

For each segment, month by month:

- *Volume.* New logos per month grow at the annual new-logo growth rate, compounding monthly, so month t has `base logos x (1 + growth)^((t-1)/12)`.
- *Price.* The ACV of a new logo steps up by the price uplift at the start of each model year (months 13-24 have one uplift).
- *New-logo ARR* = logos x ACV.
- *Rates.* The annual churn and expansion inputs are converted to monthly compounding rates: monthly churn = 1 - (1 - annual)^(1/12); monthly expansion = (1 + annual)^(1/12) - 1.
- *Expansion* and *churn* in a month are the monthly rates applied to that month's **opening** ARR (churn is a negative).
- Closing ARR = opening + new + expansion + churn. Opening ARR of month 1 is the opening input.

## 2. Subscription deferred revenue roll-forward

Both segments are billed **annually upfront** and recognised **evenly over 12 months, starting in the month of invoice**.

- *Invoices.* In month t the invoice is that month's new-logo ARR (annual price paid upfront) **plus** the renewal of the invoice raised 12 months earlier, scaled by (1 + annual expansion - annual churn). For months 1 to 12, the invoice "12 months earlier" is the opening ARR divided by 12 (opening-year customers renew evenly through the year).
- *Revenue* in month t = one twelfth of the sum of the invoices raised in months t-11 to t (including months before the model starts: each of those 12 history months is opening ARR / 12).
- *Opening deferred revenue.* An invoice raised k months before the opening date (k = 0 for the last month) still has (11 - k)/12 of its value unrecognised at the opening date. Sum that across the 12 history invoices.
- Roll-forward: closing = opening + invoices - revenue.

## 3. Goodwill (IFRS 3 and US GAAP)

Use the Vaultline inputs on the `Inputs` sheet. Steps in words:

- Vaultline's opening deferred revenue is built the same way as in section 2 (opening ARR / 12 per history month, same unexpired fractions).
- Payables = opening ARR / 12 x (sum of the four cost percentages) x payables months.
- Book net assets = cash + receivables + PP&E - payables - deferred revenue.
- Enterprise value = offer multiple x opening ARR. Equity price = EV + cash acquired.
- Customer relationships and technology are percentages of EV. A deferred tax liability = tax rate x those two intangibles.
- IFRS 3 writes deferred revenue down by the haircut percentage (the liability falls, so net assets rise). US GAAP after ASU 2021-08 does not.
- Goodwill = equity price - fair value of identifiable net assets.

## Targets (Base scenario, GBP k)

| Item | Model value |
|---|---|
| ARR bridge, FY1: new-logo ARR, Enterprise | 7,731.72 |
| ARR bridge, FY1: expansion, Enterprise | 4,009.35 |
| ARR bridge, FY1: churn, Enterprise | -1,878.17 |
| ARR bridge: closing ARR at month 12, Enterprise | 35,862.91 |
| ARR bridge: closing ARR at month 24, Enterprise | 47,089.70 |
| ARR bridge, FY1: new-logo ARR, Mid-market | 4,467.22 |
| ARR bridge, FY1: expansion, Mid-market | 1,059.96 |
| ARR bridge, FY1: churn, Mid-market | -1,986.44 |
| ARR bridge: closing ARR at month 12, Mid-market | 17,540.74 |
| ARR bridge: closing ARR at month 24, Mid-market | 21,231.54 |
| ARR bridge: total closing ARR at month 24 | 68,321.24 |
| Deferred revenue: subscription invoices raised in FY1 (both segments) | 53,578.94 |
| Deferred revenue: subscription revenue recognised in FY1 | 47,306.01 |
| Deferred revenue: opening subscription deferred revenue | 18,333.33 |
| Deferred revenue: closing subscription deferred revenue, month 12 | 24,606.27 |
| Deferred revenue: closing subscription deferred revenue, month 24 | 31,613.92 |
| Goodwill: equity purchase price | 30,500.00 |
| Goodwill: fair value of identifiable net assets, IFRS 3 | 11,330.00 |
| Goodwill: IFRS 3 | 19,170.00 |
| Goodwill: US GAAP (ASU 2021-08) | 19,995.00 |

## When something does not match

Find the first month or line where you and the model diverge, then work out which of you is wrong. Either answer is useful: a bug in the model belongs in `docs/ai_usage.md`; a gap in your understanding belongs in your own notes before you defend this model to anyone. Record the outcome below.

| Date | Section | Matched? | Notes / what you found |
|---|---|---|---|
| | ARR bridge | | |
| | Deferred revenue | | |
| | Goodwill | | |
