# Northgate Controls — FY1 Q3 business review

*Synthetic data. Base scenario. GBP. Source: `Business_Review` and `Budget_vs_Forecast` tabs of `keel_saas_model.xlsx`. The Q3 position is month 9 of the model.*

## Headline

ARR is on course to grow about 34% over FY1 (£40.0m to £53.4m) and the business is still loss-making, as planned. The year is tracking **£1.1m below budget on exit ARR** and **£1.2m below budget on EBITDA**. The ARR gap is fully explained by churn and new-logo volume. It is not a pricing problem.

| At the end of Q3 (month 9) | Value |
|---|---|
| Closing ARR | £50.0m (opening £40.0m) |
| Net revenue retention (annualised) | 102.7% |
| Gross revenue retention (annualised) | 91.9% |
| CAC payback | 18.8 months |
| LTV / CAC | 4.5x |
| Burn multiple | 0.24x |
| Gross margin | 76.6% |
| EBITDA margin (month 9) | −11.7% |
| Cash | £22.8m |

## What moved against budget (full-year view)

| Closing ARR, budget to forecast | £k |
|---|---|
| Budget closing ARR | 54,530 |
| Volume: 209 new customers vs 226 budgeted | −994 |
| Price: average new-customer ACV £58k vs £57k budgeted | +263 |
| Expansion | +569 |
| Churn: £3.9m lost vs £2.9m budgeted | −965 |
| **Forecast closing ARR** | **53,404** |

- **Churn is the biggest miss.** The budget assumed about £2.9m of churned ARR. Gross churn is running at 9.7% of opening ARR (Mid-market 12%, Enterprise 6%).
- **New-logo volume is 17 customers short.** That is the second-largest miss.
- **Pricing and expansion are ahead of budget.** They recover about £0.8m of the gap, which suggests the shortfall is a volume and retention problem, not a discounting one.
- **Revenue is £0.8m (1.7%) below budget** for the year. About 0.5 of each ARR effect reaches FY1 revenue, using the mid-year convention in the bridge.

## EBITDA

| FY1 EBITDA, budget to forecast | £k |
|---|---|
| Budget EBITDA | −5,750 |
| Revenue | −833 |
| Cost of sales | −215 |
| Sales & marketing | +465 |
| Research & development | −390 |
| General & administrative | −250 |
| **Forecast EBITDA** | **−6,972** |

The variance widens each quarter (Q1 −£0.19m, Q2 −£0.23m, Q3 −£0.35m, Q4 −£0.45m), because the revenue shortfall compounds while hiring runs ahead of the budgeted profile. **The favourable S&M line is partly an accounting effect:** commission is now capitalised under IFRS 15, so cash S&M is higher than the P&L shows.

## Risks

1. **Retention in Mid-market.** Mid-market NRR is below 100% on its own (7% expansion against 12% churn). Each extra point of churn in both segments takes about £1.2m off month-24 ARR (engine-run; see `Lever_Analysis`).
2. **CAC payback is 19 months, drifting up** as the sales team grows faster than gross new ARR. It would breach a 24-month limit in the Downside case (27 months).
3. **Hiring ahead of revenue.** Adding 0.5 S&M hires per month costs about £1.1m of month-24 cash (engine-run; see `Lever_Analysis`).
4. **Funding is not a near-term issue in Base**, but see the Downside below.

## Downside case alongside the point estimate

| Month 24 | Base | Downside |
|---|---|---|
| ARR | £68.3m | £50.2m |
| NRR (month 12) | 102.8% | 93.8% |
| CAC payback (month 12) | 18.8 months | 27.4 months |
| FY2 EBITDA | −£3.1m | −£7.6m |
| Cash | £22.6m | £9.5m |

In the Downside the business never runs out of cash within the 24 months (lowest month-end cash £9.5m), but it burns through about £14.5m of its £24m opening cash and NRR falls below 100%, so the installed base shrinks even before new sales. **Recommended triggers:** slow S&M and R&D hiring if NRR falls below 100% or CAC payback passes 24 months for two consecutive months.

## Caveats

- The budget used here is synthetic and was written to differ from the Base forecast in a plausible way.
- Revenue bridge splits use a mid-year convention; the "timing / other" residual absorbs what the ARR effects do not explain.
- The Downside and lever figures come from full recalculations of the workbook (static tables, checked for staleness on the `Checks` tab).
