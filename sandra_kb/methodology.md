# Sandra Methodology Knowledge Base

Sources:

- `BMD5302_Robo_Adviser_Report.docx`, extracted from the user-provided report
- `Model.xlsm`, especially `1_Questionnaire`, `10_Question_Bank`, `11_Scoring`, `12_Optimizer`, `14_Methodology`, and `15_QID_Bounds`

This knowledge base is explanatory context. Live recommendations must still come from `Model.xlsm`.

## System Overview

The project combines three parts:

- An efficient-frontier engine built from five years of monthly NAV data for ten FSMOne-listed funds.
- A risk-aversion questionnaire that converts investor preferences into an Arrow-Pratt risk-aversion coefficient `A`.
- An Excel/VBA platform that uses the questionnaire result and Solver to produce a utility-maximising portfolio.

The intended user journey is:

1. Draw 10 questionnaire items, one from each category.
2. Record one answer per question.
3. Let the workbook calculate the final risk-aversion score and investor profile.
4. Choose whether short selling is allowed.
5. Run the optimizer.
6. Review the final table, efficient-frontier chart, and allocation chart.

## Data And Fund Universe

The report uses monthly Net Asset Value data from April 2021 to April 2026, giving 60 observations per fund.

Monthly returns are computed as:

```text
R_t = P_t / P_(t-1) - 1
```

Statistics are annualised:

- Mean monthly returns are multiplied by 12.
- Monthly variances are multiplied by 12.
- Monthly standard deviations are multiplied by sqrt(12).

The risk-free rate used in the report is 3.50% per annum, aligned with the Singapore MAS-bill rate during the analysis period.

The ten-fund universe spans bonds, gold, equities, emerging markets, REITs, and global income:

| Fund | Asset class | Ann. return | Std dev | Sharpe |
| --- | --- | ---: | ---: | ---: |
| Amova SG Bond | SG Bond | 2.33% | 4.81% | -0.242 |
| BK World Gold | Commodity (Gold) | 26.54% | 33.33% | 0.691 |
| ES Asian Bond | Asian Bond | -6.38% | 7.72% | -1.280 |
| ES Asian Equity | Asian Equity | 6.85% | 18.48% | 0.182 |
| Fid World | Global Equity | 7.50% | 13.11% | 0.305 |
| Fid America | US Equity | 7.67% | 13.88% | 0.300 |
| Fid EM | Emerging Markets | 4.96% | 19.69% | 0.074 |
| LionGlobal SG | SG Equity | 15.24% | 16.29% | 0.721 |
| Manulife AP REIT | Asia-Pacific REITs | -8.68% | 17.05% | -0.714 |
| PIMCO Income | Global Bond | 2.98% | 5.73% | -0.090 |

## Efficient Frontier Methodology

The annualised covariance matrix is built from monthly fund returns. For portfolio weights `w`, annualised expected returns `mu`, covariance matrix `Sigma`, and risk-free rate `r_f`:

```text
portfolio return:   r_p = w' * mu
portfolio variance: sigma_p^2 = w' * Sigma * w
Sharpe ratio:       (r_p - r_f) / sigma_p
```

The efficient frontier is traced by solving for portfolios at 50 target return levels. The chart compares a long-only frontier with a short-selling-allowed frontier.

Interpretation:

- Bond funds have the lowest variance and low covariance with equities, so they provide much of the variance reduction in the GMVP.
- Equity funds are more correlated with one another, so diversification inside the equity bucket is more limited.
- BK World Gold has high variance and positive covariance with many equity funds; it behaves like a risk asset in this dataset despite its safe-haven label.
- Allowing short selling weakly improves the feasible frontier because it relaxes the long-only constraint.

## Report Benchmark Results

These are report benchmarks, not a substitute for a live workbook run:

| Portfolio | Return | Std dev | Sharpe |
| --- | ---: | ---: | ---: |
| GMVP, long-only | 2.47% | 4.73% | -0.218 |
| GMVP, short-selling allowed | 3.28% | 3.61% | -0.062 |
| Tangency portfolio, long-only | 18.59% | 18.89% | 0.799 |

The long-only GMVP is concentrated in Amova SG Bond and PIMCO Income. The tangency portfolio is concentrated in LionGlobal SG and BK World Gold. The report notes that the negative Sharpe ratios for the GMVP variants reflect a period when the 3.50% risk-free rate exceeded the returns achievable by minimum-variance risky portfolios.

## Utility Framework

The optimizer uses the mean-variance utility function:

```text
U = r - (A / 2) * sigma^2
```

Where:

- `r` is expected portfolio return.
- `sigma^2` is portfolio variance.
- `A` is the Arrow-Pratt coefficient of risk aversion.

A higher `A` means the investor is more sensitive to variance. The optimizer searches for portfolio weights that maximise utility for the investor's workbook-generated `A`.

The report uses the Merton risky-share formula as a calibration check:

```text
w* = (mu - r_f) / (A * sigma^2)
```

This helps justify the risk profile bands, but Sandra should not use it to override workbook outputs.

## Questionnaire Design

The question bank contains 100 items across 10 categories. Each session draws one question from each category, then shuffles the order so the user does not see the category structure.

The category weights are:

| # | Category | Weight | Role |
| ---: | --- | ---: | --- |
| 1 | Explicit Gambles | 1.00 | Utility-derived or revealed preference |
| 2 | Time Horizon and Liquidity | 0.40 | Situational moderator |
| 3 | Behavioural Loss Reactions | 0.70 | Behavioural preference |
| 4 | Portfolio Allocation Preferences | 0.70 | Behavioural preference |
| 5 | Forced Concentration | 1.00 | Revealed preference |
| 6 | Self-Description and Personality | 0.70 | Behavioural preference, Grable-Lytton anchored |
| 7 | Experience and Knowledge | 0.40 | Situational moderator |
| 8 | Goals and Purpose | 0.70 | Behavioural preference |
| 9 | Market Event Scenarios | 0.70 | Behavioural preference |
| 10 | Situational Finance | 0.40 | Situational moderator |

Weight tiers:

- Tier A, weight 1.00: utility-derived or revealed-preference questions.
- Tier B, weight 0.70: behavioural preference signals.
- Tier C, weight 0.40: situational moderators, such as capacity and constraints.

The report links the design to Grable and Lytton risk-tolerance work, while the workbook operationalises it through the question bank and scoring sheet.

## Scoring Methodology

Each answer has an implied `A` value. The scoring sheet computes a weighted mean from answered questions:

```text
raw_A = sum(weight_i * implied_A_i * answered_i) / sum(weight_i * answered_i)
```

The achievable range can vary depending on the 10 questions drawn, so the workbook rescales against the draw-specific minimum and maximum:

```text
raw_min = sum(weight_i * min_A_i * answered_i) / sum(weight_i * answered_i)
raw_max = sum(weight_i * max_A_i * answered_i) / sum(weight_i * answered_i)
A_final = 1 + 9 * (raw_A - raw_min) / (raw_max - raw_min)
```

The workbook clamps `A_final` to `[1, 10]` as a numerical safety net. The `11_Scoring` sheet includes a live audit trail with QID, category, weight, implied A, min/max A, and weighted contribution.

Sandra should explain this as: the workbook first scores the selected answers, then normalises the result so the current random question draw maps consistently into the 1-to-10 risk-aversion scale.

## Risk Profile Bands

The report maps `A_final` into five bands:

| A range | Profile | Typical risky share | Practical meaning |
| --- | --- | --- | --- |
| 1.0-2.0 | Aggressive | 80%-100% | Maximises return; high tolerance for drawdown |
| 2.0-4.0 | Moderately Aggressive | 60%-80% | Growth-oriented; comfortable with volatility |
| 4.0-6.0 | Moderate | 40%-60% | Balanced growth and capital preservation |
| 6.0-8.0 | Moderately Conservative | 20%-40% | Prefers stability; limited equity exposure |
| 8.0-10.0 | Conservative | 0%-20% | Capital preservation; very low risk tolerance |

Use these bands as explanatory context. If the workbook returns a profile, present the workbook profile first.

## Optimizer Methodology

The `12_Optimizer` sheet maximises:

```text
U = r_p - (A_final / 2) * sigma_p^2
```

Decision variables are the first nine fund weights. The tenth weight is dependent:

```text
w_10 = 1 - sum(w_1 ... w_9)
```

The return vector and covariance matrix come from Part 1 data sheets, so the optimizer uses the same inputs as the frontier analysis.

Solver variants:

| Variant | Macro | Interpretation |
| --- | --- | --- |
| Long-only | `RunOptimizer` | Weights constrained to `[0, 1]`; closer to a retail fund-platform assumption |
| Short-selling allowed | `RunOptimizerShortSelling` | Weights unrestricted; negative weights represent short positions; research/constraint-relaxation scenario |

The short-selling variant can improve utility because it relaxes a constraint. It should not be framed as more suitable unless the user explicitly chooses it and understands the practical implications.

## Platform And Workbook Notes

Important workbook sheets:

- `1_Questionnaire`: displays the random 10-question draw and records answers.
- `10_Question_Bank`: stores 100 questions, answer options, and implied A values.
- `11_Scoring`: computes raw score, rescaled `A_final`, investor profile, and audit trail.
- `12_Optimizer`: maximises utility using Solver.
- `14_Methodology`: documents the workbook workflow and scoring recipe.
- `15_QID_Bounds`: stores draw-specific min/max values used for rescaling.

Key macros:

- `RandomizeQuestions`: draws one question from each category and shuffles presentation.
- `ClearQuestions`: clears answers for a fresh assessment.
- `RunOptimizer`: runs the long-only Solver setup.
- `RunOptimizerShortSelling`: runs the unrestricted short-selling variant.
- `CalculateMVP`: generates the final calculator output and charts.

The workbook methodology sheet states that the full flow is designed to be completed in Excel with macro support enabled and Solver available.

## Ethical And Practical Limits

The report frames the platform as a decision-support tool, not a substitute for licensed financial advice.

Sandra should remind users when appropriate:

- The model is based on historical NAV data and selected assumptions.
- The questionnaire translates preferences into a structured risk score; it does not capture every financial circumstance.
- Users should reassess at least annually or after major life changes.
- For high-stakes or regulated investment decisions, users should consult a licensed financial adviser.

## Future Extension Ideas From The Report

The report suggests future improvements:

- Black-Litterman or Bayesian updating to reduce sensitivity to historical-return estimation.
- Natural-language risk profiling from open-ended investor responses.
- Dynamic rebalancing signals triggered by portfolio drift.
- Live NAV feeds for real-time frontier updates.

Treat these as roadmap ideas, not implemented features.

