# Report Context

This reference summarizes the project report `BMD5302_Robo_Adviser_Report.docx` and the workbook methodology sheets so Sandra can explain the robo-adviser without loading the full document.

The canonical runtime knowledge base for the browser/chat agent lives in `sandra_kb/`, especially `sandra_kb/methodology.md` and `sandra_kb/sandra_preprompt.md`.

## System summary

The report describes a workbook-backed robo-adviser with three linked parts:

- an efficient-frontier engine built from five years of monthly NAV data for ten FSMOne-listed funds
- a 100-question risk questionnaire that produces a personalized Arrow-Pratt risk-aversion value
- an Excel and VBA platform that turns the questionnaire result into an optimized portfolio recommendation

Sandra is the presentation layer around that workbook-backed engine. She should explain the system in a professional adviser voice, but the calculations still come from the workbook.

Sandra's intended personality is warm, practical, and methodical. She should sound human and helpful, not like a status console. Use client-facing language for users and reserve infrastructure terms for setup troubleshooting.

The stated user journey is:

1. Randomize 10 questionnaire items.
2. Record one answer per question.
3. Let the workbook derive the investor profile and risk aversion value.
4. Run the optimizer with or without short selling.
5. Review the final table, metrics, and charts.

## Utility model

The optimizer is framed around mean-variance utility:

```text
U(r, sigma) = r - (A / 2) * sigma^2
```

Where:

- `r` is expected portfolio return
- `sigma^2` is portfolio variance
- `A` is the Arrow-Pratt risk-aversion coefficient

Higher `A` means greater aversion to risk.

## Questionnaire design

The report says the questionnaire bank contains 100 items across 10 categories, with one question drawn from each category for each session. The draw is then shuffled so the user does not see category order directly.

Weight tiers:

- Tier A, weight `1.00`: utility-derived or revealed-preference questions
- Tier B, weight `0.70`: behavioral preference signals
- Tier C, weight `0.40`: situational moderators such as capacity or constraints

Named categories:

- Explicit Gambles
- Time Horizon and Liquidity
- Behavioural Loss Reactions
- Portfolio Allocation Preferences
- Forced Concentration
- Self-Description and Personality
- Experience and Knowledge
- Goals and Purpose
- Market Event Scenarios
- Situational Finance

## Risk profile bands

The report maps the final risk-aversion result into five bands:

- `1.0-2.0`: Aggressive, typical risky share `80%-100%`
- `2.0-4.0`: Moderately Aggressive, typical risky share `60%-80%`
- `4.0-6.0`: Moderate, typical risky share `40%-60%`
- `6.0-8.0`: Moderately Conservative, typical risky share `20%-40%`
- `8.0-10.0`: Conservative, typical risky share `0%-20%`

When explaining a workbook-generated profile, use these bands as report-backed context, but do not replace the actual workbook profile text.

## Short-selling interpretation

The report treats short selling as a research variant that relaxes the long-only constraint. Use that framing when explaining results:

- `RunOptimizer`: long-only, closer to a retail-fund-platform constraint set
- `RunOptimizerShortSelling`: allows negative weights and should be described as a less constrained research scenario

Do not present short selling as the default recommendation path unless the user explicitly chooses it.

## Reported platform buttons and features

The report identifies these platform features/buttons:

- `Randomise 10 Questions`
- `Clear Questionnaire`
- `Optimise (Long-Only)`
- `Optimise with Short-Selling`
- Efficient frontier charts
- Summary dashboard on `0_Summary`
- Live scoring audit on `Scoring`

For the current repo automation flow, the important button-linked macros are:

- `RandomizeQuestions`
- `RunOptimizer`
- `RunOptimizerShortSelling`
- `CalculateMVP`

## Reported benchmark findings

The report's headline findings can be used as explanatory background, not as replacement for current workbook outputs:

- GMVP long-only: return `2.47%`, standard deviation `4.73%`
- GMVP with short selling: return `3.28%`, standard deviation `3.61%`
- Tangency portfolio long-only: return `18.59%`, standard deviation `18.89%`, Sharpe ratio `0.799`

These figures are useful when the user asks about the project report or the design rationale, but any live run should be described from the workbook-generated outputs returned by the MCP flow.

## Workbook methodology notes

The workbook adds implementation details that are useful for methodology answers:

- `1_Questionnaire` displays the random 10-question draw and records answers.
- `10_Question_Bank` stores the 100 questions, option text, and implied A values.
- `11_Scoring` computes raw weighted A, draw-specific min/max, rescaled `A_final`, investor profile, and a live audit trail.
- `12_Optimizer` maximises utility using Solver with the objective `U = r - 0.5 * A * sigma^2`.
- `14_Methodology` documents the user workflow and scoring recipe.
- `15_QID_Bounds` stores the min/max implied A values used in draw-specific rescaling.

When explaining the calculation path, describe it as:

1. The workbook draws one question from each category.
2. Each answer maps to an implied A value.
3. The scoring sheet computes a weighted mean.
4. The score is rescaled to the current draw's achievable range.
5. The final A is used by Solver to maximise mean-variance utility.
