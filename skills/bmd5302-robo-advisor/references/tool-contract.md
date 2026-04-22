# Tool Contract

This skill targets the repo's `Model.xlsm` MCP workflow.

## Tool order

1. `start_investor_questionnaire`
2. `submit_investor_questionnaire_answers` when the first tool does not complete answer collection itself
3. `run_investor_mvp` or `run_investor_mvp_with_chart_images`

## Start tool

`start_investor_questionnaire(workbook_path="Model.xlsm", output_dir="notebook_outputs", visible=False, use_elicitation=True)`

Expected behavior:

- creates a session-scoped workbook copy
- runs `RandomizeQuestions`
- extracts 10 questions from `1_Questionnaire`
- returns a `session_id`
- returns structured questions keyed like `q1` through `q10`
- may complete questionnaire answer collection immediately through MCP elicitation

Important returned fields:

- `status`
- `session_id`
- `workbook_copy`
- `metadata_path`
- `questions`
- `answer_submission_format`
- `elicitation_supported`
- `next_step` when manual follow-up is required

## Answer submit tool

`submit_investor_questionnaire_answers(session_id, answers, output_dir="notebook_outputs", visible=False)`

Expected behavior:

- writes answer letters into column `F`
- reads the workbook-generated investor profile from `G21`
- returns:
  - `answers`
  - `investor_profile`
  - `creative_profile_message`
  - `next_step`

The `answers` payload should be keyed by question key, for example:

```json
{
  "q1": "b",
  "q2": "a",
  "q3": "d"
}
```

## Final optimizer tools

`run_investor_mvp(session_id, allow_short_selling=None, output_dir="notebook_outputs", visible=False, use_elicitation=True)`

`run_investor_mvp_with_chart_images(...)`

Expected behavior:

- if `allow_short_selling` is missing, the tool may ask for it through elicitation
- otherwise it returns `status="short_selling_choice_required"` with a `next_step`
- runs:
  - `RunOptimizer` for no short selling
  - `RunOptimizerShortSelling` for short selling
- writes `2_MVP_Calculator!B6` as `Yes` or `No`
- runs `CalculateMVP`
- returns the final table from `A18:D28`
- returns both chart paths from sheet `2_MVP_Calculator`

When completed, the final payload includes:

- `status`
- `session_id`
- `allow_short_selling`
- `investor_profile`
- `summary_range`
- `summary_table_matrix`
- `summary_table_columns`
- `summary_table_records`
- `chart_paths`

`run_investor_mvp_with_chart_images` additionally returns two MCP image blocks when the run completes.

## Session storage

The workflow is disk-backed, not memory-only.

Each run lives under:

- `notebook_outputs/model_sessions/<session_id>/`

Each session directory contains:

- a copied workbook
- `session.json`
- chart PNGs under `charts/`

## Chart and workbook rules

- Use workbook outputs as authoritative.
- Do not recompute table values or chart data in chat.
- The final chart names are:
  - `MVP_FrontierChart`
  - `OptimalWeight_Chart`
