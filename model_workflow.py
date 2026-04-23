from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
import json
from pathlib import Path
import re
import string
import time
from typing import Any, Callable
from uuid import uuid4

import pandas as pd
from pydantic import BaseModel, Field, create_model
import xlwings as xw

from excel_workbook_support import (
    ExcelChartExporter,
    ExcelChartSpec,
    call_vba_macro,
    log_excel_exception,
)


ADVISOR_NAME = "Sandra"


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _sanitize_filename(value: str) -> str:
    sanitized = re.sub(r"[^A-Za-z0-9]+", "_", value).strip("_").lower()
    return sanitized or "artifact"


def _normalize_option_lines(raw_value: Any) -> list[str]:
    if raw_value is None:
        return []

    if isinstance(raw_value, list):
        values = raw_value
    else:
        values = str(raw_value).splitlines()

    return [line.strip() for line in values if str(line).strip()]


def _normalize_answer_letter(value: str) -> str:
    match = re.search(r"[A-Za-z]", value)
    if not match:
        raise ValueError(f"'{value}' does not contain a valid answer choice.")
    return match.group(0).lower()


@dataclass(frozen=True, slots=True)
class ModelWorkbookContract:
    questionnaire_sheet: str = "1_Questionnaire"
    questionnaire_macro_name: str = "RandomizeQuestions"
    question_start_row: int = 9
    question_end_row: int = 18
    question_text_column: str = "D"
    question_options_column: str = "E"
    answer_column: str = "F"
    investor_profile_cell: str = "G21"
    optimizer_sheet: str = "12_Optimizer"
    no_short_macro_name: str = "RunOptimizer"
    short_macro_name: str = "RunOptimizerShortSelling"
    calculator_sheet: str = "2_MVP_Calculator"
    calculator_target_sigma_cell: str = "B5"
    calculator_stats_range: str = "B12:B15"
    short_selling_cell: str = "B6"
    calculator_macro_name: str = "CalculateMVP"
    calculator_variance_cell: str = "B14"
    calculator_weight_range: str = "C19:C28"
    summary_range: str = "A18:D28"
    chart_names: tuple[str, str] = ("MVP_FrontierChart", "OptimalWeight_Chart")

    def question_rows(self) -> range:
        return range(self.question_start_row, self.question_end_row + 1)


@dataclass(frozen=True, slots=True)
class InvestorQuestion:
    key: str
    row: int
    prompt: str
    options: list[str]
    option_letters: list[str]

    def to_dict(self) -> dict[str, Any]:
        return {
            "key": self.key,
            "row": self.row,
            "prompt": self.prompt,
            "options": list(self.options),
            "option_letters": list(self.option_letters),
            "answer_prompt": f"Choose one: {', '.join(self.option_letters)}",
        }


@dataclass(slots=True)
class QuestionnaireSessionState:
    session_id: str
    source_workbook_path: str
    workbook_copy: str
    use_source_workbook: bool
    chart_dir: str
    metadata_path: str
    created_at: str
    updated_at: str
    status: str
    questions: list[InvestorQuestion]
    answers: dict[str, str] = field(default_factory=dict)
    investor_profile: str | None = None
    allow_short_selling: bool | None = None
    chart_paths: dict[str, str] = field(default_factory=dict)
    summary_table_matrix: list[list[Any]] = field(default_factory=list)
    summary_table_columns: list[str] = field(default_factory=list)
    summary_table_records: list[dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "session_id": self.session_id,
            "source_workbook_path": self.source_workbook_path,
            "workbook_copy": self.workbook_copy,
            "use_source_workbook": self.use_source_workbook,
            "chart_dir": self.chart_dir,
            "metadata_path": self.metadata_path,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "status": self.status,
            "questions": [question.to_dict() for question in self.questions],
            "answers": dict(self.answers),
            "investor_profile": self.investor_profile,
            "allow_short_selling": self.allow_short_selling,
            "chart_paths": dict(self.chart_paths),
            "summary_table_matrix": self.summary_table_matrix,
            "summary_table_columns": self.summary_table_columns,
            "summary_table_records": self.summary_table_records,
        }

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> QuestionnaireSessionState:
        return cls(
            session_id=str(payload["session_id"]),
            source_workbook_path=str(payload["source_workbook_path"]),
            workbook_copy=str(payload["workbook_copy"]),
            use_source_workbook=bool(payload.get("use_source_workbook", False)),
            chart_dir=str(payload["chart_dir"]),
            metadata_path=str(payload["metadata_path"]),
            created_at=str(payload["created_at"]),
            updated_at=str(payload["updated_at"]),
            status=str(payload["status"]),
            questions=[
                InvestorQuestion(
                    key=str(question["key"]),
                    row=int(question["row"]),
                    prompt=str(question["prompt"]),
                    options=[str(option) for option in question["options"]],
                    option_letters=[str(letter) for letter in question["option_letters"]],
                )
                for question in payload.get("questions", [])
            ],
            answers={str(key): str(value) for key, value in payload.get("answers", {}).items()},
            investor_profile=(
                None
                if payload.get("investor_profile") is None
                else str(payload["investor_profile"])
            ),
            allow_short_selling=payload.get("allow_short_selling"),
            chart_paths={
                str(key): str(value)
                for key, value in payload.get("chart_paths", {}).items()
            },
            summary_table_matrix=[
                [value for value in row] for row in payload.get("summary_table_matrix", [])
            ],
            summary_table_columns=[
                str(value) for value in payload.get("summary_table_columns", [])
            ],
            summary_table_records=[
                {str(key): value for key, value in row.items()}
                for row in payload.get("summary_table_records", [])
            ],
        )


@dataclass(frozen=True, slots=True)
class ModelSessionPaths:
    session_id: str
    source_workbook_path: Path
    output_dir: Path
    session_dir: Path
    workbook_copy: Path
    chart_dir: Path
    metadata_path: Path

    @classmethod
    def create(
        cls,
        workbook_path: str | Path = "Model.xlsm",
        output_dir: str | Path = "notebook_outputs",
        session_id: str | None = None,
        use_source_workbook: bool = True,
    ) -> ModelSessionPaths:
        resolved_workbook_path = Path(workbook_path).expanduser().resolve()
        resolved_output_dir = Path(output_dir).expanduser().resolve()
        resolved_output_dir.mkdir(parents=True, exist_ok=True)

        resolved_session_id = session_id or uuid4().hex
        session_dir = resolved_output_dir / "model_sessions" / resolved_session_id
        chart_dir = session_dir / "charts"
        session_dir.mkdir(parents=True, exist_ok=True)
        chart_dir.mkdir(parents=True, exist_ok=True)

        return cls(
            session_id=resolved_session_id,
            source_workbook_path=resolved_workbook_path,
            output_dir=resolved_output_dir,
            session_dir=session_dir,
            workbook_copy=resolved_workbook_path,
            chart_dir=chart_dir,
            metadata_path=session_dir / "session.json",
        )

    @classmethod
    def from_state(cls, state: QuestionnaireSessionState) -> ModelSessionPaths:
        workbook_copy = Path(state.workbook_copy).expanduser().resolve()
        metadata_path = Path(state.metadata_path).expanduser().resolve()
        chart_dir = Path(state.chart_dir).expanduser().resolve()
        session_dir = metadata_path.parent
        output_dir = session_dir.parent.parent

        return cls(
            session_id=state.session_id,
            source_workbook_path=Path(state.source_workbook_path).expanduser().resolve(),
            output_dir=output_dir,
            session_dir=session_dir,
            workbook_copy=workbook_copy,
            chart_dir=chart_dir,
            metadata_path=metadata_path,
        )

    def chart_path(self, chart_name: str) -> Path:
        return self.chart_dir / f"{_sanitize_filename(chart_name)}.png"


@dataclass(frozen=True, slots=True)
class MvpWorkflowProgressStore:
    session_id: str
    progress_path: Path

    @classmethod
    def create(
        cls,
        *,
        session_id: str,
        output_dir: str | Path = "notebook_outputs",
    ) -> MvpWorkflowProgressStore:
        resolved_output_dir = Path(output_dir).expanduser().resolve()
        session_dir = resolved_output_dir / "model_sessions" / session_id
        session_dir.mkdir(parents=True, exist_ok=True)
        return cls(
            session_id=session_id,
            progress_path=session_dir / "mvp_progress.json",
        )

    def snapshot(self) -> dict[str, Any] | None:
        if not self.progress_path.exists():
            return None
        try:
            payload = json.loads(self.progress_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return None
        return payload if isinstance(payload, dict) else None

    def emit(
        self,
        *,
        stage: str,
        message: str,
        status: str = "working",
        allow_short_selling: bool | None = None,
    ) -> dict[str, Any]:
        current = self.snapshot() or {}
        sequence = int(current.get("sequence") or 0) + 1
        payload = {
            "session_id": self.session_id,
            "status": status,
            "stage": stage,
            "message": message,
            "sequence": sequence,
            "updated_at": _utc_now_iso(),
            "allow_short_selling": allow_short_selling,
        }
        self.progress_path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = self.progress_path.with_name(f"{self.progress_path.name}.tmp")
        temp_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        temp_path.replace(self.progress_path)
        return payload


class ShortSellingChoice(BaseModel):
    allow_short_selling: bool = Field(
        title="Allow short selling?",
        description="Select Yes to allow short selling, or No for long-only optimization.",
    )


class ModelWorkbookRunner:
    def __init__(self, contract: ModelWorkbookContract | None = None) -> None:
        self.contract = contract or ModelWorkbookContract()

    def start_questionnaire_session(
        self,
        workbook_path: str | Path = "Model.xlsm",
        output_dir: str | Path = "notebook_outputs",
        *,
        visible: bool = False,
        use_source_workbook: bool = True,
    ) -> QuestionnaireSessionState:
        paths = ModelSessionPaths.create(
            workbook_path=workbook_path,
            output_dir=output_dir,
            use_source_workbook=True,
        )
        if not paths.source_workbook_path.exists():
            raise FileNotFoundError(f"Workbook not found: {paths.source_workbook_path}")

        try:
            with xw.App(visible=visible, add_book=False) as app:
                app.display_alerts = False
                app.screen_updating = False

                book = app.books.open(str(paths.workbook_copy))
                try:
                    call_vba_macro(book, self.contract.questionnaire_macro_name)
                    book.app.calculate()
                    book.save()
                    questions = self._extract_questions(
                        book.sheets[self.contract.questionnaire_sheet]
                    )
                finally:
                    book.close()
        except Exception as exc:
            log_excel_exception(exc)
            raise RuntimeError(
                "Excel automation failed while starting the Model.xlsm questionnaire session. "
                "Make sure Microsoft Excel is installed, macOS has granted automation "
                "access to the current Python or terminal host, and macros are enabled "
                "when the workbook opens."
            ) from exc

        timestamp = _utc_now_iso()
        state = QuestionnaireSessionState(
            session_id=paths.session_id,
            source_workbook_path=str(paths.source_workbook_path),
            workbook_copy=str(paths.workbook_copy),
            use_source_workbook=True,
            chart_dir=str(paths.chart_dir),
            metadata_path=str(paths.metadata_path),
            created_at=timestamp,
            updated_at=timestamp,
            status="questions_generated",
            questions=questions,
        )
        self._save_session_state(state)
        return state

    def submit_answers(
        self,
        session_id: str,
        answers: dict[str, str],
        output_dir: str | Path = "notebook_outputs",
        *,
        visible: bool = False,
    ) -> QuestionnaireSessionState:
        state = self.load_session_state(session_id=session_id, output_dir=output_dir)
        paths = ModelSessionPaths.from_state(state)
        normalized_answers = self._normalize_answers(state.questions, answers)

        try:
            with xw.App(visible=visible, add_book=False) as app:
                app.display_alerts = False
                app.screen_updating = False

                book = app.books.open(str(paths.workbook_copy))
                try:
                    sheet = book.sheets[self.contract.questionnaire_sheet]
                    for question in state.questions:
                        answer_cell = f"{self.contract.answer_column}{question.row}"
                        sheet.range(answer_cell).value = normalized_answers[question.key]

                    book.app.calculate()
                    book.save()
                    investor_profile = self._read_investor_profile(sheet)
                finally:
                    book.close()
        except Exception as exc:
            log_excel_exception(exc)
            raise RuntimeError(
                "Excel automation failed while writing questionnaire answers. "
                "Make sure Microsoft Excel is installed, macOS has granted automation "
                "access to the current Python or terminal host, and macros are enabled "
                "when the copied workbook opens."
            ) from exc

        state.answers = normalized_answers
        state.investor_profile = investor_profile
        state.status = "answers_recorded"
        state.updated_at = _utc_now_iso()
        self._save_session_state(state)
        return state

    def run_mvp(
        self,
        session_id: str,
        allow_short_selling: bool,
        output_dir: str | Path = "notebook_outputs",
        *,
        visible: bool = False,
        progress_callback: Callable[[str, str], None] | None = None,
    ) -> QuestionnaireSessionState:
        state = self.load_session_state(session_id=session_id, output_dir=output_dir)
        if not state.answers:
            raise ValueError(
                "This questionnaire session does not have recorded answers yet. "
                "Submit answers before running the MVP workflow."
            )

        paths = ModelSessionPaths.from_state(state)

        def emit_progress(stage: str, message: str) -> None:
            if progress_callback is not None:
                progress_callback(stage, message)

        try:
            emit_progress("opening_excel", "Opening the workbook in Excel.")
            with xw.App(visible=visible, add_book=False) as app:
                app.display_alerts = False
                app.screen_updating = False

                book = app.books.open(str(paths.workbook_copy))
                try:
                    emit_progress("running_optimizer", "Running the optimizer sheet.")
                    self._run_optimizer(book, allow_short_selling=allow_short_selling)
                    calculator_sheet = book.sheets[self.contract.calculator_sheet]
                    # CalculateMVP's VBA uses Solver on the calculator sheet's
                    # own model (C19:C28) and assumes that sheet context is live.
                    # The optimizer macro leaves 12_Optimizer active, so sync and
                    # activate the calculator sheet first to match button-driven use.
                    emit_progress(
                        "syncing_calculator",
                        "Switching to the calculator sheet and syncing workbook formulas.",
                    )
                    book.app.calculate()
                    calculator_sheet.activate()
                    time.sleep(0.25)
                    calculator_sheet.range(self.contract.short_selling_cell).value = (
                        "Yes" if allow_short_selling else "No"
                    )
                    emit_progress(
                        "running_calculator_macro",
                        "Running the calculator macro for the final portfolio weights.",
                    )
                    call_vba_macro(book, self.contract.calculator_macro_name)
                    book.app.calculate()
                    emit_progress(
                        "waiting_for_outputs",
                        "Waiting for the workbook outputs to settle.",
                    )
                    self._wait_for_mvp_outputs(
                        calculator_sheet,
                        progress_callback=emit_progress,
                    )
                    emit_progress("saving_workbook", "Saving the workbook outputs.")
                    book.save()

                    emit_progress("reading_summary", "Reading the final portfolio table.")
                    summary_matrix, summary_columns, summary_records = self._read_summary_table(
                        calculator_sheet
                    )
                    emit_progress("exporting_charts", "Exporting the workbook charts.")
                    chart_paths = self._export_final_charts(book, paths)
                finally:
                    book.close()
        except Exception as exc:
            log_excel_exception(exc)
            raise RuntimeError(
                "Excel automation failed while running the MVP workflow. Make sure "
                "Microsoft Excel is installed, macOS has granted automation access to "
                "the current Python or terminal host, and macros are enabled when the "
                "copied workbook opens."
            ) from exc

        state.allow_short_selling = allow_short_selling
        state.chart_paths = chart_paths
        state.summary_table_matrix = summary_matrix
        state.summary_table_columns = summary_columns
        state.summary_table_records = summary_records
        state.status = "completed"
        state.updated_at = _utc_now_iso()
        self._save_session_state(state)
        emit_progress("completed", "Portfolio output is ready.")
        return state

    def load_session_state(
        self,
        session_id: str,
        output_dir: str | Path = "notebook_outputs",
    ) -> QuestionnaireSessionState:
        resolved_output_dir = Path(output_dir).expanduser().resolve()
        metadata_path = resolved_output_dir / "model_sessions" / session_id / "session.json"
        if not metadata_path.exists():
            raise FileNotFoundError(f"Questionnaire session not found: {metadata_path}")

        payload = json.loads(metadata_path.read_text(encoding="utf-8"))
        return QuestionnaireSessionState.from_dict(payload)

    def build_questionnaire_elicitation_model(
        self,
        questions: list[InvestorQuestion],
    ) -> type[BaseModel]:
        fields: dict[str, Any] = {}
        for index, question in enumerate(questions, start=1):
            option_lines = "\n".join(
                f"{letter}) {option}"
                for letter, option in zip(question.option_letters, question.options, strict=True)
            )
            fields[question.key] = (
                str,
                Field(
                    title=f"Question {index}",
                    description=(
                        f"{question.prompt}\n\nOptions:\n{option_lines}\n\n"
                        f"Reply with one letter: {', '.join(question.option_letters)}"
                    ),
                ),
            )

        return create_model("InvestorQuestionnaireAnswers", **fields)

    def serialize_start_payload(self, state: QuestionnaireSessionState) -> dict[str, Any]:
        return {
            "advisor_name": ADVISOR_NAME,
            "status": state.status,
            "session_id": state.session_id,
            "source_workbook_path": state.source_workbook_path,
            "workbook_copy": state.workbook_copy,
            "use_source_workbook": state.use_source_workbook,
            "metadata_path": state.metadata_path,
            "questions": [question.to_dict() for question in state.questions],
            "answer_submission_format": {
                "description": "Submit answers keyed by question key using answer letters.",
                "example": {
                    question.key: question.option_letters[0] for question in state.questions
                },
            },
        }

    def serialize_profile_payload(self, state: QuestionnaireSessionState) -> dict[str, Any]:
        if state.investor_profile is None:
            raise ValueError("Investor profile is not available for this session yet.")

        return {
            "advisor_name": ADVISOR_NAME,
            "status": state.status,
            "session_id": state.session_id,
            "workbook_copy": state.workbook_copy,
            "use_source_workbook": state.use_source_workbook,
            "answers": dict(state.answers),
            "investor_profile": state.investor_profile,
            "creative_profile_message": self.build_profile_message(state.investor_profile),
            "next_step": "Ask the user whether they want short selling (Yes or No).",
        }

    def serialize_final_payload(self, state: QuestionnaireSessionState) -> dict[str, Any]:
        if not state.summary_table_matrix:
            raise ValueError("Final workbook outputs are not available for this session yet.")

        return {
            "advisor_name": ADVISOR_NAME,
            "status": state.status,
            "session_id": state.session_id,
            "workbook_copy": state.workbook_copy,
            "use_source_workbook": state.use_source_workbook,
            "allow_short_selling": state.allow_short_selling,
            "investor_profile": state.investor_profile,
            "summary_range": self.contract.summary_range,
            "summary_table_matrix": state.summary_table_matrix,
            "summary_table_columns": state.summary_table_columns,
            "summary_table_records": state.summary_table_records,
            "chart_paths": dict(state.chart_paths),
        }

    @staticmethod
    def build_profile_message(investor_profile: str) -> str:
        cleaned_profile = investor_profile.strip()
        lowered_profile = cleaned_profile.lower()

        if "conservative" in lowered_profile:
            guidance = (
                "This points to a preference for capital preservation, steadier portfolio "
                "behavior, and a lower overall risk budget."
            )
        elif "moderate" in lowered_profile or "balanced" in lowered_profile:
            guidance = (
                "This suggests a balanced stance that still values growth, but with visible "
                "attention to downside control."
            )
        elif "aggressive" in lowered_profile or "growth" in lowered_profile:
            guidance = (
                "This indicates a stronger willingness to accept volatility in pursuit of "
                "higher long-term return potential."
            )
        else:
            guidance = (
                "This points to a distinct investment stance that should guide the optimizer "
                "settings and portfolio interpretation."
            )

        return (
            f"Sandra's workbook-generated assessment classifies the investor profile as "
            f"{cleaned_profile}. {guidance} The next step is to decide whether short selling "
            "should be enabled for the optimizer run."
        )

    def _save_session_state(self, state: QuestionnaireSessionState) -> None:
        metadata_path = Path(state.metadata_path).expanduser().resolve()
        metadata_path.parent.mkdir(parents=True, exist_ok=True)
        metadata_path.write_text(json.dumps(state.to_dict(), indent=2), encoding="utf-8")

    def _extract_questions(self, sheet: xw.Sheet) -> list[InvestorQuestion]:
        question_range = (
            f"{self.contract.question_text_column}{self.contract.question_start_row}:"
            f"{self.contract.question_options_column}{self.contract.question_end_row}"
        )
        rows = sheet.range(question_range).options(ndim=2).value
        if rows is None:
            raise RuntimeError("The questionnaire range is empty after question randomization.")

        questions: list[InvestorQuestion] = []
        for offset, row in enumerate(rows, start=0):
            prompt, raw_options = row
            prompt_text = "" if prompt is None else str(prompt).strip()
            options = _normalize_option_lines(raw_options)
            row_number = self.contract.question_start_row + offset

            if not prompt_text:
                raise RuntimeError(f"Question text is blank at row {row_number}.")
            if len(options) < 2:
                raise RuntimeError(
                    f"Question at row {row_number} does not have at least two options."
                )

            option_letters = list(string.ascii_lowercase[: len(options)])
            questions.append(
                InvestorQuestion(
                    key=f"q{offset + 1}",
                    row=row_number,
                    prompt=prompt_text,
                    options=options,
                    option_letters=option_letters,
                )
            )

        expected_questions = len(list(self.contract.question_rows()))
        if len(questions) != expected_questions:
            raise RuntimeError(
                f"Expected {expected_questions} questions but extracted {len(questions)}."
            )

        return questions

    def _normalize_answers(
        self,
        questions: list[InvestorQuestion],
        answers: dict[str, str],
    ) -> dict[str, str]:
        normalized_answers: dict[str, str] = {}

        for question in questions:
            raw_answer = answers.get(question.key)
            if raw_answer is None:
                raise ValueError(f"Missing answer for {question.key}.")

            answer_text = str(raw_answer).strip()
            answer_letter = _normalize_answer_letter(answer_text)

            if answer_letter in question.option_letters:
                normalized_answers[question.key] = answer_letter
                continue

            lowered_answer_text = answer_text.lower()
            for letter, option in zip(question.option_letters, question.options, strict=True):
                if lowered_answer_text == option.lower():
                    normalized_answers[question.key] = letter
                    break
            else:
                raise ValueError(
                    f"'{raw_answer}' is not valid for {question.key}. "
                    f"Expected one of: {', '.join(question.option_letters)}."
                )

        return normalized_answers

    def _read_investor_profile(self, sheet: xw.Sheet) -> str:
        value = sheet.range(self.contract.investor_profile_cell).value
        if value is None or not str(value).strip():
            raise RuntimeError(
                f"Investor profile cell {self.contract.investor_profile_cell} is blank."
            )
        return str(value).strip()

    def _run_optimizer(self, book: xw.Book, *, allow_short_selling: bool) -> None:
        optimizer_sheet = book.sheets[self.contract.optimizer_sheet]
        optimizer_sheet.activate()
        macro_name = (
            self.contract.short_macro_name
            if allow_short_selling
            else self.contract.no_short_macro_name
        )
        call_vba_macro(book, macro_name)

    def _read_summary_table(
        self,
        sheet: xw.Sheet,
    ) -> tuple[list[list[Any]], list[str], list[dict[str, Any]]]:
        matrix = sheet.range(self.contract.summary_range).options(ndim=2).value
        if matrix is None:
            return [], [], []

        normalized_matrix = [[value for value in row] for row in matrix]
        if not normalized_matrix:
            return [], [], []

        first_row = normalized_matrix[0]
        has_header_row = all(
            isinstance(value, str) and value.strip()
            for value in first_row
        )

        if has_header_row:
            columns = [str(value) for value in first_row]
            data_rows = normalized_matrix[1:]
        else:
            columns = [f"column_{index}" for index in range(1, len(first_row) + 1)]
            data_rows = normalized_matrix

        dataframe = pd.DataFrame(data_rows, columns=columns)
        return normalized_matrix, columns, dataframe.to_dict(orient="records")

    def _wait_for_mvp_outputs(
        self,
        sheet: xw.Sheet,
        *,
        timeout_seconds: float = 15.0,
        poll_interval_seconds: float = 0.25,
        progress_callback: Callable[[str, str], None] | None = None,
    ) -> None:
        deadline = time.monotonic() + timeout_seconds
        next_progress_ping = time.monotonic() + 2.0
        previous_snapshot: tuple[Any, ...] | None = None
        stable_snapshot_reads = 0
        last_stats_values: Any = None
        last_weight_values: Any = None
        last_summary_matrix: Any = None

        # Wait for the workbook-owned calculator stats, weights, and final
        # summary output to settle before Python reads the final table.
        while True:
            sheet.book.app.calculate()

            stats_values = sheet.range(self.contract.calculator_stats_range).options(
                ndim=1
            ).value
            summary_matrix = sheet.range(self.contract.summary_range).options(ndim=2).value
            weight_values = sheet.range(self.contract.calculator_weight_range).options(
                ndim=1
            ).value

            snapshot = (
                self._freeze_workbook_value(stats_values),
                self._freeze_workbook_value(summary_matrix),
                self._freeze_workbook_value(weight_values),
            )
            if snapshot == previous_snapshot:
                stable_snapshot_reads += 1
            else:
                stable_snapshot_reads = 1
                previous_snapshot = snapshot

            last_stats_values = stats_values
            last_weight_values = weight_values
            last_summary_matrix = summary_matrix
            if stable_snapshot_reads >= 2 and self._mvp_snapshot_ready(
                stats_values=stats_values,
                weight_values=weight_values,
                summary_matrix=summary_matrix,
            ):
                return

            if time.monotonic() >= deadline:
                raise RuntimeError(
                    "Workbook outputs did not settle after CalculateMVP. "
                    f"Stats range {self.contract.calculator_stats_range}={last_stats_values!r}, "
                    f"weight range {self.contract.calculator_weight_range}="
                    f"{last_weight_values!r}, summary range "
                    f"{self.contract.summary_range}={last_summary_matrix!r}."
                )

            if progress_callback is not None and time.monotonic() >= next_progress_ping:
                progress_callback(
                    "waiting_for_outputs",
                    "Excel is still settling the final portfolio table and chart inputs.",
                )
                next_progress_ping = time.monotonic() + 2.0

            time.sleep(poll_interval_seconds)

    @staticmethod
    def _freeze_workbook_value(value: Any) -> Any:
        if isinstance(value, list):
            return tuple(ModelWorkbookRunner._freeze_workbook_value(item) for item in value)
        if isinstance(value, tuple):
            return tuple(ModelWorkbookRunner._freeze_workbook_value(item) for item in value)
        if isinstance(value, float):
            return round(value, 12)
        return value

    @staticmethod
    def _mvp_snapshot_ready(
        *,
        stats_values: Any,
        weight_values: Any,
        summary_matrix: Any,
    ) -> bool:
        return (
            ModelWorkbookRunner._numeric_sequence_ready(stats_values)
            and ModelWorkbookRunner._numeric_sequence_ready(weight_values)
            and ModelWorkbookRunner._summary_matrix_ready(summary_matrix)
        )

    @staticmethod
    def _numeric_sequence_ready(values: Any) -> bool:
        if not isinstance(values, list) or not values:
            return False
        return all(
            isinstance(value, (int, float)) and not isinstance(value, bool)
            for value in values
        )

    @staticmethod
    def _summary_matrix_ready(matrix: Any) -> bool:
        if not isinstance(matrix, list) or len(matrix) < 2:
            return False
        rows = [row for row in matrix if isinstance(row, list)]
        if len(rows) != len(matrix):
            return False
        header = rows[0]
        if not header or not all(str(value).strip() for value in header):
            return False
        return any(any(value not in (None, "") for value in row) for row in rows[1:])

    def _export_final_charts(
        self,
        book: xw.Book,
        paths: ModelSessionPaths,
    ) -> dict[str, str]:
        exported_paths: dict[str, str] = {}
        for chart_name in self.contract.chart_names:
            chart_path = paths.chart_path(chart_name)
            exporter = ExcelChartExporter(
                ExcelChartSpec(
                    sheet_name=self.contract.calculator_sheet,
                    chart_name=chart_name,
                )
            )
            exporter.export(book, chart_path)
            exported_paths[chart_name] = str(chart_path)

        return exported_paths


def accepted_answers_from_elicitation(payload: BaseModel) -> dict[str, str]:
    return {str(key): str(value) for key, value in payload.model_dump().items()}
