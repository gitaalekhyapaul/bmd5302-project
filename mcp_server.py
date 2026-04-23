from __future__ import annotations

import argparse
import logging
import time
from typing import Any

import mcp.types as mcp_types
from mcp.server.fastmcp import FastMCP  # type: ignore[import-not-found]
from mcp.server.fastmcp import Context  # type: ignore[import-not-found]
from mcp.server.fastmcp.utilities.types import Image  # type: ignore[import-not-found]

from model_workflow import (
    ADVISOR_NAME,
    ModelWorkbookContract,
    ModelWorkbookRunner,
    ShortSellingChoice,
    accepted_answers_from_elicitation,
)

logger = logging.getLogger(__name__)


def _configure_logging(level: int) -> None:
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
        datefmt="%Y-%m-%dT%H:%M:%S",
        force=True,
    )
    logging.getLogger("mcp").setLevel(level)


def _supports_elicitation(context: Context | None) -> bool:
    if context is None:
        logger.debug("elicitation check: no MCP context; treating as unsupported")
        return False

    supported = context.session.check_client_capability(
        mcp_types.ClientCapabilities(
            elicitation=mcp_types.ElicitationCapability(),
        )
    )
    logger.debug("elicitation check: client capability=%s", supported)
    return supported


def _log_tool_done(name: str, started: float, summary: str) -> None:
    logger.info(
        "%s completed in %.3fs — %s", name, time.perf_counter() - started, summary
    )


def _build_server(host: str, port: int, streamable_http_path: str) -> FastMCP:
    mcp = FastMCP(
        "sandra-robo-advisor-mcp",
        host=host,
        port=port,
        streamable_http_path=streamable_http_path,
    )

    @mcp.tool()
    async def start_investor_questionnaire(
        workbook_path: str = "Model.xlsm",
        output_dir: str = "notebook_outputs",
        visible: bool = False,
        use_elicitation: bool = True,
        use_source_workbook: bool = True,
        context: Context | None = None,
    ) -> dict[str, Any]:
        """Start Sandra's questionnaire session and optionally elicit answers."""
        t0 = time.perf_counter()
        enforced_use_source_workbook = True
        logger.info(
            "tool start_investor_questionnaire workbook_path=%r output_dir=%r "
            "visible=%s use_elicitation=%s use_source_workbook=%s (requested=%s)",
            workbook_path,
            output_dir,
            visible,
            use_elicitation,
            enforced_use_source_workbook,
            use_source_workbook,
        )
        runner = ModelWorkbookRunner()
        state = runner.start_questionnaire_session(
            workbook_path=workbook_path,
            output_dir=output_dir,
            visible=visible,
            use_source_workbook=enforced_use_source_workbook,
        )
        logger.debug(
            "questionnaire session started session_id=%s questions=%d",
            state.session_id,
            len(state.questions),
        )
        payload = runner.serialize_start_payload(state)
        payload["elicitation_supported"] = _supports_elicitation(context)
        payload["llm_question_display_instructions"] = (
            "Display all questions and all answer options verbatim before requesting "
            "any answers. Do not summarize or omit options."
        )

        if not use_elicitation or not payload["elicitation_supported"]:
            payload["next_step"] = (
                "Collect the questionnaire answers from the user and call "
                "`submit_investor_questionnaire_answers`."
            )
            payload["manual_question_display_format"] = (
                "Use this format for each question:\n"
                "Question {n} ({question_key}): {question_prompt}\n"
                "Options:\n"
                "- A) ...\n"
                "- B) ...\n"
                "- C) ...\n"
                "- D) ...\n"
                "Answer format: return one letter per question using the question key, "
                "for example {'q1': 'B', 'q2': 'D'}."
            )
            _log_tool_done(
                "start_investor_questionnaire",
                t0,
                f"session_id={state.session_id} path=manual_answers "
                f"elicitation_supported={payload['elicitation_supported']}",
            )
            return payload

        schema = runner.build_questionnaire_elicitation_model(state.questions)
        n_fields = len(getattr(schema, "model_fields", ()))
        logger.info(
            "eliciting questionnaire answers session_id=%s schema_fields=%d",
            state.session_id,
            n_fields,
        )
        elicitation_result = await context.elicit(
            message=(
                "Sandra is ready to begin the Model.xlsm investor questionnaire. "
                "Please choose one answer letter for each question."
            ),
            schema=schema,
        )
        payload["elicitation_action"] = elicitation_result.action
        logger.info(
            "questionnaire elicitation returned action=%r session_id=%s",
            elicitation_result.action,
            state.session_id,
        )
        if logger.isEnabledFor(logging.DEBUG):
            logger.debug("elicitation raw data: %r", elicitation_result.data)

        if elicitation_result.action != "accept":
            payload["next_step"] = (
                "Collect the questionnaire answers from the user in chat and call "
                "`submit_investor_questionnaire_answers`."
            )
            payload["manual_question_display_format"] = (
                "Use this format for each question:\n"
                "Question {n} ({question_key}): {question_prompt}\n"
                "Options:\n"
                "- A) ...\n"
                "- B) ...\n"
                "- C) ...\n"
                "- D) ...\n"
                "Answer format: return one letter per question using the question key, "
                "for example {'q1': 'B', 'q2': 'D'}."
            )
            _log_tool_done(
                "start_investor_questionnaire",
                t0,
                f"session_id={state.session_id} elicitation_action={elicitation_result.action!r}",
            )
            return payload

        answered_state = runner.submit_answers(
            session_id=state.session_id,
            answers=accepted_answers_from_elicitation(elicitation_result.data),
            output_dir=output_dir,
            visible=visible,
        )
        payload.update(runner.serialize_profile_payload(answered_state))
        payload["input_method"] = "elicitation"
        _log_tool_done(
            "start_investor_questionnaire",
            t0,
            f"session_id={state.session_id} path=elicitation_accepted "
            f"profile_present={'investor_profile' in payload}",
        )
        return payload

    @mcp.tool()
    def submit_investor_questionnaire_answers(
        session_id: str,
        answers: dict[str, str],
        output_dir: str = "notebook_outputs",
        visible: bool = False,
    ) -> dict[str, Any]:
        """Write questionnaire answers into Model.xlsm and return Sandra's workbook profile."""
        t0 = time.perf_counter()
        logger.info(
            "tool submit_investor_questionnaire_answers session_id=%s "
            "answer_count=%d answer_keys=%s output_dir=%r visible=%s",
            session_id,
            len(answers),
            sorted(answers.keys()),
            output_dir,
            visible,
        )
        if logger.isEnabledFor(logging.DEBUG):
            logger.debug("submit_investor_questionnaire_answers answers=%r", answers)
        runner = ModelWorkbookRunner()
        state = runner.submit_answers(
            session_id=session_id,
            answers=answers,
            output_dir=output_dir,
            visible=visible,
        )
        out = runner.serialize_profile_payload(state)
        out["llm_short_selling_instruction"] = (
            "Do NOT assume whether short selling is allowed. "
            "Ask the user directly: 'Do you want to allow short selling? (Yes/No)' "
            "and only proceed after the user explicitly answers."
        )
        out["next_step"] = (
            "Ask the user whether they want short selling (Yes/No), then call "
            "`run_investor_mvp` or `run_investor_mvp_with_chart_images` with explicit "
            "`allow_short_selling=true` or `allow_short_selling=false`."
        )
        _log_tool_done(
            "submit_investor_questionnaire_answers",
            t0,
            f"session_id={session_id} response_keys={sorted(out.keys())}",
        )
        return out

    async def _run_investor_mvp_workflow(
        session_id: str,
        allow_short_selling: bool | None,
        output_dir: str,
        visible: bool,
        use_elicitation: bool,
        context: Context | None,
        *,
        log_tool_completion: bool,
        started: float,
    ) -> dict[str, Any]:
        runner = ModelWorkbookRunner()

        if allow_short_selling is None:
            state = runner.load_session_state(
                session_id=session_id, output_dir=output_dir
            )
            logger.debug(
                "loaded session for short-selling branch profile_len=%s",
                len(state.investor_profile or ""),
            )
            payload = {
                "advisor_name": ADVISOR_NAME,
                "status": "short_selling_choice_required",
                "session_id": state.session_id,
                "investor_profile": state.investor_profile,
                "elicitation_supported": _supports_elicitation(context),
            }

            if not use_elicitation or not payload["elicitation_supported"]:
                payload["next_step"] = (
                    "Ask the user whether they want short selling (Yes or No), then call "
                    "`run_investor_mvp` again with `allow_short_selling=true` or `false`."
                )
                if log_tool_completion:
                    _log_tool_done(
                        "run_investor_mvp",
                        started,
                        "status=short_selling_choice_required path=manual_short_choice "
                        f"elicitation_supported={payload['elicitation_supported']}",
                    )
                return payload

            logger.info(
                "eliciting short-selling choice session_id=%s",
                state.session_id,
            )
            elicitation_result = await context.elicit(
                message=(
                    "Sandra is ready to run the optimizer. "
                    "Would you like to allow short selling?"
                ),
                schema=ShortSellingChoice,
            )
            payload["elicitation_action"] = elicitation_result.action
            logger.info(
                "short-selling elicitation action=%r session_id=%s",
                elicitation_result.action,
                state.session_id,
            )
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(
                    "short-selling elicitation data=%r", elicitation_result.data
                )
            if elicitation_result.action != "accept":
                payload["next_step"] = (
                    "Ask the user whether they want short selling (Yes or No), then call "
                    "`run_investor_mvp` again with an explicit boolean choice."
                )
                if log_tool_completion:
                    _log_tool_done(
                        "run_investor_mvp",
                        started,
                        "status=short_selling_choice_required elicitation_action="
                        f"{elicitation_result.action!r}",
                    )
                return payload

            allow_short_selling = bool(elicitation_result.data.allow_short_selling)
            logger.info(
                "short-selling choice from elicitation allow_short_selling=%s",
                allow_short_selling,
            )

        logger.info(
            "running MVP workbook steps session_id=%s allow_short_selling=%s",
            session_id,
            allow_short_selling,
        )
        final_state = runner.run_mvp(
            session_id=session_id,
            allow_short_selling=allow_short_selling,
            output_dir=output_dir,
            visible=visible,
        )
        out = runner.serialize_final_payload(final_state)
        logger.debug("run_investor_mvp raw payload keys=%s", sorted(out.keys()))
        if log_tool_completion:
            _log_tool_done(
                "run_investor_mvp",
                started,
                f"status={out.get('status')!r} session_id={session_id}",
            )
        return out

    # @mcp.tool()
    async def run_investor_mvp(
        session_id: str,
        allow_short_selling: bool | None = None,
        output_dir: str = "notebook_outputs",
        visible: bool = False,
        use_elicitation: bool = True,
        context: Context | None = None,
    ) -> dict[str, Any]:
        """Run Sandra's final optimizer and MVP flow for a questionnaire session."""
        t0 = time.perf_counter()
        logger.info(
            "tool run_investor_mvp session_id=%s allow_short_selling=%r "
            "output_dir=%r visible=%s use_elicitation=%s",
            session_id,
            allow_short_selling,
            output_dir,
            visible,
            use_elicitation,
        )
        return await _run_investor_mvp_workflow(
            session_id,
            allow_short_selling,
            output_dir,
            visible,
            use_elicitation,
            context,
            log_tool_completion=True,
            started=t0,
        )

    @mcp.tool(structured_output=False)
    async def run_investor_mvp_with_chart_images(
        session_id: str,
        allow_short_selling: bool | None = None,
        output_dir: str = "notebook_outputs",
        visible: bool = False,
        use_elicitation: bool = True,
        context: Context | None = None,
    ) -> list[dict[str, Any] | Image]:
        """Run Sandra's final Model.xlsm MVP flow and return summary + sheet 2 chart images.

        LLM presentation contract:
        1) Display the full `final_summary_table` to the user.
        2) Then present both chart images to the user.
        3) Charts are returned as MCP `Image` objects in response parts 2 and 3,
           not as file paths inside the JSON payload.
        """
        t0 = time.perf_counter()
        logger.info(
            "tool run_investor_mvp_with_chart_images session_id=%s allow_short_selling=%r "
            "output_dir=%r visible=%s use_elicitation=%s",
            session_id,
            allow_short_selling,
            output_dir,
            visible,
            use_elicitation,
        )
        mvp_started = time.perf_counter()
        payload = await _run_investor_mvp_workflow(
            session_id,
            allow_short_selling,
            output_dir,
            visible,
            use_elicitation,
            context,
            log_tool_completion=False,
            started=mvp_started,
        )
        logger.debug(
            "run_investor_mvp_with_chart_images MVP phase done in %.3fs",
            time.perf_counter() - mvp_started,
        )
        payload["llm_presentation_instructions"] = (
            "Display the entire `final_summary_table` first. "
            "After the table is fully shown, present the chart images. "
            "Important: chart images are MCP Image objects in this tool response "
            "(parts 2 and 3), not JSON file paths."
        )
        if payload.get("status") != "completed":
            _log_tool_done(
                "run_investor_mvp_with_chart_images",
                t0,
                f"early_return status={payload.get('status')!r} parts=1",
            )
            return [payload]

        contract = ModelWorkbookContract()
        paths = payload.get("chart_paths", {})
        logger.info(
            "attaching chart images session_id=%s charts=%s",
            session_id,
            [paths.get(n) for n in contract.chart_names],
        )
        payload.pop("chart_paths", None)
        payload["chart_transport"] = (
            "Images are attached as MCP Image objects in response parts 2 and 3."
        )
        payload["chart_names"] = list(contract.chart_names)
        _log_tool_done(
            "run_investor_mvp_with_chart_images",
            t0,
            "status=completed parts=3 (payload + 2 images)",
        )
        first_chart = paths.get(contract.chart_names[0])
        second_chart = paths.get(contract.chart_names[1])
        if not first_chart or not second_chart:
            raise RuntimeError(
                "Expected chart paths were not produced by the workbook run for "
                f"{contract.chart_names[0]!r} and {contract.chart_names[1]!r}."
            )
        return [
            payload,
            Image(path=first_chart),
            Image(path=second_chart),
        ]

    @mcp.tool()
    def get_model_workbook_contract() -> dict[str, Any]:
        """Return the active Model.xlsm workbook contract used by Sandra's investor tools."""
        t0 = time.perf_counter()
        logger.info("tool get_model_workbook_contract")
        contract = ModelWorkbookContract()
        out = {
            "advisor_name": ADVISOR_NAME,
            "questionnaire_sheet": contract.questionnaire_sheet,
            "questionnaire_macro_name": contract.questionnaire_macro_name,
            "question_range": (
                f"A{contract.question_start_row}:F{contract.question_end_row}"
            ),
            "question_text_column": contract.question_text_column,
            "question_options_column": contract.question_options_column,
            "answer_column": contract.answer_column,
            "investor_profile_cell": contract.investor_profile_cell,
            "optimizer_sheet": contract.optimizer_sheet,
            "no_short_macro_name": contract.no_short_macro_name,
            "short_macro_name": contract.short_macro_name,
            "calculator_sheet": contract.calculator_sheet,
            "short_selling_cell": contract.short_selling_cell,
            "calculator_macro_name": contract.calculator_macro_name,
            "summary_range": contract.summary_range,
            "chart_names": list(contract.chart_names),
        }
        _log_tool_done(
            "get_model_workbook_contract",
            t0,
            f"chart_names={out['chart_names']}",
        )
        return out

    return mcp


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run Sandra's workbook-backed MCP server over HTTP.",
    )
    parser.add_argument(
        "--transport",
        choices=("streamable-http", "sse"),
        default="streamable-http",
        help="MCP transport to serve.",
    )
    parser.add_argument(
        "--host",
        default="0.0.0.0",
        help="Host for HTTP transports.",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port for HTTP transports.",
    )
    parser.add_argument(
        "--mount-path",
        default=None,
        help="Optional mount path (used by SSE transport).",
    )
    parser.add_argument(
        "--streamable-http-path",
        default="/mcp",
        help="HTTP path for streamable HTTP transport.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging level for this process (default INFO).",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Shorthand for --log-level DEBUG.",
    )
    return parser


def main() -> None:
    args = _build_parser().parse_args()
    level = logging.DEBUG if args.verbose else getattr(logging, args.log_level)
    _configure_logging(level)
    logger.info(
        "starting Sandra MCP host=%s port=%s transport=%s mount_path=%r "
        "streamable_http_path=%s log_level=%s",
        args.host,
        args.port,
        args.transport,
        args.mount_path,
        args.streamable_http_path,
        logging.getLevelName(level),
    )
    mcp = _build_server(
        host=args.host,
        port=args.port,
        streamable_http_path=args.streamable_http_path,
    )
    mcp.run(transport=args.transport, mount_path=args.mount_path)


if __name__ == "__main__":
    main()
