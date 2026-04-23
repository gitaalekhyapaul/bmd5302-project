from __future__ import annotations

import argparse
import base64
from contextlib import closing
from datetime import datetime, timezone
import html
import json
import logging
import os
from pathlib import Path
import re
import sqlite3
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
from sandra_chat_server import SandraChatMemory, register_sandra_chat_backend_tools

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent
APP_RESOURCE_URI = "ui://sandra-investment-chat/mcp-app.html"
APP_RESOURCE_MIME_TYPE = "text/html;profile=mcp-app"
APP_DEFAULT_DB_PATH = "notebook_outputs/sandra_app_memory.sqlite3"
LOG_MAX_STRING = 400
LOG_MAX_ITEMS = 12
LOG_MAX_DEPTH = 4


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _sanitize_for_log(
    value: Any,
    *,
    depth: int = 0,
    max_depth: int = LOG_MAX_DEPTH,
    max_items: int = LOG_MAX_ITEMS,
    max_string: int = LOG_MAX_STRING,
) -> Any:
    if depth >= max_depth:
        return f"<{type(value).__name__}>"
    if isinstance(value, dict):
        items = list(value.items())
        sanitized = {
            str(key): _sanitize_for_log(
                item,
                depth=depth + 1,
                max_depth=max_depth,
                max_items=max_items,
                max_string=max_string,
            )
            for key, item in items[:max_items]
        }
        if len(items) > max_items:
            sanitized["__truncated_keys__"] = len(items) - max_items
        return sanitized
    if isinstance(value, (list, tuple, set)):
        items = list(value)
        sanitized_items = [
            _sanitize_for_log(
                item,
                depth=depth + 1,
                max_depth=max_depth,
                max_items=max_items,
                max_string=max_string,
            )
            for item in items[:max_items]
        ]
        if len(items) > max_items:
            sanitized_items.append(f"... ({len(items) - max_items} more items)")
        return sanitized_items
    if isinstance(value, str):
        compact = re.sub(r"\s+", " ", value).strip()
        if len(compact) > max_string:
            return f"{compact[:max_string]}... <len={len(compact)}>"
        return compact
    if isinstance(value, Path):
        return str(value)
    return value


def _log_payload(
    level: int,
    message: str,
    payload: Any,
    *,
    exc_info: bool = False,
) -> None:
    logger.log(
        level,
        "%s | payload=%s",
        message,
        json.dumps(_sanitize_for_log(payload), default=str, sort_keys=True),
        exc_info=exc_info,
    )


def _load_dotenv_file(env_path: str | Path = PROJECT_ROOT / ".env") -> None:
    """Load repo-local .env values without overriding the parent environment."""
    path = Path(env_path).expanduser()
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if key.startswith("export "):
            key = key.removeprefix("export ").strip()
        value = value.strip().strip('"').strip("'")
        if key:
            os.environ.setdefault(key, value)


def _app_memory_db_path() -> Path:
    configured = os.environ.get("SANDRA_APP_DB_PATH", APP_DEFAULT_DB_PATH)
    path = Path(configured).expanduser()
    if not path.is_absolute():
        path = PROJECT_ROOT / path
    return path.resolve()


def _connect_app_memory_db() -> sqlite3.Connection:
    path = _app_memory_db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS app_threads (
            thread_id TEXT PRIMARY KEY,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            label TEXT NOT NULL,
            state_json TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS app_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            thread_id TEXT NOT NULL,
            created_at TEXT NOT NULL,
            role TEXT NOT NULL,
            event_type TEXT NOT NULL,
            content TEXT NOT NULL,
            session_id TEXT,
            payload_json TEXT NOT NULL,
            FOREIGN KEY(thread_id) REFERENCES app_threads(thread_id)
        )
        """
    )
    conn.commit()
    return conn


def _ensure_app_thread(conn: sqlite3.Connection, thread_id: str) -> None:
    now = _utc_now_iso()
    conn.execute(
        """
        INSERT INTO app_threads (thread_id, created_at, updated_at, label, state_json)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(thread_id) DO NOTHING
        """,
        (thread_id, now, now, "Sandra Investment Chat", "{}"),
    )
    conn.commit()


def _get_app_thread_state(conn: sqlite3.Connection, thread_id: str) -> dict[str, Any]:
    row = conn.execute(
        "SELECT state_json FROM app_threads WHERE thread_id = ?",
        (thread_id,),
    ).fetchone()
    if row is None:
        return {}
    try:
        state = json.loads(str(row["state_json"]))
    except json.JSONDecodeError:
        return {}
    return state if isinstance(state, dict) else {}


def _update_app_thread_state(
    thread_id: str,
    updates: dict[str, Any],
) -> dict[str, Any]:
    with closing(_connect_app_memory_db()) as conn:
        _ensure_app_thread(conn, thread_id)
        state = _get_app_thread_state(conn, thread_id)
        state.update(updates)
        conn.execute(
            "UPDATE app_threads SET updated_at = ?, state_json = ? WHERE thread_id = ?",
            (_utc_now_iso(), json.dumps(state, default=str), thread_id),
        )
        conn.commit()
        _log_payload(
            logging.DEBUG,
            "app_memory.update_state",
            {"thread_id": thread_id, "updates": updates, "state": state},
        )
        return state


def _append_app_memory_event(
    *,
    thread_id: str,
    role: str,
    content: str,
    event_type: str = "message",
    session_id: str | None = None,
    payload: dict[str, Any] | None = None,
) -> dict[str, Any]:
    with closing(_connect_app_memory_db()) as conn:
        _ensure_app_thread(conn, thread_id)
        now = _utc_now_iso()
        conn.execute(
            """
            INSERT INTO app_events
                (thread_id, created_at, role, event_type, content, session_id, payload_json)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                thread_id,
                now,
                role,
                event_type,
                content,
                session_id,
                json.dumps(payload or {}, default=str),
            ),
        )
        conn.execute(
            "UPDATE app_threads SET updated_at = ? WHERE thread_id = ?",
            (now, thread_id),
        )
        conn.commit()
        event_count = conn.execute(
            "SELECT COUNT(*) AS count FROM app_events WHERE thread_id = ?",
            (thread_id,),
        ).fetchone()["count"]
    result = {
        "thread_id": thread_id,
        "event_count": int(event_count),
        "database_path": str(_app_memory_db_path()),
    }
    _log_payload(
        logging.DEBUG,
        "app_memory.append_event",
        {
            "thread_id": thread_id,
            "role": role,
            "event_type": event_type,
            "session_id": session_id,
            "content": content,
            "payload": payload,
            "result": result,
        },
    )
    return result


def _get_app_memory_snapshot(thread_id: str, limit: int = 40) -> dict[str, Any]:
    with closing(_connect_app_memory_db()) as conn:
        _ensure_app_thread(conn, thread_id)
        state = _get_app_thread_state(conn, thread_id)
        rows = conn.execute(
            """
            SELECT id, created_at, role, event_type, content, session_id, payload_json
            FROM app_events
            WHERE thread_id = ?
            ORDER BY id DESC
            LIMIT ?
            """,
            (thread_id, limit),
        ).fetchall()
        event_count = conn.execute(
            "SELECT COUNT(*) AS count FROM app_events WHERE thread_id = ?",
            (thread_id,),
        ).fetchone()["count"]

    events = []
    for row in reversed(rows):
        try:
            payload = json.loads(str(row["payload_json"]))
        except json.JSONDecodeError:
            payload = {}
        events.append(
            {
                "id": int(row["id"]),
                "created_at": str(row["created_at"]),
                "role": str(row["role"]),
                "event_type": str(row["event_type"]),
                "content": str(row["content"]),
                "session_id": row["session_id"],
                "payload": payload,
            }
        )

    result = {
        "thread_id": thread_id,
        "database_path": str(_app_memory_db_path()),
        "event_count": int(event_count),
        "recent_events": events,
        "state": state,
    }
    _log_payload(
        logging.DEBUG,
        "app_memory.snapshot",
        {"thread_id": thread_id, "limit": limit, "result": result},
    )
    return result


def _app_tool_meta(visibility: list[str] | None = None) -> dict[str, Any]:
    ui_meta: dict[str, Any] = {"resourceUri": APP_RESOURCE_URI}
    if visibility:
        ui_meta["visibility"] = visibility
    return {
        "ui": ui_meta,
        "ui/resourceUri": APP_RESOURCE_URI,
    }


def _read_sandra_app_html() -> str:
    dist_html = PROJECT_ROOT / "mcp_app" / "dist" / "mcp-app.html"
    if dist_html.exists():
        return dist_html.read_text(encoding="utf-8")

    return f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Sandra Investment Chat</title>
    <style>
      body {{
        margin: 0;
        min-height: 100vh;
        display: grid;
        place-items: center;
        background: #06110f;
        color: #f5efe2;
        font-family: Avenir Next, Helvetica Neue, sans-serif;
      }}
      main {{
        max-width: 720px;
        padding: 32px;
        border: 1px solid rgba(216, 196, 146, 0.24);
        border-radius: 24px;
        background: rgba(11, 30, 26, 0.82);
      }}
      code {{ color: #f0d98a; }}
    </style>
  </head>
  <body>
    <main>
      <h1>Sandra Investment Chat</h1>
      <p>The visual Sandra app has not been built yet.</p>
      <p>
        Run <code>npm --prefix mcp_app run build</code> from the repo root,
        then restart Sandra's local workbook service.
      </p>
      <p>UI resource: <code>{html.escape(APP_RESOURCE_URI)}</code></p>
    </main>
  </body>
</html>"""


def _clean_questionnaire_option_text(option: Any, letter: Any) -> str:
    """Hide workbook scoring/letter artifacts while preserving the submitted value."""
    text = str(option).strip()
    letter_text = str(letter).strip()
    if len(letter_text) == 1:
        text = re.sub(
            rf"^\s*(?:\(\s*{re.escape(letter_text)}\s*\)|"
            rf"{re.escape(letter_text)}\s*[\).:-])\s*",
            "",
            text,
            count=1,
            flags=re.IGNORECASE,
        )
    text = re.sub(
        r"\s*(?:(?:-|–|—)+>|→|⇒)\s*A\s*=\s*[-+]?\d[\d,.]*\s*$",
        "",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"\s*(?:[,;|]\s*)?A\s*=\s*[-+]?\d[\d,.]*\s*$",
        "",
        text,
        flags=re.IGNORECASE,
    )
    return text.strip()


def _render_questionnaire_form(payload: dict[str, Any]) -> str:
    questions = payload.get("questions", [])
    session_id = html.escape(str(payload.get("session_id", "")))
    question_count = len(questions) if isinstance(questions, list) else 0
    parts = [
        '<div class="questionnaire-card">',
        '<div class="questionnaire-header">',
        "<div>",
        "<h3>Investor questionnaire</h3>",
        "<p>Sandra will submit your selections to the workbook exactly as required.</p>",
        "</div>",
        f'<div class="answer-meter" id="answer-meter">0/{question_count}</div>',
        "</div>",
        (
            f'<form id="sandra-questionnaire-form" class="question-stack" '
            f'data-session-id="{session_id}" data-question-count="{question_count}">'
        ),
    ]

    if isinstance(questions, list):
        for index, question in enumerate(questions, start=1):
            if not isinstance(question, dict):
                continue
            key = html.escape(str(question.get("key", f"q{index}")))
            prompt = html.escape(str(question.get("prompt", "")))
            options = question.get("options", [])
            letters = question.get("option_letters", [])
            if not isinstance(options, list):
                options = []
            if not isinstance(letters, list) or len(letters) != len(options):
                letters = [chr(ord("a") + i) for i in range(len(options))]

            parts.extend(
                [
                    f'<fieldset class="question-block" aria-labelledby="{key}-prompt">',
                    f'<p class="question-prompt" id="{key}-prompt">{index}. {prompt}</p>',
                    '<div class="option-grid">',
                ]
            )
            for letter, option in zip(letters, options, strict=False):
                value = html.escape(str(letter).lower())
                option_text = html.escape(_clean_questionnaire_option_text(option, letter))
                input_id = f"{key}-{value}"
                parts.extend(
                    [
                        f'<label class="option-row" for="{input_id}">',
                        (
                            f'<input id="{input_id}" type="radio" name="{key}" '
                            f'value="{value}" required>'
                        ),
                        f"<span>{option_text}</span>",
                        "</label>",
                    ]
                )
            parts.extend(["</div>", "</fieldset>"])

    parts.extend(
        [
            '<button class="primary-button" id="submit-questionnaire" type="submit" disabled>',
            "Submit answers to Sandra",
            "</button>",
            "</form>",
            "</div>",
        ]
    )
    return "".join(parts)


def _chart_images_from_paths(chart_paths: dict[str, str]) -> list[dict[str, str]]:
    images = []
    for chart_name, chart_path in chart_paths.items():
        path = Path(chart_path)
        if not path.exists():
            continue
        images.append(
            {
                "name": chart_name,
                "mime_type": "image/png",
                "data": base64.b64encode(path.read_bytes()).decode("ascii"),
            }
        )
    return images


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
    _load_dotenv_file()
    mcp = FastMCP(
        "sandra-robo-advisor-mcp",
        host=host,
        port=port,
        streamable_http_path=streamable_http_path,
    )

    @mcp.resource(
        APP_RESOURCE_URI,
        name="sandra_investment_chat_app",
        title="Sandra Investment Chat",
        description="Professional visual interface for Sandra's guided investment flow.",
        mime_type=APP_RESOURCE_MIME_TYPE,
        meta={
            "ui": {
                "csp": {
                    "resourceDomains": [],
                    "connectDomains": [],
                },
                "prefersBorder": False,
            }
        },
    )
    def sandra_investment_chat_app() -> str:
        """Return the bundled Sandra investment chat MCP App HTML."""
        return _read_sandra_app_html()

    @mcp.tool(
        name="open_sandra_investment_chat",
        title="Open Sandra Investment Chat",
        description=(
            "Open Sandra's guided investment chat. "
            "Use this when the user wants the easy form-based experience."
        ),
        meta=_app_tool_meta(["model"]),
    )
    def open_sandra_investment_chat(thread_id: str = "default") -> dict[str, Any]:
        """Launch the Sandra MCP App and initialize the local SQLite memory thread."""
        _log_payload(
            logging.INFO,
            "tool open_sandra_investment_chat request",
            {"thread_id": thread_id},
        )
        memory = SandraChatMemory()
        snapshot = memory.snapshot(thread_id)
        greeting = (
            "Sandra is ready to guide the investment consultation. "
            "Open the visual chat to begin, or continue here for a text-only flow."
        )
        memory.append_event(
            thread_id=thread_id,
            role="assistant",
            event_type="app_opened",
            content=greeting,
        )
        payload = {
            "advisor_name": ADVISOR_NAME,
            "thread_id": thread_id,
            "message": greeting,
            "ui_resource_uri": APP_RESOURCE_URI,
            "memory": {
                "database_path": snapshot["database_path"],
                "event_count": snapshot["event_count"],
            },
        }
        _log_payload(logging.INFO, "tool open_sandra_investment_chat response", payload)
        return payload

    @mcp.tool(
        name="sandra_app_memory_snapshot",
        title="Sandra App Memory Snapshot",
        description="Return the local SQLite memory snapshot for Sandra's app UI.",
        meta=_app_tool_meta(["app"]),
    )
    def sandra_app_memory_snapshot(
        thread_id: str = "default",
        limit: int = 40,
    ) -> dict[str, Any]:
        """Return recent app events and durable state from the local SQLite memory DB."""
        payload = _get_app_memory_snapshot(thread_id=thread_id, limit=limit)
        _log_payload(
            logging.INFO,
            "tool sandra_app_memory_snapshot response",
            {"thread_id": thread_id, "limit": limit, "payload": payload},
        )
        return payload

    @mcp.tool(
        name="sandra_app_record_chat_event",
        title="Record Sandra App Chat Event",
        description="Persist a Sandra app chat event or user note into local SQLite memory.",
        meta=_app_tool_meta(["app"]),
    )
    def sandra_app_record_chat_event(
        thread_id: str,
        role: str,
        content: str,
        event_type: str = "message",
        session_id: str | None = None,
        payload: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Append a UI chat event to Sandra's local SQLite memory."""
        _log_payload(
            logging.INFO,
            "tool sandra_app_record_chat_event request",
            {
                "thread_id": thread_id,
                "role": role,
                "event_type": event_type,
                "session_id": session_id,
                "content": content,
                "payload": payload,
            },
        )
        result = _append_app_memory_event(
            thread_id=thread_id,
            role=role,
            event_type=event_type,
            content=content,
            session_id=session_id,
            payload=payload,
        )
        _log_payload(logging.INFO, "tool sandra_app_record_chat_event response", result)
        return result

    @mcp.tool(
        name="sandra_app_start_questionnaire_form",
        title="Start Sandra App Questionnaire Form",
        description=(
            "App-only tool that runs Model.xlsm RandomizeQuestions and returns "
            "server-rendered questionnaire form HTML for Sandra's UI."
        ),
        meta=_app_tool_meta(["app"]),
    )
    def sandra_app_start_questionnaire_form(
        thread_id: str = "default",
        workbook_path: str = "Model.xlsm",
        output_dir: str = "notebook_outputs",
        visible: bool = False,
    ) -> dict[str, Any]:
        """Start the workbook questionnaire and render the UI form on the server."""
        t0 = time.perf_counter()
        logger.info(
            "tool sandra_app_start_questionnaire_form thread_id=%s workbook_path=%r "
            "output_dir=%r visible=%s",
            thread_id,
            workbook_path,
            output_dir,
            visible,
        )
        _log_payload(
            logging.INFO,
            "tool sandra_app_start_questionnaire_form request",
            {
                "thread_id": thread_id,
                "workbook_path": workbook_path,
                "output_dir": output_dir,
                "visible": visible,
            },
        )
        runner = ModelWorkbookRunner()
        state = runner.start_questionnaire_session(
            workbook_path=workbook_path,
            output_dir=output_dir,
            visible=visible,
            use_source_workbook=True,
        )
        payload = runner.serialize_start_payload(state)
        payload["thread_id"] = thread_id
        payload["form_html"] = _render_questionnaire_form(payload)
        _update_app_thread_state(
            thread_id,
            {
                "stage": "questionnaire",
                "session_id": state.session_id,
                "source_workbook_path": state.source_workbook_path,
            },
        )
        _append_app_memory_event(
            thread_id=thread_id,
            role="assistant",
            event_type="questionnaire_started",
            content="Sandra rendered the Model.xlsm investor questionnaire form.",
            session_id=state.session_id,
            payload={
                "question_count": len(state.questions),
                "source_workbook_path": state.source_workbook_path,
            },
        )
        _log_tool_done(
            "sandra_app_start_questionnaire_form",
            t0,
            f"thread_id={thread_id} session_id={state.session_id} questions={len(state.questions)}",
        )
        _log_payload(logging.INFO, "tool sandra_app_start_questionnaire_form response", payload)
        return payload

    @mcp.tool(
        name="sandra_app_submit_questionnaire_form",
        title="Submit Sandra App Questionnaire Form",
        description=(
            "App-only tool that writes Sandra UI form answers to Model.xlsm and "
            "returns the workbook-generated investor profile."
        ),
        meta=_app_tool_meta(["app"]),
    )
    def sandra_app_submit_questionnaire_form(
        thread_id: str,
        session_id: str,
        answers: dict[str, str],
        output_dir: str = "notebook_outputs",
        visible: bool = False,
    ) -> dict[str, Any]:
        """Submit UI form answers into the workbook and return Sandra's profile payload."""
        t0 = time.perf_counter()
        logger.info(
            "tool sandra_app_submit_questionnaire_form thread_id=%s session_id=%s "
            "answer_count=%d",
            thread_id,
            session_id,
            len(answers),
        )
        _log_payload(
            logging.INFO,
            "tool sandra_app_submit_questionnaire_form request",
            {
                "thread_id": thread_id,
                "session_id": session_id,
                "answers": answers,
                "output_dir": output_dir,
                "visible": visible,
            },
        )
        runner = ModelWorkbookRunner()
        state = runner.submit_answers(
            session_id=session_id,
            answers=answers,
            output_dir=output_dir,
            visible=visible,
        )
        payload = runner.serialize_profile_payload(state)
        payload["thread_id"] = thread_id
        payload["next_ui_step"] = "Ask the user to choose whether short selling is allowed."
        _update_app_thread_state(
            thread_id,
            {
                "stage": "profile",
                "session_id": session_id,
                "investor_profile": state.investor_profile,
                "answers": state.answers,
            },
        )
        _append_app_memory_event(
            thread_id=thread_id,
            role="assistant",
            event_type="profile_generated",
            content=payload.get(
                "creative_profile_message",
                "Sandra generated the investor profile.",
            ),
            session_id=session_id,
            payload={
                "answers": state.answers,
                "investor_profile": state.investor_profile,
            },
        )
        _log_tool_done(
            "sandra_app_submit_questionnaire_form",
            t0,
            f"thread_id={thread_id} session_id={session_id} "
            f"profile_present={bool(state.investor_profile)}",
        )
        _log_payload(logging.INFO, "tool sandra_app_submit_questionnaire_form response", payload)
        return payload

    @mcp.tool(
        name="sandra_app_run_mvp",
        title="Run Sandra App MVP",
        description=(
            "App-only tool that runs Model.xlsm optimizer/calculator macros and "
            "returns the final summary plus chart images for Sandra's UI."
        ),
        meta=_app_tool_meta(["app"]),
    )
    def sandra_app_run_mvp(
        thread_id: str,
        session_id: str,
        allow_short_selling: bool,
        output_dir: str = "notebook_outputs",
        visible: bool = False,
    ) -> dict[str, Any]:
        """Run the final workbook flow and return UI-ready charts as base64 PNGs."""
        t0 = time.perf_counter()
        logger.info(
            "tool sandra_app_run_mvp thread_id=%s session_id=%s allow_short_selling=%s",
            thread_id,
            session_id,
            allow_short_selling,
        )
        _log_payload(
            logging.INFO,
            "tool sandra_app_run_mvp request",
            {
                "thread_id": thread_id,
                "session_id": session_id,
                "allow_short_selling": allow_short_selling,
                "output_dir": output_dir,
                "visible": visible,
            },
        )
        runner = ModelWorkbookRunner()
        try:
            state = runner.run_mvp(
                session_id=session_id,
                allow_short_selling=allow_short_selling,
                output_dir=output_dir,
                visible=visible,
            )
        except Exception:
            _log_payload(
                logging.ERROR,
                "tool sandra_app_run_mvp failed",
                {
                    "thread_id": thread_id,
                    "session_id": session_id,
                    "allow_short_selling": allow_short_selling,
                    "output_dir": output_dir,
                    "visible": visible,
                },
                exc_info=True,
            )
            raise
        payload = runner.serialize_final_payload(state)
        payload["thread_id"] = thread_id
        payload["chart_images"] = _chart_images_from_paths(payload.get("chart_paths", {}))
        payload["llm_presentation_instructions"] = (
            "Display the entire final summary table before the chart images. "
            "State that workbook Ann. Return values are model assumptions, not guarantees."
        )
        _update_app_thread_state(
            thread_id,
            {
                "stage": "completed",
                "session_id": session_id,
                "allow_short_selling": allow_short_selling,
                "summary_table_records": state.summary_table_records,
                "chart_paths": state.chart_paths,
            },
        )
        _append_app_memory_event(
            thread_id=thread_id,
            role="assistant",
            event_type="mvp_completed",
            content="Sandra completed the workbook optimizer flow and rendered final outputs.",
            session_id=session_id,
            payload={
                "allow_short_selling": allow_short_selling,
                "summary_row_count": len(state.summary_table_records),
                "chart_names": list(state.chart_paths.keys()),
            },
        )
        _log_tool_done(
            "sandra_app_run_mvp",
            t0,
            f"thread_id={thread_id} session_id={session_id} status={payload.get('status')!r}",
        )
        _log_payload(logging.INFO, "tool sandra_app_run_mvp response", payload)
        return payload

    register_sandra_chat_backend_tools(mcp)

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
        _log_payload(
            logging.INFO,
            "tool start_investor_questionnaire request",
            {
                "workbook_path": workbook_path,
                "output_dir": output_dir,
                "visible": visible,
                "use_elicitation": use_elicitation,
                "use_source_workbook": enforced_use_source_workbook,
                "requested_use_source_workbook": use_source_workbook,
            },
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
            _log_payload(logging.INFO, "tool start_investor_questionnaire response", payload)
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
        _log_payload(logging.INFO, "tool start_investor_questionnaire response", payload)
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
        _log_payload(
            logging.INFO,
            "tool submit_investor_questionnaire_answers request",
            {
                "session_id": session_id,
                "answers": answers,
                "output_dir": output_dir,
                "visible": visible,
            },
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
        _log_payload(logging.INFO, "tool submit_investor_questionnaire_answers response", out)
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
        _log_payload(
            logging.INFO,
            "_run_investor_mvp_workflow request",
            {
                "session_id": session_id,
                "allow_short_selling": allow_short_selling,
                "output_dir": output_dir,
                "visible": visible,
                "use_elicitation": use_elicitation,
                "context_present": context is not None,
            },
        )

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
                _log_payload(logging.INFO, "_run_investor_mvp_workflow response", payload)
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
                _log_payload(logging.INFO, "_run_investor_mvp_workflow response", payload)
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
        try:
            final_state = runner.run_mvp(
                session_id=session_id,
                allow_short_selling=allow_short_selling,
                output_dir=output_dir,
                visible=visible,
            )
        except Exception:
            _log_payload(
                logging.ERROR,
                "_run_investor_mvp_workflow workbook execution failed",
                {
                    "session_id": session_id,
                    "allow_short_selling": allow_short_selling,
                    "output_dir": output_dir,
                    "visible": visible,
                },
                exc_info=True,
            )
            raise
        out = runner.serialize_final_payload(final_state)
        logger.debug("run_investor_mvp raw payload keys=%s", sorted(out.keys()))
        if log_tool_completion:
            _log_tool_done(
                "run_investor_mvp",
                started,
                f"status={out.get('status')!r} session_id={session_id}",
            )
        _log_payload(logging.INFO, "_run_investor_mvp_workflow response", out)
        return out

    @mcp.tool()
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
        _log_payload(
            logging.INFO,
            "tool run_investor_mvp request",
            {
                "session_id": session_id,
                "allow_short_selling": allow_short_selling,
                "output_dir": output_dir,
                "visible": visible,
                "use_elicitation": use_elicitation,
            },
        )
        payload = await _run_investor_mvp_workflow(
            session_id,
            allow_short_selling,
            output_dir,
            visible,
            use_elicitation,
            context,
            log_tool_completion=True,
            started=t0,
        )
        _log_payload(logging.INFO, "tool run_investor_mvp response", payload)
        return payload

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
        _log_payload(
            logging.INFO,
            "tool run_investor_mvp_with_chart_images request",
            {
                "session_id": session_id,
                "allow_short_selling": allow_short_selling,
                "output_dir": output_dir,
                "visible": visible,
                "use_elicitation": use_elicitation,
            },
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
            _log_payload(logging.INFO, "tool run_investor_mvp_with_chart_images response", payload)
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
        _log_payload(logging.INFO, "tool run_investor_mvp_with_chart_images response", payload)
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
            "calculator_target_sigma_cell": contract.calculator_target_sigma_cell,
            "calculator_stats_range": contract.calculator_stats_range,
            "short_selling_cell": contract.short_selling_cell,
            "calculator_macro_name": contract.calculator_macro_name,
            "calculator_variance_cell": contract.calculator_variance_cell,
            "calculator_weight_range": contract.calculator_weight_range,
            "summary_range": contract.summary_range,
            "chart_names": list(contract.chart_names),
        }
        _log_tool_done(
            "get_model_workbook_contract",
            t0,
            f"chart_names={out['chart_names']}",
        )
        _log_payload(logging.INFO, "tool get_model_workbook_contract response", out)
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
