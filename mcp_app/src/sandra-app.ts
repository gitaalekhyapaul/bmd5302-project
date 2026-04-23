import {
  App,
  applyDocumentTheme,
  applyHostFonts,
  applyHostStyleVariables,
  type McpUiHostContext,
} from "@modelcontextprotocol/ext-apps";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import DOMPurify from "dompurify";
import { marked } from "marked";
import "./sandra-app.css";

type ToolPayload = Record<string, unknown>;

type ChartImage = {
  name: string;
  mime_type: string;
  data: string;
};

type RuntimeMode = "browser" | "mcp";

type StreamHandlers = {
  onStatus?: (message: string) => void;
  onToken?: (token: string) => void;
  onResult?: (payload: ToolPayload) => void;
};

const THREAD_STORAGE_KEY = "sandra-chat-thread-id";
const runtimeMode: RuntimeMode = window.parent === window ? "browser" : "mcp";
const threadId = loadThreadId();
let currentSessionId = "";

const shell = document.getElementById("app-shell") as HTMLElement;
const chatScroll = document.getElementById("chat-scroll") as HTMLElement;
const memoryStatus = document.getElementById("memory-status") as HTMLElement;
const startButton = document.getElementById("start-chat-button") as HTMLButtonElement;
const chatInput = document.getElementById("chat-input") as HTMLInputElement;
const sendChatButton = document.getElementById("send-chat-button") as HTMLButtonElement;
const fullscreenButton = document.getElementById("fullscreen-button") as HTMLButtonElement;

const app = new App({ name: "Sandra Investment Chat", version: "1.0.0" });
marked.setOptions({ breaks: true, gfm: true });
let chartLightboxFallbackMaximized = false;

function loadThreadId(): string {
  try {
    const existing = window.localStorage.getItem(THREAD_STORAGE_KEY);
    if (existing) {
      return existing;
    }
    const generated =
      window.crypto && "randomUUID" in window.crypto
        ? `sandra-${window.crypto.randomUUID()}`
        : `sandra-${Date.now()}`;
    window.localStorage.setItem(THREAD_STORAGE_KEY, generated);
    return generated;
  } catch {
    return "default";
  }
}

function extractPayload(result: CallToolResult): ToolPayload {
  if (result.structuredContent && typeof result.structuredContent === "object") {
    const structured = result.structuredContent as ToolPayload;
    if (structured.result && typeof structured.result === "object") {
      return structured.result as ToolPayload;
    }
    return structured;
  }

  const textPart = result.content?.find((part) => part.type === "text");
  if (textPart && "text" in textPart) {
    try {
      return JSON.parse(textPart.text) as ToolPayload;
    } catch {
      return { message: textPart.text };
    }
  }

  return {};
}

function asString(value: unknown, fallback = ""): string {
  return typeof value === "string" ? value : fallback;
}

function messageFromPayload(payload: ToolPayload, fallback: string): string {
  return (
    asString(payload.assistant_message) ||
    asString(payload.message) ||
    asString(payload.error) ||
    fallback
  );
}

function messageFromUnknown(error: unknown): string {
  const raw = error instanceof Error ? error.message : String(error);
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
      return messageFromPayload(parsed as ToolPayload, raw);
    }
  } catch {
    // Keep the original error text when it is not a JSON payload.
  }
  return raw;
}

function asRecordArray(value: unknown): Record<string, unknown>[] {
  return Array.isArray(value)
    ? value.filter((item): item is Record<string, unknown> => {
        return Boolean(item) && typeof item === "object" && !Array.isArray(item);
      })
    : [];
}

function setStage(stage: string) {
  document.querySelectorAll<HTMLElement>("#stage-list li").forEach((item) => {
    item.classList.toggle("active", item.dataset.stage === stage);
  });
}

function scrollToBottom() {
  chatScroll.scrollTop = chatScroll.scrollHeight;
}

function renderMarkdown(value: string): string {
  const parsed = marked.parse(value, { async: false }) as string;
  return DOMPurify.sanitize(parsed);
}

function markdownBlock(value: string, className = ""): string {
  const classes = ["markdown", className].filter(Boolean).join(" ");
  return `<div class="${classes}">${renderMarkdown(value)}</div>`;
}

function appendMessage(role: "sandra" | "user", html: string): HTMLElement {
  const article = document.createElement("article");
  article.className = `message ${role === "sandra" ? "sandra-message" : "user-message"} entrance`;
  article.innerHTML = `
    <div class="avatar">${role === "sandra" ? "S" : "Y"}</div>
    <div class="bubble">
      <p class="speaker">${role === "sandra" ? "Sandra" : "You"}</p>
      ${html}
    </div>
  `;
  chatScroll.appendChild(article);
  scrollToBottom();
  return article.querySelector<HTMLElement>(".bubble") ?? article;
}

function appendMarkdownMessage(
  role: "sandra" | "user",
  markdown: string,
  className = "",
): HTMLElement {
  return appendMessage(role, markdownBlock(markdown, className));
}

function appendLoadingMessage(markdown: string): HTMLElement {
  const initialSummary = escapeHtml(markdown.replace(/\s+/g, " ").trim());
  const bubble = appendMessage(
    "sandra",
    `
      <div class="loader-shell" data-progress-shell aria-live="polite">
        <button class="loader-toggle" data-progress-toggle type="button" aria-expanded="true">
          <span class="loader-orb" aria-hidden="true">
            <span></span>
          </span>
          <span class="loader-summary" data-progress-summary>${initialSummary}</span>
          <span class="loader-toggle-label" data-progress-toggle-label>Hide log</span>
        </button>
        <div class="loader-body" data-progress-details>
          ${markdownBlock(markdown, "status-line loader-message")}
          <ol class="mini-log" data-mini-log>
            <li>Queued in Sandra's workflow</li>
          </ol>
        </div>
      </div>
    `,
  );
  bubble.classList.add("is-loading");
  const toggle = bubble.querySelector<HTMLButtonElement>("[data-progress-toggle]");
  toggle?.addEventListener("click", () => {
    const shell = bubble.querySelector<HTMLElement>("[data-progress-shell]");
    setProgressCollapsed(bubble, !shell?.classList.contains("is-collapsed"));
  });
  return bubble;
}

function addMiniLog(bubble: HTMLElement, entry: string) {
  const log = bubble.querySelector<HTMLOListElement>("[data-mini-log]");
  if (!log || !entry) {
    return;
  }
  const normalized = entry.replace(/\s+/g, " ").trim();
  const last = log.lastElementChild?.textContent?.trim();
  if (!normalized || normalized === last) {
    return;
  }
  const item = document.createElement("li");
  item.textContent = normalized;
  log.appendChild(item);
  while (log.children.length > 3) {
    log.firstElementChild?.remove();
  }
}

function stopLoading(bubble: HTMLElement) {
  bubble.classList.remove("is-loading");
  bubble.querySelector<HTMLElement>("[data-progress-shell]")?.classList.add("is-complete");
}

function updateProgressSummary(bubble: HTMLElement, summary: string) {
  const summaryNode = bubble.querySelector<HTMLElement>("[data-progress-summary]");
  if (!summaryNode) {
    return;
  }
  const normalized = summary.replace(/\s+/g, " ").trim();
  if (!normalized) {
    return;
  }
  summaryNode.textContent = normalized;
  summaryNode.title = normalized;
}

function setProgressCollapsed(bubble: HTMLElement, collapsed: boolean) {
  const shell = bubble.querySelector<HTMLElement>("[data-progress-shell]");
  const toggle = bubble.querySelector<HTMLButtonElement>("[data-progress-toggle]");
  const label = bubble.querySelector<HTMLElement>("[data-progress-toggle-label]");
  if (!shell || !toggle || !label) {
    return;
  }
  shell.classList.toggle("is-collapsed", collapsed);
  bubble.classList.toggle("progress-collapsed", collapsed);
  toggle.setAttribute("aria-expanded", String(!collapsed));
  label.textContent = collapsed ? "Show log" : "Hide log";
  scrollToBottom();
}

function finishStatusBubble(bubble: HTMLElement, summary: string) {
  stopLoading(bubble);
  addMiniLog(bubble, summary);
  updateProgressSummary(bubble, summary);
  setProgressCollapsed(bubble, true);
}

function replaceLoadingWithMarkdown(bubble: HTMLElement, markdown: string, className = "") {
  bubble.classList.remove("is-loading", "progress-collapsed");
  bubble.querySelector<HTMLElement>("[data-progress-shell]")?.remove();
  bubble.insertAdjacentHTML("beforeend", markdownBlock(markdown, className));
  scrollToBottom();
}

function setBubbleMarkdown(bubble: HTMLElement, markdown: string, className = "") {
  const existing =
    bubble.querySelector<HTMLElement>(".loader-message") ??
    bubble.querySelector<HTMLElement>(".markdown");
  if (existing) {
    existing.className = ["markdown", className].filter(Boolean).join(" ");
    existing.innerHTML = renderMarkdown(markdown);
  } else {
    bubble.insertAdjacentHTML("beforeend", markdownBlock(markdown, className));
  }
  updateProgressSummary(bubble, markdown);
  scrollToBottom();
}

function appendStatus(text: string): HTMLElement {
  return appendLoadingMessage(text);
}

function appendError(error: unknown) {
  const message = messageFromUnknown(error);
  appendMarkdownMessage("sandra", `I could not complete that step. ${message}`, "error-text");
}

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function chartImageSrc(image: ChartImage): string {
  return `data:${image.mime_type};base64,${image.data}`;
}

function chartDownloadName(name: string): string {
  const cleaned = name.replace(/[^a-z0-9._-]+/gi, "_").replace(/^_+|_+$/g, "");
  return `${cleaned || "sandra-chart"}.png`;
}

async function callTool(name: string, args: ToolPayload = {}): Promise<ToolPayload> {
  if (runtimeMode === "browser") {
    return callBrowserApi(name, args);
  }
  const result = await app.callServerTool({ name, arguments: args });
  return extractPayload(result);
}

async function callBrowserApi(name: string, args: ToolPayload = {}): Promise<ToolPayload> {
  let response: Response;
  if (name === "sandra_chat_turn") {
    response = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(args),
    });
  } else if (name === "sandra_chat_memory_snapshot") {
    const params = new URLSearchParams({
      thread_id: asString(args.thread_id, threadId),
      limit: String(args.limit ?? 40),
    });
    response = await fetch(`/api/memory?${params.toString()}`);
  } else if (name === "sandra_chat_record_event") {
    response = await fetch("/api/record-event", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(args),
    });
  } else {
    throw new Error(`Unsupported browser API call: ${name}`);
  }

  const payload = (await response.json()) as ToolPayload;
  if (!response.ok) {
    throw new Error(messageFromPayload(payload, "Sandra chat API request failed."));
  }
  return payload;
}

async function callChatTurn(
  args: ToolPayload,
  handlers: StreamHandlers = {},
): Promise<ToolPayload> {
  if (runtimeMode === "browser") {
    return callBrowserChatStream(args, handlers);
  }
  return callTool("sandra_chat_turn", args);
}

async function callBrowserChatStream(
  args: ToolPayload,
  handlers: StreamHandlers,
): Promise<ToolPayload> {
  const response = await fetch("/api/chat/stream", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(args),
  });
  if (!response.ok || !response.body) {
    const text = await response.text();
    let errorMessage = text || "Sandra chat stream failed.";
    try {
      const payload = JSON.parse(text) as ToolPayload;
      errorMessage = messageFromPayload(payload, errorMessage);
    } catch {
      // Keep the original response text when it is not JSON.
    }
    throw new Error(errorMessage);
  }

  const decoder = new TextDecoder();
  const reader = response.body.getReader();
  let buffer = "";
  let resultPayload: ToolPayload = {};

  function processEventBlock(block: string) {
    const lines = block.split(/\r?\n/);
    let eventName = "message";
    const dataLines: string[] = [];
    for (const line of lines) {
      if (line.startsWith("event:")) {
        eventName = line.slice("event:".length).trim();
      } else if (line.startsWith("data:")) {
        dataLines.push(line.slice("data:".length).trimStart());
      }
    }
    if (!dataLines.length) {
      return;
    }

    const payload = JSON.parse(dataLines.join("\n")) as ToolPayload;
    if (eventName === "status") {
      handlers.onStatus?.(asString(payload.message));
      return;
    }
    if (eventName === "token") {
      handlers.onToken?.(asString(payload.text));
      return;
    }
    if (eventName === "result") {
      resultPayload = payload;
      handlers.onResult?.(payload);
      return;
    }
    if (eventName === "error") {
      throw new Error(messageFromPayload(payload, "Sandra chat stream failed."));
    }
  }

  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    buffer += decoder.decode(value, { stream: true });
    let boundary = buffer.indexOf("\n\n");
    while (boundary !== -1) {
      const block = buffer.slice(0, boundary);
      buffer = buffer.slice(boundary + 2);
      processEventBlock(block);
      boundary = buffer.indexOf("\n\n");
    }
  }
  buffer += decoder.decode();
  if (buffer.trim()) {
    processEventBlock(buffer);
  }

  return resultPayload;
}

function handleHostContextChanged(ctx: McpUiHostContext) {
  if (ctx.theme) {
    applyDocumentTheme(ctx.theme);
  }
  if (ctx.styles?.variables) {
    applyHostStyleVariables(ctx.styles.variables);
  }
  if (ctx.styles?.css?.fonts) {
    applyHostFonts(ctx.styles.css.fonts);
  }
  if (ctx.safeAreaInsets) {
    shell.style.paddingTop = `${ctx.safeAreaInsets.top}px`;
    shell.style.paddingRight = `${ctx.safeAreaInsets.right}px`;
    shell.style.paddingBottom = `${ctx.safeAreaInsets.bottom}px`;
    shell.style.paddingLeft = `${ctx.safeAreaInsets.left}px`;
  }
}

function attachQuestionnaireHandlers(container: HTMLElement) {
  const form = container.querySelector<HTMLFormElement>("#sandra-questionnaire-form");
  const submit = container.querySelector<HTMLButtonElement>("#submit-questionnaire");
  const meter = container.querySelector<HTMLElement>("#answer-meter");

  if (!form || !submit || !meter) {
    return;
  }

  const formElement = form;
  const submitButton = submit;
  const meterElement = meter;

  currentSessionId = formElement.dataset.sessionId ?? "";
  const total = Number(formElement.dataset.questionCount ?? "10");

  function updateMeter() {
    const formData = new FormData(formElement);
    const answered = Array.from(new Set(Array.from(formData.keys()))).length;
    meterElement.textContent = `${answered}/${total}`;
    submitButton.disabled = answered !== total;
  }

  formElement.addEventListener("change", updateMeter);
  formElement.addEventListener("submit", async (event) => {
    event.preventDefault();
    const answers = Object.fromEntries(new FormData(formElement).entries());
    submitButton.disabled = true;
    appendMarkdownMessage("user", "I have completed the investor questionnaire.");
    const statusBubble = appendStatus(
      "Thank you. I am writing those answers into Model.xlsm and reading the workbook-generated investor profile.",
    );

    try {
      const payload = await callChatTurn(
        {
          thread_id: threadId,
          user_message: "I have completed the investor questionnaire.",
          action: "submit_questionnaire",
          session_id: currentSessionId,
          answers,
        },
        {
          onStatus: (message) => {
            setBubbleMarkdown(statusBubble, message, "status-line loader-message");
            addMiniLog(statusBubble, message);
          },
        },
      );
      finishStatusBubble(
        statusBubble,
        asString(payload.status) === "configuration_required"
          ? "Needs operator attention"
          : "Workbook profile returned",
      );
      setStage("profile");
      renderProfile(payload);
    } catch (error) {
      finishStatusBubble(statusBubble, "Profile step stopped");
      appendError(error);
      submitButton.disabled = false;
    }
  });

  updateMeter();
}

function renderQuestionnaire(payload: ToolPayload) {
  if (asString(payload.status) === "configuration_required") {
    appendMarkdownMessage(
      "sandra",
      messageFromPayload(payload, "Sandra needs the workbook MCP server before this step can continue."),
    );
    return;
  }

  const formHtml = asString(payload.form_html);
  if (!formHtml) {
    const message = asString(
      payload.assistant_message,
      "The server did not return a questionnaire form.",
    );
    appendMarkdownMessage("sandra", message);
    return;
  }
  appendMessage("sandra", formHtml);
  const lastBubble = chatScroll.querySelector<HTMLElement>(".message:last-child .bubble");
  if (lastBubble) {
    attachQuestionnaireHandlers(lastBubble);
  }
}

function renderProfile(payload: ToolPayload) {
  if (asString(payload.status) === "configuration_required") {
    appendMarkdownMessage(
      "sandra",
      messageFromPayload(payload, "Sandra needs the workbook MCP server before this step can continue."),
    );
    return;
  }

  const profileMessage = asString(
    payload.creative_profile_message,
    "Your investor profile has been produced by the workbook.",
  );
  const investorProfile = asString(payload.investor_profile);

  appendMessage(
    "sandra",
    `
      <h3>Your workbook profile</h3>
      ${markdownBlock(profileMessage)}
      ${investorProfile ? markdownBlock(investorProfile, "status-line") : ""}
      ${markdownBlock("Before I run the optimizer, please make the short-selling choice explicitly.")}
      <div class="short-choice">
        <button class="choice-button" data-short-selling="false" type="button">No short selling</button>
        <button class="choice-button positive" data-short-selling="true" type="button">Allow short selling</button>
      </div>
    `,
  );

  const lastBubble = chatScroll.querySelector<HTMLElement>(".message:last-child .bubble");
  lastBubble?.querySelectorAll<HTMLButtonElement>("[data-short-selling]").forEach((button) => {
    button.addEventListener("click", () => {
      const allowShortSelling = button.dataset.shortSelling === "true";
      void runOptimizer(allowShortSelling);
    });
  });
}

function renderResultTable(records: Record<string, unknown>[]) {
  if (!records.length) {
    return "<p class=\"status-line\">No final summary table was returned.</p>";
  }

  const columns = Object.keys(records[0]);
  const header = columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("");
  const rows = records
    .map((row) => {
      const cells = columns
        .map((column) => `<td>${escapeHtml(String(row[column] ?? ""))}</td>`)
        .join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");

  return `<table class="result-table"><thead><tr>${header}</tr></thead><tbody>${rows}</tbody></table>`;
}

function renderCharts(images: ChartImage[]) {
  if (!images.length) {
    return "";
  }

  const figures = images
    .map((image, index) => {
      const src = chartImageSrc(image);
      const title = escapeHtml(image.name);
      return `
        <figure>
          <button class="chart-preview" data-chart-index="${index}" type="button">
            <img src="${src}" alt="${title}" />
            <figcaption>
              <span>${title}</span>
              <span class="chart-open-hint">Inspect</span>
            </figcaption>
          </button>
        </figure>
      `;
    })
    .join("");

  return `<div class="chart-grid">${figures}</div>`;
}

function ensureChartLightbox(): HTMLElement {
  const existing = document.getElementById("chart-lightbox");
  if (existing) {
    return existing;
  }

  const lightbox = document.createElement("div");
  lightbox.id = "chart-lightbox";
  lightbox.className = "chart-lightbox";
  lightbox.hidden = true;
  lightbox.innerHTML = `
    <div class="chart-backdrop" data-chart-close></div>
    <section class="chart-panel" role="dialog" aria-modal="true" aria-labelledby="chart-lightbox-title">
      <header class="chart-toolbar">
        <div>
          <p class="eyebrow">Workbook chart</p>
          <h3 id="chart-lightbox-title">Chart preview</h3>
        </div>
        <div class="chart-actions">
          <a class="chart-control" data-chart-download href="#" download>Download PNG</a>
          <button class="chart-control" data-chart-maximize type="button">Maximize</button>
          <button class="chart-control" data-chart-close type="button">Close</button>
        </div>
      </header>
      <div class="chart-viewport">
        <img data-chart-image alt="" />
      </div>
    </section>
  `;
  document.body.appendChild(lightbox);
  lightbox.querySelectorAll<HTMLElement>("[data-chart-close]").forEach((control) => {
    control.addEventListener("click", closeChartLightbox);
  });
  lightbox.querySelector<HTMLButtonElement>("[data-chart-maximize]")?.addEventListener("click", () => {
    void toggleChartLightboxFullscreen();
  });
  return lightbox;
}

function syncChartLightboxFullscreenLabel() {
  const lightbox = document.getElementById("chart-lightbox");
  if (!lightbox) {
    return;
  }
  const maximized = document.fullscreenElement === lightbox || chartLightboxFallbackMaximized;
  lightbox.classList.toggle("is-maximized", maximized);
  const button = lightbox.querySelector<HTMLButtonElement>("[data-chart-maximize]");
  if (button) {
    button.textContent = maximized ? "Restore" : "Maximize";
  }
}

async function toggleChartLightboxFullscreen() {
  const lightbox = ensureChartLightbox();
  if (document.fullscreenElement === lightbox) {
    await document.exitFullscreen();
    chartLightboxFallbackMaximized = false;
    syncChartLightboxFullscreenLabel();
    return;
  }
  try {
    await lightbox.requestFullscreen();
    chartLightboxFallbackMaximized = false;
  } catch {
    chartLightboxFallbackMaximized = !chartLightboxFallbackMaximized;
  }
  syncChartLightboxFullscreenLabel();
}

function closeChartLightbox() {
  const lightbox = document.getElementById("chart-lightbox");
  if (!lightbox) {
    return;
  }
  if (document.fullscreenElement === lightbox) {
    void document.exitFullscreen();
  }
  chartLightboxFallbackMaximized = false;
  lightbox.hidden = true;
  lightbox.classList.remove("is-open", "is-maximized");
}

function openChartLightbox(image: ChartImage) {
  const lightbox = ensureChartLightbox();
  const src = chartImageSrc(image);
  const title = image.name;
  const imageNode = lightbox.querySelector<HTMLImageElement>("[data-chart-image]");
  const titleNode = lightbox.querySelector<HTMLElement>("#chart-lightbox-title");
  const download = lightbox.querySelector<HTMLAnchorElement>("[data-chart-download]");
  if (imageNode) {
    imageNode.src = src;
    imageNode.alt = title;
  }
  if (titleNode) {
    titleNode.textContent = title;
  }
  if (download) {
    download.href = src;
    download.download = chartDownloadName(title);
  }
  lightbox.hidden = false;
  lightbox.classList.add("is-open");
  syncChartLightboxFullscreenLabel();
}

function attachChartHandlers(container: HTMLElement, images: ChartImage[]) {
  container.querySelectorAll<HTMLButtonElement>("[data-chart-index]").forEach((button) => {
    button.addEventListener("click", () => {
      const index = Number(button.dataset.chartIndex);
      const image = images[index];
      if (image) {
        openChartLightbox(image);
      }
    });
  });
}

async function runOptimizer(allowShortSelling: boolean) {
  setStage("optimizer");
  appendMessage(
    "user",
    markdownBlock(allowShortSelling ? "Allow short selling." : "Do not allow short selling."),
  );
  const statusBubble = appendStatus(
    "I am running the optimizer macros and calculator sheet in Model.xlsm. Excel may briefly come forward while workbook charts are exported.",
  );

  try {
    const payload = await callChatTurn(
      {
        thread_id: threadId,
        user_message: allowShortSelling
          ? "Allow short selling and run the optimizer."
          : "Do not allow short selling and run the optimizer.",
        action: "run_mvp",
        session_id: currentSessionId,
        allow_short_selling: allowShortSelling,
      },
      {
        onStatus: (message) => {
          setBubbleMarkdown(statusBubble, message, "status-line loader-message");
          addMiniLog(statusBubble, message);
        },
      },
    );
    finishStatusBubble(
      statusBubble,
      asString(payload.status) === "configuration_required"
        ? "Needs operator attention"
        : "Optimizer artifacts returned",
    );
    if (asString(payload.status) === "configuration_required") {
      appendMarkdownMessage(
        "sandra",
        messageFromPayload(payload, "Sandra needs the workbook MCP server before this step can continue."),
      );
      return;
    }
    const records = asRecordArray(payload.summary_table_records);
    const chartImages = (Array.isArray(payload.chart_images) ? payload.chart_images : []) as ChartImage[];

    const resultBubble = appendMessage(
      "sandra",
      `
        <h3>Workbook optimizer output</h3>
        ${markdownBlock("The final table and charts below were generated from Model.xlsm. The annual return values are workbook model assumptions, not guarantees.")}
        ${renderResultTable(records)}
        ${renderCharts(chartImages)}
      `,
    );
    attachChartHandlers(resultBubble, chartImages);
  } catch (error) {
    finishStatusBubble(statusBubble, "Optimizer step stopped");
    appendError(error);
  }
}

async function startConsultation() {
  startButton.disabled = true;
  setStage("questionnaire");
  appendMarkdownMessage("user", "Start the consultation.");
  const statusBubble = appendStatus(
    "I am opening Model.xlsm, randomizing the questionnaire, and rendering the answer form from the workbook output.",
  );

  try {
    const payload = await callChatTurn(
      {
        thread_id: threadId,
        user_message: "Start Sandra's workbook-backed investor questionnaire.",
        action: "start_questionnaire",
      },
      {
        onStatus: (message) => {
          setBubbleMarkdown(statusBubble, message, "status-line loader-message");
          addMiniLog(statusBubble, message);
        },
      },
    );
    finishStatusBubble(
      statusBubble,
      asString(payload.status) === "configuration_required"
        ? "Needs operator attention"
        : "Questionnaire form ready",
    );
    renderQuestionnaire(payload);
  } catch (error) {
    finishStatusBubble(statusBubble, "Questionnaire step stopped");
    appendError(error);
    startButton.disabled = false;
  }
}

async function sendChatMessage() {
  const message = chatInput.value.trim();
  if (!message) {
    return;
  }
  chatInput.value = "";
  sendChatButton.disabled = true;
  appendMarkdownMessage("user", message);
  const assistantBubble = appendStatus(
    "I am reviewing that with the current conversation memory and workbook-grounded rules.",
  );

  try {
    let streamedText = "";
    let sawToken = false;
    const payload = await callChatTurn(
      {
        thread_id: threadId,
        user_message: message,
        action: "message",
        session_id: currentSessionId || null,
      },
      {
        onStatus: (statusMessage) => {
          if (!sawToken) {
            setBubbleMarkdown(assistantBubble, statusMessage, "status-line loader-message");
            addMiniLog(assistantBubble, statusMessage);
          }
        },
        onToken: (token) => {
          sawToken = true;
          if (!streamedText) {
            replaceLoadingWithMarkdown(assistantBubble, "");
          }
          streamedText += token;
          setBubbleMarkdown(assistantBubble, streamedText);
        },
      },
    );
    const finalMessage = asString(payload.assistant_message, "I am ready to continue.");
    if (!sawToken) {
      replaceLoadingWithMarkdown(assistantBubble, finalMessage);
    } else if (streamedText.trim() !== finalMessage.trim()) {
      setBubbleMarkdown(assistantBubble, finalMessage);
    }
    memoryStatus.textContent = "Conversation saved";
  } catch (error) {
    finishStatusBubble(assistantBubble, "Response stopped");
    appendError(error);
  } finally {
    sendChatButton.disabled = false;
    chatInput.focus();
  }
}

async function sendSafeLog(level: "info" | "error", data: unknown) {
  if (runtimeMode === "mcp") {
    await app.sendLog({ level, data });
    return;
  }
  if (level === "error") {
    console.error(data);
  } else {
    console.info(data);
  }
}

app.onteardown = async () => ({});
app.onerror = (error) => appendError(error);
app.onhostcontextchanged = handleHostContextChanged;
app.ontoolresult = (result) => {
  const payload = extractPayload(result);
  const returnedThreadId = asString(payload.thread_id);
  if (returnedThreadId) {
    memoryStatus.textContent = "SQLite ready";
  }
};

startButton.addEventListener("click", () => {
  void startConsultation();
});

sendChatButton.addEventListener("click", () => {
  void sendChatMessage();
});

chatInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    void sendChatMessage();
  }
});

fullscreenButton.addEventListener("click", async () => {
  if (runtimeMode === "browser") {
    if (!document.fullscreenElement) {
      await document.documentElement.requestFullscreen();
    } else {
      await document.exitFullscreen();
    }
    return;
  }
  try {
    await app.requestDisplayMode({ mode: "fullscreen" });
  } catch {
    await sendSafeLog("info", "Host did not accept fullscreen request.");
  }
});

document.addEventListener("fullscreenchange", syncChartLightboxFullscreenLabel);

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeChartLightbox();
  }
});

async function initializeSharedMemoryStatus() {
  try {
    const payload = await callTool("sandra_chat_memory_snapshot", {
      thread_id: threadId,
    });
    const count = Number(payload.event_count ?? 0);
    const state = payload.state as ToolPayload | undefined;
    currentSessionId = asString(state?.session_id, currentSessionId);
    memoryStatus.textContent = count > 0 ? `${count} saved events` : "SQLite ready";
  } catch (error) {
    memoryStatus.textContent = "Memory unavailable";
    await sendSafeLog("error", String(error));
  }
}

async function initializeMcpMode() {
  await app.connect();
  const ctx = app.getHostContext();
  if (ctx) {
    handleHostContextChanged(ctx);
  }
  await initializeSharedMemoryStatus();
}

async function initializeBrowserMode() {
  fullscreenButton.textContent = "Focus view";
  await initializeSharedMemoryStatus();
}

if (runtimeMode === "mcp") {
  initializeMcpMode().catch((error) => appendError(error));
} else {
  initializeBrowserMode().catch((error) => appendError(error));
}
