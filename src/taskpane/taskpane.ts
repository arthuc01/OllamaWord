import { TEXT_ACTIONS } from "../shared/actions";

// The task pane is served over HTTPS, so the browser expects helper calls to
// use HTTPS too. The helper still talks to Ollama over http://127.0.0.1:11434.
// Use localhost because office-addin-dev-certs trusts https://localhost.
const HELPER_BASE_URL = "https://localhost:8008";

interface HelperModel {
  name: string;
  size?: string;
  modifiedAt?: string;
}

interface HelperErrorResponse {
  error?: string;
}

interface CursorContext {
  beforeCursor: string;
  afterCursor: string;
}

type Mode = "selection" | "cursor";
type LastOperation = "transform" | "draft" | "insert" | undefined;

const outputTextArea = document.getElementById("output-text") as HTMLTextAreaElement;
const customPromptTextArea = document.getElementById("custom-prompt") as HTMLTextAreaElement;
const selectionCount = document.getElementById("selection-count") as HTMLDivElement;
const actionSelect = document.getElementById("action-select") as HTMLSelectElement;
const modelSelect = document.getElementById("model-select") as HTMLSelectElement;
const targetLengthSelect = document.getElementById("target-length-select") as HTMLSelectElement;
const toneSelect = document.getElementById("tone-select") as HTMLSelectElement;
const activeModelLabel = document.getElementById("active-model-label") as HTMLSpanElement;
const statusArea = document.getElementById("status-area") as HTMLDivElement;
const statusMessage = document.getElementById("status-message") as HTMLDivElement;
const spinner = document.getElementById("spinner") as HTMLDivElement;

const selectionModeButton = document.getElementById("selection-mode-button") as HTMLButtonElement;
const cursorModeButton = document.getElementById("cursor-mode-button") as HTMLButtonElement;
const selectionModePanel = document.getElementById("selection-mode-panel") as HTMLDivElement;
const cursorModePanel = document.getElementById("cursor-mode-panel") as HTMLDivElement;
const runButton = document.getElementById("run-button") as HTMLButtonElement;
const draftAtCursorButton = document.getElementById("draft-at-cursor-button") as HTMLButtonElement;
const writeAtCursorButton = document.getElementById("write-at-cursor-button") as HTMLButtonElement;
const reloadModelsButton = document.getElementById("reload-models-button") as HTMLButtonElement;
const replaceButton = document.getElementById("replace-button") as HTMLButtonElement;
const insertButton = document.getElementById("insert-button") as HTMLButtonElement;
const copyButton = document.getElementById("copy-button") as HTMLButtonElement;
const runAgainButton = document.getElementById("run-again-button") as HTMLButtonElement;

let activeMode: Mode = "selection";
let lastOperation: LastOperation;

/**
 * Office.onReady fires after Office.js has loaded and connected to the host
 * application. Word APIs should only be called after this point.
 */
Office.onReady(() => {
  populateActions();
  wireButtonHandlers();
  void loadModels();
});

function populateActions() {
  actionSelect.innerHTML = "";

  for (const action of TEXT_ACTIONS) {
    const option = document.createElement("option");
    option.value = action.id;
    option.textContent = action.label;
    actionSelect.appendChild(option);
  }
}

function wireButtonHandlers() {
  selectionModeButton.addEventListener("click", () => setMode("selection"));
  cursorModeButton.addEventListener("click", () => setMode("cursor"));
  runButton.addEventListener("click", () => void runTransform());
  runAgainButton.addEventListener("click", () => void runLastOperation());
  draftAtCursorButton.addEventListener("click", () => void generateForCursor(false));
  writeAtCursorButton.addEventListener("click", () => void generateForCursor(true));
  reloadModelsButton.addEventListener("click", () => void loadModels());
  modelSelect.addEventListener("change", updateActiveModelLabel);
  replaceButton.addEventListener("click", () => void replaceCurrentSelection());
  insertButton.addEventListener("click", () => void insertBelowCurrentSelection());
  copyButton.addEventListener("click", () => void copyOutputToClipboard());

  updateOutputActions();
}

function setMode(mode: Mode) {
  activeMode = mode;
  const isSelectionMode = mode === "selection";

  selectionModeButton.classList.toggle("active", isSelectionMode);
  cursorModeButton.classList.toggle("active", !isSelectionMode);
  selectionModePanel.classList.toggle("hidden", !isSelectionMode);
  cursorModePanel.classList.toggle("hidden", isSelectionMode);
}

function setStatus(message: string, kind: "info" | "success" | "error" = "info") {
  statusMessage.textContent = message;
  statusArea.classList.toggle("success", kind === "success");
  statusArea.classList.toggle("error", kind === "error");
}

function setLoading(isLoading: boolean) {
  spinner.classList.toggle("hidden", !isLoading);

  runButton.disabled = isLoading;
  runAgainButton.disabled = isLoading;
  draftAtCursorButton.disabled = isLoading;
  writeAtCursorButton.disabled = isLoading;
  reloadModelsButton.disabled = isLoading;

  if (isLoading) {
    replaceButton.disabled = true;
    insertButton.disabled = true;
    copyButton.disabled = true;
  } else {
    updateOutputActions();
  }
}

function updateCounts(text: string) {
  const trimmedText = text.trim();
  const words = trimmedText.length === 0 ? 0 : trimmedText.split(/\s+/).length;
  selectionCount.textContent = `${words} words / ${text.length} characters`;
}

function getOutputTextOrMessage(): string | undefined {
  const text = outputTextArea.value.trim();

  if (text.length === 0) {
    setStatus("There is no output yet. Run an action first.", "error");
    return undefined;
  }

  return outputTextArea.value;
}

function updateActiveModelLabel() {
  activeModelLabel.textContent = modelSelect.value || "No model";
}

function updateOutputActions() {
  const hasOutput = outputTextArea.value.trim().length > 0;

  replaceButton.disabled = !hasOutput;
  insertButton.disabled = !hasOutput;
  copyButton.disabled = !hasOutput;
  runAgainButton.disabled = !lastOperation;
}

/**
 * Reads a helper response as JSON, but first checks the content type.
 *
 * If the helper is not running, or an old helper does not have the requested
 * route, the browser may receive an HTML error page. Calling response.json()
 * on that HTML causes the confusing "Unexpected token '<'" error. This helper
 * turns that case into a message that explains what to restart.
 */
async function readHelperJson<T>(response: Response, endpointName: string): Promise<T & HelperErrorResponse> {
  const contentType = response.headers.get("content-type") ?? "";

  if (!contentType.includes("application/json")) {
    const responseText = await response.text();
    const shortResponseText = responseText.trim().slice(0, 120);

    throw new Error(
      `The helper returned non-JSON from ${endpointName}. Restart npm run dev:helper. Response started with: ${shortResponseText}`
    );
  }

  return (await response.json()) as T & HelperErrorResponse;
}

/**
 * Word.run creates a request context for Word JavaScript API calls.
 *
 * The Word API is batched: you queue operations such as loading properties,
 * then call context.sync() to send those operations to Word and receive results.
 */
async function readCurrentSelectionText(): Promise<string> {
  return await Word.run(async (context) => {
    const selection = context.document.getSelection();

    // Text is not available immediately. load("text") tells Word which
    // property we want to read on the next context.sync().
    selection.load("text");
    await context.sync();

    return selection.text ?? "";
  });
}

async function loadModels() {
  setLoading(true);
  setStatus("Loading Ollama models...");
  modelSelect.innerHTML = "";

  try {
    const response = await fetch(`${HELPER_BASE_URL}/models`);
    const data = await readHelperJson<{ models?: HelperModel[] }>(response, "/models");

    if (!response.ok) {
      throw new Error(data.error ?? "The helper server could not load models.");
    }

    if (!Array.isArray(data.models)) {
      throw new Error("The helper server returned a malformed model list.");
    }

    if (data.models.length === 0) {
      setStatus("No Ollama models are installed. Run `ollama pull llama3.1:8b` in a terminal.", "error");
      return;
    }

    for (const model of data.models) {
      const option = document.createElement("option");
      option.value = model.name;
      option.textContent = model.size ? `${model.name} (${model.size})` : model.name;
      modelSelect.appendChild(option);
    }

    updateActiveModelLabel();
    setStatus(`Ready. Using ${modelSelect.value}.`, "success");
  } catch (error) {
    console.error(error);
    setStatus(
      error instanceof Error
        ? error.message
        : "Could not contact the local helper server at https://localhost:8008.",
      "error"
    );
  } finally {
    setLoading(false);
  }
}

async function runTransform() {
  if (!modelSelect.value) {
    setStatus("No Ollama model is available. Check that Ollama is running and has a model installed.", "error");
    return;
  }

  setMode("selection");
  setLoading(true);
  setStatus("Reading highlighted text...");

  try {
    const text = await readCurrentSelectionText();
    updateCounts(text);

    if (text.trim().length === 0) {
      setStatus("No text is highlighted. Highlight text in Word, then try again.", "error");
      return;
    }

    setStatus("Sending highlighted text to Ollama...");

    const response = await fetch(`${HELPER_BASE_URL}/transform`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: modelSelect.value,
        action: actionSelect.value,
        text,
        options: {
          preserveMeaning: true,
          targetLength: targetLengthSelect.value,
          tone: toneSelect.value
        }
      })
    });

    const data = await readHelperJson<{ text?: string }>(response, "/transform");

    if (!response.ok) {
      throw new Error(data.error ?? "The helper server could not transform the text.");
    }

    if (typeof data.text !== "string") {
      throw new Error("The helper server returned a malformed transform response.");
    }

    const generatedText = data.text;
    outputTextArea.value = generatedText;
    lastOperation = "transform";
    updateOutputActions();
    setStatus("Output ready.", "success");
  } catch (error) {
    console.error(error);
    setStatus(
      error instanceof Error
        ? error.message
        : "Could not transform the text. Check that the helper and Ollama are running.",
      "error"
    );
  } finally {
    setLoading(false);
  }
}

async function runLastOperation() {
  if (lastOperation === "transform") {
    await runTransform();
    return;
  }

  if (lastOperation === "draft") {
    await generateForCursor(false);
    return;
  }

  if (lastOperation === "insert") {
    await generateForCursor(true);
    return;
  }

  setStatus("Run an action first.", "error");
}

/**
 * Reads a small amount of context around the current cursor.
 *
 * Office.js does not expose a simple "cursor offset inside paragraph" value.
 * To keep this beginner-friendly and reliable, we use range expansion:
 *
 * - Expand the current selection/cursor to the whole containing paragraph.
 * - Create one range from paragraph start to cursor/selection start.
 * - Create another range from cursor/selection end to paragraph end.
 *
 * This gives the helper explicit before/after text, so the model knows where
 * the generated content will be inserted.
 */
async function readCursorContext(): Promise<CursorContext> {
  return await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const paragraphRange = selection.paragraphs.getFirst().getRange();
    const beforeCursorRange = paragraphRange.expandTo(selection);
    const afterCursorRange = selection.expandTo(paragraphRange);

    beforeCursorRange.load("text");
    afterCursorRange.load("text");
    selection.load("text");
    await context.sync();

    return {
      beforeCursor: removeSelectionTextFromEnd(beforeCursorRange.text, selection.text),
      afterCursor: removeSelectionTextFromStart(afterCursorRange.text, selection.text)
    };
  });
}

function removeSelectionTextFromEnd(rangeText: string, selectionText: string) {
  if (selectionText && rangeText.endsWith(selectionText)) {
    return rangeText.slice(0, -selectionText.length).trim();
  }

  return rangeText.trim();
}

function removeSelectionTextFromStart(rangeText: string, selectionText: string) {
  if (selectionText && rangeText.startsWith(selectionText)) {
    return rangeText.slice(selectionText.length).trim();
  }

  return rangeText.trim();
}

async function generateForCursor(shouldInsert: boolean) {
  const prompt = customPromptTextArea.value.trim();

  if (prompt.length === 0) {
    setStatus("Enter a prompt first.", "error");
    return;
  }

  if (!modelSelect.value) {
    setStatus("No Ollama model is available. Check that Ollama is running and has a model installed.", "error");
    return;
  }

  setMode("cursor");
  setLoading(true);
  setStatus("Reading cursor context...");

  try {
    const cursorContext = await readCursorContext();

    setStatus("Generating text...");

    const response = await fetch(`${HELPER_BASE_URL}/compose`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: modelSelect.value,
        prompt,
        beforeCursor: cursorContext.beforeCursor,
        afterCursor: cursorContext.afterCursor
      })
    });

    const data = await readHelperJson<{ text?: string }>(response, "/compose");

    if (!response.ok) {
      throw new Error(data.error ?? "The helper server could not generate text.");
    }

    if (typeof data.text !== "string") {
      throw new Error("The helper server returned a malformed compose response.");
    }

    const generatedText = data.text;
    outputTextArea.value = generatedText;
    lastOperation = shouldInsert ? "insert" : "draft";
    updateOutputActions();

    if (shouldInsert) {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();

        // If text is highlighted, this replaces it. If the cursor is simply
        // placed in the document, this inserts at that cursor position.
        selection.insertText(generatedText, Word.InsertLocation.replace);
        await context.sync();
      });
    }

    setStatus(shouldInsert ? "Generated text inserted at the cursor." : "Draft ready.", "success");
  } catch (error) {
    console.error(error);
    setStatus(
      error instanceof Error
        ? error.message
        : "Could not generate text. Check that the helper and Ollama are running.",
      "error"
    );
  } finally {
    setLoading(false);
  }
}

/**
 * Replaces the current Word selection with the output preview.
 *
 * Word uses the current selection at the moment this button is clicked. For
 * predictable results, keep the same text selected in Word until replacement.
 */
async function replaceCurrentSelection() {
  const outputText = getOutputTextOrMessage();

  if (!outputText) {
    return;
  }

  setLoading(true);
  setStatus("Replacing the current Word selection...");

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(outputText, Word.InsertLocation.replace);
      await context.sync();
    });

    setStatus("Selection replaced.", "success");
  } catch (error) {
    console.error(error);
    setStatus("Could not replace the selection in Word.", "error");
  } finally {
    setLoading(false);
  }
}

/**
 * Inserts output after the current selection.
 *
 * The leading blank lines make the result appear below the selection instead
 * of directly attached to the final selected word.
 */
async function insertBelowCurrentSelection() {
  const outputText = getOutputTextOrMessage();

  if (!outputText) {
    return;
  }

  setLoading(true);
  setStatus("Inserting output below the current Word selection...");

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(`\n\n${outputText}`, Word.InsertLocation.after);
      await context.sync();
    });

    setStatus("Output inserted below the selection.", "success");
  } catch (error) {
    console.error(error);
    setStatus("Could not insert the output in Word.", "error");
  } finally {
    setLoading(false);
  }
}

async function copyOutputToClipboard() {
  const outputText = getOutputTextOrMessage();

  if (!outputText) {
    return;
  }

  try {
    await navigator.clipboard.writeText(outputText);
    setStatus("Output copied to clipboard.", "success");
  } catch (error) {
    console.error(error);
    setStatus("Could not copy to clipboard. Select the output text and copy it manually.", "error");
  }
}
