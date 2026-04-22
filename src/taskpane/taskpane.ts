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

const outputTextArea = document.getElementById("output-text") as HTMLTextAreaElement;
const customPromptTextArea = document.getElementById("custom-prompt") as HTMLTextAreaElement;
const selectionCount = document.getElementById("selection-count") as HTMLDivElement;
const actionSelect = document.getElementById("action-select") as HTMLSelectElement;
const modelSelect = document.getElementById("model-select") as HTMLSelectElement;
const targetLengthSelect = document.getElementById("target-length-select") as HTMLSelectElement;
const toneSelect = document.getElementById("tone-select") as HTMLSelectElement;
const statusMessage = document.getElementById("status-message") as HTMLDivElement;
const spinner = document.getElementById("spinner") as HTMLDivElement;

const runButton = document.getElementById("run-button") as HTMLButtonElement;
const writeAtCursorButton = document.getElementById("write-at-cursor-button") as HTMLButtonElement;
const reloadModelsButton = document.getElementById("reload-models-button") as HTMLButtonElement;
const replaceButton = document.getElementById("replace-button") as HTMLButtonElement;
const insertButton = document.getElementById("insert-button") as HTMLButtonElement;
const copyButton = document.getElementById("copy-button") as HTMLButtonElement;
const runAgainButton = document.getElementById("run-again-button") as HTMLButtonElement;

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
  runButton.addEventListener("click", () => void runTransform());
  runAgainButton.addEventListener("click", () => void runTransform());
  writeAtCursorButton.addEventListener("click", () => void generateAndInsertAtCursor());
  reloadModelsButton.addEventListener("click", () => void loadModels());
  replaceButton.addEventListener("click", () => void replaceCurrentSelection());
  insertButton.addEventListener("click", () => void insertBelowCurrentSelection());
  copyButton.addEventListener("click", () => void copyOutputToClipboard());
}

function setStatus(message: string) {
  statusMessage.textContent = message;
}

function setLoading(isLoading: boolean) {
  spinner.classList.toggle("hidden", !isLoading);

  runButton.disabled = isLoading;
  runAgainButton.disabled = isLoading;
  writeAtCursorButton.disabled = isLoading;
  reloadModelsButton.disabled = isLoading;
}

function updateCounts(text: string) {
  const trimmedText = text.trim();
  const words = trimmedText.length === 0 ? 0 : trimmedText.split(/\s+/).length;
  selectionCount.textContent = `${words} words / ${text.length} characters`;
}

function getOutputTextOrMessage(): string | undefined {
  const text = outputTextArea.value.trim();

  if (text.length === 0) {
    setStatus("There is no output yet. Run an action first.");
    return undefined;
  }

  return outputTextArea.value;
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
    const data = (await response.json()) as { models?: HelperModel[]; error?: string };

    if (!response.ok) {
      throw new Error(data.error ?? "The helper server could not load models.");
    }

    if (!Array.isArray(data.models)) {
      throw new Error("The helper server returned a malformed model list.");
    }

    if (data.models.length === 0) {
      setStatus("No Ollama models are installed. Run `ollama pull llama3.1:8b` in a terminal.");
      return;
    }

    for (const model of data.models) {
      const option = document.createElement("option");
      option.value = model.name;
      option.textContent = model.size ? `${model.name} (${model.size})` : model.name;
      modelSelect.appendChild(option);
    }

    setStatus(`Ready. Using ${modelSelect.value}.`);
  } catch (error) {
    console.error(error);
    setStatus(
      error instanceof Error
        ? error.message
        : "Could not contact the local helper server at https://localhost:8008."
    );
  } finally {
    setLoading(false);
  }
}

async function runTransform() {
  if (!modelSelect.value) {
    setStatus("No Ollama model is available. Check that Ollama is running and has a model installed.");
    return;
  }

  setLoading(true);
  setStatus("Reading highlighted text...");

  try {
    const text = await readCurrentSelectionText();
    updateCounts(text);

    if (text.trim().length === 0) {
      setStatus("No text is highlighted. Highlight text in Word, then try again.");
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

    const data = (await response.json()) as { text?: string; error?: string };

    if (!response.ok) {
      throw new Error(data.error ?? "The helper server could not transform the text.");
    }

    if (typeof data.text !== "string") {
      throw new Error("The helper server returned a malformed transform response.");
    }

    const generatedText = data.text;
    outputTextArea.value = generatedText;
    setStatus("Output ready. Review it before inserting or replacing text.");
  } catch (error) {
    console.error(error);
    setStatus(
      error instanceof Error
        ? error.message
        : "Could not transform the text. Check that the helper and Ollama are running."
    );
  } finally {
    setLoading(false);
  }
}

/**
 * Reads a small amount of context around the current cursor.
 *
 * The simplest reliable context for a beginner-friendly Word add-in is the
 * paragraph containing the current selection/cursor. This gives the model
 * nearby writing style and topic without needing whole-document access.
 */
async function readCursorParagraphContext(): Promise<string> {
  return await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const paragraphs = selection.paragraphs;

    paragraphs.load("items/text");
    await context.sync();

    return paragraphs.items.map((paragraph) => paragraph.text).join("\n").trim();
  });
}

async function generateAndInsertAtCursor() {
  const prompt = customPromptTextArea.value.trim();

  if (prompt.length === 0) {
    setStatus("Enter a prompt first.");
    return;
  }

  if (!modelSelect.value) {
    setStatus("No Ollama model is available. Check that Ollama is running and has a model installed.");
    return;
  }

  setLoading(true);
  setStatus("Reading cursor context...");

  try {
    const cursorContext = await readCursorParagraphContext();

    setStatus("Generating text...");

    const response = await fetch(`${HELPER_BASE_URL}/compose`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: modelSelect.value,
        prompt,
        context: cursorContext
      })
    });

    const data = (await response.json()) as { text?: string; error?: string };

    if (!response.ok) {
      throw new Error(data.error ?? "The helper server could not generate text.");
    }

    if (typeof data.text !== "string") {
      throw new Error("The helper server returned a malformed compose response.");
    }

    const generatedText = data.text;
    outputTextArea.value = generatedText;

    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      // If text is highlighted, this replaces it. If the cursor is simply
      // placed in the document, this inserts at that cursor position.
      selection.insertText(generatedText, Word.InsertLocation.replace);
      await context.sync();
    });

    setStatus("Generated text inserted at the cursor.");
  } catch (error) {
    console.error(error);
    setStatus(
      error instanceof Error
        ? error.message
        : "Could not generate text. Check that the helper and Ollama are running."
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

    setStatus("Selection replaced.");
  } catch (error) {
    console.error(error);
    setStatus("Could not replace the selection in Word.");
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

    setStatus("Output inserted below the selection.");
  } catch (error) {
    console.error(error);
    setStatus("Could not insert the output in Word.");
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
    setStatus("Output copied to clipboard.");
  } catch (error) {
    console.error(error);
    setStatus("Could not copy to clipboard. Select the output text and copy it manually.");
  }
}
