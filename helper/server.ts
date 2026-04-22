import express, { Request, Response } from "express";
import https from "node:https";
import { getHttpsServerOptions } from "office-addin-dev-certs";
import { findAction, TEXT_ACTIONS } from "../src/shared/actions";

const HELPER_PORT = Number(process.env.HELPER_PORT ?? 8008);
const OLLAMA_BASE_URL = process.env.OLLAMA_BASE_URL ?? "http://127.0.0.1:11434";
const OLLAMA_TIMEOUT_MS = Number(process.env.OLLAMA_TIMEOUT_MS ?? 120_000);

const app = express();

app.use(express.json({ limit: "2mb" }));

/**
 * The task pane runs on https://localhost:3000 and calls this helper on
 * https://localhost:8008. Browsers enforce CORS between those origins, so the
 * helper explicitly allows local development requests.
 */
app.use((request, response, next) => {
  response.setHeader("Access-Control-Allow-Origin", "*");
  response.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  response.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (request.method === "OPTIONS") {
    response.sendStatus(204);
    return;
  }

  next();
});

interface OllamaModel {
  name?: string;
  model?: string;
  size?: number;
  modified_at?: string;
}

interface TransformRequest {
  model?: string;
  action?: string;
  text?: string;
  options?: {
    preserveMeaning?: boolean;
    targetLength?: string;
    tone?: string;
  };
}

interface ComposeRequest {
  model?: string;
  prompt?: string;
  context?: string;
}

/**
 * fetchWithTimeout wraps fetch with AbortController so a slow or stuck Ollama
 * request returns a useful error instead of leaving the add-in waiting forever.
 */
async function fetchWithTimeout(url: string, options: RequestInit = {}) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), OLLAMA_TIMEOUT_MS);

  try {
    return await fetch(url, {
      ...options,
      signal: controller.signal
    });
  } finally {
    clearTimeout(timeout);
  }
}

function formatBytes(bytes: number | undefined): string | undefined {
  if (!bytes || !Number.isFinite(bytes)) {
    return undefined;
  }

  const gibibytes = bytes / 1024 / 1024 / 1024;
  return `${gibibytes.toFixed(1)} GB`;
}

function sendError(response: Response, status: number, message: string, details?: unknown) {
  if (details) {
    console.error(message, details);
  } else {
    console.error(message);
  }

  response.status(status).json({
    error: message
  });
}

/**
 * Prompt builder for Ollama.
 *
 * The helper owns prompt construction so the task pane can stay simple. If you
 * later want to switch from /api/generate to /api/chat, this is also where you
 * can keep the same policy and action instructions.
 */
function buildPrompt(request: Required<Pick<TransformRequest, "action" | "text">> & TransformRequest) {
  const action = findAction(request.action);

  if (!action) {
    throw new Error(`Unknown action: ${request.action}`);
  }

  const targetLength = request.options?.targetLength ?? "similar";
  const tone = request.options?.tone ?? "neutral";

  return `You are a careful writing assistant working on text selected in a Microsoft Word document.

Task:
${action.helperInstruction}

Rules:
- Preserve the original language unless the user explicitly asks otherwise.
- Do not invent facts.
- Do not add citations unless citations are already present in the source text.
- Keep names, dates, chemical terms, and technical identifiers unchanged unless a correction is obvious.
- Avoid over-formatting unless the requested task explicitly asks for formatting.
- Return only the transformed text, with no preamble or explanation.
- Target length: ${targetLength}.
- Tone: ${tone}.

Selected text:
"""${request.text}"""
`;
}

/**
 * Prompt builder for generating new text at the current Word cursor.
 *
 * This is deliberately separate from buildPrompt because there may be no
 * selected source text. The user instruction is the source of the new content.
 */
function buildComposePrompt(request: Required<Pick<ComposeRequest, "prompt">> & Pick<ComposeRequest, "context">) {
  const contextBlock = request.context?.trim()
    ? `Nearby document context:
"""${request.context.trim()}"""

`
    : "";

  return `You are a careful writing assistant drafting text for a Microsoft Word document.

${contextBlock}User request:
${request.prompt}

Rules:
- Use the nearby document context only to understand topic, style, terminology, and continuity.
- Do not rewrite, summarize, or quote the context unless the user explicitly asks for that.
- Follow the user's request directly.
- Do not invent specific facts, names, dates, citations, or references unless the user provided them.
- If the request asks for factual detail that is not supplied, keep the wording general.
- Write polished text that can be inserted directly into the document.
- Return only the drafted text, with no preamble or explanation.
`;
}

/**
 * Older compose prompt without document context kept here as a readable example
 * of the same endpoint's minimum behavior.
 */
function buildStandaloneComposePrompt(request: Required<Pick<ComposeRequest, "prompt">>) {
  return `You are a careful writing assistant drafting text for a Microsoft Word document.

User request:
${request.prompt}

Rules:
- Follow the user's request directly.
- Do not invent specific facts, names, dates, citations, or references unless the user provided them.
- If the request asks for factual detail that is not supplied, keep the wording general.
- Write polished text that can be inserted directly into the document.
- Return only the drafted text, with no preamble or explanation.
`;
}

async function callOllamaGenerate(model: string, prompt: string): Promise<string> {
  const ollamaResponse = await fetchWithTimeout(`${OLLAMA_BASE_URL}/api/generate`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model,
      prompt,
      stream: false
    })
  });

  if (!ollamaResponse.ok) {
    const body = await ollamaResponse.text();
    throw new Error(`Ollama returned HTTP ${ollamaResponse.status}: ${body}`);
  }

  const data = (await ollamaResponse.json()) as { response?: string };

  if (typeof data.response !== "string") {
    throw new Error("Ollama returned a malformed /api/generate response.");
  }

  return data.response.trim();
}

/**
 * Alternative chat implementation kept here as a simple future switch.
 * To use it, call callOllamaChat from /transform instead of callOllamaGenerate.
 */
async function callOllamaChat(model: string, prompt: string): Promise<string> {
  const ollamaResponse = await fetchWithTimeout(`${OLLAMA_BASE_URL}/api/chat`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model,
      stream: false,
      messages: [
        {
          role: "user",
          content: prompt
        }
      ]
    })
  });

  if (!ollamaResponse.ok) {
    const body = await ollamaResponse.text();
    throw new Error(`Ollama returned HTTP ${ollamaResponse.status}: ${body}`);
  }

  const data = (await ollamaResponse.json()) as { message?: { content?: string } };

  if (typeof data.message?.content !== "string") {
    throw new Error("Ollama returned a malformed /api/chat response.");
  }

  return data.message.content.trim();
}

app.get("/health", (_request, response) => {
  response.json({
    ok: true,
    helper: "running",
    ollamaBaseUrl: OLLAMA_BASE_URL
  });
});

app.get("/actions", (_request, response) => {
  response.json({
    actions: TEXT_ACTIONS.map(({ id, label }) => ({ id, label }))
  });
});

app.get("/models", async (_request: Request, response: Response) => {
  console.log(`[models] Fetching Ollama models from ${OLLAMA_BASE_URL}/api/tags`);

  try {
    const ollamaResponse = await fetchWithTimeout(`${OLLAMA_BASE_URL}/api/tags`, {
      method: "GET"
    });

    if (!ollamaResponse.ok) {
      throw new Error(`Ollama returned HTTP ${ollamaResponse.status}`);
    }

    const data = (await ollamaResponse.json()) as { models?: OllamaModel[] };

    if (!Array.isArray(data.models)) {
      throw new Error("Ollama returned a malformed /api/tags response.");
    }

    response.json({
      models: data.models.map((model) => ({
        name: model.name ?? model.model ?? "unknown",
        size: formatBytes(model.size),
        modifiedAt: model.modified_at
      }))
    });
  } catch (error) {
    sendError(
      response,
      503,
      "Ollama does not appear to be running on localhost:11434.",
      error
    );
  }
});

app.post("/transform", async (request: Request, response: Response) => {
  const body = request.body as TransformRequest;

  if (!body.model || typeof body.model !== "string") {
    sendError(response, 400, "Missing model.");
    return;
  }

  if (!body.action || typeof body.action !== "string") {
    sendError(response, 400, "Missing action.");
    return;
  }

  if (!body.text || typeof body.text !== "string" || body.text.trim().length === 0) {
    sendError(response, 400, "Missing selected text.");
    return;
  }

  try {
    const prompt = buildPrompt({
      model: body.model,
      action: body.action,
      text: body.text,
      options: body.options
    });

    console.log(`[transform] model=${body.model} action=${body.action} chars=${body.text.length}`);

    // Start with /api/generate because it is simple. callOllamaChat above is
    // ready if you later prefer chat-style messages.
    const transformedText = await callOllamaGenerate(body.model, prompt);

    response.json({
      text: transformedText
    });
  } catch (error) {
    const message =
      error instanceof Error && error.name === "AbortError"
        ? "The Ollama request timed out. Try a smaller selection or a faster model."
        : "The helper could not transform the text with Ollama.";

    sendError(response, 500, message, error);
  }
});

app.post("/compose", async (request: Request, response: Response) => {
  const body = request.body as ComposeRequest;

  if (!body.model || typeof body.model !== "string") {
    sendError(response, 400, "Missing model.");
    return;
  }

  if (!body.prompt || typeof body.prompt !== "string" || body.prompt.trim().length === 0) {
    sendError(response, 400, "Missing prompt.");
    return;
  }

  try {
    const prompt = buildComposePrompt({
      prompt: body.prompt,
      context: body.context
    });

    console.log(
      `[compose] model=${body.model} promptChars=${body.prompt.length} contextChars=${body.context?.length ?? 0}`
    );

    const draftedText = await callOllamaGenerate(body.model, prompt);

    response.json({
      text: draftedText
    });
  } catch (error) {
    const message =
      error instanceof Error && error.name === "AbortError"
        ? "The Ollama request timed out. Try a shorter prompt or a faster model."
        : "The helper could not generate text with Ollama.";

    sendError(response, 500, message, error);
  }
});

async function startServer() {
  const httpsOptions = await getHttpsServerOptions();

  https.createServer(httpsOptions, app).listen(HELPER_PORT, "localhost", () => {
    console.log(`Ollama Word helper listening at https://localhost:${HELPER_PORT}`);
    console.log(`Ollama expected at ${OLLAMA_BASE_URL}`);
  });
}

startServer().catch((error) => {
  console.error("Could not start the helper server.", error);
  process.exit(1);
});
