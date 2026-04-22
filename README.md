# Ollama Word

A beginner-friendly Microsoft Word task pane add-in that sends Word text to a locally running Ollama model through a small local Node/Express helper.

The current UI is intentionally compact:

- Apply an action directly to currently highlighted Word text.
- Generate new text from a prompt and insert it at the current cursor position.
- Preview output before replacing, inserting below, or copying it.
- Open `Model and settings` only when you want to change model, tone, or target length.

## Architecture Summary

- Word loads the task pane from `https://localhost:3000`.
- The task pane uses Office.js to read, replace, and insert Word text.
- The task pane calls the local helper at `https://localhost:8008`.
- The helper calls Ollama at `http://127.0.0.1:11434`.
- The helper owns prompt construction, model listing, timeouts, and Ollama error handling.

This avoids having the Office add-in call Ollama directly and keeps browser/CORS behavior easier to debug.

## Folder Tree

```text
OllamaWord/
|-- assets/
|   |-- icon-16.png
|   |-- icon-32.png
|   |-- icon-64.png
|   `-- icon-80.png
|-- helper/
|   `-- server.ts
|-- src/
|   |-- commands/
|   |   |-- commands.html
|   |   `-- commands.ts
|   |-- shared/
|   |   `-- actions.ts
|   `-- taskpane/
|       |-- taskpane.css
|       |-- taskpane.html
|       `-- taskpane.ts
|-- .gitignore
|-- LICENSE
|-- manifest.xml
|-- package-lock.json
|-- package.json
|-- README.md
|-- tsconfig.json
`-- vite.config.ts
```

## Setup

1. Install Node.js 18 or newer.

2. Install Ollama and confirm it works:

   ```powershell
   ollama list
   ```

3. Pull at least one model if needed:

   ```powershell
   ollama pull llama3.1:8b
   ```

4. Install project dependencies:

   ```powershell
   npm install
   ```

5. Trust the local Office development certificate:

   ```powershell
   npx office-addin-dev-certs install
   ```

## Run Locally

Start the helper server in one terminal:

```powershell
npm run dev:helper
```

Start the task pane dev server in another terminal:

```powershell
npm run dev:addin
```

Keep both terminals open while using the add-in.

## Sideload Into Word

Word needs a trusted add-in catalog that contains `manifest.xml` directly inside the catalog folder.

Recommended local catalog:

```text
C:\OfficeAddinCatalog
```

Copy only the manifest into that folder:

```powershell
Copy-Item .\manifest.xml C:\OfficeAddinCatalog\manifest.xml -Force
```

Share the folder:

1. Right-click `C:\OfficeAddinCatalog`.
2. Open `Properties > Sharing > Advanced Sharing`.
3. Tick `Share this folder`.
4. Use the share name `OfficeAddinCatalog`.

Confirm this path opens in File Explorer and shows `manifest.xml` directly:

```text
\\localhost\OfficeAddinCatalog
```

Then in Word:

1. Open `File > Options > Trust Center > Trust Center Settings`.
2. Open `Trusted Add-in Catalogs`.
3. Add:

   ```text
   \\localhost\OfficeAddinCatalog
   ```

4. Tick `Show in Menu`.
5. Fully close and reopen Word.
6. Open `Insert > My Add-ins > Shared Folder`.
7. Choose `Ollama Word`.

The whole project does not need to be copied into `C:\OfficeAddinCatalog`. Only `manifest.xml` belongs there.

## Use The Add-In

### Transform Highlighted Text

1. Highlight text in Word.
2. Open the `Ollama Word` task pane.
3. Choose an action, such as `Rewrite`, `Summarise`, or `Academic tone`.
4. Click `Apply to highlighted text`.
5. Review the output preview.
6. Choose `Replace selection`, `Insert below`, or `Copy output`.

The add-in reads the highlighted Word text at the moment you click `Apply to highlighted text`. There is no separate refresh-selection step.

### Write At Cursor

1. Place the cursor where you want new text.
2. Enter a prompt, for example:

   ```text
   Write a short paragraph about the limitations of this method.
   ```

3. Click `Generate and insert at cursor`.

If text is highlighted, Word treats that highlighted text as the current selection, so generated text will replace it. If you want insertion only, place the cursor without highlighting text.

### Change Model And Settings

Open the collapsed `Model and settings` section to change:

- Ollama model
- target length
- tone

Use `Reload` after installing or removing Ollama models.

## Helper Endpoints

Health check:

```powershell
curl.exe -k https://localhost:8008/health
```

List Ollama models:

```powershell
curl.exe -k https://localhost:8008/models
```

Transform selected text:

```powershell
curl.exe -k https://localhost:8008/transform `
  -H "Content-Type: application/json" `
  -d "{\"model\":\"llama3.1:8b\",\"action\":\"rewrite\",\"text\":\"This are a rough sentence.\",\"options\":{\"preserveMeaning\":true,\"targetLength\":\"similar\",\"tone\":\"neutral\"}}"
```

Compose new text:

```powershell
curl.exe -k https://localhost:8008/compose `
  -H "Content-Type: application/json" `
  -d "{\"model\":\"llama3.1:8b\",\"prompt\":\"Write a short paragraph about subject X.\"}"
```

The `-k` flag is useful for command-line testing with local development certificates.

## Debugging Notes

- If `/models` says Ollama is unavailable, confirm `ollama list` works and Ollama is listening on `http://127.0.0.1:11434`.
- If the task pane cannot call the helper, confirm both `npm run dev:addin` and `npm run dev:helper` are running.
- If Word shows certificate warnings or a blank pane, rerun `npx office-addin-dev-certs install`, then restart Word.
- If `Shared Folder` says no add-ins are available, confirm `manifest.xml` is directly inside `C:\OfficeAddinCatalog`, not inside a subfolder.
- If the ribbon button does not appear, remove and re-add the trusted catalog entry, then restart Word.
- If replacement affects the wrong text, reselect the intended text before clicking `Replace selection`.
- If `Generate and insert at cursor` replaces text, check whether text is highlighted. Word replaces the current selection.
- To inspect task pane logs, right-click the task pane and use the web inspector option if available. On some Office builds, use Microsoft Edge DevTools with WebView2 debugging.
- Slow responses usually mean the selected text is large or the selected model is slow. Try a smaller selection or a smaller model.

## Development Commands

Type-check and build:

```powershell
npm run build
```

Run only TypeScript checking:

```powershell
npm run typecheck
```

## Suggested Next Improvements

- Add a setting to choose whether generated cursor text inserts or replaces when text is highlighted.
- Add tracked-changes mode by inserting suggestions instead of replacing text directly.
- Add paragraph-level tools that operate on the paragraph containing the cursor.
- Add whole-document tools with careful confirmation prompts.
- Save favourite prompts in Office roaming settings or local storage.
- Add reusable templates for academic, email, and feedback workflows.
- Add an OpenAI-compatible mode through Ollama for chat-completions-style requests.
