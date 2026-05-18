# Gemini AI for Office Model Refresh and Live API Maintenance

This runbook records the working process used to restore the official Gemini AI for Office behavior while adding automatic Google model discovery and Live API support.

## Goal

Keep the Microsoft Word add-in behaving like the official Gemini AI for Office add-in:

- document edits through Gemini function calling
- redlines and comments preserved
- web research available through the official tool loop
- Word equation objects inserted when requested
- no LocalMind or Ollama routing inside the add-in

At the same time, keep the cloud model dropdown current when Google AI Studio exposes new models.

## Project Structure

- `src/taskpane/taskpane.html`: settings UI and dropdowns
- `src/taskpane/taskpane.js`: model list, official Gemini tool loop, Live adapter, Word equation insertion
- `scripts/list-google-ai-models.ps1`: key-safe Google `ListModels` query
- `scripts/google-ai-usable-models.current.json`: last captured usable model list, without API key
- `manifest.local.xml`: local sideload manifest using `https://localhost:5443/office/taskpane.html`

## Model Refresh

The add-in now has a Settings button named `List usable Google models`.

It calls:

`https://generativelanguage.googleapis.com/v1beta/models?key={key}`

It stores usable models in browser localStorage under:

`geminiUsableGoogleModels`

Usable means:

- `generateContent`: normal official add-in flow
- `bidiGenerateContent`: Live API WebSocket flow

Run from PowerShell without saving the key:

```powershell
$env:GEMINI_API_KEY="your-key"
powershell -ExecutionPolicy Bypass -File F:\workspace_oi\miscelleneous\scripts\list-google-ai-models.ps1
```

## Live API Support

Live models such as `gemini-3.1-flash-live-preview` cannot be sent to the normal REST `generateContent` endpoint.

The add-in routes model names containing `live` to:

`callGeminiLiveAsGenerateContent()`

The Live adapter:

1. Opens a WebSocket to `BidiGenerateContent`.
2. Sends raw `setup`.
3. Sends text through `clientContent` with `turnComplete: true`.
4. Parses string or Blob WebSocket messages.
5. Converts Live output/tool calls into the normal `candidates[0].content.parts` shape expected by the official loop.

Smoke test result for the user key:

- `gemini-3.1-flash-live-preview` opened the WebSocket.
- Received `setupComplete`.
- Received `serverContent`.

## Equation Handling

The add-in includes a dedicated `insert_word_equation` tool.

User requirements:

- no location mentioned -> insert at cursor
- ambiguous famous equations -> research and ask the variant
- Word math format -> Word equation object
- chat should still explain enough
- use web search before inserting famous equations

The prompt section `EQUATION WORKFLOW` enforces this.

## Validation

Run:

```powershell
npm run build
npx office-addin-manifest validate manifest.local.xml
```

Check served files:

```powershell
$html=(Invoke-WebRequest -Uri 'https://localhost:5443/office/taskpane.html' -UseBasicParsing).Content
$js=(Invoke-WebRequest -Uri 'https://localhost:5443/office/taskpane.js' -UseBasicParsing).Content
[pscustomobject]@{
  HasModelRefresh=$html.Contains('List usable Google models')
  HasLiveModel=$html.Contains('gemini-3.1-flash-live-preview')
  HasLiveAdapter=$js.Contains('BidiGenerateContent')
  HasLocalMind=$js.Contains('LocalMind')
  HasOllama=$js.ToLower().Contains('ollama')
  HasSecret=$js.Contains('paste-the-real-key-here')
}
```

Expected:

- model refresh present
- Live model present
- Live adapter present
- LocalMind false
- Ollama false
- secret false

## Sideload Registration

Expected official add-in ID:

`3dcfdb34-70c3-4bfe-8d5e-85089afcf673`

Commands:

```powershell
Copy-Item -LiteralPath 'F:\workspace_oi\miscelleneous\gemini_in_word\Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting\manifest.local.xml' -Destination 'C:\Users\gkl99\AppData\Local\OfficeAddinCatalog\manifest.local.xml' -Force
npx office-addin-dev-settings register manifest.local.xml
npx office-addin-dev-settings registered
npx office-addin-dev-settings sideload manifest.local.xml --app Word
```

## Safety

Never save the API key in this repo. Keep it only in:

- add-in password field/localStorage
- temporary `$env:GEMINI_API_KEY`

Search before finishing:

```powershell
rg -n "paste-the-real-key-here|GEMINI_API_KEY=.*paste-the-real-key-here" F:\workspace_oi\miscelleneous -g "*" --glob "!**/node_modules/**" --glob "!**/dist/**"
```

No hits should remain.

## Installed Skill

The reusable Codex skill is saved at:

`C:\Users\gkl99\.codex\skills\gemini-office-model-refresh\SKILL.md`

Use that skill when model IDs change, when the dropdown needs refresh, when a Live model must be wired, or when the Word add-in stops showing after manifest changes.
