# Gemini AI for Microsoft Word

A Microsoft Word task-pane add-in that connects Word to Google Gemini using your own Google AI Studio API key. It can chat with the current document, apply edits with Word Track Changes, insert comments and highlights, run web-backed research, and insert Word math equation objects.

This repository is a community-maintained fork of the MIT-licensed Gemini AI for Office Word add-in by Anson Lai. It keeps the original Office/Gemini document-editing workflow and adds model discovery for newer Google AI Studio models, including Gemini 3.1 and Live API models.

## Why Use This

- No Microsoft Copilot subscription is required.
- You bring your own Google AI Studio API key.
- Edits are applied inside Word rather than only returned as chat text.
- Track Changes/redlines are supported, so you can review and accept/reject edits.
- The model dropdown can refresh from Google `ListModels`, so newly available models can appear without code changes.
- Live API models such as `gemini-3.1-flash-live-preview` are routed through a WebSocket adapter instead of the normal REST endpoint.

## What It Can Do

- Chat with the current Word document.
- Rewrite, improve, or format selected text or whole sections.
- Apply edits as redlines.
- Insert comments.
- Highlight matching text.
- Navigate to sections.
- Edit lists, tables, and document sections.
- Research facts with Google Search tools.
- Insert Word equation objects from LaTeX-style model output.

## Requirements

- Windows or macOS with Microsoft Word desktop.
- Node.js and npm.
- A Google AI Studio API key: https://aistudio.google.com/app/api-keys
- Office add-in development certificates.

## Quick Start From Source

```powershell
git clone <this-repository-url>
cd Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting
npm install
npx office-addin-dev-certs install
npm start
```

`npm start` starts the HTTPS dev server on `https://localhost:3000` and sideloads the add-in into Word using `manifest.xml`.

In Word:

1. Open the add-in task pane.
2. Click the settings gear.
3. Paste your Google AI Studio API key.
4. Click `List usable Google models`.
5. Choose fast and slow models.
6. Save settings.

## Model Support

The add-in supports two Google model paths:

- `generateContent`: the normal official Gemini API path used for chat, redlines, comments, research, and document edits.
- `bidiGenerateContent`: the Gemini Live API WebSocket path used for Live models.

The Settings button `List usable Google models` calls:

```text
https://generativelanguage.googleapis.com/v1beta/models
```

It keeps models that support either `generateContent` or `bidiGenerateContent`.

You can also list models from PowerShell without saving your key in the repo:

```powershell
$env:GEMINI_API_KEY="your-key"
powershell -ExecutionPolicy Bypass -File scripts\list-google-ai-models.ps1
```

Do not commit API keys.

## Equation Workflow

For requests such as "Insert Fourier series in Word math format here":

1. Gemini researches the equation first.
2. If common variants exist, it asks which variant to use.
3. Once clear, it inserts a Word equation object at the cursor by default.

If you explicitly ask for another location, it can insert at the beginning or end.

## Development

Build:

```powershell
npm run build
```

Validate manifest:

```powershell
npx office-addin-manifest validate manifest.xml
```

Run local dev server:

```powershell
npm start
```

Manual sideload:

1. In Word, go to `Insert > Add-ins > My Add-ins`.
2. Choose `Upload My Add-in`.
3. Select `manifest.xml`.

## Security

- The API key is entered by the user in the add-in settings.
- The key is stored by the Office webview in browser localStorage.
- The key is not included in the source code.
- Do not commit `.env`, logs, local manifests, build outputs, or captured keys.

Before publishing, run a secret scan such as:

```powershell
rg -n "paste-the-real-key-here|GEMINI_API_KEY=.*paste-the-real-key-here" . -g "*" --glob "!node_modules/**" --glob "!dist/**"
```

## Packaging Notes

For development, use `manifest.xml` with `https://localhost:3000`.

For your own production deployment, host the built files from `dist/` on an HTTPS origin and update the manifest URLs accordingly. Office add-ins require HTTPS.

This repository intentionally does not include user-specific `manifest.local.xml` files.

## Troubleshooting

Add-in does not show:

```powershell
npx office-addin-dev-settings registered
npx office-addin-dev-settings register manifest.xml
npx office-addin-dev-settings sideload manifest.xml --app Word
```

Model fails with `not found`:

- Click `List usable Google models` in Settings.
- Select only models returned for your API key.
- Live models must use the Live API path; do not manually force them into `generateContent`.

Model only explains and does not edit:

- Use an instruction that asks for a document action, such as "rewrite this paragraph with track changes".
- Make sure the selected model supports tool/function calling.

## Attribution

Original project:

- Gemini AI for Office Microsoft Word Add-in by Anson Lai
- Original repository: https://github.com/AnsonLai/Gemini-AI-for-Office-Microsoft-Word-Add-In-for-Vibe-Drafting
- License: MIT

This fork preserves the upstream MIT license and adds Google model discovery, current model dropdowns, Live API routing, and Word equation insertion improvements.

## License

MIT License. See [LICENSE](LICENSE).
