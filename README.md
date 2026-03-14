# Word Add-in to QTI (MVP)

This workspace contains a working MVP converter and a minimal Word task pane add-in scaffold.

## What this does today

- Parses a simple authoring format exported from Word text.
- Generates a QTI 1.2-style package ZIP with:
  - `imsmanifest.xml`
  - `<assessment-id>/<assessment-id>.xml`
  - `<assessment-id>/assessment_meta.xml`
- Produces LMS-importable packages for platforms that accept this structure (such as Canvas-compatible imports).

## Input format

Use the structure shown in [samples/question-bank-template.txt](samples/question-bank-template.txt):

- First line: `Title: <Assessment Title>`
- Optional global points line right after title: `Points: <N>`
  - Applies to all questions by default
- Each question starts with either:
  - `N. <question stem>`
  - `N) <question stem>`
  - `N <question stem>` (Word auto-list export format)
  - `<question stem>?` (fallback when Word strips list markers in `body.text`)
- Each choice line supports either:
  - Checkbox style: `- [x] Choice text` or `- [ ] Choice text`
  - Lettered style (upper/lower): `a. [x] Choice text`, `a) [ ] Choice text`, `A. [x] Choice text`, `B) [ ] Choice text`
  - Word lettered-list export format: `a <choice text>` / `A <choice text>`
  - Plain choice text lines under a question (for stripped-list Word output)
- You can also mark the correct choice by appending `{{ANS}}` (or `{{ans}}`) at end of the choice line.
  - Example: `B) helical {{ANS}}`
- Equations captured from Word OOXML are preserved as inline equation placeholders and emitted in Canvas equation-image HTML format in QTI item HTML.
  - Placeholder form in parsed text: `{{EQ:...}}`
  - Output form in HTML: `<img class="equation_image" ... data-equation-content="...">`
  - Current OMML-to-LaTeX conversion covers common structures such as radicals and fractions.
- You can set per-question points by appending `(Points: N)` to the question line.
  - Example: `What is valence? (Points: 2)`

Rules:

- Each question must have at least 2 choices.
- Each question must have exactly 1 correct choice.
- If `Points: N` is present, that value is the default for all questions.
- `(Points: N)` on a question overrides the global default for that question.
- If no global or per-question points are set, default is `1`.

## Run locally

```bash
npm install
npm run demo
```

or custom input/output:

```bash
npm run generate:qti -- --input samples/question-bank-template.txt --output dist/my-quiz-qti.zip
```

## Validate generated QTI against local XSDs

Validate the demo ZIP (QTI 1.2):

```bash
npm run validate:demo
```

Validate any ZIP manually:

```bash
npm run validate:qti -- --zip dist/my-quiz-qti.zip --version 1p2
```

Supported versions: `1p2`, `2p1`.

## Parser safety check

Run a negative test that verifies mixed `[x]` and `{{ANS}}` markers in one question are rejected with a clear error:

```bash
npm run check:mixed-marker
```

Run all core checks together:

```bash
npm run check:all
```

## Included Word add-in scaffold

- Manifest: [addin/manifest.xml](addin/manifest.xml)
- Task pane UI: [addin/taskpane.html](addin/taskpane.html)
- Task pane logic: [addin/taskpane.js](addin/taskpane.js)
- Dev server config: [addin/vite.config.js](addin/vite.config.js)

The task pane includes:
- `Check Questions` to verify the document can be converted (parse + XML checks) without downloading.
- `Generate QTI ZIP` to run the same checks and then download the package.

The task pane reads current Word document text and generates a downloadable QTI ZIP.
Before download, it performs a well-formed XML check on generated XML.
With `Strict mode` enabled in the task pane, it checks all generated XML files (`assessment`, `imsmanifest`, and `assessment_meta`).

## Run add-in locally

1. Install dependencies:

```bash
npm install
```

2. Start Word-ready dev flow (sideload + HTTPS server):

```bash
npm run dev:word
```

Or sideload, launch Word, and start server in one step:

```bash
npm run dev:word:open
```

If port `3000` is busy from a previous run, reset it with:

```bash
npm run stop:dev
```

Or do a full reset + sideload + open + serve in one command:

```bash
npm run dev:word:clean
```

This command copies [addin/manifest.xml](addin/manifest.xml) to the default Word sideload folder on macOS, then starts the task pane server.

3. Open Word and look for the `QTI Exporter` button in the Home tab.

4. If sideloading to a custom folder is required, set `WORD_SIDELOAD_DIR`:

```bash
WORD_SIDELOAD_DIR="/custom/wef/path" npm run sideload:word
```

5. Click `Generate QTI ZIP` in the task pane.

## Quick local release bundle

Create a handoff bundle that includes `manifest.xml` and built add-in assets:

```bash
npm run package:addin
```

Output files:
- `release/qti-word-addin-release.zip`
- `release/qti-word-addin-release/`

## Production release bundle

Create a production-ready bundle and rewrite all `https://localhost:3000` manifest URLs to your hosted base URL:

```bash
ADDIN_BASE_URL="https://your-domain.example" npm run package:addin:prod
```

Optionally verify the generated production manifest URLs:

```bash
ADDIN_BASE_URL="https://your-domain.example" npm run verify:prod:manifest
```

Production output files:
- `release/qti-word-addin-release-prod.zip`
- `release/qti-word-addin-release-prod/`

## Public launch prep

Use these docs for a structured public rollout (tailored for GitHub Pages hosting):

- Launch readiness checklist: [docs/PUBLIC_LAUNCH_CHECKLIST.md](docs/PUBLIC_LAUNCH_CHECKLIST.md)
- Release execution runbook: [docs/PUBLIC_RELEASE_RUNBOOK.md](docs/PUBLIC_RELEASE_RUNBOOK.md)
- GitHub Actions setup: [docs/GITHUB_ACTIONS_SETUP.md](docs/GITHUB_ACTIONS_SETUP.md)

GitHub Actions workflows included:
- `.github/workflows/deploy-pages.yml` (deploy static add-in site to GitHub Pages)
- `.github/workflows/create-prod-bundle.yml` (manual production bundle generation with `ADDIN_BASE_URL` input)

After each successful Pages deploy, download artifact `qti-word-addin-release-prod-from-pages` from the workflow run to get a verified production manifest bundle.

## License and legal (draft)

Use this section as a starter for distribution. Replace placeholders before release.

### Project license

This project is licensed under the **MIT License**.

See [LICENSE](LICENSE) for the full license text.

### Third-party licenses

This project uses third-party packages (for example, npm dependencies) that are licensed separately by their respective authors.

You are responsible for reviewing and complying with those licenses when distributing this add-in.

### Terms, privacy, and support

For distributed builds (especially outside internal use), publish and link:
- Terms of Use / EULA: **TBD**
- Privacy Policy: **TBD**
- Support Contact: **TBD**

### Warranty disclaimer (sample)

This software is provided "as is", without warranty of any kind, express or implied, including but not limited to merchantability, fitness for a particular purpose, and noninfringement.

Implementation details and extension ideas are in [docs/WORD_ADDIN_MVP.md](docs/WORD_ADDIN_MVP.md).
