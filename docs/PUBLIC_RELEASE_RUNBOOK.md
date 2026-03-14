# Public Release Runbook

This runbook provides a repeatable GitHub-based release process for public deployment.

## Prerequisites

- Repository is pushed to GitHub.
- GitHub Pages is enabled and publicly reachable via HTTPS.
- Production base URL is known (`https://<username>.github.io/<repo>` or org equivalent).
- You have permission to publish GitHub Pages content (directly or via GitHub Actions).
- You have access to deploy the add-in manifest (admin center, AppSource, or managed sideload process).
- Workflow file is present: `.github/workflows/deploy-pages.yml`.
- Optional packaging workflow is present: `.github/workflows/create-prod-bundle.yml`.

## Release steps

1. Run quality checks:

```bash
npm run check:all
```

2. Build production add-in assets and create production package:

```bash
ADDIN_BASE_URL="https://<username>.github.io/<repo>" npm run package:addin:prod
```

2b. Verify production manifest URLs:

```bash
ADDIN_BASE_URL="https://<username>.github.io/<repo>" npm run verify:prod:manifest
```

3. Review generated manifest:

- File: `release/qti-word-addin-release-prod/manifest.xml`
- Confirm no `https://localhost:3000` URLs remain.
- Confirm icon URLs and taskpane URLs point to GitHub Pages base URL.

4. Deploy static assets to GitHub Pages:

- Deploy folder contents from `release/qti-word-addin-release-prod/dist-addin/` to the published Pages path.
- Ensure the deployed path matches manifest URL paths (`/addin/...` and `/assets/...`).

5. Deploy manifest:

- Use `release/qti-word-addin-release-prod/manifest.xml` for your deployment channel.

6. Smoke test in Word:

- Open add-in in Word.
- Confirm the taskpane loads from GitHub Pages URL.
- Run `Check Questions` and `Generate QTI ZIP` with a sample document.
- Import resulting ZIP into target LMS for final confidence check.

## Optional: build production bundle in GitHub UI

Use this when you want production artifacts without running local commands.

1. Open GitHub repository → **Actions**.
2. Run workflow **Create Production Bundle**.
3. Provide input `addin_base_url` (example: `https://<username>.github.io/<repo>`).
4. Download artifact `qti-word-addin-release-prod` from workflow run.
5. Use the downloaded `manifest.xml` and `dist-addin` outputs for deployment.

## Automatic production bundle from Pages deploy

When `.github/workflows/deploy-pages.yml` runs successfully, it now also uploads:

- Artifact: `qti-word-addin-release-prod-from-pages`

This artifact is generated using the live deployed `page_url` and already includes verified production manifest URLs.

## GitHub Pages quick verification

- Open the taskpane URL from the manifest in a browser.
- Open one icon URL from the manifest and confirm image load.
- Confirm a JS asset under `/assets/` is reachable.
- If any URL fails, fix Pages path/repo base path and re-run packaging.

## Rollback plan

- Keep prior release bundle and prior manifest version archived.
- If a release fails, re-deploy previous static assets and previous manifest.
- Communicate rollback to users and publish a hotfix timeline.
