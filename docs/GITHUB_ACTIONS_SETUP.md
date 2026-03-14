# GitHub Actions Setup Guide

This guide explains how to use the included GitHub workflows to publish and package the add-in.

## Included workflows

- `.github/workflows/deploy-pages.yml`
  - Builds `dist-addin` and deploys it to GitHub Pages.
  - Triggers on push to `main` and manual run.
  - Also creates and uploads a production bundle artifact (`qti-word-addin-release-prod-from-pages`) using the deployed Pages URL.

- `.github/workflows/create-prod-bundle.yml`
  - Runs checks and creates production bundle artifacts.
  - Manual run with `addin_base_url` input.

## One-time GitHub setup

1. Push this repository to GitHub.
2. Open **Settings → Pages**.
3. Under **Build and deployment**, choose **Source: GitHub Actions**.
4. Ensure Actions are enabled for the repository.

## First deployment (host add-in files)

1. Push to `main` (or run workflow manually).
2. Open **Actions → Deploy Add-in to GitHub Pages**.
3. Wait until workflow succeeds.
4. Copy your Pages URL:
   - `https://<username>.github.io/<repo>` (project site), or
   - `https://<username>.github.io` (user/org site)
5. (Optional) Download artifact `qti-word-addin-release-prod-from-pages` from the same workflow run.

## Create production bundle artifact

1. Open **Actions → Create Production Bundle**.
2. Click **Run workflow**.
3. Set `addin_base_url` to your Pages URL.
4. Run and wait for success.
5. Download artifact: `qti-word-addin-release-prod`.

## Use the artifact outputs

Inside downloaded artifact, use:

- `release/qti-word-addin-release-prod/manifest.xml`
  - Deploy this manifest via Microsoft 365 admin or your chosen channel.

- `release/qti-word-addin-release-prod/dist-addin/`
  - These are the static files expected to be hosted at your `addin_base_url`.

## Verification checklist

- No `https://localhost:3000` remains in production `manifest.xml`.
- Manifest URLs resolve in browser (taskpane page, JS/CSS, icons).
- Add-in opens in Word and can run `Check Questions` and `Generate QTI ZIP`.

## Common issue

If taskpane fails to load, mismatch between `addin_base_url` and actual Pages path is the most common cause. Regenerate the production bundle with the exact live Pages URL.
