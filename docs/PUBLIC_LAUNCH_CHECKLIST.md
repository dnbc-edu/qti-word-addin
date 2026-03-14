# Public Launch Checklist

Use this checklist before distributing to the public.

## 1) GitHub repository and Pages setup

- [ ] Push the project to GitHub.
- [ ] Confirm the target branch for releases (commonly `main`).
- [ ] Enable GitHub Pages for the repository.
- [ ] Confirm the published Pages URL format:
	- Project site: `https://<username>.github.io/<repo>`
	- User/org site: `https://<username>.github.io`
- [ ] Ensure Pages URL is publicly reachable via HTTPS.
- [ ] Confirm stable URL strategy (avoid changing repo name/path after launch).

## 2) Manifest and packaging

- [ ] Bump manifest version in `addin/manifest.xml` for release.
- [ ] Generate production bundle:

```bash
ADDIN_BASE_URL="https://<username>.github.io/<repo>" npm run package:addin:prod
```

- [ ] Validate generated manifest URL replacements in `release/qti-word-addin-release-prod/manifest.xml`.
- [ ] Confirm no `https://localhost:3000` URLs remain in production manifest.
- [ ] Archive final release artifact from `release/qti-word-addin-release-prod.zip`.

## 3) Functional quality

- [ ] Run parser and XML checks:

```bash
npm run check:all
```

- [ ] Validate QTI outputs against sample teacher content.
- [ ] Verify equations render in target LMS imports.
- [ ] Verify output filename format and timestamp behavior.

## 4) Cross-platform compatibility

- [ ] Test on Word Desktop (Windows).
- [ ] Test on Word Desktop (macOS).
- [ ] Test on Word Web.
- [ ] Confirm add-in icon/ribbon labels are correct.
- [ ] Confirm taskpane assets load from GitHub Pages URL (not localhost).

## 5) Legal and trust

- [ ] Keep `LICENSE` present (MIT).
- [ ] Publish Terms of Use URL (default: `docs/TERMS_OF_USE.md`).
- [ ] Publish Privacy Policy URL (default: `docs/PRIVACY_POLICY.md`).
- [ ] Publish support contact URL/email (default: `docs/SUPPORT.md`).
- [ ] Ensure legal/support URLs are ready for app listing and user documentation.

## 6) Distribution

- [ ] Pilot with a small teacher group first.
- [ ] Capture issues and patch before broad rollout.
- [ ] Publish via Microsoft 365 admin deployment (org) and/or AppSource (public marketplace).
- [ ] Prepare announcement docs (quick start + troubleshooting).

## 7) GitHub-specific operational checks

- [ ] Add branch protection for release branch.
- [ ] Restrict direct pushes to release branch.
- [ ] Verify GitHub Actions deployment status after each release.
- [ ] Keep a release tag/changelog for each public update.

## 8) Post-launch operations

- [ ] Define versioning policy and release cadence.
- [ ] Track incidents and user feedback.
- [ ] Add regression checks for fixed bugs.
- [ ] Maintain a changelog for teacher-visible updates.
