# Contributing

Thanks for contributing to QTI Word Add-in.

## Development setup

```bash
npm install
```

## Useful commands

- `npm run check:all` — run parser and XML validation checks.
- `npm run build:addin` — build add-in assets.
- `npm run dev:word:clean` — clean sideload + local Word dev flow.
- `npm run package:addin:prod` — create production release bundle.

## Pull request expectations

- Keep changes focused and minimal.
- Preserve existing behavior unless a change is required.
- Update docs when behavior, commands, or release steps change.
- Ensure checks pass before opening a PR.

## Reporting issues

Use GitHub Issues with:

- Reproduction steps
- Expected vs actual behavior
- Platform details (start with Word Web, then macOS/Windows Desktop if relevant)
- Sample input (if shareable)