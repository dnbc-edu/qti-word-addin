# Word Add-in MVP Integration Guide

This guide shows how to wire the converter into a Word task pane add-in.

## 1) Authoring contract in Word

Use plain text with this pattern inside the document body:

```text
Title: COMMUNICATIONS ENG EXAM 1

1. Question stem...
- [x] Correct answer
- [ ] Distractor 1
- [ ] Distractor 2
```

## 2) Read document text in task pane

```javascript
async function getDocumentText() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load('text');
    await context.sync();
    return body.text;
  });
}
```

## 3) Generate ZIP

```javascript
import { generateQtiZipBuffer } from '../src/qti-generator.js';

async function buildQti() {
  const text = await getDocumentText();
  const zipBuffer = await generateQtiZipBuffer(text);
  return new Blob([zipBuffer], { type: 'application/zip' });
}
```

## 4) Download ZIP from task pane

```javascript
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}
```

## 5) Recommended UX (minimal)

- Button: `Generate QTI ZIP`
- Small status area:
  - Parsing started
  - Parsed N questions
  - ZIP generated
  - Error details if validation fails

## 6) Validation errors to surface

- No questions found
- Question with fewer than 2 choices
- Question with no correct answer
- Question with multiple correct answers

## 7) Next upgrade after MVP

- Support multi-answer questions
- Support feedback and point values
- Support image/media extraction and manifest linking
- Add QTI profile presets per LMS
