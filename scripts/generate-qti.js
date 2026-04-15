import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import { generateQtiZipBuffer, parseForValidation } from '../src/qti-generator.js';

// Use mammoth to get more faithful text extraction (including lists)
let mammoth;
try {
  mammoth = await import('mammoth');
} catch (e) {
  // defer to require for older Node resolution
  try {
    // eslint-disable-next-line global-require, import/no-extraneous-dependencies
    mammoth = require('mammoth');
  } catch (err) {
    mammoth = null;
  }
}

function parseArgs(argv) {
  const options = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (token === '--input') {
      options.input = argv[index + 1];
      index += 1;
    }
    if (token === '--output') {
      options.output = argv[index + 1];
      index += 1;
    }
  }
  return options;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.input || !args.output) {
    throw new Error('Usage: node scripts/generate-qti.js --input <path> --output <path>');
  }
  let inputText;
  const ext = path.extname(args.input || '').toLowerCase();
  if (ext === '.docx') {
    const buffer = await fs.readFile(args.input);
    inputText = await extractTextFromDocx(buffer);
  } else {
    inputText = await fs.readFile(args.input, 'utf8');
  }
  const parsed = parseForValidation(inputText);
  const zipBuffer = await generateQtiZipBuffer(inputText);

  await fs.mkdir(path.dirname(args.output), { recursive: true });
  await fs.writeFile(args.output, zipBuffer);

  console.log(`Generated ${args.output}`);
  console.log(`Title: ${parsed.title}`);
  console.log(`Questions: ${parsed.questions.length}`);
}

async function extractTextFromDocx(buffer) {
  // Try mammoth first; if it produces numbered lines, use that.
  // Prefer XML-based extraction below so we can reconstruct numbering and OMML math.
  // Mammoth is available as a helper, but its output may drop or alter equations.

  const zip = await JSZip.loadAsync(buffer);
  const docFile = zip.file('word/document.xml');
  if (!docFile) {
    throw new Error('Invalid .docx file: word/document.xml not found');
  }

  const xmlText = await docFile.async('text');

  // Parse paragraphs and reconstruct numbering from w:numPr
  const paragraphs = [];
  const pRegex = /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/gi;
  let pMatch;
  while ((pMatch = pRegex.exec(xmlText))) {
    const pInner = pMatch[1];

    const ilvlMatch = pInner.match(/<w:ilvl[^>]*w:val="?(\d+)"?[^>]*\/?>(?:<\/w:ilvl>)?/i);
    const numIdMatch = pInner.match(/<w:numId[^>]*w:val="?(\d+)"?[^>]*\/?>(?:<\/w:numId>)?/i);
    const ilvl = ilvlMatch ? Number(ilvlMatch[1]) : null;
    const numId = numIdMatch ? String(numIdMatch[1]) : null;

    let text = '';
    // Convert OMML math blocks (<m:oMath> / <m:oMathPara>) into equation placeholders
    let pProcessed = pInner.replace(/<m:oMath(?:Para)?[\s\S]*?<\/(?:m:oMath|m:oMathPara)>/gi, (mathMatch) => {
      // Handle radicals: <m:rad> with optional <m:deg> (degree) and <m:e> (radicand)
      if (/\<m:rad\b/i.test(mathMatch)) {
        const degBlock = (mathMatch.match(/<m:deg[\s\S]*?<\/m:deg>/i) || [])[0] || '';
        const eBlock = (mathMatch.match(/<m:e[\s\S]*?<\/m:e>/i) || [])[0] || '';
        const tRegex = /<m:t[^>]*>([\s\S]*?)<\/m:t>/gi;
        let m;
        let degText = '';
        let eText = '';
        if (degBlock) {
          while ((m = tRegex.exec(degBlock))) degText += m[1];
          tRegex.lastIndex = 0;
        }
        if (eBlock) {
          while ((m = tRegex.exec(eBlock))) eText += m[1];
          tRegex.lastIndex = 0;
        }
        degText = String(degText || '').trim();
        eText = String(eText || '').trim();
        if (eText) {
          const latex = degText ? `\\sqrt[${degText}]{${eText}}` : `\\sqrt{${eText}}`;
          return `{{EQ:${latex}}}`;
        }
      }

      // If this OMML math contains a fraction (<m:f>), extract numerator and denominator
      if (/\<m:f\b/i.test(mathMatch)) {
        const numBlock = (mathMatch.match(/<m:num[\s\S]*?<\/m:num>/i) || [])[0] || '';
        const denBlock = (mathMatch.match(/<m:den[\s\S]*?<\/m:den>/i) || [])[0] || '';
        const tRegex = /<m:t[^>]*>([\s\S]*?)<\/m:t>/gi;
        let m;
        let numText = '';
        let denText = '';
        if (numBlock) {
          while ((m = tRegex.exec(numBlock))) numText += m[1];
          tRegex.lastIndex = 0;
        }
        if (denBlock) {
          while ((m = tRegex.exec(denBlock))) denText += m[1];
          tRegex.lastIndex = 0;
        }
        numText = String(numText || '').trim();
        denText = String(denText || '').trim();
        if (numText && denText) {
          const latex = `\\frac{${numText}}{${denText}}`;
          return `{{EQ:${latex}}}`;
        }
      }

      // Handle integrals: <m:integral> or <m:nary> encodings with <m:sub>, <m:sup>, and <m:e>
      if (/\<m:integral\b/i.test(mathMatch) || /\<m:nary\b/i.test(mathMatch) && /<m:sub|<m:sup/i.test(mathMatch)) {
        const subBlock = (mathMatch.match(/<m:sub[\s\S]*?<\/m:sub>/i) || [])[0] || '';
        const supBlock = (mathMatch.match(/<m:sup[\s\S]*?<\/m:sup>/i) || [])[0] || '';
        const eBlock = (mathMatch.match(/<m:e[\s\S]*?<\/m:e>/i) || [])[0] || '';
        const tRegex = /<m:t[^>]*>([\s\S]*?)<\/m:t>/gi;
        let m;
        let subText = '';
        let supText = '';
        let eText = '';
        if (subBlock) {
          while ((m = tRegex.exec(subBlock))) subText += m[1];
          tRegex.lastIndex = 0;
        }
        if (supBlock) {
          while ((m = tRegex.exec(supBlock))) supText += m[1];
          tRegex.lastIndex = 0;
        }
        if (eBlock) {
          while ((m = tRegex.exec(eBlock))) eText += m[1];
          tRegex.lastIndex = 0;
        }
        subText = String(subText || '').trim();
        supText = String(supText || '').trim();
        eText = String(eText || '').trim();
        if (eText) {
          const subFmt = subText ? `_{${subText}}` : '';
          const supFmt = supText ? `^{${supText}}` : '';
          let latex = `\\int${subFmt}${supFmt} ${eText}`.trim();
          latex = latex.replace(/\s+dx\b/g, '\\,dx');
          return `{{EQ:${latex}}}`;
        }
      }

      // Handle simple superscripts / subscripts encoded as <m:sSup>, <m:sSub>, or <m:sup>/<m:sub>
      if (/\<m:sSup\b|\<m:sSub\b|<m:sup|<m:sub/i.test(mathMatch)) {
        const eBlock = (mathMatch.match(/<m:e[\s\S]*?<\/m:e>/i) || [])[0] || '';
        const supBlock = (mathMatch.match(/<m:sup[\s\S]*?<\/m:sup>/i) || (mathMatch.match(/<m:sSup[\s\S]*?<\/m:sSup>/i) || []))[0] || '';
        const subBlock = (mathMatch.match(/<m:sub[\s\S]*?<\/m:sub>/i) || (mathMatch.match(/<m:sSub[\s\S]*?<\/m:sSub>/i) || []))[0] || '';
        const tRegex = /<m:t[^>]*>([\s\S]*?)<\/m:t>/gi;
        let m;
        let eText = '';
        let supText = '';
        let subText = '';
        if (eBlock) {
          while ((m = tRegex.exec(eBlock))) eText += m[1];
          tRegex.lastIndex = 0;
        }
        if (supBlock) {
          while ((m = tRegex.exec(supBlock))) supText += m[1];
          tRegex.lastIndex = 0;
        }
        if (subBlock) {
          while ((m = tRegex.exec(subBlock))) subText += m[1];
          tRegex.lastIndex = 0;
        }
        eText = String(eText || '').trim();
        supText = String(supText || '').trim();
        subText = String(subText || '').trim();
        if (eText) {
          let latex = eText;
          if (supText) latex = `${latex}^{${supText}}`;
          if (subText) latex = `${latex}_{${subText}}`;
          return `{{EQ:${latex}}}`;
        }
      }

      // Fallback: flatten inner <m:t> text into placeholder
      const mtRegex = /<m:t[^>]*>([\s\S]*?)<\/m:t>/gi;
      let mtMatch;
      let mathText = '';
      while ((mtMatch = mtRegex.exec(mathMatch))) {
        mathText += mtMatch[1];
      }
      mathText = mathText.replace(/^[\s\u00A0]+|[\s\u00A0]+$/g, '');
      if (!mathText) return '{{EQ}}';
      const converted = convertOmmlMathToLatex(mathText);
      return `{{EQ:${converted}}}`;
    });

    // Replace <w:t> nodes with their content and keep any equation placeholders inserted above
    text = pProcessed.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/gi, '$1');
    // Remove any remaining XML tags
    text = text.replace(/<[^>]+>/g, '');
    text = text.replace(/^[ \t\n]+|[ \t\n]+$/g, '');
    paragraphs.push({ text, numId, ilvl });
  }

  const counters = {};
  const lines = [];

  for (const p of paragraphs) {
    if (!p.text) {
      continue;
    }

    if (p.numId != null) {
      counters[p.numId] = counters[p.numId] || [];
      const lvl = p.ilvl != null ? p.ilvl : 0;
      for (let i = 0; i <= lvl; i++) {
        if (counters[p.numId][i] == null) counters[p.numId][i] = 0;
      }
      counters[p.numId][lvl] += 1;
      for (let i = lvl + 1; i < counters[p.numId].length; i++) counters[p.numId][i] = 0;

      if (lvl === 0) {
        lines.push(`${counters[p.numId][0]}. ${p.text}`);
      } else {
        lines.push(`- ${p.text}`);
      }
    } else {
      lines.push(p.text);
    }
  }

  const extractedText = lines.join('\n\n').trim();
  try {
    await fs.writeFile(path.join(process.cwd(), 'tests', 'doc_extracted.txt'), extractedText, 'utf8');
  } catch (e) {
    // ignore write errors
  }

  return extractedText;
}

function convertOmmlMathToLatex(mathText) {
  let t = String(mathText || '');

  // Normalize unicode and arrows to LaTeX-friendly tokens
  // Ensure integral symbols are normalized to 'int' so regexes map them to \\int
  t = t.replace(/∫/g, 'int');
  t = t.replace(/∞/g, 'infty');
  t = t.replace(/->|→|⇒/g, '\\to');
  t = t.replace(/\s+/g, ' ').trim();

  // Limits: handle 'limx\to0x+7' and similar forms
  let m = t.match(/^lim\s*([A-Za-z0-9]+)\\to\s*([^\s]+)\s*(.*)$/i);
  if (m) {
    let varName = m[1];
    let val = m[2];
    let rest = String(m[3] || '').trim();
    const split = val.match(/^(\d+)(.+)$/);
    if (!rest && split) {
      val = split[1];
      rest = split[2] || '';
    }
    return `\\lim_{${varName}\\to ${val}} ${rest}`.trim();
  }
  m = t.match(/^lim\s*([A-Za-z0-9]+)→?\\?to?\s*([^\s]+)\s*(.*)$/i);
  if (m) {
    let varName = m[1];
    let val = m[2];
    let rest = String(m[3] || '').trim();
    const split = val.match(/^(\d+)(.+)$/);
    if (!rest && split) {
      val = split[1];
      rest = split[2] || '';
    }
    return `\\lim_{${varName}\\to ${val}} ${rest}`.trim();
  }

  // Sums: map to \sum_{...}^{...}
  m = t.match(/^sum_?([^\^\s]+)\^?([^\s]+)?\s*(.*)$/i);
  if (m && /sum/i.test(t)) {
    const sub = m[1] || '';
    const sup = m[2] || '';
    const rest = String(m[3] || '').trim();
    const subFmt = sub ? `_{${sub}}` : '';
    const supFmt = sup ? `^{${sup}}` : '';
    return `\\sum${subFmt}${supFmt} ${rest}`.trim();
  }

  // Integrals: map to \int_{a}^{b} f and ensure \\,dx spacing
  m = t.match(/^(?:int|∫)_?([^\^\s]+)\^?([^\s]+)?\s*(.*)$/i);
  if (m && /int|∫/i.test(t)) {
    const sub = m[1] || '';
    const sup = m[2] || '';
    let rest = String(m[3] || '').trim();
    const subFmt = sub ? `_{${sub}}` : '';
    const supFmt = sup ? `^{${sup}}` : '';
    let out = `\\int${subFmt}${supFmt} ${rest}`.trim();
    out = out.replace(/([^\\])\s*dx\b/g, '$1\\,dx');
    out = out.replace(/\s+dx\b/g, '\\,dx');
    return out;
  }

  // Replace 'infty' with \infty
  t = t.replace(/infty/gi, '\\infty');

  // Map common trig/log functions to LaTeX commands
  t = t.replace(/\b(arcsin|arccos|arctan|sin|cos|tan|csc|sec|cot)\b/gi, (m) => `\\${m.toLowerCase()}`);
  t = t.replace(/\b(log|ln|exp)\b/gi, (m) => `\\${m.toLowerCase()}`);

  // Fallback: if expression looks like a single identifier immediately followed by digits (e.g., "x7"), treat as exponent
  // Only match when there's no operator (+-*/^) to avoid false positives like "x+7".
  const simpleExpMatch = t.match(/^([A-Za-z]+)(\d+)$/);
  if (simpleExpMatch) {
    return `${simpleExpMatch[1]}^{${simpleExpMatch[2]}}`;
  }

  return t;
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
