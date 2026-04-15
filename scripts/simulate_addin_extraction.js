import fs from 'node:fs/promises';
import JSZip from 'jszip';
import path from 'node:path';

async function main() {
  const input = process.argv[2] || path.join('tests', 'GEAS-CHEMISTRY-CHAPTER-1.docx');
  const buffer = await fs.readFile(input);
  const zip = await JSZip.loadAsync(buffer);
  const docFile = zip.file('word/document.xml');
  if (!docFile) {
    console.error('word/document.xml not found in docx');
    process.exit(1);
  }
  const ooxml = await docFile.async('text');

  const result = ooxmlToParserTextRegexFallback(ooxml);
  console.log(result);
  await fs.writeFile(path.join('tests', 'addin_doc_extracted.txt'), result, 'utf8');
}

function decodeXmlEntities(value) {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function normalizeParagraphText(text) {
  return text
    .replace(/\u00A0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/\s*\n\s*/g, '\n')
    .trim();
}

function convertOmmlBlockToPlaceholderByRegex(ommlBlock) {
  const tokenMatches = Array.from(ommlBlock.matchAll(/<m:t[^>]*>([\s\S]*?)<\/m:t>/g));
  const tokens = tokenMatches.map((match) => decodeXmlEntities(match[1]).trim()).filter(Boolean);

  if (!tokens.length) {
    return ' {{EQ}} ';
  }

  let latex = tokens.join(' ');
  const compact = latex.replace(/\s+/g, ' ').trim();

  // Normalize integral symbol
  if (compact.includes('∫')) {
    // replace unicode with 'int' for simple handling
    latex = latex.replace(/∫/g, 'int');
  }

  // Handle limits like: lim x→0 expression  -> \\lim_{x\\to 0} expression
  if (/\blim\b/i.test(compact)) {
    const limMatch = compact.match(/lim\s*([A-Za-z0-9]+)(?:\s|\u2192|->|\\to)*([A-Za-z0-9]+)\s*(.*)/i);
    if (limMatch) {
      const varName = limMatch[1];
      const val = limMatch[2];
      const rest = limMatch[3] || '';
      latex = `\\lim_{${varName}\\to ${val}} ${rest}`.trim();
    }
  }

  // Handle simple integrals: int a b expr -> \\int_{a}^{b} expr
  if (/\bint\b/i.test(compact)) {
    const intMatch = compact.match(/int\s*([^\s]+)?\s*([^\s]+)?\s*(.*)/i);
    if (intMatch) {
      const sub = intMatch[1] || '';
      const sup = intMatch[2] || '';
      const rest = intMatch[3] || '';
      const subFmt = sub ? `_{${sub}}` : '';
      const supFmt = sup ? `^{${sup}}` : '';
      latex = `\\int${subFmt}${supFmt} ${rest}`.trim();
    }
  }

  // Simple exponent fallback: x7 -> x^{7}
  const simpleExp = compact.match(/^([A-Za-z]+)(\d+)$/);
  if (simpleExp) {
    latex = `${simpleExp[1]}^{${simpleExp[2]}}`;
  }

  return ` {{EQ:${latex}}} `;
}

function ooxmlToParserTextRegexFallback(ooxml) {
  const compact = ooxml.replace(/\r?\n/g, ' ');
  const paragraphMatches = compact.match(/<w:p[\s\S]*?<\/w:p>/g) || [];
  const lines = [];

  for (const paragraphXml of paragraphMatches) {
    let transformed = paragraphXml;

    transformed = transformed.replace(/<m:oMathPara[\s\S]*?<\/m:oMathPara>/g, (match) => convertOmmlBlockToPlaceholderByRegex(match));
    transformed = transformed.replace(/<m:oMath[\s\S]*?<\/m:oMath>/g, (match) => convertOmmlBlockToPlaceholderByRegex(match));
    transformed = transformed.replace(/<w:tab\s*\/?\s*>/g, '\t');
    transformed = transformed.replace(/<w:(?:br|cr)\s*\/?\s*>/g, '\n');
    transformed = transformed.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, '$1');
    transformed = transformed.replace(/<[^>]+>/g, ' ');
    transformed = decodeXmlEntities(transformed);

    const cleaned = normalizeParagraphText(transformed);
    if (!cleaned) {
      continue;
    }

    const splitLines = cleaned
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean);

    lines.push(...splitLines);
  }

  return lines.join('\n');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
