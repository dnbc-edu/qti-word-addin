import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';

// Polyfill DOMParser using xmldom
let DOMParser;
try {
  // eslint-disable-next-line import/no-extraneous-dependencies
  const { DOMParser: XmldomParser } = await import('xmldom');
  DOMParser = XmldomParser;
} catch (e) {
  console.error('Please run: npm install xmldom');
  process.exit(1);
}

// Minimal Node global for nodeType constants
global.Node = { ELEMENT_NODE: 1, TEXT_NODE: 3 };
global.DOMParser = DOMParser;

// Provide a minimal `document` and `window` polyfill so addin/taskpane.js can import
global.document = global.document || {
  getElementById: () => ({}),
  createElement: () => ({ setAttribute: () => {}, appendChild: () => {} }),
  body: { appendChild: () => {} }
};
global.window = global.window || {};
// Minimal Office.js polyfill so module-level `Office.onReady` calls don't fail
global.Office = global.Office || {
  onReady: (cb) => { try { cb({}); } catch (e) {} },
  context: { document: {} }
};

// Inline necessary extraction helpers from addin/taskpane.js so we can run in Node
function decodeXmlEntities(value) {
  return String(value || '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

function firstDirectChildByTag(node, namespace, localName) {
  if (!node || !node.childNodes) return null;
  for (let i = 0; i < node.childNodes.length; i += 1) {
    const child = node.childNodes[i];
    if (child.nodeType === 1 && child.namespaceURI === namespace && child.localName === localName) return child;
  }
  return null;
}

function renderChildContent(node) {
  let output = '';
  for (let i = 0; i < node.childNodes.length; i += 1) {
    const ch = node.childNodes[i];
    output += renderOmmlNodeToLatex(ch);
  }
  return output;
}

function renderRadicalToLatex(radNode) {
  const degreeNode = firstDirectChildByTag(radNode, MATH_NS, 'deg');
  const expressionNode = firstDirectChildByTag(radNode, MATH_NS, 'e');
  const degreeLatex = degreeNode ? renderChildContent(degreeNode).trim() : '';
  const expressionLatex = expressionNode ? renderChildContent(expressionNode).trim() : '';
  if (!expressionLatex) return '';
  if (degreeLatex) return `\\sqrt[${degreeLatex}]{${expressionLatex}}`;
  return `\\sqrt{${expressionLatex}}`;
}

function renderFractionToLatex(fractionNode) {
  const numeratorNode = firstDirectChildByTag(fractionNode, MATH_NS, 'num');
  const denominatorNode = firstDirectChildByTag(fractionNode, MATH_NS, 'den');
  const numeratorLatex = numeratorNode ? renderChildContent(numeratorNode).trim() : '';
  const denominatorLatex = denominatorNode ? renderChildContent(denominatorNode).trim() : '';
  if (!numeratorLatex || !denominatorLatex) return '';
  return `\\frac{${numeratorLatex}}{${denominatorLatex}}`;
}

function renderOmmlNodeToLatex(node) {
  if (!node) return '';
  if (node.nodeType === 3) return node.textContent || '';
  if (node.nodeType !== 1) return '';
  const element = node;
  if (element.namespaceURI !== MATH_NS) return renderChildContent(element);
  if (element.localName === 'oMath' || element.localName === 'oMathPara') return renderChildContent(element);
  if (element.localName === 'rad') return renderRadicalToLatex(element);
  if (element.localName === 'f') return renderFractionToLatex(element);
  if (element.localName === 't') return element.textContent || '';
  if (element.localName === 'nary') {
    const subNode = firstDirectChildByTag(element, MATH_NS, 'sub');
    const supNode = firstDirectChildByTag(element, MATH_NS, 'sup');
    const eNode = firstDirectChildByTag(element, MATH_NS, 'e');
    const subLatex = subNode ? renderChildContent(subNode).trim() : '';
    const supLatex = supNode ? renderChildContent(supNode).trim() : '';
    const exprLatex = eNode ? renderChildContent(eNode).trim() : '';
    const subFmt = subLatex ? `_{${subLatex}}` : '';
    const supFmt = supLatex ? `^{${supLatex}}` : '';
    if (!exprLatex) return '';
    return `\\int${subFmt}${supFmt} ${exprLatex}`;
  }
  if (element.localName === 'func') {
    const limElems = element.getElementsByTagNameNS(MATH_NS, 'lim');
    if (limElems && limElems.length) {
      const limNode = limElems[0];
      const limText = renderChildContent(limNode).replace(/\s+/g, ' ').trim();
      const parts = limText.split(/→|->|\\\\to|\sto|\s+to\s+/).map(p => p.trim()).filter(Boolean);
      const varName = parts[0] || '';
      const val = parts[1] || '';
      const exprNode = firstDirectChildByTag(element, MATH_NS, 'e');
      const expr = exprNode ? renderChildContent(exprNode).trim() : '';
      if (varName && val) {
        return `\\lim_{${varName}\\to ${val}} ${expr}`.trim();
      }
    }
  }
  return renderChildContent(element);
}

function xmlNodeToText(node) {
  if (!node) return '';
  if (node.nodeType === 3) return node.textContent || '';
  if (node.nodeType !== 1) return '';
  const element = node;
  if (element.namespaceURI === MATH_NS && (element.localName === 'oMath' || element.localName === 'oMathPara')) {
    const mathTokens = renderOmmlNodeToLatex(element).replace(/\s+/g, ' ').trim();
    return ` {{EQ:${mathTokens}}} `;
  }
  if (element.namespaceURI === WORD_NS && element.localName === 't') return element.textContent || '';
  if (element.namespaceURI === WORD_NS && element.localName === 'tab') return '\t';
  if (element.namespaceURI === WORD_NS && (element.localName === 'br' || element.localName === 'cr')) return '\n';
  let output = '';
  for (let i = 0; i < element.childNodes.length; i += 1) output += xmlNodeToText(element.childNodes[i]);
  return output;
}

function ooxmlToParserText(ooxml) {
  const xmlDoc = new DOMParser().parseFromString(ooxml, 'application/xml');
  const parserError = xmlDoc.getElementsByTagName('parsererror')[0];
  if (parserError) throw new Error('Failed to parse OOXML');
  const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(WORD_NS, 'p'));
  const lines = [];
  for (const paragraph of paragraphs) {
    const paragraphText = String(xmlNodeToText(paragraph) || '').replace(/\u00A0/g, ' ').replace(/[ \t]+/g, ' ').replace(/\s*\n\s*/g, '\n').trim();
    if (!paragraphText) continue;
    const splitLines = paragraphText.split('\n').map(l => l.trim()).filter(Boolean);
    lines.push(...splitLines);
  }
  return lines.join('\n');
}

function ooxmlToParserTextWithNumbering(ooxml) {
  const xmlDoc = new DOMParser().parseFromString(ooxml, 'application/xml');
  const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(WORD_NS, 'p'));
  const parsed = [];
  for (const p of paragraphs) {
    const numPr = p.getElementsByTagNameNS(WORD_NS, 'numPr')[0] || null;
    let ilvl = null;
    let numId = null;
    if (numPr) {
      const ilvlNode = numPr.getElementsByTagNameNS(WORD_NS, 'ilvl')[0];
      const numIdNode = numPr.getElementsByTagNameNS(WORD_NS, 'numId')[0];
      if (ilvlNode) ilvl = Number(ilvlNode.getAttribute('w:val') || ilvlNode.getAttribute('val') || 0);
      if (numIdNode) numId = String(numIdNode.getAttribute('w:val') || numIdNode.getAttribute('val') || '');
    }
    const text = String(xmlNodeToText(p) || '').replace(/\u00A0/g, ' ').replace(/[ \t]+/g, ' ').replace(/\s*\n\s*/g, '\n').trim();
    parsed.push({ text, numId, ilvl });
  }
  const counters = {};
  const lines = [];
  for (const p of parsed) {
    if (!p.text) continue;
    if (p.numId != null && p.numId !== '') {
      counters[p.numId] = counters[p.numId] || [];
      const lvl = p.ilvl != null ? p.ilvl : 0;
      for (let i = 0; i <= lvl; i++) if (counters[p.numId][i] == null) counters[p.numId][i] = 0;
      counters[p.numId][lvl] += 1;
      for (let i = lvl + 1; i < counters[p.numId].length; i++) counters[p.numId][i] = 0;
      if (lvl === 0) lines.push(`${counters[p.numId][0]}. ${p.text}`);
      else lines.push(`- ${p.text}`);
    } else {
      lines.push(p.text);
    }
  }
  return lines.join('\n');
}

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

  // Try numbering-aware
  let numbered = '';
  try {
    numbered = addin.ooxmlToParserTextWithNumbering(ooxml);
  } catch (e) {
    numbered = '';
  }

  // Try DOM extraction
  let dom = '';
  try {
    dom = addin.ooxmlToParserText(ooxml);
  } catch (e) {
    dom = '';
  }

  // Regex fallback
  let regex = '';
  try {
    regex = addin.ooxmlToParserTextRegexFallback(ooxml);
  } catch (e) {
    regex = '';
  }

  await fs.writeFile(path.join('tests', 'addin_dom_numbered.txt'), numbered, 'utf8');
  await fs.writeFile(path.join('tests', 'addin_dom_dom.txt'), dom, 'utf8');
  await fs.writeFile(path.join('tests', 'addin_dom_regex.txt'), regex, 'utf8');

  console.log('Wrote tests/addin_dom_numbered.txt, addin_dom_dom.txt, addin_dom_regex.txt');
}

main().catch((err) => { console.error(err); process.exit(1); });
