import { generateQtiPackageArtifacts, isXmlWellFormed, parseForValidation } from '../src/qti-generator.js';

const checkButton = document.getElementById('checkButton');
const generateButton = document.getElementById('generateButton');
const statusNode = document.getElementById('status');
const strictModeToggle = document.getElementById('strictModeToggle');
const parserPermissiveToggle = document.getElementById('parserPermissiveToggle');
const checkResultsNode = document.getElementById('checkResults');
const checkResultsSummaryNode = document.getElementById('checkResultsSummary');
const checkResultsMetaNode = document.getElementById('checkResultsMeta');
const checkResultsIssuesNode = document.getElementById('checkResultsIssues');
const checkResultsPreviewNode = document.getElementById('checkResultsPreview');
const copyCheckReportButton = document.getElementById('copyCheckReportButton');
const APP_VERSION = '2026.04.24-final15';
let latestCheckReportText = '';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

function setStatus(message) {
  statusNode.textContent = message;

  const normalizedMessage = (message || '').toLowerCase();
  statusNode.classList.remove('status-success', 'status-error');

  if (normalizedMessage.startsWith('error:')) {
    statusNode.classList.add('status-error');
    return;
  }

  if (normalizedMessage.startsWith('check passed') || normalizedMessage.startsWith('done')) {
    statusNode.classList.add('status-success');
  }
}

function clearCheckResults() {
  if (checkResultsSummaryNode) {
    checkResultsSummaryNode.textContent = '';
  }

  if (checkResultsMetaNode) {
    checkResultsMetaNode.textContent = '';
  }

  if (checkResultsIssuesNode) {
    checkResultsIssuesNode.innerHTML = '';
  }

  if (checkResultsPreviewNode) {
    checkResultsPreviewNode.textContent = '';
  }

  if (copyCheckReportButton) {
    copyCheckReportButton.disabled = true;
  }

  latestCheckReportText = '';

  if (checkResultsNode) {
    checkResultsNode.hidden = true;
    checkResultsNode.classList.remove('results-pass', 'results-error');
  }
}

function getIssueHint(issue) {
  const text = String(issue || '').toLowerCase();

  if (text.includes('must have at least 2 choices')) {
    return 'Ensure each question has at least two choice lines.';
  }

  if (text.includes('must have exactly 1 correct choice')) {
    return 'Mark exactly one choice as correct using [x] or {{ANS}}.';
  }

  if (text.includes('multiple correct choices')) {
    return 'Keep only one correct marker for each question.';
  }

  if (text.includes('no questions were parsed')) {
    return 'Use numbered stems or stems ending with ? or :, with choices on separate lines.';
  }

  if (text.includes('not well-formed')) {
    return 'Re-run in strict XML mode to pinpoint file-level XML issues.';
  }

  return '';
}

function annotateIssuesWithHints(issues) {
  return (issues || []).map((issue) => {
    if (/^Auto-fix applied:/i.test(issue)) {
      return issue;
    }

    const hint = getIssueHint(issue);
    return hint ? `${issue} Hint: ${hint}` : issue;
  });
}

function buildQuestionPreviewText(questions, limit = 3) {
  const items = Array.isArray(questions) ? questions : [];
  if (items.length === 0) {
    return '';
  }

  const shown = items.slice(0, limit);
  const lines = [`Preview: first ${shown.length} of ${items.length} parsed question(s)`];

  for (const question of shown) {
    lines.push(`Q${question.index}: ${question.stem}`);
    for (const choice of question.choices || []) {
      lines.push(`  ${choice.isCorrect ? '[x]' : '[ ]'} ${choice.text}`);
    }
  }

  return lines.join('\n');
}

function buildCheckReportText(result) {
  const lines = [
    `QTI Check Report v${APP_VERSION}`,
    `Generated: ${new Date().toISOString()}`,
    `Outcome: ${result.outcome || 'unknown'}`,
    `Summary: ${result.summary || ''}`,
    `Questions parsed: ${result.questionCount ?? 0}`,
    `Issues found: ${Array.isArray(result.issues) ? result.issues.length : 0}`,
    `Warnings: ${result.warningCount ?? 0}`,
    `Auto-fixes: ${result.autoFixCount ?? 0}`,
    `Parser mode: ${result.parserMode || 'strict'}`,
    `XML check mode: ${result.xmlMode || 'assessment-only'}`,
    `Extraction source: ${result.source || 'unknown'}`,
    `Flattened fallback: ${result.usedFlattenedFallback ? 'yes' : 'no'}`,
    `Equation placeholders: ${result.equationCount ?? 0}`
  ];

  if (Array.isArray(result.issues) && result.issues.length > 0) {
    lines.push('Issues:');
    for (const issue of result.issues) {
      lines.push(`- ${issue}`);
    }
  }

  if (result.previewText) {
    lines.push('');
    lines.push(result.previewText);
  }

  return lines.join('\n');
}

async function copyLatestCheckReport() {
  if (!latestCheckReportText) {
    setStatus('Error: No check report available to copy yet.');
    return;
  }

  try {
    let copied = false;
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(latestCheckReportText);
      copied = true;
    } else {
      const tempTextArea = document.createElement('textarea');
      tempTextArea.value = latestCheckReportText;
      document.body.appendChild(tempTextArea);
      tempTextArea.select();
      copied = Boolean(document.execCommand('copy'));
      document.body.removeChild(tempTextArea);
    }

    if (!copied) {
      throw new Error('Copy command was rejected by host runtime.');
    }

    setStatus('Check report copied to clipboard.');
  } catch {
    setStatus('Error: Unable to copy report to clipboard in this environment.');
  }
}

function parseIssuesFromErrorMessage(message) {
  const text = String(message || '').trim();
  if (!text) {
    return ['Unknown validation error.'];
  }

  const matches = text.match(/Question\s+\d+[^.]*\.?/gi) || [];
  if (matches.length > 0) {
    return [...new Set(matches.map((entry) => entry.trim()))];
  }

  const parts = text
    .split(/[;\n]/)
    .map((part) => part.trim())
    .filter(Boolean);

  return parts.length > 0 ? parts : [text];
}

function renderCheckResults(result) {
  if (!checkResultsNode || !checkResultsSummaryNode || !checkResultsMetaNode || !checkResultsIssuesNode) {
    return;
  }

  const issueCount = Array.isArray(result.issues) ? result.issues.length : 0;
  const warningCount = result.warningCount ?? 0;
  const autoFixCount = result.autoFixCount ?? 0;
  checkResultsNode.classList.remove('results-pass', 'results-error');
  checkResultsNode.classList.add(result.outcome === 'error' ? 'results-error' : 'results-pass');
  checkResultsSummaryNode.textContent = result.summary || 'Check completed.';

  const metaLines = [
    `Outcome: ${result.outcome || 'unknown'}`,
    `Questions parsed: ${result.questionCount ?? 0}`,
    `Issues found: ${issueCount}`,
    `Warnings: ${warningCount}`,
    `Auto-fixes: ${autoFixCount}`,
    `Parser mode: ${result.parserMode || 'strict'}`,
    `XML check mode: ${result.xmlMode || 'assessment-only'}`,
    `Extraction source: ${result.source || 'unknown'}`,
    `Flattened fallback: ${result.usedFlattenedFallback ? 'yes' : 'no'}`,
    `Equation placeholders: ${result.equationCount ?? 0}`
  ];
  checkResultsMetaNode.textContent = metaLines.join(' | ');
  checkResultsMetaNode.classList.toggle('warning', warningCount > 0 || autoFixCount > 0);

  checkResultsIssuesNode.innerHTML = '';
  if (issueCount > 0) {
    for (const issue of result.issues) {
      const item = document.createElement('li');
      item.textContent = issue;
      checkResultsIssuesNode.appendChild(item);
    }
  }

  if (checkResultsPreviewNode) {
    checkResultsPreviewNode.textContent = result.previewText || '';
  }

  latestCheckReportText = buildCheckReportText(result);
  if (copyCheckReportButton) {
    copyCheckReportButton.disabled = !latestCheckReportText;
  }

  checkResultsNode.hidden = false;
}

function toSafeFilename(value) {
  const normalized = (value || 'assessment')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');

  return normalized || 'assessment';
}

function buildExportFilename(title) {
  const base = toSafeFilename(title);
  const now = new Date();
  const yyyy = String(now.getFullYear());
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const min = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');
  const datePart = `${yyyy}-${mm}-${dd}`;
  const timePart = `${hh}-${min}-${ss}`;
  return `${base}-qti-${datePart}-${timePart}.zip`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}

async function getDocumentOoxml() {
  return Word.run(async (context) => {
    const body = context.document.body;
    const bodyOoxml = body.getOoxml();
    await context.sync();
    return bodyOoxml.value;
  });
}

async function getDocumentHtml() {
  return Word.run(async (context) => {
    const body = context.document.body;
    const bodyHtml = body.getHtml();
    await context.sync();
    return bodyHtml.value;
  });
}

function countEquationPlaceholders(text) {
  const matches = text.match(/\{\{EQ(?::[\s\S]*?)?\}\}(?!\})/g);
  return matches ? matches.length : 0;
}

function equationNodeToPlaceholder(equationNode) {
  const mathTokens = renderOmmlNodeToLatex(equationNode)
    .replace(/\s+/g, ' ')
    .trim();

  if (!mathTokens) {
    return ' {{EQ}} ';
  }

  return ` {{EQ:${mathTokens}}} `;
}

function directChildrenByTag(node, namespace, localName) {
  return Array.from(node.childNodes).filter(
    (child) => child.nodeType === Node.ELEMENT_NODE && child.namespaceURI === namespace && child.localName === localName
  );
}

function firstDirectChildByTag(node, namespace, localName) {
  return directChildrenByTag(node, namespace, localName)[0] || null;
}

function renderChildContent(node) {
  let output = '';
  for (const child of node.childNodes) {
    output += renderOmmlNodeToLatex(child);
  }
  return output;
}

function renderRadicalToLatex(radNode) {
  const degreeNode = firstDirectChildByTag(radNode, MATH_NS, 'deg');
  const expressionNode = firstDirectChildByTag(radNode, MATH_NS, 'e');

  const degreeLatex = degreeNode ? renderChildContent(degreeNode).trim() : '';
  const expressionLatex = expressionNode ? renderChildContent(expressionNode).trim() : '';

  if (!expressionLatex) {
    return '';
  }

  if (degreeLatex) {
    return `\\sqrt[${degreeLatex}]{${expressionLatex}}`;
  }

  return `\\sqrt{${expressionLatex}}`;
}

function renderFractionToLatex(fractionNode) {
  const numeratorNode = firstDirectChildByTag(fractionNode, MATH_NS, 'num');
  const denominatorNode = firstDirectChildByTag(fractionNode, MATH_NS, 'den');

  const numeratorLatex = numeratorNode ? renderChildContent(numeratorNode).trim() : '';
  const denominatorLatex = denominatorNode ? renderChildContent(denominatorNode).trim() : '';

  if (!numeratorLatex || !denominatorLatex) {
    return '';
  }

  return `\\frac{${numeratorLatex}}{${denominatorLatex}}`;
}

function renderOmmlNodeToLatex(node) {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent || '';
  }

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return '';
  }

  const element = node;
  if (element.namespaceURI !== MATH_NS) {
    return renderChildContent(element);
  }

  if (element.localName === 'oMath' || element.localName === 'oMathPara') {
    return renderChildContent(element);
  }

  if (element.localName === 'rad') {
    return renderRadicalToLatex(element);
  }

  if (element.localName === 'f') {
    return renderFractionToLatex(element);
  }

  if (element.localName === 'nary') {
    // Handle n-ary operators (integrals, sums). Default to integral when symbol not explicit.
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
    // Handle function-like constructs; specifically lim with limLow/lim
    // Find any descendant <m:lim> element
    const limElems = element.getElementsByTagNameNS(MATH_NS, 'lim');
    if (limElems && limElems.length) {
      const limNode = limElems[0];
      const limText = renderChildContent(limNode).replace(/\s+/g, ' ').trim();
      // parse var and value: e.g. "x→0" or "x->0" or "x to 0"
      const parts = limText.split(/→|->|\\to|\sto|\s+to\s+/).map(p => p.trim()).filter(Boolean);
      const varName = parts[0] || '';
      const val = parts[1] || '';
      const exprNode = firstDirectChildByTag(element, MATH_NS, 'e');
      const expr = exprNode ? renderChildContent(exprNode).trim() : '';
      if (varName && val) {
        return `\\lim_{${varName}\\to ${val}} ${expr}`.trim();
      }
    }

    // Handle named functions like sin, cos, tan, log, ln, exp, etc.
    const fNameNode = firstDirectChildByTag(element, MATH_NS, 'fName') || (element.getElementsByTagNameNS ? element.getElementsByTagNameNS(MATH_NS, 'fName')[0] : null);
    const eNode = firstDirectChildByTag(element, MATH_NS, 'e') || (element.getElementsByTagNameNS ? element.getElementsByTagNameNS(MATH_NS, 'e')[0] : null);
    if (fNameNode) {
      const rawName = renderChildContent(fNameNode).trim();
      const arg = eNode ? renderChildContent(eNode).trim() : '';
      if (rawName) {
        const name = rawName.toLowerCase();
        const fnMap = {
          sin: '\\sin', cos: '\\cos', tan: '\\tan', csc: '\\csc', sec: '\\sec', cot: '\\cot',
          arcsin: '\\arcsin', arccos: '\\arccos', arctan: '\\arctan', log: '\\log', ln: '\\ln', exp: '\\exp'
        };
        const fnLatex = fnMap[name] || `\\operatorname{${rawName}}`;
        return arg ? `${fnLatex}\\left(${arg}\\right)` : fnLatex;
      }
    }
  }

  if (element.localName === 't') {
    return element.textContent || '';
  }

  // Handle simple superscript/subscript constructs like <m:sSup>, <m:sSub>, <m:sup>, <m:sub>
  if (element.localName === 'sSup' || element.localName === 'sSub' || element.localName === 'sup' || element.localName === 'sub') {
    const eNode = firstDirectChildByTag(element, MATH_NS, 'e') || firstDirectChildByTag(element, MATH_NS, 'base') || null;
    const supNode = firstDirectChildByTag(element, MATH_NS, 'sup') || firstDirectChildByTag(element, MATH_NS, 'sup') || null;
    const subNode = firstDirectChildByTag(element, MATH_NS, 'sub') || firstDirectChildByTag(element, MATH_NS, 'sSub') || null;

    const base = eNode ? renderChildContent(eNode).trim() : '';
    const sup = supNode ? renderChildContent(supNode).trim() : '';
    const sub = subNode ? renderChildContent(subNode).trim() : '';

    if (!base) return '';
    let out = base;
    if (sup) out = `${out}^{${sup}}`;
    if (sub) out = `${out}_{${sub}}`;
    return out;
  }

  return renderChildContent(element);
}

function xmlNodeToText(node) {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent || '';
  }

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return '';
  }

  const element = node;
  if (element.namespaceURI === MATH_NS && (element.localName === 'oMath' || element.localName === 'oMathPara')) {
    return equationNodeToPlaceholder(element);
  }

  if (element.namespaceURI === WORD_NS && element.localName === 't') {
    return element.textContent || '';
  }

  if (element.namespaceURI === WORD_NS && element.localName === 'tab') {
    return '\t';
  }

  if (element.namespaceURI === WORD_NS && (element.localName === 'br' || element.localName === 'cr')) {
    return '\n';
  }

  let output = '';
  for (const childNode of element.childNodes) {
    output += xmlNodeToText(childNode);
  }

  return output;
}

function normalizeParagraphText(text) {
  return text
    .replace(/\u00A0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/\s*\n\s*/g, '\n')
    .trim();
}

function ooxmlToParserText(ooxml) {
  const xmlDoc = new DOMParser().parseFromString(ooxml, 'application/xml');
  const parserError = xmlDoc.querySelector('parsererror');
  if (parserError) {
    throw new Error(`Failed to parse document OOXML: ${parserError.textContent?.trim() || 'Unknown XML parser error'}`);
  }

  const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(WORD_NS, 'p'));
  const lines = [];

  for (const paragraph of paragraphs) {
    const paragraphText = normalizeParagraphText(xmlNodeToText(paragraph));
    if (!paragraphText) {
      continue;
    }

    const splitLines = paragraphText
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean);

    lines.push(...splitLines);
  }

  return lines.join('\n');
}

function ooxmlToParserTextWithNumbering(ooxml) {
  const xmlDoc = new DOMParser().parseFromString(ooxml, 'application/xml');
  const parserError = xmlDoc.querySelector('parsererror');
  if (parserError) {
    throw new Error(`Failed to parse document OOXML: ${parserError.textContent?.trim() || 'Unknown XML parser error'}`);
  }

  const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(WORD_NS, 'p'));
  const parsed = [];

  for (const p of paragraphs) {
    const numPr = p.getElementsByTagNameNS(WORD_NS, 'numPr')[0] || null;
    let ilvl = null;
    let numId = null;
    if (numPr) {
      const ilvlNode = numPr.getElementsByTagNameNS(WORD_NS, 'ilvl')[0];
      const numIdNode = numPr.getElementsByTagNameNS(WORD_NS, 'numId')[0];
      if (ilvlNode) {
        ilvl = Number(ilvlNode.getAttribute('w:val') || ilvlNode.getAttribute('val') || 0);
      }
      if (numIdNode) {
        numId = String(numIdNode.getAttribute('w:val') || numIdNode.getAttribute('val') || '');
      }
    }

    const text = normalizeParagraphText(xmlNodeToText(p));
    parsed.push({ text, numId, ilvl });
  }

  const counters = {};
  const lines = [];

  for (const p of parsed) {
    if (!p.text) continue;
    if (p.numId != null && p.numId !== '') {
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

  return lines.join('\n');
}

function decodeXmlEntities(value) {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
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

  // Handle limits like: lim x→0 expression  -> \lim_{x\to 0} expression
  if (/\blim\b/i.test(compact)) {
    const limMatch = compact.match(/lim\s*([A-Za-z0-9]+)(?:\s|\u2192|->|\\to)*([A-Za-z0-9]+)\s*(.*)/i);
    if (limMatch) {
      const varName = limMatch[1];
      const val = limMatch[2];
      const rest = limMatch[3] || '';
      latex = `\\lim_{${varName}\\to ${val}} ${rest}`.trim();
    }
  }

  // Handle simple integrals: int a b expr -> \int_{a}^{b} expr
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
  // Combine token + space + number into exponent: "x 7" or "(a+b) 2" -> "x^{7}" or "(a+b)^{2}"
  latex = latex.replace(/(\S+)\s+(\d+)/g, (m, p1, p2) => {
    if (p1.includes('\\')) return m;
    return `${p1}^{${p2}}`;
  });
  // Ensure integral token is converted to LaTeX '\\int' when not already formatted
  if (/\bint\b/i.test(latex) && !/\\\\?int/.test(latex)) {
    latex = latex.replace(/\bint\b/gi, '\\int');
  }
  if (ommlBlock.includes('<m:rad')) {
    if (tokens.length >= 2) {
      latex = `\\sqrt[${tokens[0]}]{${tokens.slice(1).join(' ')}}`;
    } else {
      latex = `\\sqrt{${tokens[0]}}`;
    }
  } else if (ommlBlock.includes('<m:f')) {
    const numeratorMatch = ommlBlock.match(/<m:num>[\s\S]*?<m:t[^>]*>([\s\S]*?)<\/m:t>[\s\S]*?<\/m:num>/);
    const denominatorMatch = ommlBlock.match(/<m:den>[\s\S]*?<m:t[^>]*>([\s\S]*?)<\/m:t>[\s\S]*?<\/m:den>/);
    const numerator = numeratorMatch ? decodeXmlEntities(numeratorMatch[1]).trim() : '';
    const denominator = denominatorMatch ? decodeXmlEntities(denominatorMatch[1]).trim() : '';
    if (numerator && denominator) {
      latex = `\\frac{${numerator}}{${denominator}}`;
    }
  }

  // Handle function tokens from regex fallback: sin, cos, tan, log, ln, exp, etc.
  const funcMatch = compact.match(/^(sin|cos|tan|csc|sec|cot|arcsin|arccos|arctan|log|ln|exp)\b\s*(.*)$/i);
  if (funcMatch) {
    const fn = funcMatch[1].toLowerCase();
    const rest = funcMatch[2] || '';
    const fnMap = { sin: '\\sin', cos: '\\cos', tan: '\\tan', csc: '\\csc', sec: '\\sec', cot: '\\cot', arcsin: '\\arcsin', arccos: '\\arccos', arctan: '\\arctan', log: '\\log', ln: '\\ln', exp: '\\exp' };
    const fnLatex = fnMap[fn] || `\\operatorname{${fn}}`;
    const arg = rest.trim();
    latex = arg ? `${fnLatex}(${arg})` : fnLatex;
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

function htmlEquationToPlaceholder(element) {
  const dataEquation = element.getAttribute('data-equation-content') || '';
  const title = element.getAttribute('title') || '';
  const alt = element.getAttribute('alt') || '';

  let latex = dataEquation.trim();
  if (!latex && alt.toLowerCase().startsWith('latex:')) {
    latex = alt.slice(6).trim();
  }

  if (!latex) {
    latex = title.trim();
  }

  if (!latex) {
    return ' {{EQ}} ';
  }

  return ` {{EQ:${latex}}} `;
}

function htmlNodeToText(node) {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent || '';
  }

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return '';
  }

  const element = node;
  if (element.tagName === 'BR') {
    return '\n';
  }

  if (element.tagName === 'IMG') {
    const className = element.getAttribute('class') || '';
    const src = element.getAttribute('src') || '';
    const alt = element.getAttribute('alt') || '';
    const title = element.getAttribute('title') || '';
    const looksLikeEquation =
      className.includes('equation_image')
      || element.hasAttribute('data-equation-content')
      || src.includes('equation_images')
      || /^latex\s*:/i.test(alt)
      || title.includes('\\')
      || alt.includes('\\');

    if (looksLikeEquation) {
      return htmlEquationToPlaceholder(element);
    }
  }

  if (element.tagName === 'MATH') {
    const mathText = element.textContent?.replace(/\s+/g, ' ').trim() || '';
    if (mathText) {
      return ` {{EQ:${mathText}}} `;
    }
  }

  let text = '';
  for (const child of element.childNodes) {
    text += htmlNodeToText(child);
  }
  return text;
}

function htmlToParserText(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const lines = [];
  const blocks = doc.querySelectorAll('p, li, div');

  if (blocks.length === 0) {
    const text = normalizeParagraphText(htmlNodeToText(doc.body));
    return text;
  }

  for (const block of blocks) {
    const blockText = normalizeParagraphText(htmlNodeToText(block));
    if (!blockText) {
      continue;
    }

    const splitLines = blockText
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean);

    lines.push(...splitLines);
  }

  return lines.join('\n');
}

async function getDocumentParserText() {
  const candidates = [];

  try {
    const ooxmlRaw = await getDocumentOoxml();
    const normalizedOoxmlRaw = ooxmlRaw.includes('<w:p') ? ooxmlRaw : decodeXmlEntities(ooxmlRaw);

    // Prefer OOXML extraction that preserves numbering (matches CLI behavior)
    try {
      const numberedText = ooxmlToParserTextWithNumbering(normalizedOoxmlRaw);
      candidates.push({ source: 'ooxml-numbering', rawText: numberedText, equationCount: countEquationPlaceholders(numberedText), ooxml: normalizedOoxmlRaw });
    } catch {
    }

    try {
      const domText = ooxmlToParserText(normalizedOoxmlRaw);
      candidates.push({ source: 'ooxml-dom', rawText: domText, equationCount: countEquationPlaceholders(domText), ooxml: normalizedOoxmlRaw });
    } catch {
    }

    try {
      const regexText = ooxmlToParserTextRegexFallback(normalizedOoxmlRaw);
      candidates.push({ source: 'ooxml-regex', rawText: regexText, equationCount: countEquationPlaceholders(regexText), ooxml: normalizedOoxmlRaw });
    } catch {
    }
  } catch {
  }

  try {
    const html = await getDocumentHtml();
    const htmlText = htmlToParserText(html);
    candidates.push({ source: 'html', rawText: htmlText, equationCount: countEquationPlaceholders(htmlText) });
  } catch {
  }

  try {
    const fallbackText = await Word.run(async (context) => {
      const body = context.document.body;
      body.load('text');
      await context.sync();
      return body.text;
    });

    candidates.push({ source: 'text', rawText: fallbackText, equationCount: countEquationPlaceholders(fallbackText) });
  } catch {
  }

  const validCandidates = candidates.filter((candidate) => (candidate.rawText || '').trim().length > 0);
  validCandidates.sort((left, right) => {
    if (right.equationCount !== left.equationCount) {
      return right.equationCount - left.equationCount;
    }

    return right.rawText.length - left.rawText.length;
  });

  const selected = validCandidates[0] || { source: 'text', rawText: '', equationCount: 0 };
  return {
    rawText: selected.rawText,
    source: selected.source,
    ooxml: selected.ooxml || null
  };
}

function setActionButtonsDisabled(isDisabled) {
  checkButton.disabled = isDisabled;
  generateButton.disabled = isDisabled;
}

function validateXmlArtifacts(artifacts, strictModeEnabled) {
  const checks = [
    { label: 'assessment XML', xml: artifacts.assessmentXml }
  ];

  if (strictModeEnabled) {
    checks.push({ label: 'manifest XML', xml: artifacts.manifestXml });
    checks.push({ label: 'assessment_meta XML', xml: artifacts.assessmentMetaXml });
  }

  for (const check of checks) {
    const result = isXmlWellFormed(check.xml);
    if (!result.ok) {
      throw new Error(`${check.label} is not well-formed. ${result.message}`);
    }
  }
}

async function handleGenerate() {
  setActionButtonsDisabled(true);
  clearCheckResults();

  try {
    const { parsed, artifacts, strictModeEnabled, equationCount, source } = await runPreflightChecks({
      includeZipData: true,
      includeDebugInfo: true
    });
    validateXmlArtifacts(artifacts, strictModeEnabled);

    const filename = buildExportFilename(parsed.title);
    const blob = new Blob([artifacts.zipData], { type: 'application/zip' });
    downloadBlob(blob, filename);

    setStatus(`Done (v${APP_VERSION}). Downloaded ${filename} (equations detected: ${equationCount}, source: ${source}).`);
  } catch (error) {
    setStatus(`Error: ${error?.message || 'Unknown error'}`);
  } finally {
    setActionButtonsDisabled(false);
  }
}

async function runPreflightChecks(options = {}) {
  const includeZipData = options.includeZipData !== false;
  const includeDebugInfo = Boolean(options.includeDebugInfo);

  setStatus('Reading document...');
  const { rawText, source, ooxml } = await getDocumentParserText();
  const equationCount = countEquationPlaceholders(rawText);
  const parserPermissiveEnabled = Boolean(parserPermissiveToggle?.checked);
  const parserMode = parserPermissiveEnabled ? 'permissive' : 'strict';
  const parserDiagnostics = {
    warnings: [],
    autoFixes: [],
    usedFlattenedFallback: false
  };

  const strictModeEnabled = Boolean(strictModeToggle?.checked);
  const xmlMode = strictModeEnabled ? 'strict-all-xml' : 'assessment-only';
  const preflightContext = {
    equationCount,
    source,
    parserMode,
    xmlMode,
    parserDiagnostics
  };

  setStatus(
    source === 'text'
      ? 'Using text fallback extraction (equations may be reduced).'
      : `Detected ${equationCount} equation placeholder(s) via ${source}.`
  );

  setStatus(`Validating questions (${parserMode} parser)...`);
  let parsed;
  try {
    parsed = parseForValidation(rawText, {
      permissive: parserPermissiveEnabled,
      diagnostics: parserDiagnostics
    });
  } catch (error) {
    error.preflightContext = preflightContext;
    throw error;
  }

  setStatus(`Generating package for ${parsed.questions.length} question(s)...`);
  let artifacts;
  try {
    const debugRawText = (source && source.startsWith('ooxml')) ? (ooxml || rawText) : rawText;
    artifacts = await generateQtiPackageArtifacts(rawText, {
      parsed,
      skipZipData: !includeZipData,
      debugInfo: includeDebugInfo
        ? {
          enabled: true,
          source,
          equationCount,
          rawText: debugRawText
        }
        : undefined
    });
  } catch (error) {
    error.preflightContext = preflightContext;
    throw error;
  }

  setStatus(
    strictModeEnabled
      ? 'Strict mode enabled: checking all generated XML files...'
      : 'Checking generated assessment XML...'
  );

  return {
    parsed,
    artifacts,
    strictModeEnabled,
    equationCount,
    source,
    parserMode,
    xmlMode,
    parserDiagnostics
  };
}

async function handleCheckQuestions() {
  setActionButtonsDisabled(true);
  clearCheckResults();

  try {
    const {
      parsed,
      artifacts,
      strictModeEnabled,
      equationCount,
      source,
      parserMode,
      xmlMode,
      parserDiagnostics
    } = await runPreflightChecks({ includeZipData: false, includeDebugInfo: false });
    validateXmlArtifacts(artifacts, strictModeEnabled);

    renderCheckResults({
      outcome: 'pass',
      summary: parserDiagnostics.autoFixes.length > 0
        ? 'Check passed with auto-fixes. Review adjustments below before generating ZIP.'
        : 'Check passed. Questions are ready for QTI ZIP generation.',
      issues: annotateIssuesWithHints([
        ...parserDiagnostics.warnings.map((item) => `Warning: ${item}`),
        ...parserDiagnostics.autoFixes.map((item) => `Auto-fix applied: ${item}`)
      ]),
      questionCount: parsed.questions.length,
      warningCount: parserDiagnostics.warnings.length,
      autoFixCount: parserDiagnostics.autoFixes.length,
      parserMode,
      xmlMode,
      source,
      equationCount,
      usedFlattenedFallback: Boolean(parserDiagnostics.usedFlattenedFallback),
      previewText: buildQuestionPreviewText(parsed.questions)
    });

    setStatus(
      `Check passed (v${APP_VERSION}). ${parsed.questions.length} question(s) are ready for QTI ZIP generation (equations detected: ${equationCount}, source: ${source}).`
    );
  } catch (error) {
    const errorMessage = error?.message || 'Unknown error';
    const preflightContext = error?.preflightContext || {};
    const parserDiagnostics = preflightContext.parserDiagnostics || { warnings: [], autoFixes: [], usedFlattenedFallback: false };
    const issues = annotateIssuesWithHints([
      ...parseIssuesFromErrorMessage(errorMessage),
      ...parserDiagnostics.warnings.map((item) => `Warning: ${item}`),
      ...parserDiagnostics.autoFixes.map((item) => `Auto-fix applied: ${item}`)
    ]);

    renderCheckResults({
      outcome: 'error',
      summary: 'Check failed. Review the detected issues below.',
      issues,
      questionCount: 0,
      warningCount: parserDiagnostics.warnings.length,
      autoFixCount: parserDiagnostics.autoFixes.length,
      parserMode: preflightContext.parserMode || 'strict',
      xmlMode: preflightContext.xmlMode || (Boolean(strictModeToggle?.checked) ? 'strict-all-xml' : 'assessment-only'),
      source: preflightContext.source || 'unknown',
      equationCount: preflightContext.equationCount ?? 0,
      usedFlattenedFallback: Boolean(parserDiagnostics.usedFlattenedFallback),
      previewText: ''
    });

    setStatus(`Error: ${errorMessage}`);
  } finally {
    setActionButtonsDisabled(false);
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus('This add-in only runs in Microsoft Word.');
    return;
  }

  setActionButtonsDisabled(false);
  setStatus(`Ready (v${APP_VERSION}). Click “Check Questions” or “Generate QTI ZIP”.`);
  checkButton.addEventListener('click', handleCheckQuestions);
  generateButton.addEventListener('click', handleGenerate);
  copyCheckReportButton?.addEventListener('click', copyLatestCheckReport);
});
