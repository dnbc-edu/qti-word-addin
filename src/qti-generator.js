import JSZip from 'jszip';

function getRandomToken() {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID().replaceAll('-', '');
  }

  return `${Date.now().toString(16)}${Math.random().toString(16).slice(2)}`;
}

async function shortStableHash(value) {
  const input = String(value);

  if (globalThis.crypto?.subtle && globalThis.TextEncoder) {
    const encoded = new TextEncoder().encode(input);
    const digest = await globalThis.crypto.subtle.digest('SHA-256', encoded);
    const bytes = Array.from(new Uint8Array(digest));
    return bytes.map((byte) => byte.toString(16).padStart(2, '0')).join('').slice(0, 32);
  }

  let hash = 0;
  for (let index = 0; index < input.length; index += 1) {
    hash = (hash << 5) - hash + input.charCodeAt(index);
    hash |= 0;
  }

  return Math.abs(hash).toString(16).padStart(32, '0').slice(0, 32);
}

function escapeXml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function toHtmlParagraph(text) {
  return `&lt;p&gt;${renderTextWithEquationPlaceholders(text)}&lt;/p&gt;`;
}

function normalizeLatexForCanvas(value) {
  let latex = String(value || '').trim();

  if ((latex.startsWith('\\(') && latex.endsWith('\\)')) || (latex.startsWith('\\[') && latex.endsWith('\\]'))) {
    latex = latex.slice(2, -2).trim();
  }

  latex = latex.replace(/\\\\(?=[A-Za-z()[\]{}])/g, '\\');
  return latex || 'equation';
}

function renderCanvasEquationImage(latex) {
  const normalizedLatex = normalizeLatexForCanvas(latex);
  const encodedLatex = encodeURIComponent(encodeURIComponent(normalizedLatex));
  const src = `/equation_images/${encodedLatex}?scale=1`;

  return `&lt;img class=&quot;equation_image&quot; title=&quot;${escapeXml(normalizedLatex)}&quot; src=&quot;${escapeXml(src)}&quot; alt=&quot;LaTeX: ${escapeXml(normalizedLatex)}&quot; data-equation-content=&quot;${escapeXml(normalizedLatex)}&quot; data-ignore-a11y-check=&quot;&quot;&gt;`;
}

function normalizeLatexDelimitersToEquationPlaceholders(text) {
  const input = String(text || '');

  const replaceLatexDelimited = (source, openDelimiter, closeDelimiter) => {
    const openPattern = new RegExp(`${openDelimiter}\\s*`, 'g');
    let result = '';
    let index = 0;

    while (index < source.length) {
      openPattern.lastIndex = index;
      const openMatch = openPattern.exec(source);
      if (!openMatch) {
        result += source.slice(index);
        break;
      }

      const start = openMatch.index;
      const contentStart = openPattern.lastIndex;
      result += source.slice(index, start);

      const closeRegex = new RegExp(`\\s*${closeDelimiter}`, 'g');
      closeRegex.lastIndex = contentStart;
      const closeMatch = closeRegex.exec(source);

      if (!closeMatch) {
        result += source.slice(start);
        break;
      }

      const content = source.slice(contentStart, closeMatch.index).trim();
      result += content ? `{{EQ:${content}}}` : '{{EQ}}';
      index = closeRegex.lastIndex;
    }

    return result;
  };

  let normalized = input;
  normalized = replaceLatexDelimited(normalized, String.raw`\\+\(`, String.raw`\\+\)`);
  normalized = replaceLatexDelimited(normalized, String.raw`\\+\[`, String.raw`\\+\]`);
  return normalized;
}

function renderTextWithEquationPlaceholders(text) {
  const normalizedText = normalizeLatexDelimitersToEquationPlaceholders(text);
  const placeholderRegex = /\{\{EQ(?::([\s\S]*?))?\}\}(?!\})/g;
  let result = '';
  let lastIndex = 0;

  for (const match of normalizedText.matchAll(placeholderRegex)) {
    const matchIndex = match.index ?? 0;
    const before = normalizedText.slice(lastIndex, matchIndex);
    result += escapeXml(before);

    const equationRaw = (match[1] || '').trim() || 'equation';
    result += renderCanvasEquationImage(equationRaw);

    lastIndex = matchIndex + match[0].length;
  }

  result += escapeXml(normalizedText.slice(lastIndex));
  return result;
}

function extractAnswerMarker(text) {
  const markerRegex = /\s*\{\{ans\}\}\s*$/i;
  const hasAnswerMarker = markerRegex.test(text);
  const cleanedText = text.replace(markerRegex, '').trim();

  return {
    cleanedText,
    hasAnswerMarker
  };
}

function parseQuestionLine(line) {
  const match = line.match(/^(\d+)(?:[.)])?(?:\s+|\t+)?(.+)$/);
  if (!match) {
    const fallbackParsedStem = extractPointsFromStem(line);
    // Avoid treating list/choice lines (e.g., "- Choice?" or "• Choice?") as question stems
    if (/^[-•*\u2013\u2014\u2212]\s*/.test(fallbackParsedStem.stem)) {
      return null;
    }

    // Accept unnumbered stems that end with '?' and prompt-style stems that
    // end with ':' (common in imported DOCX content).
    if (!/[?:]\s*$/.test(fallbackParsedStem.stem)) {
      return null;
    }

    return {
      index: null,
      stem: fallbackParsedStem.stem,
      pointsPossible: fallbackParsedStem.pointsPossible
    };
  }

  const parsedStem = extractPointsFromStem(match[2]);

  return {
    index: Number(match[1]),
    stem: parsedStem.stem,
    pointsPossible: parsedStem.pointsPossible
  };
}

function extractPointsFromStem(stemText) {
  const pointsMatch = stemText.match(/\s*\(\s*points\s*:\s*([0-9]+(?:\.[0-9]+)?)\s*\)\s*$/i);
  if (!pointsMatch) {
    return {
      stem: stemText.trim(),
      pointsPossible: null
    };
  }

  const pointsValue = Number(pointsMatch[1]);
  return {
    stem: stemText.replace(pointsMatch[0], '').trim(),
    pointsPossible: Number.isFinite(pointsValue) && pointsValue > 0 ? pointsValue : null
  };
}

function parseLetteredChoiceLine(line) {
  // Handle cases where the choice text might be immediately after the letter/closing paren without a space,
  // as well as multiple choices on the same line (if we've split by space elsewhere).
  // But here we rely on the line-by-line input.
  const match = line.match(/^([a-z])(?:[.)])(?:\s+|\t+)?(?:\[(x| )\](?:\s+|\t+)?)?(.*)$/i);
  if (!match) {
    return null;
  }

  return {
    letter: match[1],
    checkboxMark: match[2] || '',
    text: match[3]
  };
}

function parseNumberedChoiceLine(line) {
  const match = line.match(/^(\d+)(?:[.)])(?:\s+|\t+)?(?:\[(x| )\](?:\s+|\t+)?)?(.+)$/i);
  if (!match) {
    return null;
  }

  return {
    number: Number(match[1]),
    checkboxMark: match[2] || '',
    text: match[3]
  };
}

function looksLikeQuestionStem(text) {
  const normalized = String(text || '').trim();
  return /[?:]\s*$/.test(normalized) || /\(\s*points\s*:/i.test(normalized);
}

function shouldTreatAsNumberedChoice(numberedChoice, currentQuestion) {
  if (!currentQuestion) {
    return false;
  }

  const expectedChoiceNumber = currentQuestion.choices.length + 1;
  if (numberedChoice.number !== expectedChoiceNumber) {
    return false;
  }

  if (looksLikeQuestionStem(numberedChoice.text)) {
    return false;
  }

  return true;
}

function parseGenericChoiceLine(line) {
  const match = line.match(/^(?:[-•*\u2013\u2014\u2212]\s*)?(?:\[(x| )\](?:\s+|\t+))?(.+)$/i);
  if (!match) {
    return null;
  }

  return {
    checkboxMark: match[1] || '',
    text: match[2]
  };
}

function resolveMarkedBy(checkboxMarked, markerMarked) {
  if (checkboxMarked && markerMarked) {
    return 'both';
  }

  if (checkboxMarked) {
    return 'checkbox';
  }

  if (markerMarked) {
    return 'ans-marker';
  }

  return 'none';
}

function isBulletPrefixedChoiceLine(line) {
  return /^[-•*\u2013\u2014\u2212]\s*/.test(String(line || '').trim());
}

function parseQuestionBank(inputText, options = {}) {
  const diagnostics = options.diagnostics || {
    warnings: [],
    autoFixes: [],
    usedFlattenedFallback: false
  };

  const normalizedInput = inputText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = normalizedInput.split(/\n/);
  // Simple parse mode: treat any line ending with '?' as a stem and the
  // subsequent non-empty lines until the next stem as choices. This is a
  // tolerant fallback useful for flattened or paragraph-based extractions.
  if (options.simpleParse) {
    return parseQuestionBankSimple(lines, { ...options, diagnostics });
  }

  const looksFlattened = /[a-z]\)\S/i.test(normalizedInput);
  const shouldTryFallbackFromError = (message) => {
    if (/No questions were parsed/i.test(message)) {
      return true;
    }

    // Only attempt flattened fallback on validation errors when input strongly
    // resembles flattened text. This prevents good structured parses from
    // being replaced by an inferior fallback parse.
    if ((/must have at least 2 choices/i.test(message) || /must have exactly 1 correct choice/i.test(message)) && looksFlattened) {
      return true;
    }

    return false;
  };

  const scoreParseQuality = (parsed) => {
    const questions = parsed?.questions || [];
    const questionCount = questions.length;
    const totalChoices = questions.reduce((sum, q) => sum + (q.choices?.length || 0), 0);
    const lowChoiceQuestions = questions.filter((q) => (q.choices?.length || 0) < 2).length;

    return {
      questionCount,
      totalChoices,
      lowChoiceQuestions
    };
  };

  const isFallbackBetter = (primaryParsedPermissive, fallbackParsed) => {
    const primaryScore = scoreParseQuality(primaryParsedPermissive);
    const fallbackScore = scoreParseQuality(fallbackParsed);

    if (fallbackScore.questionCount !== primaryScore.questionCount) {
      return fallbackScore.questionCount > primaryScore.questionCount;
    }

    if (fallbackScore.lowChoiceQuestions !== primaryScore.lowChoiceQuestions) {
      return fallbackScore.lowChoiceQuestions < primaryScore.lowChoiceQuestions;
    }

    return fallbackScore.totalChoices > primaryScore.totalChoices;
  };

  try {
    return parseQuestionBankFromLines(lines, { ...options, diagnostics });
  } catch (error) {
    const message = String(error?.message || '');
    if (!shouldTryFallbackFromError(message)) {
      throw error;
    }

    const fallbackLines = splitFlattenedWordInput(normalizedInput);
    diagnostics.usedFlattenedFallback = true;
    const fallbackParsed = parseQuestionBankFromLines(fallbackLines, { ...options, diagnostics });

    // If primary strict parsing failed but primary permissive parsing still has
    // better structure than fallback, keep strict behavior and surface the
    // original error instead of silently accepting a lower-quality fallback.
    try {
      const primaryParsedPermissive = parseQuestionBankFromLines(lines, {
        ...options,
        permissive: true,
        diagnostics: { warnings: [], autoFixes: [], usedFlattenedFallback: false }
      });
      if (!isFallbackBetter(primaryParsedPermissive, fallbackParsed)) {
        throw error;
      }
    } catch (primaryPermissiveError) {
      // If permissive primary parsing also fails, fallbackParsed remains the
      // best available result.
    }

    return fallbackParsed;
  }
}

function splitFlattenedWordInput(inputText) {
  const flattened = inputText
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n+/g, ' ')
    .trim();

  if (!flattened) {
    return [];
  }

  let withBreaks = flattened;
  withBreaks = withBreaks.replace(/\s+(?=Title:\s*)/gi, '\n');
  withBreaks = withBreaks.replace(/\s+(?=Points:\s*[0-9])/gi, '\n');
  withBreaks = withBreaks.replace(/\s+(?=\d+[.)]\s*)/g, '\n');
  withBreaks = withBreaks.replace(/\s+(?=[a-z][.)]\s*)/gi, '\n');
  withBreaks = withBreaks.replace(/\s+(?=-\s*\[(?:x| )\]\s*)/gi, '\n');

  // Also split marker runs that were glued directly to previous tokens.
  withBreaks = withBreaks.replace(/([^\n])(?=\d+[.)]\s*)/g, '$1\n');
  withBreaks = withBreaks.replace(/([^\n])(?=[a-z][.)]\s*)/gi, '$1\n');
    // Handle inline lettered choices glued without whitespace: "a)Choiceb)Choice"
    withBreaks = withBreaks.replace(/([^\n])(?=[a-z]\))/gi, '$1\n');
    // Also split when a choice label immediately follows choice text (e.g. "Profilingb)" -> insert break before "b)")
    withBreaks = withBreaks.replace(/([a-z0-9])([a-z])\)/gi, '$1\n$2)');

  return withBreaks
    .split(/\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function parseQuestionBankFromLines(lines, options = {}) {
  let title = 'Word QTI Assessment';
  let documentDefaultPoints = 1;
  const questions = [];
  const diagnostics = options.diagnostics || {
    warnings: [],
    autoFixes: [],
    usedFlattenedFallback: false
  };

  let currentQuestion = null;

  for (const rawLine of lines) {
    const normalizedLine = rawLine.replace(/\u00A0/g, ' ');
    const line = normalizedLine.trim();
    if (!line) {
      continue;
    }

    const titleMatch = line.match(/^Title:\s*(.+)$/i);
    if (titleMatch) {
      title = titleMatch[1].trim();
      continue;
    }

    const globalPointsMatch = line.match(/^Points:\s*([0-9]+(?:\.[0-9]+)?)\s*$/i);
    if (globalPointsMatch && !currentQuestion && questions.length === 0) {
      const parsedPoints = Number(globalPointsMatch[1]);
      if (Number.isFinite(parsedPoints) && parsedPoints > 0) {
        documentDefaultPoints = parsedPoints;
      }
      continue;
    }

    const numberedChoice = parseNumberedChoiceLine(line);
    if (numberedChoice && shouldTreatAsNumberedChoice(numberedChoice, currentQuestion)) {
      const parsedChoiceText = extractAnswerMarker(numberedChoice.text);
      const checkboxMarked = numberedChoice.checkboxMark.toLowerCase() === 'x';
      const markerMarked = parsedChoiceText.hasAnswerMarker;
      const markedBy = resolveMarkedBy(checkboxMarked, markerMarked);

      currentQuestion.choices.push({
        text: parsedChoiceText.cleanedText,
        isCorrect: checkboxMarked || markerMarked,
        markedBy
      });
      continue;
    }

    const question = parseQuestionLine(line);
    if (question) {
      // If the parser returned a fallback question (no numeric index) and
      // there is an open current question with no choices yet, treat this
      // as a continuation of the current question's stem rather than
      // starting a new question. This handles cases where Word split a
      // multi-part stem containing question marks into separate lines.
      if (question.index == null && currentQuestion && currentQuestion.choices.length === 0) {
        currentQuestion.stem = `${currentQuestion.stem} ${question.stem}`.trim();
        // If the fallback provided points, prefer it only if current has none
        if (question.pointsPossible != null && (currentQuestion.pointsPossible == null)) {
          currentQuestion.pointsPossible = question.pointsPossible;
        }
        continue;
      }

      if (currentQuestion) {
        questions.push(currentQuestion);
      }

      currentQuestion = {
        index: question.index,
        stem: question.stem,
        pointsPossible: question.pointsPossible,
        choices: []
      };
      continue;
    }

    const checkboxChoiceMatch = line.match(/^-?\s*\[(x| )\]\s+(.+)$/i);
    if (checkboxChoiceMatch && currentQuestion) {
      const parsedChoiceText = extractAnswerMarker(checkboxChoiceMatch[2]);
      const checkboxMarked = checkboxChoiceMatch[1].toLowerCase() === 'x';
      const markerMarked = parsedChoiceText.hasAnswerMarker;

      const markedBy = resolveMarkedBy(checkboxMarked, markerMarked);

      currentQuestion.choices.push({
        text: parsedChoiceText.cleanedText,
        isCorrect: checkboxMarked || markerMarked,
        markedBy
      });
      continue;
    }

    const letteredChoice = parseLetteredChoiceLine(line);
    if (letteredChoice && currentQuestion) {
      const parsedChoiceText = extractAnswerMarker(letteredChoice.text);
      const checkboxMarked = letteredChoice.checkboxMark.toLowerCase() === 'x';
      const markerMarked = parsedChoiceText.hasAnswerMarker;

      const markedBy = resolveMarkedBy(checkboxMarked, markerMarked);

      currentQuestion.choices.push({
        text: parsedChoiceText.cleanedText,
        isCorrect: checkboxMarked || markerMarked,
        markedBy
      });
      continue;
    }

    const genericChoice = parseGenericChoiceLine(line);
    if (genericChoice && currentQuestion) {
      if (currentQuestion.choices.length > 0 && !isBulletPrefixedChoiceLine(line)) {
        const lastChoice = currentQuestion.choices[currentQuestion.choices.length - 1];
        lastChoice.text += ' ' + genericChoice.text;
        const parsedChoiceText = extractAnswerMarker(lastChoice.text);
        lastChoice.text = parsedChoiceText.cleanedText;
        if (parsedChoiceText.hasAnswerMarker) {
          lastChoice.isCorrect = true;
          lastChoice.markedBy = resolveMarkedBy(false, true);
        }
        continue;
      }
      const parsedChoiceText = extractAnswerMarker(genericChoice.text);
      const checkboxMarked = genericChoice.checkboxMark.toLowerCase() === 'x';
      const markerMarked = parsedChoiceText.hasAnswerMarker;
      const markedBy = resolveMarkedBy(checkboxMarked, markerMarked);

      currentQuestion.choices.push({
        text: parsedChoiceText.cleanedText,
        isCorrect: checkboxMarked || markerMarked,
        markedBy
      });
    }
  }

  if (currentQuestion) {
    questions.push(currentQuestion);
  }

  questions.forEach((question, index) => {
    if (!question.index) {
      question.index = index + 1;
    }

    if (question.pointsPossible == null) {
      const parsedStem = extractPointsFromStem(question.stem);
      question.stem = parsedStem.stem;
      question.pointsPossible = parsedStem.pointsPossible ?? documentDefaultPoints;
    }
  });
  for (const question of questions) {
    if (options.debugParse) {
      console.log(`Checking question ${question.index}: ${question.choices.length} choices`);
    }
    if (question.choices.length < 2) {
      throw new Error(`Question ${question.index} must have at least 2 choices.`);
    }

    let markedChoices = question.choices.filter((choice) => choice.isCorrect);
    let correctCount = markedChoices.length;

    if (correctCount > 1) {
      const hasCheckboxMark = markedChoices.some((choice) => choice.markedBy === 'checkbox' || choice.markedBy === 'both');
      const hasAnsMarker = markedChoices.some((choice) => choice.markedBy === 'ans-marker' || choice.markedBy === 'both');

      if (!options.permissive) {
        if (hasCheckboxMark && hasAnsMarker) {
          throw new Error(
            `Question ${question.index} has multiple correct choices marked across [x] and {{ANS}} styles. Keep only one correct answer.`
          );
        }
        throw new Error(`Question ${question.index} has multiple correct choices. Keep only one correct answer.`);
      }

      if (options.debugParse) {
        try {
          console.error('DEBUG multiple correct choices for question', question.index, JSON.stringify(markedChoices.map(c => ({ text: c.text, isCorrect: c.isCorrect, markedBy: c.markedBy })), null, 2));
        } catch (e) {
          // ignore debug-print failures
        }
      }

      // AUTO-FIX (permissive only): keep first marked choice and clear the rest.
      const firstMarked = markedChoices[0];
      question.choices.forEach((c) => { c.isCorrect = false; });
      firstMarked.isCorrect = true;
      const autoFixMessage = `Question ${question.index}: had ${correctCount} marked choices; keeping first marked answer.`;
      diagnostics.autoFixes.push(autoFixMessage);
      console.warn(`AUTO-FIX: (permissive) ${autoFixMessage} "${firstMarked.text}"`);

      // Recompute markedChoices/correctCount for downstream validation
      markedChoices = question.choices.filter((choice) => choice.isCorrect);
      correctCount = markedChoices.length;
    }

    if (correctCount !== 1) {
      if (options.permissive) {
        if (correctCount === 0 && question.choices.length > 0) {
          question.choices[0].isCorrect = true;
          const autoFixMessage = `Question ${question.index}: had no marked choices; defaulting to first choice.`;
          diagnostics.autoFixes.push(autoFixMessage);
          console.warn(`AUTO-FIX: (permissive) ${autoFixMessage} "${question.choices[0].text}"`);
        } else if (correctCount > 1) {
          const firstMarked = markedChoices[0];
          question.choices.forEach((c) => { c.isCorrect = false; });
          firstMarked.isCorrect = true;
          const autoFixMessage = `Question ${question.index}: had ${correctCount} marked choices; keeping first marked answer.`;
          diagnostics.autoFixes.push(autoFixMessage);
          console.warn(`AUTO-FIX: (permissive) ${autoFixMessage} "${firstMarked.text}"`);
        }
        // recompute
        markedChoices = question.choices.filter((choice) => choice.isCorrect);
        correctCount = markedChoices.length;
        if (correctCount !== 1) {
          throw new Error(`Question ${question.index} must have exactly 1 correct choice. Use [x] or {{ANS}} once.`);
        }
      } else {
        throw new Error(`Question ${question.index} must have exactly 1 correct choice. Use [x] or {{ANS}} once.`);
      }
    }

    question.choices = question.choices.map(({ text, isCorrect }) => ({ text, isCorrect }));
  }

  if (!questions.length) {
    throw new Error('No questions were parsed. Check the input format.');
  }

  return { title, questions, diagnostics };
}

function parseQuestionBankSimple(lines, options = {}) {
  const questions = [];
  let current = null;
  let qIndex = 0;
  const diagnostics = options.diagnostics || {
    warnings: [],
    autoFixes: [],
    usedFlattenedFallback: false
  };

  const isStemLine = (l) => String(l || '').trim().endsWith('?');

  for (const raw of lines) {
    const line = String(raw || '').replace(/\u00A0/g, ' ').trim();
    if (!line) continue;

    if (isStemLine(line)) {
      if (current) questions.push(current);
      qIndex += 1;
      current = { index: qIndex, stem: line, pointsPossible: null, choices: [] };
      continue;
    }

    // treat as a choice if we have a current stem
    if (current) {
      const parsed = extractAnswerMarker(line);
      const checkboxMatch = line.match(/^\s*\[(x| )\]\s*/i);
      const checkboxMarked = Boolean(checkboxMatch && checkboxMatch[1].toLowerCase() === 'x');
      current.choices.push({ text: parsed.cleanedText, isCorrect: checkboxMarked || parsed.hasAnswerMarker, markedBy: checkboxMarked ? 'checkbox' : (parsed.hasAnswerMarker ? 'ans-marker' : 'none') });
      continue;
    }

    // otherwise ignore
  }

  if (current) questions.push(current);

  if (!questions.length) throw new Error('No questions were parsed. Check the input format.');

  return { title: 'Word QTI Assessment', questions, diagnostics };
}

function createAssessmentXml(assessmentId, title, questions) {
  const itemsXml = questions
    .map((question) => {
      const questionToken = getRandomToken();
      const itemId = `question_${question.index}_${questionToken}`;
      const choiceIds = question.choices.map((choice, choiceIndex) => {
        const choiceToken = getRandomToken();
        return {
          id: `choice_${question.index}_${choiceIndex + 1}_${choiceToken}`,
          text: choice.text,
          isCorrect: choice.isCorrect
        };
      });

      const originalAnswerIds = choiceIds.map((choice) => choice.id).join(',');
      const correctChoice = choiceIds.find((choice) => choice.isCorrect);

      const responseLabels = choiceIds
        .map(
          (choice) => `
              <response_label ident="${choice.id}">
                <material>
                  <mattext texttype="text/html">${toHtmlParagraph(choice.text)}</mattext>
                </material>
              </response_label>`
        )
        .join('');

      return `
      <item ident="${itemId}" title="Question ${question.index}">
        <itemmetadata>
          <qtimetadata>
            <qtimetadatafield>
              <fieldlabel>question_type</fieldlabel>
              <fieldentry>multiple_choice_question</fieldentry>
            </qtimetadatafield>
            <qtimetadatafield>
              <fieldlabel>points_possible</fieldlabel>
              <fieldentry>${(Number(question.pointsPossible ?? 1)).toFixed(1)}</fieldentry>
            </qtimetadatafield>
            <qtimetadatafield>
              <fieldlabel>original_answer_ids</fieldlabel>
              <fieldentry>${originalAnswerIds}</fieldentry>
            </qtimetadatafield>
            <qtimetadatafield>
              <fieldlabel>assessment_question_identifierref</fieldlabel>
              <fieldentry>question_ref_${itemId}</fieldentry>
            </qtimetadatafield>
          </qtimetadata>
        </itemmetadata>
        <presentation>
          <material>
            <mattext texttype="text/html">${toHtmlParagraph(question.stem)}</mattext>
          </material>
          <response_lid ident="response1" rcardinality="Single">
            <render_choice>${responseLabels}
            </render_choice>
          </response_lid>
        </presentation>
        <resprocessing>
          <outcomes>
            <decvar maxvalue="100" minvalue="0" varname="SCORE" vartype="Decimal" />
          </outcomes>
          <respcondition continue="No">
            <conditionvar>
              <varequal respident="response1">${correctChoice.id}</varequal>
            </conditionvar>
            <setvar action="Set" varname="SCORE">100</setvar>
          </respcondition>
        </resprocessing>
      </item>`;
    })
    .join('');

  return `<?xml version='1.0' encoding='UTF-8'?>
<questestinterop xmlns="http://www.imsglobal.org/xsd/ims_qtiasiv1p2" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.imsglobal.org/xsd/ims_qtiasiv1p2 http://www.imsglobal.org/xsd/ims_qtiasiv1p2p1.xsd">
  <assessment ident="${assessmentId}" title="${escapeXml(title)}">
    <qtimetadata>
      <qtimetadatafield>
        <fieldlabel>cc_maxattempts</fieldlabel>
        <fieldentry>1</fieldentry>
      </qtimetadatafield>
    </qtimetadata>
    <section ident="root_section">${itemsXml}
    </section>
  </assessment>
</questestinterop>
`;
}

function createManifestXml(assessmentId, folderName, dateIso) {
  return `<?xml version='1.0' encoding='UTF-8'?>
<manifest identifier="manifest_${assessmentId}" xmlns="http://www.imsglobal.org/xsd/imsccv1p1/imscp_v1p1" xmlns:lom="http://ltsc.ieee.org/xsd/imsccv1p1/LOM/resource" xmlns:imsmd="http://www.imsglobal.org/xsd/imsmd_v1p2" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.imsglobal.org/xsd/imsccv1p1/imscp_v1p1 http://www.imsglobal.org/xsd/imscp_v1p1.xsd http://ltsc.ieee.org/xsd/imsccv1p1/LOM/resource http://www.imsglobal.org/profile/cc/ccv1p1/LOM/ccv1p1_lomresource_v1p0.xsd http://www.imsglobal.org/xsd/imsmd_v1p2 http://www.imsglobal.org/xsd/imsmd_v1p2p2.xsd">
  <metadata>
    <schema>IMS Content</schema>
    <schemaversion>1.1.3</schemaversion>
    <imsmd:lom>
      <imsmd:general>
        <imsmd:title>
          <imsmd:string>QTI assessment generated by Word add-in MVP</imsmd:string>
        </imsmd:title>
      </imsmd:general>
      <imsmd:lifeCycle>
        <imsmd:contribute>
          <imsmd:date>
            <imsmd:dateTime>${dateIso}</imsmd:dateTime>
          </imsmd:date>
        </imsmd:contribute>
      </imsmd:lifeCycle>
      <imsmd:rights>
        <imsmd:copyrightAndOtherRestrictions>
          <imsmd:value>yes</imsmd:value>
        </imsmd:copyrightAndOtherRestrictions>
        <imsmd:description>
          <imsmd:string>Private (Copyrighted)</imsmd:string>
        </imsmd:description>
      </imsmd:rights>
    </imsmd:lom>
  </metadata>
  <organizations />
  <resources>
    <resource identifier="${assessmentId}" type="imsqti_xmlv1p2">
      <file href="${folderName}/${assessmentId}.xml" />
      <dependency identifierref="${assessmentId}_dependency" />
    </resource>
    <resource identifier="${assessmentId}_dependency" type="associatedcontent/imscc_xmlv1p1/learning-application-resource" href="${folderName}/assessment_meta.xml">
      <file href="${folderName}/assessment_meta.xml" />
    </resource>
  </resources>
</manifest>
`;
}

function createAssessmentMetaXml(title, dateIso) {
  return `<?xml version='1.0' encoding='UTF-8'?>
<quiz identifier="i${getRandomToken()}">
  <title>${escapeXml(title)}</title>
  <description></description>
  <shuffle_answers>false</shuffle_answers>
  <scoring_policy>keep_highest</scoring_policy>
  <hide_results></hide_results>
  <quiz_type>assignment</quiz_type>
  <points_possible></points_possible>
  <require_lockdown_browser>false</require_lockdown_browser>
  <show_correct_answers>true</show_correct_answers>
  <anonymous_submissions>false</anonymous_submissions>
  <could_be_locked>false</could_be_locked>
  <allowed_attempts>1</allowed_attempts>
  <one_question_at_a_time>false</one_question_at_a_time>
  <cant_go_back>false</cant_go_back>
  <access_code></access_code>
  <ip_filter></ip_filter>
  <due_at></due_at>
  <lock_at></lock_at>
  <unlock_at></unlock_at>
  <published>false</published>
  <one_time_results>false</one_time_results>
  <show_correct_answers_last_attempt>false</show_correct_answers_last_attempt>
  <only_visible_to_overrides>false</only_visible_to_overrides>
  <module_locked>false</module_locked>
  <created_at>${dateIso}</created_at>
  <updated_at>${dateIso}</updated_at>
</quiz>
`;
}

export function isXmlWellFormed(xmlText) {
  if (typeof DOMParser === 'undefined') {
    return {
      ok: true,
      message: 'DOMParser is not available in this runtime.'
    };
  }

  const document = new DOMParser().parseFromString(xmlText, 'application/xml');
  const parserError = document.querySelector('parsererror');

  if (parserError) {
    return {
      ok: false,
      message: parserError.textContent?.trim() || 'XML parsing error'
    };
  }

  return {
    ok: true,
    message: 'XML is well-formed.'
  };
}

export async function generateQtiPackageArtifacts(rawInputText, options = {}) {
  const parsed = options.parsed || parseQuestionBank(rawInputText, options);
  const { title, questions } = parsed;
  const hash = await shortStableHash(`${title}:${questions.length}`);
  const assessmentId = `word_qti_assessment_${hash}`;
  const folderName = assessmentId;
  const dateIso = new Date().toISOString().slice(0, 10);
  const skipZipData = Boolean(options.skipZipData);

  const assessmentXml = createAssessmentXml(assessmentId, title, questions);
  const manifestXml = createManifestXml(assessmentId, folderName, dateIso);
  const assessmentMetaXml = createAssessmentMetaXml(title, dateIso);

  let zipData = null;
  if (!skipZipData) {
    const zip = new JSZip();
    zip.file('imsmanifest.xml', manifestXml);
    zip.file(`${folderName}/${assessmentId}.xml`, assessmentXml);
    zip.file(`${folderName}/assessment_meta.xml`, assessmentMetaXml);

    if (options.debugInfo?.enabled) {
      const debugLines = [
        `source=${options.debugInfo.source || 'unknown'}`,
        `equationCount=${options.debugInfo.equationCount ?? 0}`,
        '--- parser input ---',
        options.debugInfo.rawText || ''
      ];

      zip.file('debug/extraction.txt', debugLines.join('\n'));
    }

    zipData = await zip.generateAsync({ type: 'uint8array' });
  }

  return {
    title,
    questions,
    assessmentId,
    folderName,
    manifestXml,
    assessmentXml,
    assessmentMetaXml,
    zipData
  };
}

export async function generateQtiZipBuffer(rawInputText, options = {}) {
  const artifacts = await generateQtiPackageArtifacts(rawInputText, options);
  return artifacts.zipData;
}

export function parseForValidation(rawInputText, options = {}) {
  return parseQuestionBank(rawInputText, options);
}
