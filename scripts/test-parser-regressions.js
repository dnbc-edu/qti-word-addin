#!/usr/bin/env node
import assert from 'node:assert/strict';
import { parseForValidation } from '../src/qti-generator.js';

function run(name, fn) {
  try {
    fn();
    console.log(`PASS ${name}`);
  } catch (error) {
    console.error(`FAIL ${name}`);
    throw error;
  }
}

run('Stem ending with colon is treated as question stem', () => {
  const text = `
1. A customer record needs correction. The bank must comply under the:
- Right to be Informed
- Right to Rectification {{ANS}}
- Right to Data Portability
- Right to Damages
`;

  const parsed = parseForValidation(text);
  assert.equal(parsed.questions.length, 1);
  assert.equal(parsed.questions[0].choices.length, 4);
  assert.equal(parsed.questions[0].choices.find((c) => c.isCorrect)?.text, 'Right to Rectification');
});

run('Flattened inline lettered choices are split via fallback parser', () => {
  const text = '1. Which one is correct? a) Alpha b) Bravo {{ANS}} c) Charlie d) Delta';
  const parsed = parseForValidation(text);
  assert.equal(parsed.questions.length, 1);
  assert.equal(parsed.questions[0].choices.length, 4);
  assert.equal(parsed.questions[0].choices.find((c) => c.isCorrect)?.text, 'Bravo');
});

run('Choice lines ending with question mark are not treated as stems', () => {
  const text = `
1. Which line is the actual answer?
- Is this a trick choice?
- No, this is the actual answer {{ANS}}
- Another distractor
- Last distractor
2. Which number comes next?
- 1
- 2
- 3 {{ANS}}
- 4
`;
  const parsed = parseForValidation(text);
  assert.equal(parsed.questions.length, 2);
});

run('Strict mode fails on multiple marked correct answers', () => {
  const text = `
1. Which one is a fruit?
1) [x] Apple
2) [x] Carrot
3) Potato
4) Celery
`;
  assert.throws(() => parseForValidation(text), /multiple correct choices/i);
});

run('Permissive mode keeps only first marked answer', () => {
  const text = `
1. Which one is a fruit?
1) [x] Apple
2) [x] Carrot
3) Potato
4) Celery
`;
  const parsed = parseForValidation(text, { permissive: true });
  const correct = parsed.questions[0].choices.filter((c) => c.isCorrect);
  assert.equal(correct.length, 1);
  assert.equal(correct[0].text, 'Apple');
});

run('Strict mode fails when no answer is marked', () => {
  const text = `
1. Which one is correct?
1) A
2) B
3) C
4) D
`;
  assert.throws(() => parseForValidation(text), /exactly 1 correct choice/i);
});

run('Permissive mode defaults to first answer when none is marked', () => {
  const text = `
1. Which one is correct?
1) A
2) B
3) C
4) D
`;
  const parsed = parseForValidation(text, { permissive: true });
  const correct = parsed.questions[0].choices.filter((c) => c.isCorrect);
  assert.equal(correct.length, 1);
  assert.equal(correct[0].text, 'A');
});

console.log('All parser regression tests passed.');