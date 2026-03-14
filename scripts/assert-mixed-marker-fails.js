import { spawn } from 'node:child_process';

const expectedFragment = 'multiple correct choices marked across [x] and {{ANS}} styles';

function runGenerator() {
  return new Promise((resolve) => {
    const child = spawn(
      'node',
      [
        'scripts/generate-qti.js',
        '--input',
        'samples/question-bank-mixed-marker-negative.txt',
        '--output',
        'dist/should-not-exist.zip'
      ],
      { stdio: ['ignore', 'pipe', 'pipe'] }
    );

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (chunk) => {
      stdout += String(chunk);
    });

    child.stderr.on('data', (chunk) => {
      stderr += String(chunk);
    });

    child.on('close', (code) => {
      resolve({ code, stdout, stderr });
    });

    child.on('error', (error) => {
      resolve({ code: 127, stdout: '', stderr: error.message });
    });
  });
}

async function main() {
  const result = await runGenerator();
  const combined = `${result.stdout}\n${result.stderr}`;

  if (result.code === 0) {
    console.error('Expected generation to fail, but it succeeded.');
    process.exit(1);
  }

  if (!combined.includes(expectedFragment)) {
    console.error('Generation failed, but expected error message fragment was not found.');
    console.error(`Expected fragment: ${expectedFragment}`);
    console.error('--- actual output ---');
    console.error(combined.trim());
    process.exit(1);
  }

  console.log('PASS: mixed-marker conflict is rejected with expected error message.');
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
