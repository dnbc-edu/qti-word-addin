import fs from 'node:fs/promises';
import path from 'node:path';
import { generateQtiZipBuffer, parseForValidation } from '../src/qti-generator.js';

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

  const inputText = await fs.readFile(args.input, 'utf8');
  const parsed = parseForValidation(inputText);
  const zipBuffer = await generateQtiZipBuffer(inputText);

  await fs.mkdir(path.dirname(args.output), { recursive: true });
  await fs.writeFile(args.output, zipBuffer);

  console.log(`Generated ${args.output}`);
  console.log(`Title: ${parsed.title}`);
  console.log(`Questions: ${parsed.questions.length}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
