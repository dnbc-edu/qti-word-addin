import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import JSZip from 'jszip';

const SCHEMA_MAP = {
  '1p2': 'schemas/qti-1p2/xsd/ims_qtiasiv1p2p1.xsd',
  '2p1': 'schemas/qti-2p1/xsd/imsqti_v2p1.xsd'
};

function parseArgs(argv) {
  const args = {
    version: '1p2'
  };

  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];

    if (token === '--zip') {
      args.zip = argv[i + 1];
      i += 1;
      continue;
    }

    if (token === '--version') {
      args.version = argv[i + 1];
      i += 1;
    }
  }

  return args;
}

function runCommand(command, commandArgs) {
  return new Promise((resolve) => {
    const child = spawn(command, commandArgs, { stdio: ['ignore', 'pipe', 'pipe'] });

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

function findAssessmentXmlEntry(zip) {
  const files = Object.keys(zip.files);
  const candidates = files.filter((name) => {
    if (!name.endsWith('.xml')) {
      return false;
    }

    if (name === 'imsmanifest.xml') {
      return false;
    }

    if (name.endsWith('/assessment_meta.xml')) {
      return false;
    }

    return true;
  });

  if (!candidates.length) {
    return null;
  }

  return candidates[0];
}

async function extractAssessmentXml(zipPath) {
  const zipBuffer = await fs.readFile(zipPath);
  const zip = await JSZip.loadAsync(zipBuffer);
  const entryName = findAssessmentXmlEntry(zip);

  if (!entryName) {
    throw new Error('Could not find assessment XML inside ZIP.');
  }

  const entry = zip.file(entryName);
  if (!entry) {
    throw new Error(`ZIP entry not found: ${entryName}`);
  }

  const content = await entry.async('string');
  return { entryName, content };
}

async function writeTempXml(content) {
  const filePath = path.join(os.tmpdir(), `qti-validation-${Date.now()}.xml`);
  await fs.writeFile(filePath, content, 'utf8');
  return filePath;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.zip) {
    throw new Error('Usage: node scripts/validate-qti.js --zip <path-to-qti-zip> [--version 1p2|2p1]');
  }

  const schemaRelative = SCHEMA_MAP[args.version];
  if (!schemaRelative) {
    throw new Error(`Unsupported --version '${args.version}'. Use 1p2 or 2p1.`);
  }

  const schemaPath = path.resolve(process.cwd(), schemaRelative);

  try {
    await fs.access(schemaPath);
  } catch {
    throw new Error(`Schema not found: ${schemaPath}. Run npm run download:xsds first.`);
  }

  const check = await runCommand('xmllint', ['--version']);
  if (check.code !== 0) {
    throw new Error('xmllint is not available. Install libxml2 tools (e.g., brew install libxml2).');
  }

  const { entryName, content } = await extractAssessmentXml(args.zip);
  const tempXmlPath = await writeTempXml(content);

  const result = await runCommand('xmllint', ['--noout', '--schema', schemaPath, tempXmlPath]);

  await fs.unlink(tempXmlPath).catch(() => {});

  if (result.code !== 0) {
    console.error(`Validation failed for ${entryName}`);
    if (result.stdout.trim()) {
      console.error(result.stdout.trim());
    }
    if (result.stderr.trim()) {
      console.error(result.stderr.trim());
    }
    process.exit(1);
  }

  console.log(`Validation succeeded for ${entryName}`);
  console.log(`Schema: ${schemaPath}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
