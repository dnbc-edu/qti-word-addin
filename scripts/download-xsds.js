import fs from 'node:fs/promises';
import path from 'node:path';

const WORKSPACE = process.cwd();

const TARGETS = [
  {
    name: 'qti-1p2',
    rootUrl: 'https://www.imsglobal.org/xsd/ims_qtiasiv1p2p1.xsd',
    outputDir: path.join(WORKSPACE, 'schemas', 'qti-1p2')
  },
  {
    name: 'qti-2p1',
    rootUrl: 'https://www.imsglobal.org/xsd/imsqti_v2p1.xsd',
    outputDir: path.join(WORKSPACE, 'schemas', 'qti-2p1')
  }
];

function extractSchemaLocations(xsdText) {
  const matches = [];
  const regex = /schemaLocation\s*=\s*"([^"]+)"/g;
  let match;

  while ((match = regex.exec(xsdText)) !== null) {
    matches.push(match[1]);
  }

  return matches;
}

function toRelativePath(urlString) {
  const url = new URL(urlString);
  const cleaned = url.pathname.replace(/^\/+/, '');
  return cleaned || 'index.xsd';
}

async function fetchText(url) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} for ${url}`);
  }

  return response.text();
}

async function writeXsd(outputDir, sourceUrl, content) {
  const relativePath = toRelativePath(sourceUrl);
  const destination = path.join(outputDir, relativePath);
  await fs.mkdir(path.dirname(destination), { recursive: true });
  await fs.writeFile(destination, content, 'utf8');
  return destination;
}

async function downloadSchemaSet(rootUrl, outputDir) {
  const visited = new Set();
  const queue = [rootUrl];

  while (queue.length) {
    const current = queue.shift();
    if (visited.has(current)) {
      continue;
    }

    visited.add(current);

    let xsdText;
    try {
      xsdText = await fetchText(current);
    } catch (error) {
      console.warn(`Skip ${current}: ${error.message}`);
      continue;
    }

    await writeXsd(outputDir, current, xsdText);

    const locations = extractSchemaLocations(xsdText);
    for (const location of locations) {
      let resolved;
      try {
        resolved = new URL(location, current).toString();
      } catch {
        continue;
      }

      if (!resolved.endsWith('.xsd')) {
        continue;
      }

      if (!visited.has(resolved)) {
        queue.push(resolved);
      }
    }
  }

  return visited.size;
}

async function main() {
  for (const target of TARGETS) {
    const count = await downloadSchemaSet(target.rootUrl, target.outputDir);
    console.log(`[${target.name}] downloaded ${count} schema file(s) to ${target.outputDir}`);
  }
}

main().catch((error) => {
  console.error(`Download failed: ${error.message}`);
  process.exit(1);
});
