import fs from 'node:fs/promises';
import path from 'node:path';

function normalizeBaseUrl(rawBaseUrl) {
  if (!rawBaseUrl) {
    throw new Error('ADDIN_BASE_URL is required.');
  }

  let parsedUrl;
  try {
    parsedUrl = new URL(rawBaseUrl);
  } catch {
    throw new Error(`Invalid ADDIN_BASE_URL value: ${rawBaseUrl}`);
  }

  if (parsedUrl.protocol !== 'https:') {
    throw new Error('ADDIN_BASE_URL must use https://');
  }

  return parsedUrl.origin + parsedUrl.pathname.replace(/\/$/, '');
}

async function main() {
  const workspaceRoot = process.cwd();
  const manifestPath = path.join(
    workspaceRoot,
    'release',
    'qti-word-addin-release-prod',
    'manifest.xml'
  );

  const baseUrl = normalizeBaseUrl(process.env.ADDIN_BASE_URL || '');
  const manifestXml = await fs.readFile(manifestPath, 'utf8');

  if (manifestXml.includes('https://localhost:3000')) {
    throw new Error('Production manifest still contains localhost URLs.');
  }

  if (!manifestXml.includes(baseUrl)) {
    throw new Error(`Production manifest does not contain expected base URL: ${baseUrl}`);
  }

  const requiredPaths = [
    `${baseUrl}/addin/taskpane.html`,
    `${baseUrl}/addin/assets/icon-16.png`,
    `${baseUrl}/addin/assets/icon-32.png`,
    `${baseUrl}/addin/assets/icon-80.png`
  ];

  for (const requiredPath of requiredPaths) {
    if (!manifestXml.includes(requiredPath)) {
      throw new Error(`Missing expected manifest URL: ${requiredPath}`);
    }
  }

  console.log(`Production manifest verification passed: ${manifestPath}`);
}

main().catch((error) => {
  console.error(`Manifest verification failed: ${error.message}`);
  process.exit(1);
});
