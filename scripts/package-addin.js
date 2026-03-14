import fs from 'node:fs/promises';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);

function parseArgs(argv) {
  const args = { prod: false, baseUrl: '' };

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (token === '--prod') {
      args.prod = true;
      continue;
    }

    if (token === '--base-url') {
      args.baseUrl = argv[index + 1] ?? '';
      index += 1;
    }
  }

  return args;
}

function normalizeBaseUrl(rawBaseUrl) {
  if (!rawBaseUrl) {
    return '';
  }

  let url;
  try {
    url = new URL(rawBaseUrl);
  } catch {
    throw new Error(`Invalid ADDIN_BASE_URL or --base-url value: ${rawBaseUrl}`);
  }

  if (url.protocol !== 'https:') {
    throw new Error('ADDIN_BASE_URL must use https://');
  }

  return url.origin + url.pathname.replace(/\/$/, '');
}

function rewriteManifestForBaseUrl(manifestXml, baseUrl) {
  return manifestXml.replace(/https:\/\/localhost:3000/g, baseUrl);
}

async function pathExists(targetPath) {
  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const effectiveBaseUrl = normalizeBaseUrl(args.baseUrl || process.env.ADDIN_BASE_URL || '');
  const isProdPackage = args.prod || Boolean(effectiveBaseUrl);

  if (args.prod && !effectiveBaseUrl) {
    throw new Error('Production packaging requires ADDIN_BASE_URL or --base-url.');
  }

  const workspaceRoot = process.cwd();
  const manifestPath = path.join(workspaceRoot, 'addin', 'manifest.xml');
  const distAddinPath = path.join(workspaceRoot, 'dist-addin');
  const releaseDir = path.join(workspaceRoot, 'release');
  const bundleName = isProdPackage ? 'qti-word-addin-release-prod' : 'qti-word-addin-release';
  const bundleDir = path.join(releaseDir, bundleName);
  const zipPath = path.join(releaseDir, `${bundleName}.zip`);

  const hasManifest = await pathExists(manifestPath);
  const hasDist = await pathExists(distAddinPath);

  if (!hasManifest) {
    throw new Error('Missing addin/manifest.xml.');
  }

  if (!hasDist) {
    throw new Error('Missing dist-addin. Run "npm run build:addin" first.');
  }

  await fs.rm(bundleDir, { recursive: true, force: true });
  await fs.rm(zipPath, { force: true });
  await fs.mkdir(bundleDir, { recursive: true });

  const manifestXml = await fs.readFile(manifestPath, 'utf8');
  const outputManifestXml = effectiveBaseUrl
    ? rewriteManifestForBaseUrl(manifestXml, effectiveBaseUrl)
    : manifestXml;

  await fs.writeFile(path.join(bundleDir, 'manifest.xml'), outputManifestXml, 'utf8');
  await fs.cp(distAddinPath, path.join(bundleDir, 'dist-addin'), { recursive: true });

  await execFileAsync('zip', ['-rq', zipPath, bundleName], { cwd: releaseDir });

  console.log(`Release bundle created: ${zipPath}`);
  console.log(`Bundle contents directory: ${bundleDir}`);
  if (effectiveBaseUrl) {
    console.log(`Manifest URLs rewritten to: ${effectiveBaseUrl}`);
  }
}

main().catch((error) => {
  console.error(`Packaging failed: ${error.message}`);
  process.exit(1);
});
