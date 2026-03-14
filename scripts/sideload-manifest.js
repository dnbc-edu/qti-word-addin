import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

const defaultMacWordSideloadDir = path.join(
  os.homedir(),
  'Library',
  'Containers',
  'com.microsoft.Word',
  'Data',
  'Documents',
  'wef'
);

const defaultMacSharedOfficeSideloadDir = path.join(
  os.homedir(),
  'Library',
  'Group Containers',
  'UBF8T346G9.Office',
  'wef'
);

async function ensureDir(dirPath) {
  await fs.mkdir(dirPath, { recursive: true });
}

async function copyManifest(sourcePath, destinationPath) {
  await fs.copyFile(sourcePath, destinationPath);
}

async function main() {
  const workspaceRoot = process.cwd();
  const sourceManifest = path.join(workspaceRoot, 'addin', 'manifest.xml');
  const configuredSideloadDir = process.env.WORD_SIDELOAD_DIR;
  const sideloadDirs = configuredSideloadDir
    ? [configuredSideloadDir]
    : [defaultMacWordSideloadDir, defaultMacSharedOfficeSideloadDir];

  for (const sideloadDir of sideloadDirs) {
    const destinationManifest = path.join(sideloadDir, 'qti-exporter-manifest.xml');
    await ensureDir(sideloadDir);
    await copyManifest(sourceManifest, destinationManifest);
    console.log(`Manifest copied to: ${destinationManifest}`);
  }

  console.log('Restart Word if the add-in does not appear immediately.');
}

main().catch((error) => {
  console.error(`Sideload failed: ${error.message}`);
  process.exit(1);
});
