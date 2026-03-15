import fs from 'node:fs/promises';
import path from 'node:path';

async function main() {
  const workspaceRoot = process.cwd();
  const sourceAssetsDir = path.join(workspaceRoot, 'addin', 'assets');
  const targetAssetsDir = path.join(workspaceRoot, 'dist-addin', 'addin', 'assets');

  await fs.mkdir(targetAssetsDir, { recursive: true });

  const iconFiles = ['icon-16.png', 'icon-32.png', 'icon-80.png'];

  for (const iconFile of iconFiles) {
    const sourcePath = path.join(sourceAssetsDir, iconFile);
    const targetPath = path.join(targetAssetsDir, iconFile);
    await fs.copyFile(sourcePath, targetPath);
  }

  console.log('Copied add-in icon assets to dist-addin/addin/assets.');
}

main().catch((error) => {
  console.error(`Failed to copy add-in static assets: ${error.message}`);
  process.exit(1);
});
