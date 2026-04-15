import fs from 'node:fs/promises';
import JSZip from 'jszip';
import path from 'node:path';

async function main() {
  const input = process.argv[2] || path.join('tests', 'qti_test_generated.zip');
  const output = process.argv[3] || path.join('tests', 'qti_canvas_compatible.zip');

  const buffer = await fs.readFile(input);
  const zip = await JSZip.loadAsync(buffer);

  const manifestFile = zip.file('imsmanifest.xml');
  if (!manifestFile) throw new Error('imsmanifest.xml missing');
  let manifest = await manifestFile.async('text');

  // Find the resource file href for first imsqti resource
  const hrefMatch = manifest.match(/<resource[^>]*type="imsqti_xmlv1p2"[^>]*>[\s\S]*?<file href="([^"]+)"\s*\/?>/i);
  if (!hrefMatch) {
    throw new Error('Unable to locate imsqti resource href in manifest');
  }
  const originalHref = hrefMatch[1];

  // Read assessment xml from original href
  const assessmentFile = zip.file(originalHref);
  if (!assessmentFile) throw new Error(`Assessment XML ${originalHref} not found in zip`);
  const assessmentXml = await assessmentFile.async('nodebuffer');

  // Update manifest: change type to imsqti_xmlv1p2p1 and update file href to 'assessment.xml'
  manifest = manifest.replace(/type="imsqti_xmlv1p2"/g, 'type="imsqti_xmlv1p2p1"');
  manifest = manifest.replace(new RegExp(originalHref.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), 'assessment.xml');

  // Create new zip with modified manifest and root assessment.xml
  const outZip = new JSZip();
  outZip.file('imsmanifest.xml', manifest, { binary: false });
  outZip.file('assessment.xml', assessmentXml, { binary: true });

  const outBuffer = await outZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  await fs.writeFile(output, outBuffer);
  console.log('Wrote', output);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
