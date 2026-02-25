#!/usr/bin/env node
/**
 * Diagnostic script — run from the app/ folder:
 *   node diagnose-msapp.js "/path/to/your/solution.zip"
 *
 * Shows exactly what the parser finds inside each .msapp.
 */
const JSZip = require('jszip');
const yaml  = require('js-yaml');
const fs    = require('fs');
const path  = require('path');

const solutionPath = process.argv[2];
if (!solutionPath) {
  console.error('Usage: node diagnose-msapp.js <path-to-solution.zip>');
  process.exit(1);
}

async function run() {
  console.log('\n=== Opening solution ZIP:', solutionPath);
  const solutionBuf = fs.readFileSync(solutionPath);
  const solutionZip = await JSZip.loadAsync(solutionBuf);

  const msappFiles = Object.keys(solutionZip.files)
    .filter(e => e.endsWith('.msapp') && !solutionZip.files[e].dir);

  console.log('Found .msapp files:', msappFiles.length);
  if (msappFiles.length === 0) {
    console.log('All entries:', Object.keys(solutionZip.files).slice(0, 30));
    return;
  }

  for (const msappPath of msappFiles) {
    console.log('\n--- .msapp:', msappPath);
    const buf = Buffer.from(await solutionZip.files[msappPath].async('arraybuffer'));
    console.log('  Buffer size:', buf.length, 'bytes');

    let msappZip;
    try {
      msappZip = await JSZip.loadAsync(buf);
    } catch(e) {
      console.log('  ERROR opening as ZIP:', e.message);
      continue;
    }

    const entries = Object.keys(msappZip.files);
    console.log('  Total entries inside .msapp:', entries.length);
    console.log('  ALL entries:');
    entries.forEach(e => console.log('    ' + e));

    // Show Src/ contents
    const srcFiles = entries.filter(e => e.toLowerCase().startsWith('src/'));
    console.log('  Src/ files:', srcFiles);

    const yamlFiles = srcFiles.filter(e => {
      const lower = e.toLowerCase();
      return lower.endsWith('.pa.yaml') || lower.endsWith('.fx.yaml') || lower.endsWith('.yaml');
    });
    console.log('  YAML files detected:', yamlFiles.length);

    // Parse each YAML file
    for (const yf of yamlFiles) {
      const file = msappZip.file(yf);
      if (!file) { console.log('    [skip] cannot read:', yf); continue; }
      const content = await file.async('string');
      let doc;
      try { doc = yaml.load(content); }
      catch(e) { console.log('    [YAML PARSE ERROR]', yf, e.message); continue; }

      const topKeys = doc && typeof doc === 'object' ? Object.keys(doc) : [];
      const hasScreens = 'Screens' in (doc || {});
      const screenNames = hasScreens ? Object.keys(doc.Screens) : [];

      console.log(`    ${yf}`);
      console.log(`      top-level keys: ${JSON.stringify(topKeys)}`);
      console.log(`      has "Screens" key: ${hasScreens}`);
      if (hasScreens) console.log(`      screen names: ${JSON.stringify(screenNames)}`);
    }
  }
}

run().catch(e => { console.error('Fatal:', e); process.exit(1); });
