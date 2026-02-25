/**
 * diagnose-sarif.js
 * Dumps the raw SARIF structure from an .msapp or .zip file so we can see
 * exactly what fields Power Apps AppChecker emits.
 *
 * Usage:
 *   node diagnose-sarif.js path/to/yourapp.msapp
 *   node diagnose-sarif.js path/to/solution.zip
 */

const JSZip = require('jszip');
const fs    = require('fs');
const path  = require('path');

const filePath = process.argv[2];
if (!filePath) { console.error('Usage: node diagnose-sarif.js <file.msapp|file.zip>'); process.exit(1); }

(async () => {
  const buf = fs.readFileSync(filePath);

  async function inspectZip(data, label) {
    const zip = await JSZip.loadAsync(data);
    const pathMap = new Map();
    for (const k of Object.keys(zip.files)) pathMap.set(k.replace(/\\/g, '/'), k);

    // find .msapp files inside if this is a solution zip
    const msappFiles = Array.from(pathMap.keys()).filter(e => e.endsWith('.msapp'));
    if (msappFiles.length) {
      console.log(`\n${label} — contains ${msappFiles.length} .msapp file(s), recursing...`);
      for (const mp of msappFiles) {
        const msappBuf = await zip.file(pathMap.get(mp)).async('nodebuffer');
        await inspectZip(msappBuf, mp);
      }
      return;
    }

    // find SARIF
    const sarifKey = Array.from(pathMap.keys()).find(e => e.endsWith('.sarif'));
    if (!sarifKey) { console.log(`\n${label} — NO .sarif file found`); return; }

    console.log(`\n════════════════════════════════`);
    console.log(`File: ${label}`);
    console.log(`SARIF path inside ZIP: ${sarifKey}`);

    const raw = await zip.file(pathMap.get(sarifKey)).async('string');
    const sarif = JSON.parse(raw);
    const runs = sarif.runs ?? [];

    for (let ri = 0; ri < runs.length; ri++) {
      const run = runs[ri];
      const results = run.results ?? [];
      const rules   = run.tool?.driver?.rules ?? [];

      console.log(`\nRun ${ri}: ${results.length} results, ${rules.length} rules in driver`);

      // Show first 3 rules
      rules.slice(0, 3).forEach((r, i) => {
        console.log(`  Rule[${i}]: id=${r.id}  shortDesc=${r.shortDescription?.text ?? '—'}`);
        console.log(`           defaultConfig.level=${r.defaultConfiguration?.level ?? '—'}`);
      });

      // Show first 5 results in full detail
      console.log(`\n--- First 5 results (raw structure) ---`);
      results.slice(0, 5).forEach((res, i) => {
        console.log(`\n  Result[${i}]:`);
        console.log(`    ruleId   : ${res.ruleId ?? '(absent)'}`);
        console.log(`    level    : ${res.level ?? '(absent — SARIF default = warning)'}`);
        console.log(`    kind     : ${res.kind ?? '(absent)'}`);

        // message
        const msg = res.message;
        if (typeof msg === 'string') console.log(`    message  : "${msg}"`);
        else if (msg) {
          console.log(`    message.text    : ${msg.text ?? '—'}`);
          console.log(`    message.id      : ${msg.id ?? '—'}`);
          console.log(`    message.arguments: ${JSON.stringify(msg.arguments ?? [])}`);
        }

        // locations
        const locs = res.locations ?? [];
        locs.forEach((loc, li) => {
          console.log(`    location[${li}] keys: ${Object.keys(loc).join(', ')}`);

          // physical
          const phys = loc.physicalLocation;
          if (phys) {
            console.log(`      physicalLocation.artifactLocation.uri   : ${phys.artifactLocation?.uri ?? '—'}`);
            console.log(`      physicalLocation.artifactLocation.uriBaseId: ${phys.artifactLocation?.uriBaseId ?? '—'}`);
            console.log(`      physicalLocation.region.startLine       : ${phys.region?.startLine ?? '—'}`);
            console.log(`      physicalLocation.region.startColumn     : ${phys.region?.startColumn ?? '—'}`);
          }

          // logical
          const logicals = loc.logicalLocations ?? [];
          logicals.forEach((ll, lli) => {
            console.log(`      logicalLocation[${lli}].name              : ${ll.name ?? '—'}`);
            console.log(`      logicalLocation[${lli}].fullyQualifiedName: ${ll.fullyQualifiedName ?? '—'}`);
            console.log(`      logicalLocation[${lli}].kind              : ${ll.kind ?? '—'}`);
          });
        });

        // properties (custom severity etc.)
        if (res.properties) {
          console.log(`    properties: ${JSON.stringify(res.properties)}`);
        }
      });

      // Tally level values across all results
      const levelCounts = {};
      for (const r of results) {
        const lv = r.level ?? '(absent)';
        levelCounts[lv] = (levelCounts[lv] ?? 0) + 1;
      }
      console.log(`\n--- Level distribution across all ${results.length} results ---`);
      Object.entries(levelCounts).forEach(([lv, cnt]) => console.log(`    ${lv}: ${cnt}`));

      // Location field availability
      let hasPhysical = 0, hasLogical = 0, hasNeither = 0;
      for (const r of results) {
        const loc = (r.locations ?? [])[0] ?? {};
        if (loc.physicalLocation?.artifactLocation?.uri) hasPhysical++;
        else if ((loc.logicalLocations ?? []).length) hasLogical++;
        else hasNeither++;
      }
      console.log(`\n--- Location field usage ---`);
      console.log(`    physicalLocation.artifactLocation.uri present: ${hasPhysical}`);
      console.log(`    logicalLocations present (no physUri):         ${hasLogical}`);
      console.log(`    neither:                                       ${hasNeither}`);
    }
})().catch(err => { console.error(err); process.exit(1); });
