/**
 * diagnose-datasources.js
 * Dumps the full structure of References/DataSources.json from an .msapp or
 * solution .zip so we can see every field Power Apps exports.
 *
 * Usage:
 *   node diagnose-datasources.js path/to/app.msapp
 *   node diagnose-datasources.js path/to/solution.zip
 */

const JSZip = require('jszip');
const fs    = require('fs');

const filePath = process.argv[2];
if (!filePath) { console.error('Usage: node diagnose-datasources.js <file>'); process.exit(1); }

(async () => {
  const buf = fs.readFileSync(filePath);

  async function inspectZip(data, label) {
    const zip = await JSZip.loadAsync(data);
    const pathMap = new Map();
    for (const k of Object.keys(zip.files)) pathMap.set(k.replace(/\\/g, '/'), k);

    // Recurse into .msapp files inside a solution zip
    const msappFiles = Array.from(pathMap.keys()).filter(e => e.endsWith('.msapp'));
    if (msappFiles.length) {
      for (const mp of msappFiles) {
        const msappBuf = await zip.file(pathMap.get(mp)).async('nodebuffer');
        await inspectZip(msappBuf, mp);
      }
      return;
    }

    const dsKey = Array.from(pathMap.keys()).find(
      e => e.toLowerCase() === 'references/datasources.json'
    );
    if (!dsKey) { console.log(`\n${label} — no References/DataSources.json found`); return; }

    const raw  = await zip.file(pathMap.get(dsKey)).async('string');
    const json = JSON.parse(raw);

    console.log(`\n════════════════════════════════════════`);
    console.log(`App: ${label}`);
    console.log(`File: ${dsKey}`);

    const entries = Array.isArray(json) ? json : (json.DataSources ?? json.dataSources ?? []);
    console.log(`Total data source entries: ${entries.length}\n`);

    entries.forEach((ds, i) => {
      console.log(`── Entry [${i}] ──────────────────────`);

      // Top-level keys
      const keys = Object.keys(ds);
      console.log(`  Top-level keys: ${keys.join(', ')}`);
      console.log(`  Name     : ${ds.Name ?? ds.name ?? '—'}`);
      console.log(`  Kind     : ${ds.Kind ?? ds.kind ?? '—'}`);
      console.log(`  TableName: ${ds.TableName ?? ds.tableName ?? '—'}`);
      console.log(`  DatasetName: ${ds.DatasetName ?? ds.datasetName ?? '—'}`);

      // Look for anything that looks like a schema / field list
      const schemaFields = keys.filter(k => {
        const lk = k.toLowerCase();
        return lk.includes('field') || lk.includes('column') || lk.includes('schema')
          || lk.includes('property') || lk.includes('attribute') || lk.includes('metadata');
      });

      if (schemaFields.length) {
        for (const sf of schemaFields) {
          const val = ds[sf];
          if (Array.isArray(val)) {
            console.log(`  ${sf} (array, ${val.length} items):`);
            val.slice(0, 5).forEach((item, fi) => {
              if (typeof item === 'object' && item !== null) {
                console.log(`    [${fi}] keys: ${Object.keys(item).join(', ')}`);
                // Print all scalar values
                for (const [k, v] of Object.entries(item)) {
                  if (typeof v !== 'object') console.log(`         ${k}: ${v}`);
                }
              } else {
                console.log(`    [${fi}]: ${item}`);
              }
            });
            if (val.length > 5) console.log(`    … and ${val.length - 5} more`);
          } else if (typeof val === 'object' && val !== null) {
            console.log(`  ${sf} (object): ${JSON.stringify(val).substring(0, 200)}`);
          } else {
            console.log(`  ${sf}: ${val}`);
          }
        }
      } else {
        console.log(`  (no field/column/schema keys found)`);
      }

      // Print all remaining scalar values for full picture
      console.log(`  Other scalar values:`);
      for (const [k, v] of Object.entries(ds)) {
        if (typeof v !== 'object' && !['Name','name','Kind','kind','TableName','tableName','DatasetName','datasetName'].includes(k)) {
          console.log(`    ${k}: ${v}`);
        }
      }

      // Print nested object keys (one level deep)
      for (const [k, v] of Object.entries(ds)) {
        if (typeof v === 'object' && v !== null && !Array.isArray(v) && !schemaFields.includes(k)) {
          console.log(`  ${k} (object): keys = ${Object.keys(v).join(', ')}`);
        }
      }

      console.log('');
    });
  }

  await inspectZip(buf, filePath);
})().catch(e => { console.error(e); process.exit(1); });
