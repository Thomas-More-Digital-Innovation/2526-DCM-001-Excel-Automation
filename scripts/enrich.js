// Same pipeline as the n8n workflow, from the command line, using the per-header mappings saved in Postgres.
// Run: npm run enrich -- path/to/file.xlsx   (map the template once in the portal first)
const fs = require('fs');
const path = require('path');
const { enrichWorkbook } = require('../src/lib');

async function main() {
  const filePath = process.argv[2];
  if (!filePath) { console.error('Usage: npm run enrich -- <file.xlsx>'); process.exit(1); }

  const workbook = await enrichWorkbook(fs.readFileSync(filePath), null, process.env, console.log);

  const outPath = filePath.replace(/\.xlsx?$/i, '_enriched.xlsx');
  await workbook.xlsx.writeFile(outPath);
  console.log(`Enriched file saved to: ${outPath}`);
}

main().catch(err => { console.error(err); process.exit(1); });
