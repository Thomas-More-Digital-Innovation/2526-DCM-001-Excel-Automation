// Run: npm test   (offline — no Akeneo, OpenAI or Postgres needed)
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { loadWorkbook, extractHeadersAndRows, buildSlotMap, resolveValue, cleanMapping, readUpload, buildSuggestColumns, prepareForWrite, headerKeys, loadHeaderMappings, saveHeaderMappings } = require('../src/lib');
const { buildWorkflow } = require('../scripts/build-workflow');

const root = path.join(__dirname, '..');
const sample = fs.readFileSync(path.join(root, 'Docs/InitialFile.xlsx'));

test('finds header row, data rows and SKUs in the Floralux template', async () => {
  const workbook = await loadWorkbook(sample);
  const r = extractHeadersAndRows(workbook.worksheets[0]);
  assert.strictEqual(r.headerRowIndex, 23);
  assert.strictEqual(r.dataStartRow, 25);
  assert.strictEqual(r.headerCols.length, 136);
  assert.strictEqual(r.rows.length, 14);
  assert.deepStrictEqual(r.rows.slice(0, 3).map(x => x.sku), ['1007143', '1007144', '1007086']);
  assert.deepStrictEqual(r.headerCols.slice(0, 3).map(h => h.header), ['IDNLF', 'MARA-BISMT', 'MATNR']);
});

test('hints combine every description row above the header, but not the example row below it', async () => {
  const workbook = await loadWorkbook(sample);
  const r = extractHeadersAndRows(workbook.worksheets[0]);
  const idnlf = r.headerCols.find(h => h.header === 'IDNLF');
  assert.deepStrictEqual(idnlf.hint.split('\n'), [
    'INFORMATION BASE UNIT',
    'Same as article code on packing slip and invoice',
    'Artikelcode leverancier',
    'Code article fournisseur',
    'Supplier article code',
  ]);
  assert.ok(!r.headerCols.some(h => h.hint && h.hint.includes('Floralux N.V.')), 'branding rows must not become hints');
});

test('SKU column is found by its header, and left empty (null) when no header names it', async () => {
  const ExcelJS = require('exceljs');
  // Like real templates: data rows are sparser than the header row (they get filled by us).
  const build = headers => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.addRow(headers);
    ws.addRow(['A-100', 'Rood']);
    ws.addRow(['A-200', 'Blauw']);
    return ws;
  };
  const named = extractHeadersAndRows(build(['Artikelnummer', 'Kleur', 'Gewicht', 'Merk', 'EAN']));
  assert.strictEqual(named.skuColNum, 1);
  assert.deepStrictEqual(named.rows.map(r => r.sku), ['A-100', 'A-200']);

  const unnamed = extractHeadersAndRows(build(['Code', 'Kleur', 'Gewicht', 'Merk', 'EAN']));
  assert.strictEqual(unnamed.skuColNum, null);
  assert.deepStrictEqual(unnamed.rows, []);
});

const units = [
  { unit_type: 'CE', unit_used: 'BOT', Aantal_CE: '1', EAN: '111' },
  { unit_type: 'X', unit_used: 'KAR', Aantal_CE: '12', EAN: '222' },
  { unit_type: 'X', unit_used: 'EP', Aantal_CE: '480', EAN: '333' },
];
const product = {
  units_by_type: [{ locale: null, scope: null, data: units }],
  title: [{ locale: 'nl_BE', scope: null, data: 'Mierenspray' }, { locale: 'fr_BE', scope: null, data: 'Spray fourmis' }],
  weight: [{ locale: null, scope: null, data: { amount: '1.250', unit: 'KILOGRAM' } }],
  fostplus: [{ locale: null, scope: null, data: true }],
  __categories: ['birchmeier_sprayers'],
};

test('packing units: CE is the base, the rest go largest-first into pallet/carton/box', () => {
  const map = buildSlotMap(product);
  assert.strictEqual(map.base.unit_used, 'BOT');
  assert.strictEqual(map[4].unit_used, 'EP');
  assert.strictEqual(map[3].unit_used, 'KAR');
  assert.strictEqual(map[2], undefined);
});

test('mapping entries resolve to the right value', () => {
  assert.strictEqual(resolveValue(null, product), null);
  assert.strictEqual(resolveValue({ attr: 'title', locale: 'fr_BE' }, product), 'Spray fourmis');
  assert.strictEqual(resolveValue({ attr: 'weight' }, product), 1.25);
  assert.strictEqual(resolveValue({ attr: 'fostplus', yesno: true }, product), 'Y');
  assert.strictEqual(resolveValue({ attr: 'missing', yesno: true }, product), 'N');
  assert.strictEqual(resolveValue({ attr: 'missing' }, product), null);
  assert.strictEqual(resolveValue({ calc: 'base_unit' }, product), 'BOT');
  assert.strictEqual(resolveValue({ calc: 'pack_3_ean' }, product), '222');
  assert.strictEqual(resolveValue({ calc: 'pack_2_ean' }, product), null);
  assert.strictEqual(resolveValue({ calc: 'min_order_qty' }, product), '12');
  assert.strictEqual(resolveValue({ calc: 'origin_from_category' }, product), 'DE');
});

test('portal mappings are cleaned: unknown columns, calc fields, junk and the SKU column are dropped', () => {
  const headerCols = [1, 2, 3, 4, 5].map(colNum => ({ header: 'H', colNum }));
  const cleaned = cleanMapping({
    1: { attr: 'sku' },
    2: { attr: 'title', locale: 'nl_BE', yesno: 'yes', extra: 1 },
    3: { calc: 'no_such_field' },
    4: { calc: 'base_unit' },
    9: { attr: 'title' },
  }, headerCols, 1);
  assert.deepStrictEqual(cleaned, { 1: null, 2: { attr: 'title', locale: 'nl_BE' }, 3: null, 4: { calc: 'base_unit' }, 5: null });
});

test('OpenAI only gets header texts and hints, never row data, and not the SKU column', async () => {
  const workbook = await loadWorkbook(sample);
  const r = extractHeadersAndRows(workbook.worksheets[0]);
  const columns = buildSuggestColumns(r.headerCols, r.skuColNum);
  assert.strictEqual(columns.length, r.headerCols.length - 1);
  for (const c of columns) assert.deepStrictEqual(Object.keys(c).filter(k => !['col', 'name', 'hint'].includes(k)), []);
  const sent = JSON.stringify(columns);
  for (const { sku } of r.rows) assert.ok(!sent.includes(sku), 'SKU ' + sku + ' leaked into the OpenAI payload');
});

test('written file has no empty conditional formatting (Excel would show a repair prompt)', async () => {
  const JSZip = require('jszip');
  const workbook = await loadWorkbook(sample);
  prepareForWrite(workbook);
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  for (const name of Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
    const xml = await zip.file(name).async('string');
    assert.ok(!/<conditionalFormatting[^>]*\/>/.test(xml), name + ' has an empty <conditionalFormatting/>');
  }
});

test('header keys number repeated headers left to right', () => {
  const keys = headerKeys([{ header: 'EAN', colNum: 3 }, { header: 'MATNR', colNum: 4 }, { header: 'EAN', colNum: 9 }]);
  assert.deepStrictEqual(keys, [
    { colNum: 3, header: 'EAN', occurrence: 1 },
    { colNum: 4, header: 'MATNR', occurrence: 1 },
    { colNum: 9, header: 'EAN', occurrence: 2 },
  ]);
});

// Needs Postgres: node --env-file=.env --test   (skipped by plain `npm test`)
test('header mappings are remembered per header, across templates', { skip: !process.env.MAPPINGS_DB_URL }, async () => {
  const { Client } = require('pg');
  const tag = 'TEST_' + Date.now() + '_';
  const templateA = [{ header: tag + 'SKU', colNum: 1 }, { header: tag + 'NAAM', colNum: 2 }, { header: tag + 'LEEG', colNum: 3 }];
  try {
    await saveHeaderMappings(process.env, templateA, { skuCol: 1, mapping: { 1: null, 2: { attr: 'title', locale: 'nl_BE' }, 3: null } });
    // Different template: same NAAM header in another column, plus a header never seen before.
    const known = await loadHeaderMappings(process.env, [{ header: tag + 'NIEUW', colNum: 1 }, { header: tag + 'NAAM', colNum: 5 }, { header: tag + 'SKU', colNum: 7 }, { header: tag + 'LEEG', colNum: 8 }]);
    assert.deepStrictEqual(known, {
      5: { entry: { attr: 'title', locale: 'nl_BE' }, isSku: false },
      7: { entry: null, isSku: true },
      8: { entry: null, isSku: false },
    });
  } finally {
    const db = new Client({ connectionString: process.env.MAPPINGS_DB_URL });
    await db.connect();
    await db.query('DELETE FROM header_mappings WHERE header LIKE $1', [tag + '%']);
    await db.end();
  }
});

test('uploads are read from the n8n webhook item', () => {
  const up = readUpload({ binary: { data: { data: sample.toString('base64'), fileName: 'x.xlsx' } } });
  assert.strictEqual(up.fileName, 'x.xlsx');
  assert.strictEqual(up.buffer.length, sample.length);
  assert.throws(() => readUpload({ binary: {} }), /No file received/);
});

test('workflow.json is built from src/ (run `npm run build` if this fails)', () => {
  // Ignore CRLF: git autocrlf rewrites line endings on Windows checkouts.
  const lf = s => s.replace(/\r\n/g, '\n');
  assert.strictEqual(lf(fs.readFileSync(path.join(root, 'workflow.json'), 'utf8')), lf(buildWorkflow()));
});

test('built n8n Code nodes compile', () => {
  // Same shape n8n uses: code body inside an async function.
  const AsyncFunction = (async () => {}).constructor;
  for (const node of JSON.parse(buildWorkflow()).nodes.filter(n => n.parameters.jsCode !== undefined)) {
    assert.doesNotThrow(() => new AsyncFunction('require', '$input', '$env', node.parameters.jsCode), node.name);
  }
});
