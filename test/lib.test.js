// Run: npm test   (offline — no Akeneo, OpenAI or Postgres needed)
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { loadWorkbook, extractHeadersAndRows, buildSlotMap, resolveValue, cleanMapping, readUpload, buildSuggestColumns } = require('../src/lib');
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
