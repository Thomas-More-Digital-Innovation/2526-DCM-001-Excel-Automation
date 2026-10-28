// Shared logic for the CLI (scripts/enrich.js) and the n8n Code nodes.
// n8n gets this file inlined by scripts/build-workflow.js — edit here, then run `npm run build`.
// Everything below the cli-only marker at the bottom is stripped for n8n.
//
// Which Excel column gets which Akeneo value is NOT decided here: it is a per-template mapping,
// stored in Postgres and edited in the portal. Keys are column numbers (headers can repeat),
// and each entry is one of:
//   null                                   → leave the column untouched
//   { attr: 'code', locale?: 'nl_BE', yesno?: true }  → an Akeneo attribute value
//   { calc: 'pack_3_ean' }                 → a computed field (see CALC_FIELDS)
const ExcelJS = require('exceljs');
const https = require('https');
const http = require('http');
const crypto = require('crypto');
const { URL } = require('url');
const { Client } = require('pg');

function checkEnv(env, keys) {
  keys.forEach(k => { if (!env[k]) throw new Error('Missing env var: ' + k); });
}
const AKENEO_ENV = ['AKENEO_BASE_URL', 'AKENEO_CLIENT_ID', 'AKENEO_CLIENT_SECRET', 'AKENEO_USERNAME', 'AKENEO_PASSWORD'];

// Plain http(s) instead of fetch: fetch is not guaranteed inside the n8n Code node sandbox.
function makeRequest(url, options) {
  options = options || {};
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const isHttps = urlObj.protocol === 'https:';
    const body = options.body ? Buffer.from(options.body) : null;
    const reqOptions = {
      hostname: urlObj.hostname,
      port: urlObj.port ? Number(urlObj.port) : (isHttps ? 443 : 80),
      path: urlObj.pathname + (urlObj.search || ''),
      method: options.method || 'GET',
      headers: Object.assign({}, options.headers || {}),
    };
    if (body) reqOptions.headers['Content-Length'] = body.length;
    const req = (isHttps ? https : http).request(reqOptions, res => {
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => {
        const text = Buffer.concat(chunks).toString('utf8');
        resolve({
          ok: res.statusCode >= 200 && res.statusCode < 300,
          status: res.statusCode,
          json: () => Promise.resolve(JSON.parse(text)),
          text: () => Promise.resolve(text),
        });
      });
      res.on('error', reject);
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

// ── Cell / value helpers ──────────────────────────────────────────────

function getCellText(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'string') { const s = v.trim(); return s || null; }
  if (v && typeof v === 'object') {
    if (typeof v.getTime === 'function') return v.toISOString();
    if (v.richText) { const s = v.richText.map(r => r.text).join('').trim(); return s || null; }
    if (v.text != null) {
      const t = typeof v.text === 'string' ? v.text
        : (v.text.richText ? v.text.richText.map(r => r.text).join('') : String(v.text));
      return t.trim() || null;
    }
    if (v.result !== undefined) return v.result != null ? String(v.result).trim() || null : null;
    if (v.formula !== undefined) return null;
  }
  return String(v).trim() || null;
}

function extractValue(attrValues) {
  if (!Array.isArray(attrValues) || attrValues.length === 0) return null;
  const hit = attrValues.find(v => v.locale === null && v.scope === null)
    || attrValues.find(v => v.locale === 'en_US')
    || attrValues[0];
  if (!hit) return null;
  const data = hit.data;
  if (data && typeof data === 'object' && !Array.isArray(data) && 'amount' in data) {
    const num = parseFloat(data.amount);
    return isNaN(num) ? data.amount : num;
  }
  if (Array.isArray(data)) {
    if (data.length === 0) return null;
    const first = data[0];
    if (typeof first === 'string') return first;  // e.g. country of origin ["BE"]
    if (first && typeof first === 'object' && 'amount' in first) {
      const num = parseFloat(first.amount);
      return isNaN(num) ? first.amount : num;
    }
    return String(first);
  }
  return data;
}

// Extracts the value for a specific locale from a localised Akeneo attribute.
function extractLocalizedValue(attrValues, locale) {
  if (!Array.isArray(attrValues) || attrValues.length === 0) return null;
  let hit = attrValues.find(v => v.locale === locale);
  if (!hit) {
    const prefix = locale.split('_')[0];
    hit = attrValues.find(v => v.locale && v.locale.split('_')[0] === prefix);
  }
  if (!hit) return null;
  const data = hit.data;
  if (typeof data === 'string') return data.trim() || null;
  return data;
}

function toYesNo(v) {
  if (v === true || v === 1) return 'Y';
  if (typeof v === 'string' && /^(true|ja|yes|y|1|x)$/i.test(v.trim())) return 'Y';
  return 'N';
}

// ── Computed fields (selectable in the mapping like an attribute) ─────

// Number of base units contained in a packing row (Aantal_CE).
function unitQty(u) { const n = parseFloat(u && u.Aantal_CE); return isNaN(n) ? 0 : n; }

function unitField(unit, field) {
  const val = unit && unit[field];
  return (val !== undefined && val !== null && val !== '') ? val : null;
}

// Ranks the units_by_type rows into packing slots WITHOUT relying on the per-product
// unit_used code (BOT/CS/BAG/KAR/EP… vary by product). Uses unit_type "CE" (consumer/base
// unit) plus the quantity hierarchy: slot 2 = inner box, 3 = outer carton, 4 = pallet,
// 5 = truck. The remaining rows are ranked largest-first: pallet, then outer carton, then
// inner box (truck only when 4+ levels exist).
function buildSlotMap(productValues) {
  const attrVals = productValues['units_by_type'];
  if (!Array.isArray(attrVals) || attrVals.length === 0) return {};
  const hit = attrVals.find(v => v.locale === null && v.scope === null) || attrVals[0];
  if (!hit || !Array.isArray(hit.data) || hit.data.length === 0) return {};
  const rows = hit.data.slice();
  let base = rows.find(u => String(u.unit_type || '').toUpperCase() === 'CE');
  let rest = base ? rows.filter(u => u !== base) : rows.filter(u => unitQty(u) > 1);
  if (!base) base = rows.slice().sort((a, b) => unitQty(a) - unitQty(b))[0] || null;
  rest = rest.slice().sort((a, b) => unitQty(b) - unitQty(a)); // largest first
  const order = rest.length >= 4 ? [5, 4, 3, 2] : [4, 3, 2];
  const map = { base };
  rest.forEach((u, i) => { if (order[i]) map[order[i]] = u; });
  return map;
}

const CALC_FIELDS = {
  base_unit: {
    label: 'Packing: base unit code (CE)',
    get: v => unitField(buildSlotMap(v).base, 'unit_used'),
  },
  min_order_qty: {
    label: 'Packing: min. order qty (outer carton, else pallet)',
    get: v => { const m = buildSlotMap(v); const c = unitField(m[3], 'Aantal_CE'); return c !== null ? c : unitField(m[4], 'Aantal_CE'); },
  },
  origin_from_category: {
    label: 'Country of origin from category (Birchmeier → DE, else BE)',
    get: v => (v.__categories || []).some(c => /birchmeier|berchemeier/i.test(String(c))) ? 'DE' : 'BE',
  },
};
const PACK_SLOTS = { 2: 'inner box', 3: 'outer carton', 4: 'pallet', 5: 'truck' };
const PACK_FIELDS = { unit: ['unit_used', 'unit code'], qty: ['Aantal_CE', 'base units'], ean: ['EAN', 'EAN'] };
for (const [slot, slotName] of Object.entries(PACK_SLOTS)) {
  for (const [key, [field, fieldName]] of Object.entries(PACK_FIELDS)) {
    CALC_FIELDS['pack_' + slot + '_' + key] = {
      label: 'Packing ' + slot + ' (' + slotName + '): ' + fieldName,
      get: v => unitField(buildSlotMap(v)[slot], field),
    };
  }
}

// The value a mapping entry produces for one product, or null to leave the cell alone.
function resolveValue(entry, values) {
  if (!entry) return null;
  if (entry.calc) return CALC_FIELDS[entry.calc] ? CALC_FIELDS[entry.calc].get(values) : null;
  const raw = values[entry.attr];
  const v = entry.locale ? extractLocalizedValue(raw, entry.locale) : extractValue(raw);
  if (entry.yesno) return toYesNo(v);  // flags are always written: missing = N
  return v;
}

// ── Akeneo ────────────────────────────────────────────────────────────

async function getAkeneoToken(env) {
  checkEnv(env, AKENEO_ENV);
  const credentials = Buffer.from(env.AKENEO_CLIENT_ID + ':' + env.AKENEO_CLIENT_SECRET).toString('base64');
  const res = await makeRequest(env.AKENEO_BASE_URL + '/api/oauth/v1/token', {
    method: 'POST',
    headers: { 'Authorization': 'Basic ' + credentials, 'Content-Type': 'application/json' },
    body: JSON.stringify({ grant_type: 'password', username: env.AKENEO_USERNAME, password: env.AKENEO_PASSWORD }),
  });
  if (!res.ok) throw new Error('Akeneo token failed: ' + res.status + ' ' + await res.text());
  const data = await res.json();
  if (!data.access_token) throw new Error('Akeneo OAuth response missing access_token');
  return data.access_token;
}

// Follows Akeneo's _links.next pagination and returns all items.
async function fetchAllPages(url, token, label) {
  const items = [];
  let nextUrl = url;
  for (let pages = 0; nextUrl && pages < 50; pages++) {
    const res = await makeRequest(nextUrl, { headers: { Authorization: 'Bearer ' + token } });
    if (!res.ok) throw new Error('Akeneo ' + label + ' error: ' + res.status + ' ' + await res.text());
    const data = await res.json();
    items.push(...((data._embedded && data._embedded.items) || []));
    nextUrl = (data._links && data._links.next && data._links.next.href) || null;
  }
  return items;
}

async function fetchAkeneoAttributes(env, token) {
  const items = await fetchAllPages(env.AKENEO_BASE_URL + '/api/rest/v1/attributes?limit=100', token, 'attributes');
  return items.map(a => ({ code: a.code, label: (a.labels && (a.labels.en_US || a.labels.en)) || a.code, type: a.type, localizable: !!a.localizable }));
}

async function fetchAkeneoLocales(env, token) {
  const search = encodeURIComponent(JSON.stringify({ enabled: [{ operator: '=', value: true }] }));
  const items = await fetchAllPages(env.AKENEO_BASE_URL + '/api/rest/v1/locales?limit=100&search=' + search, token, 'locales');
  return items.map(l => l.code);
}

// Sort variants smallest-first using any numeric quantity attribute,
// so the first row with a SKU = single unit, the second = box, etc.
function sortVariantsBySize(variants) {
  const qtyKeywords = ['qty', 'quantity', 'hoeveelheid', 'aantal', 'menge', 'quantite', 'umrez'];
  function getQty(variant) {
    for (const [code, vals] of Object.entries(variant.values || {})) {
      if (qtyKeywords.some(kw => code.toLowerCase().indexOf(kw) !== -1)) {
        const val = vals && vals[0] && vals[0].data;
        if (typeof val === 'number') return val;
        if (typeof val === 'string' && !isNaN(parseFloat(val))) return parseFloat(val);
      }
    }
    return 0;
  }
  return variants.slice().sort((a, b) => getQty(a) - getQty(b));
}

// Expand a single product into one view per units_by_type row so that
// repeated SKU rows in the Excel get the correct unit-tier EAN.
// `categories` is injected as a private __categories key for computed fields.
function buildUnitViews(productValues, categories) {
  const base = Object.assign({}, productValues, { __categories: categories || [] });
  const attrVals = base['units_by_type'];
  if (!Array.isArray(attrVals) || attrVals.length === 0) return [{ values: base }];
  const hit = attrVals.find(v => v.locale === null && v.scope === null) || attrVals[0];
  if (!hit || !Array.isArray(hit.data) || hit.data.length <= 1) return [{ values: base }];
  return hit.data.map(unitRow => {
    const override = unitRow.EAN ? { ean_ce: [{ locale: null, scope: null, data: String(unitRow.EAN) }] } : {};
    return { values: Object.assign({}, base, override) };
  });
}

async function fetchProductData(env, sku, token) {
  const headers = { Authorization: 'Bearer ' + token };
  const productRes = await makeRequest(env.AKENEO_BASE_URL + '/api/rest/v1/products/' + encodeURIComponent(sku), { headers });
  if (productRes.ok) { const p = await productRes.json(); return buildUnitViews(p.values || {}, p.categories || []); }
  if (productRes.status !== 404) throw new Error('Akeneo product error for ' + sku + ': ' + productRes.status);

  const modelRes = await makeRequest(env.AKENEO_BASE_URL + '/api/rest/v1/product-models/' + encodeURIComponent(sku), { headers });
  if (!modelRes.ok) { console.warn('Not found in Akeneo: ' + sku); return []; }

  const model = await modelRes.json();
  const modelValues = model.values || {};
  const modelCats = model.categories || [];
  const search = encodeURIComponent(JSON.stringify({ parent: [{ operator: '=', value: sku }] }));
  let variants;
  try {
    variants = await fetchAllPages(env.AKENEO_BASE_URL + '/api/rest/v1/products?search=' + search + '&limit=100', token, 'variants');
  } catch (e) {
    console.warn(e.message);
    variants = [];
  }
  if (variants.length === 0) return [{ values: Object.assign({}, modelValues, { __categories: modelCats }) }];
  return sortVariantsBySize(variants.map(v => ({
    values: Object.assign({}, modelValues, v.values || {}, { __categories: (v.categories && v.categories.length ? v.categories : modelCats) }),
  })));
}

// ── Mapping: validation, AI suggestion, storage ───────────────────────

// Portal input is untrusted: keep only well-formed entries for headers that exist.
// The SKU column is the lookup key, so it is never written.
function cleanMapping(mapping, headerCols, skuCol) {
  const out = {};
  for (const { colNum } of headerCols) {
    const e = colNum !== skuCol && mapping && mapping[colNum];
    if (e && typeof e.calc === 'string' && CALC_FIELDS[e.calc]) out[colNum] = { calc: e.calc };
    else if (e && typeof e.attr === 'string' && e.attr) {
      out[colNum] = { attr: e.attr };
      if (typeof e.locale === 'string' && e.locale) out[colNum].locale = e.locale;
      if (e.yesno === true) out[colNum].yesno = true;
    } else out[colNum] = null;
  }
  return out;
}

// What OpenAI sees of the template: header texts and hints only, never values from the data rows.
function buildSuggestColumns(headerCols, skuCol) {
  return headerCols.filter(h => h.colNum !== skuCol).map(h => {
    const c = { col: h.colNum, name: h.header };
    if (h.hint) c.hint = h.hint.replace(/\n/g, ' | ');
    return c;
  });
}

// First draft for a template we have not seen before. A person reviews it in the portal.
async function suggestMapping(env, headerCols, skuCol, attributes, locales) {
  checkEnv(env, ['OPENAI_API_KEY']);
  const columns = buildSuggestColumns(headerCols, skuCol);
  const calcFields = Object.entries(CALC_FIELDS).map(([code, f]) => ({ code, label: f.label }));
  const systemPrompt = 'You map the columns of a supplier Excel template to product data from Akeneo PIM.\n' +
    'Map as many columns as you reasonably can: descriptions, commodity group, product hierarchy, supplier, brand, material, colour, classifications, weights, dimensions, packing levels, flags. Leave a column null only when nothing plausibly fits.\n' +
    'For every column return exactly one of:\n' +
    '- {"attr": "<attribute code>"} — an attribute from akeneo_attributes. For localizable attributes add "locale" (one of locales) matching the column\'s language, e.g. a column suffixed _NL or "omschrijving NL" → a nl_* locale.\n' +
    '- {"attr": "<attribute code>", "yesno": true} — for Y/N flag columns (e.g. recycling fee / packaging-tax flags), mapped to the matching boolean or flag attribute.\n' +
    '- {"calc": "<code>"} — one of calc_fields, for packing-unit columns (base unit, units per box/carton/pallet, EAN per packing level, minimum order quantity) and country of origin.\n' +
    '- null — when nothing fits, and ALWAYS for: prices of any kind, date columns (valid-from, ddmmjjjj, datum), and identifiers the customer assigns themselves (their own internal article number).\n' +
    '"hint" (when present) is the template\'s own description rows (often the same text in several languages); it describes expected values or format. ' +
    'The column "name" is usually a SAP field code (e.g. MATNR, EAN11, MEINH, UMREZ) and is the strongest signal: when the hint contradicts the name, trust the name — templates contain copy-paste mistakes in their descriptions. ' +
    'Paired SAP quantity fields: UMREZ is the number of base units in a packing level; UMREN is its denominator (normally 1), so never map UMREN to a quantity. Code lists like "PC = PIECE | INB = BOX" mean a select/text attribute, never a metric. Only use codes from the lists — never invent one.\n' +
    'Return ONLY JSON keyed by each column\'s "col" number: { "<col>": <entry or null>, ... } covering every column.';
  const res = await makeRequest('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + env.OPENAI_API_KEY, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'gpt-4o',
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: JSON.stringify({ excel_columns: columns, akeneo_attributes: attributes, calc_fields: calcFields, locales }) },
      ],
      response_format: { type: 'json_object' },
      temperature: 0,
    }),
  });
  if (!res.ok) throw new Error('OpenAI error: ' + res.status + ' ' + await res.text());
  const data = await res.json();
  const content = data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
  if (!content) throw new Error('OpenAI returned no content');
  let raw;
  try { raw = JSON.parse(content); } catch (e) { throw new Error('OpenAI returned non-JSON: ' + content.substring(0, 200)); }
  // Drop anything that points at a code that does not exist.
  const byCode = new Map(attributes.map(a => [a.code, a]));
  const mapping = cleanMapping(raw, headerCols, skuCol);
  for (const [col, e] of Object.entries(mapping)) {
    if (!e || e.calc) continue;
    const attr = byCode.get(e.attr);
    if (!attr) mapping[col] = null;
    else if (!attr.localizable || locales.indexOf(e.locale) === -1) delete e.locale;
  }
  return mapping;
}

// Same headers in the same order = same template.
function headerFingerprint(headerCols) {
  return crypto.createHash('sha256').update(JSON.stringify(headerCols.map(h => h.header))).digest('hex');
}

async function withDb(env, fn) {
  checkEnv(env, ['MAPPINGS_DB_URL']);
  const client = new Client({ connectionString: env.MAPPINGS_DB_URL });
  await client.connect();
  try {
    await client.query(`CREATE TABLE IF NOT EXISTS column_mappings (
      fingerprint TEXT PRIMARY KEY,
      headers     JSONB NOT NULL,
      sku_col     INTEGER,
      mapping     JSONB NOT NULL,
      file_name   TEXT,
      created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
      updated_at  TIMESTAMPTZ NOT NULL DEFAULT now()
    )`);
    return await fn(client);
  } finally {
    await client.end();
  }
}

async function loadMapping(env, fingerprint) {
  return withDb(env, async db => {
    const { rows } = await db.query('SELECT sku_col, mapping, updated_at FROM column_mappings WHERE fingerprint = $1', [fingerprint]);
    return rows[0] ? { skuCol: rows[0].sku_col, mapping: rows[0].mapping, updatedAt: rows[0].updated_at } : null;
  });
}

async function saveMapping(env, fingerprint, headerCols, choice, fileName) {
  await withDb(env, db => db.query(
    `INSERT INTO column_mappings (fingerprint, headers, sku_col, mapping, file_name)
     VALUES ($1, $2, $3, $4, $5)
     ON CONFLICT (fingerprint) DO UPDATE
       SET sku_col = EXCLUDED.sku_col, mapping = EXCLUDED.mapping, file_name = EXCLUDED.file_name, updated_at = now()`,
    [fingerprint, JSON.stringify(headerCols.map(h => h.header)), choice.skuCol, JSON.stringify(choice.mapping), fileName || null]));
}

// ── Extract: find header row, hint row, SKU column and data rows ──────

const SKIP_ROW_PATS = [/\bvoorbeeld\b/i, /\bexemple\b/i, /\bexample\b/i, /\bcontrole\b/i, /\bonderliggende\b/i];

const SKU_COL_PATS = [
  /^sku$/i,
  /\bartikel(?:nr|num(?:mer)?|ref(?:erentie)?|code)\b/i,
  /\bref\.?\s*(?:leverancier|article|art\.?|fournisseur)\b/i,
  /^EINA-IDNLF$/i, /^MARA-MATNR$/i,
  /\barticle\s*(?:nr|num(?:ber)?|no\.?|code|ref)\b/i,
  /\bproduct\s*(?:code|num(?:ber)?|ref)\b/i,
  /\bitem\s*(?:code|num(?:ber)?|ref)\b/i,
];

function rowIsSkip(texts) {
  return SKIP_ROW_PATS.some(p => texts.slice(0, 3).some(t => p.test(t)));
}

function isBarrierRow(ri) {
  const col1 = ri.cells.find(x => x.c === 1);
  return !!(col1 && SKIP_ROW_PATS.some(p => p.test(col1.t)) && ri.filled >= 5);
}

function sapCodeScore(texts) {
  if (!texts.length) return 0;
  const n = texts.filter(t => /^[A-Z][A-Z0-9]{1,9}(-[A-Z][A-Z0-9]{1,9})?(\([A-Z0-9_()+]+\))?$/.test(t)).length;
  return n / texts.length;
}

// Data rows with a value in the SKU column, skipping example/check rows.
function collectRows(worksheet, layout, skuColNum) {
  const rows = [];
  for (let r = layout.dataStartRow; r <= worksheet.rowCount; r++) {
    const firstCells = [];
    for (let c = 1; c <= Math.min(layout.maxCols, 5); c++) { const t = getCellText(worksheet.getRow(r).getCell(c).value); if (t) firstCells.push(t); }
    if (rowIsSkip(firstCells)) continue;
    const sku = getCellText(worksheet.getRow(r).getCell(skuColNum).value);
    if (sku === null) continue;
    rows.push({ sku, rowNum: r });
  }
  return rows;
}

function extractHeadersAndRows(worksheet) {
  const totalRows = worksheet.rowCount;

  let maxCols = 0;
  for (let r = 1; r <= Math.min(50, totalRows); r++) {
    for (let c = 200; c >= 1; c--) {
      if (getCellText(worksheet.getRow(r).getCell(c).value) !== null) { if (c > maxCols) maxCols = c; break; }
    }
  }
  if (maxCols === 0) throw new Error('Excel sheet appears empty in first 50 rows');

  const rowInfo = [];
  for (let r = 1; r <= Math.min(100, totalRows); r++) {
    const cells = [];
    for (let c = 1; c <= maxCols; c++) {
      const t = getCellText(worksheet.getRow(r).getCell(c).value);
      if (t !== null) cells.push({ c, t });
    }
    const texts = cells.map(x => x.t);
    rowInfo.push({ r, filled: cells.length, ratio: cells.length / maxCols, cells, texts, isSkip: rowIsSkip(texts) });
  }

  let lastBarrierRow = 0;
  for (const ri of rowInfo) {
    if (ri.r > 50) break;
    if (isBarrierRow(ri)) lastBarrierRow = ri.r;
  }

  let dataStartRow;
  if (lastBarrierRow > 0) {
    dataStartRow = lastBarrierRow + 1;
    while (dataStartRow <= totalRows) {
      const ri = rowInfo.find(x => x.r === dataStartRow);
      if (!ri || ri.filled > 0) break;
      dataStartRow++;
    }
  } else {
    const firstDense = rowInfo.find(ri => ri.ratio >= 0.5)
      || rowInfo.find(ri => ri.ratio >= 0.3)
      || rowInfo.find(ri => ri.filled > 0);
    if (!firstDense) throw new Error('No header row detected');
    const threshold = Math.max(0.3, firstDense.ratio * 0.8);
    let lastCluster = firstDense.r;
    dataStartRow = firstDense.r + 1;
    for (const ri of rowInfo) {
      if (ri.r <= firstDense.r || ri.r > firstDense.r + 20) continue;
      if (ri.ratio > threshold) { lastCluster = ri.r; }
      else if (ri.filled > 0 && getCellText(worksheet.getRow(ri.r).getCell(1).value) !== null) { dataStartRow = ri.r; break; }
    }
    if (dataStartRow === firstDense.r + 1 && lastCluster > firstDense.r) {
      dataStartRow = lastCluster + 1;
      for (const ri of rowInfo) {
        if (ri.r >= dataStartRow && ri.r <= dataStartRow + 5 && ri.filled > 0) { dataStartRow = ri.r; break; }
      }
    }
  }

  let headerRegion = rowInfo.filter(ri => ri.r < dataStartRow && !ri.isSkip && ri.filled >= 2);
  if (headerRegion.length === 0) headerRegion = rowInfo.filter(ri => ri.r < dataStartRow && ri.filled >= 1);
  if (headerRegion.length === 0) throw new Error('No header row found (dataStartRow=' + dataStartRow + ', totalRows=' + totalRows + ')');
  const bestHeader = headerRegion.reduce((best, ri) => {
    const s = sapCodeScore(ri.texts), bs = sapCodeScore(best.texts);
    return (s > bs || (s === bs && ri.filled > best.filled)) ? ri : best;
  }, headerRegion[0]);
  const headerRowIndex = bestHeader.r;

  // Every well-filled row above the header (section titles, explanations, translations) becomes
  // part of the column's hint, top to bottom. Mostly-empty rows (branding, address) are skipped,
  // and so are rows below the header: those hold example values, not descriptions.
  let hintRows = headerRegion.filter(ri => ri.r < headerRowIndex && ri.ratio >= 0.3);
  if (hintRows.length === 0) hintRows = rowInfo.filter(ri => ri.r === headerRowIndex - 1 && !ri.isSkip);
  const hintMap = {};
  for (const ri of hintRows) {
    for (const x of ri.cells) {
      const lines = hintMap[x.c] || (hintMap[x.c] = []);
      if (lines[lines.length - 1] !== x.t) lines.push(x.t);  // merged cells repeat the same text
    }
  }

  const headerCols = bestHeader.cells.map(x => {
    const entry = { header: x.t, colNum: x.c };
    if (hintMap[x.c]) entry.hint = hintMap[x.c].join('\n');
    return entry;
  });

  // SKU column = first header (or hint) that names it. No guessing beyond that:
  // when nothing matches it stays null and the client picks it in the portal.
  const skuHeader = headerCols.find(col => [col.header, col.hint || ''].some(t => SKU_COL_PATS.some(p => p.test(t))));
  const skuColNum = skuHeader ? skuHeader.colNum : null;

  const layout = { headerCols, headerRowIndex, dataStartRow, maxCols, skuColNum };
  layout.rows = skuColNum === null ? [] : collectRows(worksheet, layout, skuColNum);
  return layout;
}

// ── Pipelines used by the n8n nodes and the CLI ──────────────────────

// Reads the uploaded .xlsx from an n8n webhook item.
function readUpload(item) {
  const binary = item.binary || {};
  const keys = Object.keys(binary);
  if (keys.length === 0) throw new Error('No file received. Upload with: curl -X POST <url> -F "data=@file.xlsx"');
  const meta = binary['data'] || binary[keys[0]];
  if (!meta || typeof meta.data !== 'string') throw new Error('Uploaded file has unexpected structure: ' + JSON.stringify(meta).substring(0, 200));
  if (meta.data.length < 100 && /^[a-f0-9-]+$/i.test(meta.data.trim())) {
    throw new Error('n8n is running in filesystem binary mode. Add N8N_BINARY_DATA_MODE=default AND N8N_DEFAULT_BINARY_DATA_MODE=default to docker-compose and restart.');
  }
  return { buffer: Buffer.from(meta.data, 'base64'), fileName: meta.fileName || 'enriched.xlsx' };
}

async function loadWorkbook(buffer) {
  if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4B) throw new Error('Not a valid .xlsx file. .xls is not supported — save as .xlsx first.');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

// Everything the portal needs to show the mapping screen.
async function analyzeUpload(buffer, fileName, env) {
  const workbook = await loadWorkbook(buffer);
  const layout = extractHeadersAndRows(workbook.worksheets[0]);
  const fingerprint = headerFingerprint(layout.headerCols);
  const token = await getAkeneoToken(env);
  const [attributes, locales, saved] = await Promise.all([fetchAkeneoAttributes(env, token), fetchAkeneoLocales(env, token), loadMapping(env, fingerprint)]);
  const skuCol = saved && saved.skuCol ? saved.skuCol : layout.skuColNum;
  const samples = {};
  const firstRow = layout.rows[0] && workbook.worksheets[0].getRow(layout.rows[0].rowNum);
  if (firstRow) layout.headerCols.forEach(h => { samples[h.colNum] = getCellText(firstRow.getCell(h.colNum).value); });
  return {
    fileName,
    fingerprint,
    headerRowIndex: layout.headerRowIndex,
    dataStartRow: layout.dataStartRow,
    rowCount: layout.rows.length,
    columns: layout.headerCols,
    saved: !!saved,
    savedAt: saved ? saved.updatedAt : null,
    skuCol,
    samples,
    mapping: saved ? cleanMapping(saved.mapping, layout.headerCols, skuCol) : await suggestMapping(env, layout.headerCols, skuCol, attributes, locales),
    options: { attributes, locales, calcFields: Object.entries(CALC_FIELDS).map(([code, f]) => ({ code, label: f.label })) },
  };
}

// Fills the first worksheet. `choice` = { skuCol, mapping } from the portal (saved to the DB
// first), or null to use the mapping already saved for this template.
async function enrichWorkbook(workbook, choice, fileName, env, log) {
  log = log || (() => {});
  const worksheet = workbook.worksheets[0];
  const layout = extractHeadersAndRows(worksheet);
  const fingerprint = headerFingerprint(layout.headerCols);

  const noSkuCol = 'No SKU column chosen — pick the article number column in the portal first.';
  if (choice) {
    const skuCol = layout.headerCols.some(h => h.colNum === Number(choice.skuCol)) ? Number(choice.skuCol) : layout.skuColNum;
    if (skuCol === null) throw new Error(noSkuCol);
    choice = { skuCol, mapping: cleanMapping(choice.mapping, layout.headerCols, skuCol) };
    await saveMapping(env, fingerprint, layout.headerCols, choice, fileName);
  } else {
    choice = await loadMapping(env, fingerprint);
    if (!choice) throw new Error('No saved column mapping for this template yet — open it in the portal and save a mapping first.');
    if (!choice.skuCol) throw new Error(noSkuCol);
  }

  const rows = choice.skuCol !== layout.skuColNum ? collectRows(worksheet, layout, choice.skuCol) : layout.rows;
  log('Header row ' + layout.headerRowIndex + ', data from row ' + layout.dataStartRow + ', ' + rows.length + ' rows');

  const token = await getAkeneoToken(env);
  const skuMap = new Map();
  for (const row of rows) {
    if (!skuMap.has(row.sku)) skuMap.set(row.sku, []);
    skuMap.get(row.sku).push(row.rowNum);
  }
  for (const [sku, rowNumbers] of skuMap) {
    const variantList = await fetchProductData(env, sku, token);
    log(sku + ': ' + variantList.length + ' variant(s)');
    if (variantList.length === 0) continue;
    rowNumbers.forEach((rowNum, i) => {
      const values = variantList[i % variantList.length].values;
      for (const { colNum } of layout.headerCols) {
        const value = resolveValue(choice.mapping[colNum], values);
        if (value !== null && value !== undefined) worksheet.getRow(rowNum).getCell(colNum).value = value;
      }
    });
  }
  prepareForWrite(workbook);
}

// Conditional-format rule types ExcelJS can write (cf-rule-xform render()). Any other type, e.g.
// "duplicateValues", is written as an empty <conditionalFormatting/> that Excel has to repair.
const WRITABLE_CF_TYPES = ['expression', 'cellIs', 'top10', 'aboveAverage', 'dataBar', 'colorScale', 'iconSet', 'containsText', 'timePeriod'];

// ExcelJS cannot write some things it reads; strip them so writeBuffer() does not throw
// and Excel opens the result without a repair prompt.
function prepareForWrite(workbook) {
  workbook.worksheets.forEach(ws => {
    // 'Shared Formula master must exist above and or left of clone'. Clone cells are
    // { sharedFormula, result } with no formula key — keep their result as a static value.
    ws.eachRow({ includeEmpty: false }, row => {
      row.eachCell({ includeEmpty: false }, cell => {
        const v = cell.value;
        if (v && typeof v === 'object') {
          if ('sharedFormula' in v) cell.value = (v.result !== undefined && v.result !== null) ? v.result : null;
          else if ('formula' in v) cell.value = { formula: v.formula, result: v.result };
        }
      });
    });
    // Drop rules ExcelJS can't serialize: unsupported types, and x14 extended rules (formula types without formulae).
    if (ws.conditionalFormattings) {
      ws.conditionalFormattings = ws.conditionalFormattings
        .map(cf => {
          cf.rules = (cf.rules || []).filter(r => WRITABLE_CF_TYPES.includes(r.type)
            && (r.formulae !== undefined || ['expression', 'cellIs', 'containsText', 'timePeriod'].indexOf(r.type) === -1));
          return cf;
        })
        .filter(cf => cf.rules.length > 0);
    }
  });
}

// @cli-only
module.exports = { loadWorkbook, extractHeadersAndRows, analyzeUpload, enrichWorkbook, readUpload, resolveValue, cleanMapping, headerFingerprint, buildSlotMap, buildSuggestColumns, prepareForWrite, CALC_FIELDS };
