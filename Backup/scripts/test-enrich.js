// Run: AKENEO_BASE_URL=... AKENEO_CLIENT_ID=... AKENEO_CLIENT_SECRET=... AKENEO_USERNAME=... AKENEO_PASSWORD=... OPENAI_API_KEY=... node scripts/test-enrich.js path/to/file.xlsx
const ExcelJS = require('exceljs');
const fs = require('fs');

const AKENEO_BASE_URL = process.env.AKENEO_BASE_URL;
const AKENEO_CLIENT_ID = process.env.AKENEO_CLIENT_ID;
const AKENEO_CLIENT_SECRET = process.env.AKENEO_CLIENT_SECRET;
const AKENEO_USERNAME = process.env.AKENEO_USERNAME;
const AKENEO_PASSWORD = process.env.AKENEO_PASSWORD;
const OPENAI_API_KEY = process.env.OPENAI_API_KEY;

// Columns that must never be written by the AI
const BLOCKED_HEADERS = new Set([
  'MATNR',            // Col D: SAP article number — customer fills this in
  'MEAN-EAN11_1',     // Col H-L: no per-language extra EAN codes for DCM
  'MEAN-EAN11_2',
  'MEAN-EAN11_3',
  'MEAN-EAN11_4',
  'MEAN-EAN11_5',
  'Purchasing price', // Col DU: never set prices
  'STPRS',            // Col EH: standard price
  'Sales price 1',    // Col DY
  'Sales price 2',    // Col DZ
  'Sales price 3',    // Col EA
  'VKP0_S600',        // Col EB-EF: sales prices per location
  'VKP0_S601',
  'VKP0_S602',
  'VKP0_S603',
  'VKP0_S604',
]);

// Locale to use when extracting each MAKTX commercial description column
const MAKTX_LOCALES = {
  'MAKTX_NL': 'nl_BE',
  'MAKTX_FR': 'fr_BE',
  'MAKTX_EN': 'en_GB',
  'MAKTX_DE': 'de_DE',
};

// MAKTX columns always read from 'title' attribute (commercial catalog name), never from OpenAI mapping.
const MAKTX_HARDCODED_ATTR = 'title';

// ── helpers ──────────────────────────────────────────────────────────

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

// ── Packing unit helpers ─────────────────────────────────────────────

// Maps a unit_used code to the correct Floralux packing slot:
//   2 = inner carton (CK-CO), 3 = outer carton (CP-CT), 4 = pallet (CU-CY), 5 = truck
// DCM products have no inner carton, so outer carton starts at slot 3.
function getUnitSlot(unitType) {
  if (!unitType) return null;
  const ut = String(unitType).toUpperCase().trim();
  if (/^(EP|PAL|PLT|PALLET)/.test(ut)) return 4;
  if (/^(KAR|OD|BOX|CASE|KARTON|OMDOOS|CS|OUT)/.test(ut)) return 3;
  if (/^(ZUM|TRK|TRUCK|LKW)/.test(ut)) return 5;
  // Base-unit types (ST, BOT, PC, etc.) and anything unrecognised get no packing slot.
  return null;
}

// Returns a field from the packing unit that belongs to the given slot number.
function extractUnitBySlot(productValues, slotNum, fieldName) {
  const attrVals = productValues['units_by_type'];
  if (!Array.isArray(attrVals) || attrVals.length === 0) return null;
  const hit = attrVals.find(v => v.locale === null && v.scope === null) || attrVals[0];
  if (!hit || !Array.isArray(hit.data) || hit.data.length === 0) return null;
  for (const unit of hit.data) {
    if (getUnitSlot(unit.unit_used) === slotNum) {
      const val = unit[fieldName];
      return (val !== undefined && val !== null && val !== '') ? val : null;
    }
  }
  return null;
}

// Returns undefined when no hardcoded rule matches (caller falls through to OpenAI mapping).
// Returns null when a rule matches but no data exists (caller skips writing).
function getHardcodedValue(header, hint, productValues) {
  const candidates = [header, hint || ''];
  for (const name of candidates) {
    let m;
    if ((m = /\bMEINH_(\d+)\b/i.exec(name))) return extractUnitBySlot(productValues, +m[1], 'unit_used');
    if ((m = /\bUMREZ_(\d+)\b/i.exec(name))) return extractUnitBySlot(productValues, +m[1], 'Aantal_CE');
    if ((m = /\bEAN11_(\d+)\b/i.exec(name)))  return extractUnitBySlot(productValues, +m[1], 'EAN');
    // Min order quantity = base units per outer carton
    if (/\bEINE-MINBM\b/i.test(name)) return extractUnitBySlot(productValues, 3, 'Aantal_CE');
  }
  return undefined;
}

// Expand a single product into one view per units_by_type row so that
// repeated SKU rows in the Excel get the correct unit-tier EAN.
function buildUnitViews(productValues) {
  const attrVals = productValues['units_by_type'];
  if (!Array.isArray(attrVals) || attrVals.length === 0) return [{ values: productValues }];
  const hit = attrVals.find(v => v.locale === null && v.scope === null) || attrVals[0];
  if (!hit || !Array.isArray(hit.data) || hit.data.length <= 1) return [{ values: productValues }];
  return hit.data.map(unitRow => {
    const override = unitRow.EAN
      ? { ean_ce: [{ locale: null, scope: null, data: String(unitRow.EAN) }] }
      : {};
    return { values: { ...productValues, ...override } };
  });
}

// ── Akeneo ───────────────────────────────────────────────────────────

async function getAkeneoToken() {
  const credentials = Buffer.from(`${AKENEO_CLIENT_ID}:${AKENEO_CLIENT_SECRET}`).toString('base64');
  const res = await fetch(`${AKENEO_BASE_URL}/api/oauth/v1/token`, {
    method: 'POST',
    headers: { 'Authorization': `Basic ${credentials}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ grant_type: 'password', username: AKENEO_USERNAME, password: AKENEO_PASSWORD })
  });
  if (!res.ok) {
    const body = await res.text();
    console.error('Akeneo token error body:', body);
    throw new Error(`Akeneo OAuth token request failed: ${res.status}`);
  }
  const data = await res.json();
  if (!data.access_token) throw new Error('Akeneo OAuth response missing access_token');
  return data.access_token;
}

async function fetchAkeneoAttributes(token) {
  const attributes = [];
  let nextUrl = `${AKENEO_BASE_URL}/api/rest/v1/attributes?limit=100&with_count=false`;
  let pages = 0;
  while (nextUrl && pages < 50) {
    const res = await fetch(nextUrl, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
      const body = await res.text();
      console.error('Akeneo attributes error body:', body);
      throw new Error(`Akeneo attributes error: ${res.status}`);
    }
    const data = await res.json();
    (data._embedded?.items || []).forEach(a => attributes.push({
      code: a.code,
      label_en: a.labels?.en_US || a.labels?.en || a.code,
      type: a.type
    }));
    nextUrl = data._links?.next?.href || null;
    pages++;
  }
  return attributes;
}

async function fetchProductData(sku, token) {
  const headers = { Authorization: `Bearer ${token}` };
  const productRes = await fetch(`${AKENEO_BASE_URL}/api/rest/v1/products/${encodeURIComponent(sku)}`, { headers });
  if (productRes.ok) {
    const p = await productRes.json();
    return buildUnitViews(p.values || {});
  }
  if (productRes.status !== 404) throw new Error(`Akeneo product error for ${sku}: ${productRes.status}`);

  const modelRes = await fetch(`${AKENEO_BASE_URL}/api/rest/v1/product-models/${encodeURIComponent(sku)}`, { headers });
  if (!modelRes.ok) {
    console.warn(`Akeneo product-model not found for "${sku}": ${modelRes.status}`);
    return [];
  }

  const model = await modelRes.json();
  const modelValues = model.values || {};
  const search = encodeURIComponent(JSON.stringify({ parent: [{ operator: '=', value: sku }] }));
  const varRes = await fetch(`${AKENEO_BASE_URL}/api/rest/v1/products?search=${search}&limit=25`, { headers });
  if (!varRes.ok) {
    console.warn(`Akeneo variants fetch failed for model "${sku}": ${varRes.status}`);
    return [{ values: modelValues }];
  }

  const varData = await varRes.json();
  const variants = varData._embedded?.items || [];
  if (variants.length === 0) return [{ values: modelValues }];
  return variants.map(v => ({ values: { ...modelValues, ...(v.values || {}) } }));
}

// ── OpenAI ───────────────────────────────────────────────────────────

async function getColumnMapping(headerCols, attributes) {
  const columns = headerCols.map(h => h.hint
    ? { name: h.header, hint: h.hint.replace(/\n/g, ' | ') }
    : { name: h.header }
  );
  const systemPrompt =
    'You are a data mapping expert. Map Excel column headers to Akeneo attribute codes.\n' +
    'Rules:\n' +
    '- The "hint" field (when present) describes the expected values or format for that column.\n' +
    '- If the hint lists code options (e.g. "PC = PIECE | INB = BOX"), map to a simpleselect or text attribute — NEVER a metric attribute.\n' +
    '- If the hint is a date format (e.g. "ddmmjjjj"), map to a date attribute.\n' +
    '- If the hint is a numeric format (e.g. "max 3 decimals"), map to a decimal or metric attribute.\n' +
    '- For MAKTX_NL, MAKTX_FR, MAKTX_EN, MAKTX_DE: map to the COMMERCIAL product name (catalog/marketing name, not SAP internal name). Prefer attributes whose code contains "name" or "description" and that have locale-specific values (localizable=true).\n' +
    '- NEVER map any column to a price attribute (purchase price, sales price, standard price, recommended price, prijs, prix, stprs, einkaufspreis). Return null for any price-related column.\n' +
    'Return ONLY valid JSON: { "<column_name>": "<akeneo_attribute_code_or_null>" } for every column in the input.';
  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: { Authorization: `Bearer ${OPENAI_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'gpt-4o',
      messages: [
        { role: 'system', content: systemPrompt },
        {
          role: 'user',
          content: JSON.stringify({
            excel_columns: columns,
            akeneo_attributes: attributes.slice(0, 300).map(a => ({ code: a.code, label: a.label_en, type: a.type }))
          })
        }
      ],
      response_format: { type: 'json_object' },
      temperature: 0
    })
  });
  if (!res.ok) {
    const oaiBody = await res.text();
    console.error('OpenAI error body:', oaiBody);
    throw new Error(`OpenAI error: ${res.status}`);
  }
  const data = await res.json();
  const content = data.choices?.[0]?.message?.content;
  if (!content) throw new Error('OpenAI returned no content in response');
  try { return JSON.parse(content); }
  catch { throw new Error('OpenAI returned non-JSON content: ' + content.substring(0, 200)); }
}

// ── getCellText ───────────────────────────────────────────────────────

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

// ── Extract ───────────────────────────────────────────────────────────

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
  return col1 && SKIP_ROW_PATS.some(p => p.test(col1.t)) && ri.filled >= 5;
}

function sapCodeScore(texts) {
  if (!texts.length) return 0;
  const n = texts.filter(t => /^[A-Z][A-Z0-9]{1,9}(-[A-Z][A-Z0-9]{1,9})?(\([A-Z0-9_()+]+\))?$/.test(t)).length;
  return n / texts.length;
}

async function extractHeadersAndRows(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.worksheets[0];
  const totalRows = worksheet.rowCount;

  let maxCols = 0;
  for (let r = 1; r <= Math.min(50, totalRows); r++) {
    for (let c = 200; c >= 1; c--) {
      if (getCellText(worksheet.getRow(r).getCell(c).value) !== null) { if (c > maxCols) maxCols = c; break; }
    }
  }
  if (maxCols === 0) throw new Error('Empty sheet');

  const rowInfo = [];
  for (let r = 1; r <= Math.min(100, totalRows); r++) {
    const cells = [];
    let filled = 0;
    for (let c = 1; c <= maxCols; c++) {
      const t = getCellText(worksheet.getRow(r).getCell(c).value);
      if (t !== null) { filled++; cells.push({ c, t }); }
    }
    const texts = cells.map(x => x.t);
    rowInfo.push({ r, filled, ratio: filled / maxCols, cells, texts, isSkip: rowIsSkip(texts) });
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
    if (!firstDense) throw new Error('No header row found');
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
  if (headerRegion.length === 0) throw new Error('No header row found (dataStartRow=' + dataStartRow + ')');
  const bestHeader = headerRegion.reduce((best, ri) => {
    const s = sapCodeScore(ri.texts), bs = sapCodeScore(best.texts);
    return (s > bs || (s === bs && ri.filled > best.filled)) ? ri : best;
  }, headerRegion[0]);
  const headerRowIndex = bestHeader.r;

  const hintMap = {};
  const firstDenseInRegion = headerRegion.find(ri => ri.ratio >= 0.5 && ri.r !== headerRowIndex);
  const hintSource = firstDenseInRegion || rowInfo.find(ri => ri.r === headerRowIndex - 1 && !ri.isSkip);
  if (hintSource) {
    for (const { c, t } of hintSource.cells) hintMap[c] = t;
  }

  const headerCols = bestHeader.cells.map(({ c, t }) => {
    const entry = { header: t, colNum: c };
    if (hintMap[c]) entry.hint = hintMap[c];
    return entry;
  });

  let skuColNum = null;
  for (const col of headerCols) {
    if ([col.header, col.hint || ''].some(t => SKU_COL_PATS.some(p => p.test(t)))) { skuColNum = col.colNum; break; }
  }
  if (skuColNum === null) {
    const sample = rowInfo.filter(ri => ri.r >= dataStartRow && !ri.isSkip).slice(0, 20);
    let best = -1;
    for (const col of headerCols) {
      const vals = sample.map(ri => { const x = ri.cells.find(c => c.c === col.colNum); return x ? x.t : null; }).filter(v => v !== null);
      if (!vals.length) continue;
      const unique = new Set(vals).size;
      const alnum  = vals.filter(v => /^[A-Za-z0-9_\-.]+$/.test(v) && v.length <= 20).length;
      const score  = (unique / vals.length) * (alnum / vals.length) * (vals.length / Math.max(sample.length, 1));
      if (score > best) { best = score; skuColNum = col.colNum; }
    }
  }
  if (skuColNum === null) skuColNum = 1;

  const rows = [];
  for (let r = dataStartRow; r <= totalRows; r++) {
    const firstCells = [];
    for (let c = 1; c <= Math.min(maxCols, 5); c++) {
      const t = getCellText(worksheet.getRow(r).getCell(c).value);
      if (t) firstCells.push(t);
    }
    if (rowIsSkip(firstCells)) continue;
    const sku = getCellText(worksheet.getRow(r).getCell(skuColNum).value);
    if (sku === null) continue;
    rows.push({ sku, rowNum: r });
  }

  return { workbook, worksheet, headerCols, headerRowIndex, dataStartRow, rows };
}

// ── Main ─────────────────────────────────────────────────────────────

async function main() {
  const filePath = process.argv[2];
  if (!filePath) { console.error('Usage: node scripts/test-enrich.js <file.xlsx>'); process.exit(1); }

  if (!AKENEO_BASE_URL)     throw new Error('Missing env var: AKENEO_BASE_URL');
  if (!AKENEO_CLIENT_ID)    throw new Error('Missing env var: AKENEO_CLIENT_ID');
  if (!AKENEO_CLIENT_SECRET) throw new Error('Missing env var: AKENEO_CLIENT_SECRET');
  if (!AKENEO_USERNAME)     throw new Error('Missing env var: AKENEO_USERNAME');
  if (!AKENEO_PASSWORD)     throw new Error('Missing env var: AKENEO_PASSWORD');
  if (!OPENAI_API_KEY)      throw new Error('Missing env var: OPENAI_API_KEY');

  console.log('Authenticating with Akeneo...');
  const akeneoToken = await getAkeneoToken();
  console.log('Akeneo authentication successful');

  const buffer = fs.readFileSync(filePath);
  console.log('Extracting headers and rows...');
  const { workbook, worksheet, headerCols, rows } = await extractHeadersAndRows(buffer);

  console.log(`Found ${headerCols.length} headers, ${rows.length} data rows`);

  console.log('Fetching Akeneo attributes...');
  const attributes = await fetchAkeneoAttributes(akeneoToken);
  console.log(`Fetched ${attributes.length} Akeneo attributes`);

  console.log('Calling OpenAI for column mapping...');
  const columnMapping = await getColumnMapping(headerCols, attributes);
  console.log('Column mapping:', JSON.stringify(columnMapping, null, 2));

  const skuMap = new Map();
  for (const row of rows) {
    if (!skuMap.has(row.sku)) skuMap.set(row.sku, []);
    skuMap.get(row.sku).push(row.rowNum);
  }

  console.log(`Processing ${skuMap.size} unique SKUs...`);

  for (const [sku, rowNumbers] of skuMap) {
    console.log(`  Fetching ${sku}...`);
    const variantList = await fetchProductData(sku, akeneoToken);
    console.log(`    Got ${variantList.length} variant(s)`);

    if (variantList.length === 0) continue;

    for (let i = 0; i < rowNumbers.length; i++) {
      const rowNum = rowNumbers[i];
      const variant = variantList[i % variantList.length];

      for (const { header, hint, colNum } of headerCols) {
        if (BLOCKED_HEADERS.has(header)) continue;
        const hardcoded = getHardcodedValue(header, hint, variant.values);
        if (hardcoded !== undefined) {
          if (hardcoded !== null) worksheet.getRow(rowNum).getCell(colNum).value = hardcoded;
          continue;
        }
        const akeneoCode = MAKTX_LOCALES[header] ? MAKTX_HARDCODED_ATTR : columnMapping[header];
        if (!akeneoCode || !variant.values[akeneoCode]) continue;
        const locale = MAKTX_LOCALES[header];
        const value = locale
          ? extractLocalizedValue(variant.values[akeneoCode], locale)
          : extractValue(variant.values[akeneoCode]);
        if (value !== null && value !== undefined) {
          worksheet.getRow(rowNum).getCell(colNum).value = value;
        }
      }
    }
  }

  // Flatten shared formulas — prevents ExcelJS write error:
  // 'Shared Formula master must exist above and or left of clone'
  // Clone cells use { sharedFormula: 'addr', result } with NO formula key — must be handled separately.
  workbook.worksheets.forEach(ws => {
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        if (v && typeof v === 'object') {
          if ('sharedFormula' in v) {
            // Clone cell: drop shared-formula reference, keep result as static value
            cell.value = (v.result !== undefined && v.result !== null) ? v.result : null;
          } else if ('formula' in v) {
            // Master or regular formula: strip shareType/ref metadata
            cell.value = { formula: v.formula, result: v.result };
          }
        }
      });
    });
  });

  // Strip x14 extended CF rules that ExcelJS can't serialize
  workbook.worksheets.forEach(ws => {
    if (ws.conditionalFormattings) {
      ws.conditionalFormattings = ws.conditionalFormattings
        .map(cf => { cf.rules = (cf.rules || []).filter(r => r.formulae !== undefined || !['expression','cellIs','containsText','timePeriod'].includes(r.type)); return cf; })
        .filter(cf => cf.rules.length > 0);
    }
  });

  const outPath = filePath.replace(/\.xlsx?$/i, '_enriched.xlsx');
  await workbook.xlsx.writeFile(outPath);
  console.log(`\nEnriched file saved to: ${outPath}`);
}

main().catch(err => { console.error(err); process.exit(1); });
