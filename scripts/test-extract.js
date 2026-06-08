// Run: node scripts/test-extract.js path/to/file.xlsx
const ExcelJS = require('exceljs');
const fs = require('fs');

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

async function extractHeadersAndRows(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.worksheets[0];
  const totalRows = worksheet.rowCount;

  // Column extent — scan up to 200 cols across first 50 rows
  let maxCols = 0;
  for (let r = 1; r <= Math.min(50, totalRows); r++) {
    for (let c = 200; c >= 1; c--) {
      if (getCellText(worksheet.getRow(r).getCell(c).value) !== null) {
        if (c > maxCols) maxCols = c;
        break;
      }
    }
  }
  if (maxCols === 0) throw new Error('Excel sheet appears empty in first 50 rows');

  // Fill ratios for first 60 rows
  const rowFills = [];
  for (let r = 1; r <= Math.min(60, totalRows); r++) {
    let filled = 0;
    for (let c = 1; c <= maxCols; c++) {
      if (getCellText(worksheet.getRow(r).getCell(c).value) !== null) filled++;
    }
    rowFills.push({ r, filled, ratio: filled / maxCols });
  }

  // Primary header = max-fill row in the first dense region
  const firstDense = rowFills.find(rf => rf.ratio >= 0.5);
  if (!firstDense) throw new Error('No header row detected (no row ≥50% fill in first 60 rows)');

  const regionEnd = Math.min(firstDense.r + 20, 60);
  let headerFill = firstDense;
  for (const rf of rowFills) {
    if (rf.r >= firstDense.r && rf.r <= regionEnd && rf.filled > headerFill.filled) headerFill = rf;
  }
  const headerRowIndex = headerFill.r;

  // Hint row = the row immediately before the primary header (if inside the header region)
  const hintMap = {};
  const hintRowIdx = headerRowIndex - 1;
  if (hintRowIdx >= firstDense.r) {
    for (let c = 1; c <= maxCols; c++) {
      const text = getCellText(worksheet.getRow(hintRowIdx).getCell(c).value);
      if (text !== null) hintMap[c] = text;
    }
  }

  // Build headerCols (with hints where available)
  const headerCols = [];
  for (let c = 1; c <= maxCols; c++) {
    const text = getCellText(worksheet.getRow(headerRowIndex).getCell(c).value);
    if (text !== null) {
      const entry = { header: text, colNum: c };
      if (hintMap[c]) entry.hint = hintMap[c];
      headerCols.push(entry);
    }
  }

  // Data start: hybrid threshold (60% when no section rows precede the header, 40% otherwise)
  const headerThreshold = (firstDense.r === headerRowIndex) ? 0.6 : 0.4;
  let dataStartRow = headerRowIndex + 1;
  let lastClusterRow = headerRowIndex;

  for (let r = headerRowIndex + 1; r <= Math.min(headerRowIndex + 20, totalRows); r++) {
    const rfItem = rowFills.find(x => x.r === r);
    const ratio = rfItem ? rfItem.ratio : 0;
    if (ratio > headerThreshold) {
      lastClusterRow = r;
    } else {
      const col1 = getCellText(worksheet.getRow(r).getCell(1).value);
      if (col1 !== null) { dataStartRow = r; break; }
    }
  }
  if (dataStartRow === headerRowIndex + 1 && lastClusterRow > headerRowIndex) {
    dataStartRow = lastClusterRow + 1;
    for (let r = dataStartRow; r <= Math.min(dataStartRow + 5, totalRows); r++) {
      if (getCellText(worksheet.getRow(r).getCell(1).value) !== null) { dataStartRow = r; break; }
    }
  }

  // Collect data rows (col1 non-empty = article code / SKU)
  const rows = [];
  for (let r = dataStartRow; r <= totalRows; r++) {
    const sku = getCellText(worksheet.getRow(r).getCell(1).value);
    if (sku === null) continue;
    rows.push({ sku, rowNum: r });
  }

  return { headerCols, headerRowIndex, dataStartRow, rows };
}

const filePath = process.argv[2];
if (!filePath) { console.error('Usage: node scripts/test-extract.js <file.xlsx>'); process.exit(1); }

extractHeadersAndRows(filePath)
  .then(result => {
    console.log('Header row index:', result.headerRowIndex);
    console.log('Data start row:', result.dataStartRow);
    console.log('Headers found:', result.headerCols.length);
    result.headerCols.forEach((h, i) => {
      const hint = h.hint ? `  hint: "${h.hint.replace(/\n/g, ' | ').substring(0, 60)}"` : '';
      console.log(`  ${i}  col ${h.colNum}: "${h.header.replace(/\n/g, '\\n').substring(0, 60)}"${hint}`);
    });
    console.log('Data rows found:', result.rows.length);
    result.rows.slice(0, 5).forEach(r => console.log(`  row ${r.rowNum}: SKU="${r.sku}"`));
  })
  .catch(err => { console.error(err.message); process.exit(1); });
