// n8n "Enrich & Write Excel" node body (POST /webhook/excel-enrichment). src/lib.js is prepended by scripts/build-workflow.js.
// Optional form field "mapping" = JSON { skuCol, mapping } from the portal; without it the saved mapping is used.
const item = $input.first();
const upload = readUpload(item);
const body = (item.json && item.json.body) || {};
const choice = body.mapping ? JSON.parse(body.mapping) : null;

const workbook = await loadWorkbook(upload.buffer);
await enrichWorkbook(workbook, choice, upload.fileName, $env);

const base64Data = Buffer.from(await workbook.xlsx.writeBuffer()).toString('base64');
const fileName = upload.fileName;
return [{ json: { success: true, fileName }, binary: { data: { data: base64Data, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName, fileExtension: 'xlsx' } } }];
