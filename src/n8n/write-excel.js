// "Write Excel": puts the values into the uploaded file and returns it as binary for "Respond File".
const upload = readUpload($('Webhook').first());
const workbook = await applyWrites(upload.buffer, $input.first().json.writes);
const base64Data = Buffer.from(await workbook.xlsx.writeBuffer()).toString('base64');
const fileName = upload.fileName;
return [{ json: { fileName, cellsWritten: $input.first().json.writes.length }, binary: { data: { data: base64Data, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName, fileExtension: 'xlsx' } } }];
