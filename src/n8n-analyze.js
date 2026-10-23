// n8n "Analyze" node body (POST /webhook/excel-analyze). src/lib.js is prepended by scripts/build-workflow.js.
const upload = readUpload($input.first());
return [{ json: await analyzeUpload(upload.buffer, upload.fileName, $env) }];
