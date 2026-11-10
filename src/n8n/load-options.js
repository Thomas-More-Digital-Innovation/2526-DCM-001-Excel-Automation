// "Options & Saved Mapping": Akeneo attributes plus what earlier files decided per header; `unknownCols` drives the IF.
const upload = readUpload($('Analyze Webhook').first());
return [{ json: await loadOptions($env, $input.first().json.layout, upload.buffer) }];
