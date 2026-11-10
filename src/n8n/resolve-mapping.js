// "Resolve Mapping": the portal's choice (form field "mapping" = JSON { skuCol, mapping }, remembered per header),
// or without it the mappings saved per header. Outputs the mapping and the rows to fill.
const webhook = $('Webhook').first();
const body = webhook.json.body || {};
const choice = body.mapping ? JSON.parse(body.mapping) : null;
const { layout } = $input.first().json;
const resolved = await resolveChoice($env, layout, choice, readUpload(webhook).buffer);
return [{ json: { headerCols: layout.headerCols, mapping: resolved.mapping, rows: resolved.rows } }];
