// "Analyze: Read Layout" and "Enrich: Read Layout": headers, hints, SKU column, data rows and example values.
return [{ json: { layout: await readLayout(readUpload($input.first()).buffer) } }];
