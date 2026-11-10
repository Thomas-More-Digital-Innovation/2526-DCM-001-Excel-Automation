// "Fetch Products": looks up every SKU in Akeneo; outputs only the cell values to write.
const { rows, headerCols, mapping } = $input.first().json;
return [{ json: { writes: await fetchCellWrites($env, rows, headerCols, mapping) } }];
