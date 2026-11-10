// Inlines src/lib.js + the node bodies into workflow.json's Code nodes.
// Run: npm run build   (then re-import workflow.json into n8n)
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(root, f), 'utf8');
const parts = read('src/lib.js').split('// @cli-only');
if (parts.length !== 2) throw new Error('src/lib.js must contain exactly one "// @cli-only" marker');
const lib = parts[0];

// Code node name in workflow.json → its body in src/n8n/.
const NODE_SOURCES = {
  'Analyze: Read Layout': 'src/n8n/read-layout.js',
  'Options & Saved Mapping': 'src/n8n/load-options.js',
  'Suggest Mapping (OpenAI)': 'src/n8n/suggest-mapping.js',
  'Enrich: Read Layout': 'src/n8n/read-layout.js',
  'Resolve Mapping': 'src/n8n/resolve-mapping.js',
  'Fetch Products': 'src/n8n/fetch-products.js',
  'Write Excel': 'src/n8n/write-excel.js',
};

function buildWorkflow() {
  const workflow = JSON.parse(read('workflow.json'));
  for (const [name, file] of Object.entries(NODE_SOURCES)) {
    const node = workflow.nodes.find(n => n.name === name);
    if (!node) throw new Error('workflow.json has no node named "' + name + '"');
    node.parameters.jsCode = lib + '\n' + read(file);
  }
  return JSON.stringify(workflow, null, 2) + '\n';
}

module.exports = { buildWorkflow };

if (require.main === module) {
  fs.writeFileSync(path.join(root, 'workflow.json'), buildWorkflow());
  console.log('workflow.json updated');
}
