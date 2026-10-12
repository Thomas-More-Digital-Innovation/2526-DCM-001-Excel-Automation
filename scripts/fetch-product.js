// Run: node -r dotenv/config scripts/fetch-product.js [SKU]
// or:  export $(cat .env | xargs) && node scripts/fetch-product.js [SKU]
const SKU = process.argv[2] || '1007143';

const BASE_URL       = process.env.AKENEO_BASE_URL;
const CLIENT_ID      = process.env.AKENEO_CLIENT_ID;
const CLIENT_SECRET  = process.env.AKENEO_CLIENT_SECRET;
const USERNAME       = process.env.AKENEO_USERNAME;
const PASSWORD       = process.env.AKENEO_PASSWORD;

['AKENEO_BASE_URL','AKENEO_CLIENT_ID','AKENEO_CLIENT_SECRET','AKENEO_USERNAME','AKENEO_PASSWORD'].forEach(k => {
  if (!process.env[k]) { console.error(`Missing env var: ${k}`); process.exit(1); }
});

async function getToken() {
  const creds = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64');
  const res = await fetch(`${BASE_URL}/api/oauth/v1/token`, {
    method: 'POST',
    headers: { Authorization: `Basic ${creds}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ grant_type: 'password', username: USERNAME, password: PASSWORD }),
  });
  if (!res.ok) throw new Error(`Token request failed: ${res.status} ${await res.text()}`);
  const { access_token } = await res.json();
  if (!access_token) throw new Error('No access_token in response');
  return access_token;
}

async function main() {
  console.log(`Authenticating...`);
  const token = await getToken();
  const headers = { Authorization: `Bearer ${token}` };

  console.log(`Fetching product "${SKU}"...`);
  let res = await fetch(`${BASE_URL}/api/rest/v1/products/${encodeURIComponent(SKU)}`, { headers });

  if (res.ok) {
    const product = await res.json();
    console.log('\n=== Product ===');
    console.log(JSON.stringify(product, null, 2));
    return;
  }

  if (res.status !== 404) throw new Error(`Product fetch failed: ${res.status} ${await res.text()}`);

  console.log(`Not found as product, trying product-model...`);
  res = await fetch(`${BASE_URL}/api/rest/v1/product-models/${encodeURIComponent(SKU)}`, { headers });

  if (!res.ok) {
    console.error(`Not found as product or product-model: ${res.status}`);
    process.exit(1);
  }

  const model = await res.json();
  console.log('\n=== Product Model ===');
  console.log(JSON.stringify(model, null, 2));
}

main().catch(err => { console.error(err.message); process.exit(1); });
