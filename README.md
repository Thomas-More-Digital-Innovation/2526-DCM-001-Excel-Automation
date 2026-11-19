# DCM Excel Portal

> **Proof of concept.** Tested with one supplier template (`Docs/InitialFile.xlsx`) and a small set of products.
> Not everything may work as expected. There is no login, no HTTPS and no backup: see the limitations in the IT guide.

Fills supplier Excel templates with product data from Akeneo. The user uploads a template in the portal, checks which
Akeneo field fills each column (OpenAI suggests one for new headers; it never gets product data), and downloads the
filled file. Choices are saved per column header in Postgres and reused next time.

## Deploy

Needs Docker with Docker Compose, and outbound access to Akeneo and `api.openai.com`.

```sh
cp .env.example .env      # fill in every value
docker compose up -d
```

- Portal: `http://<server>:3000`
- n8n editor: `http://<server>:5678` (create the owner account on first visit)

The n8n workflow (`workflow.json`) is imported and activated on every start. Nothing to do in n8n.

## Change the code

```sh
npm install
# edit src/lib.js, src/n8n/*.js or portal/index.html
npm run build             # copies src/ into workflow.json
npm test
docker compose restart n8n
```

Edit code in the repository, not in the n8n editor: a restart overwrites the workflow with `workflow.json`.

## Docs

- `Docs/DCM Excel Portal.docx`: user guide for the portal
- `Docs/DCM Developer Docs.docx`: IT guide (deployment, configuration, where everything is, known limitations)
