# MISB Manager

A password-protected fund management website for Cloudflare Workers, adapted from the original MISB Fund Tracker. Keeps the original portfolio calculations and Excel workflows.

Includes Fund Overview, Issuer Exposure, Receivables, Profit Analytics, Active Portfolio, Account & Ledger, Cash Projection, Monitoring, MISB Reports, Simulation Copy, and Data Sources.

## Architecture

- Python Workers / Flask: calculations, authenticated API and Excel generation.
- Workers Static Assets: responsive HTML, CSS and JavaScript.
- D1: monitoring, payment marks, issuer settings, cashflow plans and workbook version pointers.
- Private R2: uploaded workbooks and previous versions. Files are never served as public assets.
- Shared workspace password and signed eight-hour HttpOnly session. This is a single shared workspace, without per-user roles or a user audit trail.

## Preview on this computer

The original ZIP is retained unchanged. If it has not already been extracted:

```powershell
Expand-Archive MISB_Fund_Tracker_simulation_copy_FIXED_v2.zip -DestinationPath legacy
```

With Flask, openpyxl and defusedxml installed:

```powershell
python scripts/local.py --seed legacy/MISB_Fund_Tracker/data
```

Open **http://127.0.0.1:5056**. This loopback-only preview skips sign-in. Data is saved in `.local/`; repeating the seed command does not replace workbooks. Do not expose the development server to the internet.

## Cloudflare local development

Install Node.js 22+ and [uv](https://docs.astral.sh/uv/getting-started/installation/), then:

```powershell
npm ci
uv sync
Copy-Item .dev.vars.example .dev.vars
```

Edit `.dev.vars`: set `APP_PASSWORD` to a private password of at least 12 characters and `SESSION_SECRET` to at least 32 random characters. The file is ignored by Git. Do not overwrite an existing `.dev.vars` if it already contains your settings.

```powershell
npx wrangler d1 migrations apply misb-manager --local
uv run pywrangler dev
```

Open **http://localhost:8787**, sign in and import the original spreadsheets through **Data Sources**. Local Cloudflare storage is independent of the loopback preview and production storage.

## Host on Cloudflare

This deploys a full application on **Workers**, not a static-only Pages project. Your account must have Workers, D1 and R2 enabled. No cloud resources are created by the local build or test commands.

1. Install dependencies above, then authenticate:

   ```powershell
   npx wrangler login
   ```

2. Create the database and private workbook bucket:

   ```powershell
   npx wrangler d1 create misb-manager
   npx wrangler r2 bucket create misb-manager-files
   ```

   Add the returned `database_id` to the `d1_databases` entry in `wrangler.jsonc`. If you change the bucket name, update `r2_buckets` too.

3. Apply the schema:

   ```powershell
   npx wrangler d1 migrations apply misb-manager --remote
   ```

4. Validate and deploy:

   ```powershell
   uv run pywrangler deploy --dry-run
   uv run pywrangler deploy
   ```

   The app denies workspace access until its secrets are configured.

5. Set production secrets using the interactive prompts:

   ```powershell
   npx wrangler secret put APP_PASSWORD
   npx wrangler secret put SESSION_SECRET
   ```

6. Open the `workers.dev` URL printed by deployment. Sign in, then import `transactions.xlsx`, `simulation.xlsx`, and optionally `midas_projections.xlsx` from the extracted ZIP under **Data Sources**. Set the simulation cut-off date from `data/source_meta.json`. Existing monitoring/payment records are not automatically uploaded; see below.

7. Optionally add your domain in Cloudflare Dashboard → Workers & Pages → misb-manager → Settings → Domains & Routes.

Future deployments preserve D1 and R2 data. Changing `SESSION_SECRET` signs everyone out. For named users and organization sign-in, Cloudflare Access can be added later.

### Carry over existing monitoring and payment records

Generate a private SQL file from the original database:

```powershell
python scripts/export_records.py legacy/MISB_Fund_Tracker/data/misb_tracker.db
npx wrangler d1 execute misb-manager --remote --file .local/legacy-records.sql
```

Use this once on a new database before editing cloud records. It imports monitoring, settings, payment marks and cashflow plans; workbook uploads still go through Data Sources. Review the generated file before remote import, since it overwrites records with matching keys.

## Workbook handling

Uploads accept `.xlsx` files up to 10 MB, with bounded archive expansion. Required columns and calculations are validated before publishing the new version. Rejected uploads leave the current version intact. Each successful upload keeps its predecessor in R2 under `sources/`; no retention deletion runs automatically. D1's `sources` table identifies the current versions.

Excel exports preserve the simulation template's sheets, formulas, styling and Remarks. Excel recalculates formulas when opening the updated simulation. The bi-weekly report is a current monitoring snapshot: its report date labels the report and does not filter historical activity. Simulation Copy supports its original ledger cut-off workflow.

## Verification

```powershell
uv run python -m unittest discover -s tests -v
npx playwright test
uv run pywrangler deploy --dry-run
```

Python tests compare totals, profit schedules and projections against the original app, and test persistence, login, first imports, rejected uploads and exports. Browser tests use installed Microsoft Edge and check every view and mobile layout against the seeded local preview. Run the preview seed command first. These commands do not deploy to production.

For the storage integration check, start `uv run pywrangler dev --port 8791`, then run `python scripts/check_worker.py`. This signs in using `.dev.vars`, imports the original workbooks into the **local** emulator and checks all Excel exports. It replaces the emulator's current workbook versions, so use test data only. The script only connects to localhost.

References: [Flask on Workers](https://developers.cloudflare.com/workers/languages/python/packages/flask/), [Python packages](https://developers.cloudflare.com/workers/languages/python/packages/), [D1 bindings](https://developers.cloudflare.com/d1/worker-api/).
