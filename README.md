# MISB Manager

A password-protected fund management website for Cloudflare Workers, adapted from the original MISB Fund Tracker. Keeps the original portfolio calculations and Excel workflows.

Includes Fund Overview, Issuer Exposure, Receivables, Profit Analytics, Active Portfolio, Account & Ledger, Cash Projection, Monitoring, MISB Reports, Simulation Copy, and Data Sources.

## Architecture

- Python Workers / Flask: calculations, authenticated API and Excel generation.
- Workers Static Assets: responsive HTML, CSS and JavaScript.
- D1: monitoring, payment marks, issuer settings, cashflow plans, uploaded workbooks and previous workbook versions. Files are never served as public assets. No R2 account or bucket is required.
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

### Git-connected Cloudflare Builds

In the **misbmanager** Worker's **Settings → Builds**, use these commands with the repository root as the root directory:

| Setting | Value |
| --- | --- |
| Build command | `npm run build` |
| Deploy command | `npm run deploy` |
| Non-production branch deploy command (if enabled) | `npm run deploy:preview` |

The build command prepares the Python dependencies and validates the Worker with a deployment dry run. It does not publish anything. The deploy command uses Pywrangler to package Python dependencies before publishing. Preview uploads create a version without promoting it to production.

If a build reports `Missing script: "build"`, ensure the commit being built contains the updated `package.json`, then retry. The npm install-script notices in the supplied log were warnings; the missing build script caused that failure. Cloud resource and secret setup below is still required.

See [Cloudflare Builds configuration](https://developers.cloudflare.com/workers/ci-cd/builds/configuration/) and [Python dependency packaging](https://developers.cloudflare.com/workers/languages/python/packages/).

### Initial resource setup

This deploys a full application on **Workers**, not a static-only Pages project. Your account must have Workers and D1 enabled. No cloud resources are created by the local build or test commands.

1. Install dependencies above, then authenticate:

   ```powershell
   npx wrangler login
   ```

2. Create the database (skip this if `misb-manager` already exists):

   ```powershell
   npx wrangler d1 create misb-manager
   ```

   Add the returned `database_id` to the `d1_databases` entry in `wrangler.jsonc`.

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

7. Optionally add your domain in Cloudflare Dashboard → Workers & Pages → misbmanager → Settings → Domains & Routes.

### Updating an existing installation

The application now stores workbooks entirely in D1. Apply all migrations before deploying:

```powershell
npx wrangler d1 migrations apply misb-manager --remote
npm run deploy
```

Migration `0002_workbook_storage.sql` adds workbook storage without altering existing monitoring or payment records. It has already been applied to this project's configured cloud database. Fresh installations still need both migrations. The Worker is named `misbmanager`; its database is `misb-manager`.

Existing local preview files are imported into SQLite automatically on the next `scripts/local.py` startup. If workbooks were uploaded to an earlier R2 installation, re-import those originals through Data Sources; remote R2 data is not copied or deleted automatically.

### Carry over existing monitoring and payment records

Generate a private SQL file from the original database:

```powershell
python scripts/export_records.py legacy/MISB_Fund_Tracker/data/misb_tracker.db
npx wrangler d1 execute misb-manager --remote --file .local/legacy-records.sql
```

Use this once on a new database before editing cloud records. It imports monitoring, settings, payment marks and cashflow plans; workbook uploads still go through Data Sources. Review the generated file before remote import, since it overwrites records with matching keys.

## Workbook handling

### Account Statement PDF

Open **Account Statement PDF**, upload the Cofundr transaction-log `.xlsx`, choose an inclusive cut-off date (or leave it blank for the latest transaction date), and click **Prepare statement**. Review the summary and select **Download account statement PDF**. Uploads replace the workspace transaction ledger and retain the previous workbook version in D1.

The PDF follows the supplied `Account Statement YTD 2Sep26.pdf`: US Letter pages, embedded Calibri fonts, Cofundr letterhead, investor 5490 / Amanahraya Trustees Berhad, five summary values, and the original six-column table. The first page holds 29 transactions and continuation pages hold 55, without repeating the header, matching the reference. Despite the reference title saying "YEAR TO DAY", it includes activity since October 2025; the generator likewise includes all completed entries in the supplied log through the chosen cut-off, not just the current calendar year.

Amounts use decimal arithmetic. Deposits and withdrawals use only successful Deposit Approved and Withdrawal Approved entries, preserving their Excel timestamps and balances. Duplicate Deposit and Withdrawal request entries are excluded. The separate RM1 balance deduction accompanying an approved withdrawal is retained as Withdrawal Fee. Balance discontinuities are flagged in the preview; excluded entries are never used to adjust approved balances. No transaction amounts or dates are sourced from the embedded, sanitized artwork template.

The supplied September 22 log, cut off at September 2, reproduces all 322 reference rows and summary totals, except the explicitly selected Excel withdrawal timestamp. PDF tests cover date filters, deposits, withdrawal fees, decimal totals, pagination, authenticating API requests, and rejecting a download if the source changed after preview.

Uploads accept `.xlsx` files up to 10 MB, with bounded archive expansion. Required columns and calculations are validated before publishing the new version. Rejected uploads leave the current version intact. Workbooks are stored as ordered, base64-encoded chunks in D1, with a size and SHA-256 integrity check. The complete workbook and current-version pointer are saved in one atomic transaction. Chunking preserves the 10 MB upload allowance while staying within D1 row limits. Each successful upload retains its predecessor in the database; no retention deletion runs automatically. Base64 encoding adds approximately one-third storage overhead. D1's `sources` table identifies the current versions.

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
