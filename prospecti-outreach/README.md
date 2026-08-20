# Prospecți · outreach list

A single-page outreach tool for the August 2026 prospect campaign, deployed as one Cloudflare Worker. The Worker serves a self-contained page plus the full prospect list, and stores per-lead outreach status in a Cloudflare D1 database.

Live: `https://prospecti-august-2026.catalin-932.workers.dev`

## What is committed

- `index.html` : the whole front end (list, search, county + status filters, per-lead detail page). No data is embedded here; it is fetched at runtime.
- `build.py` : parses the CRM spreadsheet into `data.json` and bundles it with `index.html` into `worker.js`.
- `.gitignore` : keeps the source spreadsheet, the generated `data.json` / `worker.js`, and any credentials out of git. The list holds real company phone numbers and emails, so it stays out of version control.

## Build

```bash
python3 -m pip install openpyxl
python3 build.py path/to/export.xlsx    # writes data.json and worker.js
```

The spreadsheet must have a `BD CAMPANIE` sheet with the campaign columns (Companie, CUI, Cifră de afaceri, Oraș, Județ, Status detaliat, Descriere prioritati campanie, ...). `build.py` also pulls phone numbers, emails, priorities and the eligibility facts out of the free-text notes column so the lead page can present contact details up front.

Bump `BUILD` in `index.html` whenever the data changes, so browsers fetch a fresh `data.json?b=<n>` instead of a cached copy.

## Deploy

Deploy is a single Worker upload. It needs, as environment variables, the Cloudflare account id, credentials, and the D1 database id (never commit these):

```bash
curl -X PUT \
  -H "X-Auth-Email: $CF_EMAIL" -H "X-Auth-Key: $CF_KEY" \
  "https://api.cloudflare.com/client/v4/accounts/$CF_ACCT/workers/scripts/prospecti-august-2026" \
  -F "metadata={\"main_module\":\"worker.js\",\"compatibility_date\":\"2026-06-01\",\"bindings\":[{\"type\":\"d1\",\"name\":\"DB\",\"id\":\"$DB_ID\"}]};type=application/json" \
  -F 'worker.js=@worker.js;type=application/javascript+module'
```

The workers.dev route is enabled once via `POST .../workers/scripts/prospecti-august-2026/subdomain` with `{"enabled":true}`.

## Storage

A D1 database bound as `DB` holds one table:

```sql
CREATE TABLE IF NOT EXISTS lead_status (cui TEXT PRIMARY KEY, status TEXT NOT NULL, updated_at TEXT NOT NULL);
```

Endpoints served by the Worker:

- `GET /data.json` : the full list (no-store, so status edits and data rebuilds show immediately).
- `GET /api/statuses` : `{ cui: status }` for every lead that has been touched.
- `POST /api/status` : body `{ cui, status }`, upserts the status. Rejects any status outside the allowed set.

## Outreach statuses

`Not contacted` (default), `Contacted`, `No answer`, `Follow up`, `Interested`, `Meeting`, `In progress`, `Won`, `Lost`.

## Not done yet

Each lead page has an "Îmbogățire date" section reserved for enrichment (firmographics, signals, scoring) to be added later. The page and the data model are built to take it without restructuring.

## Note on access

The page and the `POST /api/status` endpoint are currently public to anyone with the URL. Put the Worker behind Cloudflare Access or a shared passphrase before sharing it beyond the team.
