# pe-sub-jobs

Spring Boot 4.1 / Java 21 / Spring Batch 6 ingestion service for the PE Sub Borrowing Base Platform. It reads CSV files and posts them to `pe-sub-api`. Runs at **`http://localhost:3003`**.

It holds no database connection. `pe-sub-api` owns the schema and audits every write. Spring Batch runs on an in-memory job repository, so there is no run history and no restart-from-failure — instead every feed is a full-file re-run that the API upserts or skips, so replaying is always safe.

## Stack

- Java 21, Spring Boot 4.1.0, Maven 3.9
- Spring Batch 6 — chunked reads, skip policy, in-memory `ResourcelessJobRepository`
- `RestClient` (`PeSubApiClient`) — one POST per 50-row chunk
- Logback — daily rolling log, gzipped, 30-day retention

## Jobs

### `facility-ingest`

Posts facility rows to `POST /api/facilities/ingest`. The API upserts on facility name.

```
agent_bank, name, account_number, loan_amount, maturity_date, bank_status, bank_status_date
```

- `loan_amount` — plain decimal, no `$` or commas
- dates — `YYYY-MM-DD` or `M/d/yyyy`; blank means null
- new facilities start at `status = 'Not Started'`, `conc_limit_m = 25.00`; existing rows keep their platform status

### `lp-master-ingest`

Posts LP Master rows to `POST /api/lp-master/ingest`. The API upserts on investor name.

```
investor_name, parent, spv, investor_type, institutional_or_hnw, region_location,
investment_grade, sp_rating, moodys_rating, fitch_rating, aum, nav, pension_assets,
funding_ratio, ubs_lp_category, ubs_default_advance_rate, ubs_default_concentration_limit, notes
```

- `spv`, `investment_grade` — `true` / `false`
- `sp_rating`, `moodys_rating`, `fitch_rating` — rating strings; blank is stored as `""`
- `high_quality` is **not** a column. The LP DB Export dropped it in the 2026-08-18 format and
  nothing else supplies it, so pe-sub-api keeps its own column on the schema default (`TRUE`)
  rather than being fed a fabricated value.
- `investor_type`, `region_location` and `funding_ratio` are written **blank** for the same reason.
  The API reads blank as "not resubmitted" and leaves any value already on the record intact.
- `aum` / `nav` / `pension_assets` — exactly one is populated per row, chosen by the export's
  `LP Size Criteria` (`AUM` → `aum`, `NAV` → `nav`, `Assets` → `pension_assets`).

### `lp-records-seed`

Posts LP-to-facility rows to `POST /api/lpRecords/seed`, creating `lp_records` the same way the ingestion wizard does. Each row carries the full 32-column per-LP set from the LP DB Export; a legacy 7-column file still parses, since the reader pads blanks.

Row values win on the server; the matching LP Master record only fills blanks. The API inserts only when that (facility, investor) pair has no row yet, so re-running never overwrites anything committed through the Shadow BB flow.

## BB template import

`BbTemplateDirectoryImporter` watches `data/bb-templates` for `*.xlsx` and posts each new or changed file to `POST /api/bb-templates/import?mode=upsert`. It scans at startup and every 30s, and waits for a file to sit still for 2s before reading it, so half-copied files are skipped.

## Startup

The app waits for `pe-sub-api` to answer `GET /api/ping` (served only once its Flyway migrations finish), then runs the three feeds in order:

1. `facility-ingest`
2. `lp-master-ingest`
3. `lp-records-seed` — last, because the API resolves its rows against the first two

Each job logs `status / readCount / writeCount / skipCount`. A failed job is logged and the rest still run. More than 10 skipped rows marks a job `FAILED`.

Set `INGEST_RUN_ON_STARTUP=false` to skip the startup feeds — useful when `pe-sub-api` is down, and what the tests do. The `/jobs` endpoints still work either way.

## Data

`data/out/` is the default input: the three CSVs written by `scripts/lp_db_extract.py`.

`data/mock/` is a clean hand-kept fixture — 68 facilities, 37 LP Master records covering all nine UBS LP Categories, and 50 LP-to-facility rows covering every UBS and agent category, with three rows below the 40% funded split so the matrix `lt40` rates get exercised. Every categorical value is a canonical `classification_config` option and every concentration limit carries an explicit `%`, because `MoneyValues.concLimit` tells a percent from a dollar cap by magnitude — a bare `0.05` would land as 0.05%, not 5%.

Concentration limits and UBS advance rates in the fixtures come from `bb_criteria_matrix`, the team's authoritative source. The `agent_*` columns hold the agent bank's own diverging figures — that gap is what the platform measures.

`data/reference/` holds the editable lists the extract normalizes against: the BB criteria matrix, the rate floor map, agent LP categories, UBS LP classifications (`ubs_lp_categories.csv`), and the agent advance rate per category (`agent_rate_map.csv`). The last two arrived with the 2026-08-18 export format — the export now states the UBS classification outright instead of it being derived, and `agent_rate_map.csv` backstops its `Agent Advance Rate` column, supplying the rate from the row's agent category only when that cell is blank. `investor_types.csv` / `investor_type_aliases.csv` are no longer read by the extract (the Investor Type column is gone from the feed); they stay because they mirror `classification_config.INVESTOR_TYPE_OPTS` for the platform.

## Scripts

`scripts/` holds Python utilities, run by hand:

| Script | Does |
|---|---|
| `lp_db_generate.py` | Writes a simulated LP DB Export into `data/import/`. Its chaos monkey degrades values to realistic manual-entry quality — name drift, `A minus` ratings, unit mix-ups, NAV ranges — while leaving cash and identity columns clean. The same `CHAOS_SEED` gives the same output. |
| `lp_db_extract.py` | Reads the LP DB Export plus `AgentBankSummaryRpt.xlsx` from `data/import/` and writes the three CSVs into `data/out/`. Takes its input as-is, the way it must for a real export, and prints unmatched and variance counts to the console. |
| `parse_excel_templates.py` | Turns one agent BB workbook into a BB template. |
| `parse_agent_bb_directory.py` | The same, over a directory tree where each subfolder is an agent bank. |

The Agent Bank Summary is a printed report, not a table: the agent bank sits on a group-header row above the facilities it covers, and each group ends with a subtotal line. The extract carries the agent name down onto its rows, drops subtotal and reprinted rows, and disambiguates a borrower name reused under a second account by appending the account number. Two facility columns are not in the report — `ubs_participation` is left blank, and `collateral_date` comes from the export's `BBDate`. `bank_status` is Active when the account appears in the export and Inactive otherwise; export accounts the report omits become `"Unknown"`-bank placeholders, so no LP record is rejected.

## Getting started

```bash
# pe-sub-api must be running at PE_SUB_API_URL (default http://localhost:3001)
mvn spring-boot:run
```

To boot against the mock fixtures instead of the extract output:

```bash
FACILITY_INGEST_FILE=data/mock/facilities.csv \
LP_MASTER_INGEST_FILE=data/mock/lp_master.csv \
LP_FACILITY_SEEDS_FILE=data/mock/lp_facility_seeds.csv \
mvn spring-boot:run
```

## REST API

Any job can be triggered after startup:

```
POST /jobs/facility-ingest?filePath=<path>
POST /jobs/lp-master-ingest?filePath=<path>
POST /jobs/lp-records-seed?filePath=<path>
```

```json
{
  "executionId": 1,
  "status": "COMPLETED",
  "exitCode": "COMPLETED",
  "readCount": 65,
  "writeCount": 65,
  "skipCount": 0
}
```

## Environment variables

| Variable | Default | Description |
|---|---|---|
| `PORT` | `3003` (local) | HTTP port |
| `LOG_PATH` | `logs` | Log directory |
| `PE_SUB_API_URL` | `http://localhost:3001` (local) | `pe-sub-api` base URL — target of every feed and template upsert |
| `FACILITY_INGEST_FILE` | `data/out/facilities.csv` | Facilities CSV for startup ingest |
| `LP_MASTER_INGEST_FILE` | `data/out/lp_master.csv` | LP Master CSV for startup ingest |
| `LP_FACILITY_SEEDS_FILE` | `data/out/lp_facility_seeds.csv` | LP-facility seed CSV for startup ingest |
| `INGEST_RUN_ON_STARTUP` | `true` | Run the feeds on startup |
| `INGEST_SCHEMA_WAIT_TIMEOUT` | `30s` | How long to wait for `/api/ping` |
| `INGEST_SCHEMA_WAIT_INTERVAL` | `2s` | Poll interval while waiting |
| `BB_TEMPLATE_IMPORT_ENABLED` | `true` | Watch for BB template workbooks |
| `BB_TEMPLATE_IMPORT_DIR` | `data/bb-templates` | Directory scanned for `*.xlsx` |
| `BB_TEMPLATE_SCAN_INTERVAL` | `30s` | How often to rescan |
| `BB_TEMPLATE_STABLE_AGE` | `2s` | File must be this old before import |

## Logging

Logs go to `$LOG_PATH/pe-sub-jobs.log`, rotated daily into `archived/pe-sub-jobs.YYYY-MM-DD.log.gz` (30 days, 2 GB cap). The console mirrors the file.

## Build

```bash
mvn package              # fat JAR → target/pe-sub-jobs-1.0.0.jar
mvn package -DskipTests
java -jar target/pe-sub-jobs-1.0.0.jar
```

## Testing

No database, so no test database. Anything that boots the Spring context extends `IntegrationTestBase`, which swaps `PeSubApiClient` for a Mockito mock. Job tests run the real reader and processor against temp CSVs and assert what reaches the client. The write semantics — upsert, skip, merge — belong to `pe-sub-api` and are tested there.

## Layout

```
src/main/java/com/ubs/pesubjobs/
  PeSubJobsApplication.java        entry point
  JobStartupRunner.java            waits for /api/ping, then runs the feeds
  BbTemplateDirectoryImporter.java watches data/bb-templates, posts to the API
  client/                          RestClient wrapper for pe-sub-api
  config/                          job definitions + @ConfigurationProperties
  controller/                      POST /jobs/{jobName}
  exception/                       ProblemDetail error responses
  model/                           raw CSV rows and their parsed counterparts
  processor/                       date/decimal/boolean parsing; null skips a bad row
src/main/resources/                application.yml (+ local/dev/qa/prod), logback-spring.xml
data/                              import/ · out/ · mock/ · reference/ · bb-templates/
scripts/                           Python generator, extract and BB template parsers
```
