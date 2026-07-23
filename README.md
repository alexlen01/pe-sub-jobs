# pe-sub-jobs

Spring Boot 4.1 / Java 25 / Spring Batch 6 data ingestion service for the PE Sub Borrowing Base Platform. Reads flat-file CSV exports and posts them to `pe-sub-api`'s SERVICE-gated bulk endpoints. Runs at **`http://localhost:3003`**.

**DB-less by design:** this service holds no database connection and issues no SQL. `pe-sub-api`
owns the schema; every feed writes through its REST endpoints (`POST /api/facilities/ingest`,
`POST /api/lp-master/ingest`, `POST /api/lpRecords/seed`,
`PATCH /api/config/cls-conc-limit-defaults`), which also audit the writes. Spring Batch runs on
an in-memory `ResourcelessJobRepository` — no `BATCH_JOB_*` meta-tables. Run history is not
persisted and restart-from-failure is unsupported; every feed is an idempotent full-file re-run
(the API upserts/skips server-side), so replaying is always safe.

## Stack

- Java 25, Spring Boot 4.1.0, Maven 3.9
- Spring Batch 6 — chunk-oriented processing, fault-tolerant skip policy, `ResourcelessJobRepository`
- `RestClient` (`PeSubApiClient`) — one POST per 50-row chunk to `pe-sub-api`; no JDBC, no JPA
- Logback — daily rolling log to `pe-sub-jobs.log`, gzip-archived, 30-day retention

## Jobs

### `facility-ingest`

Reads a CSV of facility records, parses dates/decimals, and posts them to
`POST /api/facilities/ingest` (upsert by facility name, server-side).

**CSV columns (header row required):**

```
agent_bank, name, account_number, loan_amount, maturity_date, bank_status, bank_status_date
```

- `loan_amount` — plain decimal (no `$` or commas)
- `maturity_date`, `bank_status_date` — ISO `YYYY-MM-DD` or `M/d/yyyy`; blank = `NULL`
- `name` is the upsert key
- New facilities get `status = 'Not Started'` and `conc_limit_m = 25.00`; existing rows preserve their platform `status`

### `lp-master-ingest`

Reads a CSV of LP Master records, parses booleans/ratings, and posts them to
`POST /api/lp-master/ingest` (upsert by investor name, server-side).

**CSV columns (header row required):**

```
investor_name, parent, spv, high_qty, investor_type, inst_vs_hnw, region_location,
investment_grade, sp, mdy, fitch, aum, nav, pension, pension_funded,
ubs_classification, ubs_default_adv_rate, ubs_default_conc_limit, notes
```

- `spv`, `high_qty`, `investment_grade` — `true` / `false` strings
- `sp`, `mdy`, `fitch` — rating strings; blank stored as `""`
- `investor_name` is the upsert key

### `cls-conc-limits-ingest`

Reads a CSV of per-classification concentration-limit defaults and merges it via
`PATCH /api/config/cls-conc-limit-defaults` into the `cls_conc_limit_defaults` config map —
the same map edited on the UI's Config screen (Per-LP Concentration Limit Defaults card) and
used by the BB engine's fallback chain (per-LP limit → class default → facility limit).

**CSV columns (header row required):**

```
classification, limit_pct
```

- `limit_pct` — a **single** percent of total uncalled capital (`7.5` or `7.5%`); rows
  outside 0–100 or unparseable are skipped. The source workbook
  (`pe-sub-docs/Concentration_Limits.xls`) states each class as a range — feed the chosen
  bound as one number here, not the raw `"15.0 – 20.0"` range string (which would fail the
  numeric parse and skip the row)
- Rows **merge by classification key**: fed classes are overwritten, unfed classes are
  left untouched. En/em dashes in labels are normalized to hyphens
- The API persists the merge and refreshes its in-memory config cache in the same call —
  no follow-up `/api/config/reload` is needed

## Startup behaviour

All three jobs run automatically once the application is up, in sequence:

1. `facility-ingest` against `ingest.facility-file`
2. `lp-master-ingest` against `ingest.lp-master-file`
3. `lp-records-seed` against `ingest.lp-facility-seeds-file`

The seed job runs **after** facilities and LP Master because the API resolves its rows against both. Each job logs `status / readCount / writeCount / skipCount` on completion. A failure on one job is caught and logged; the remaining jobs still run. The skip limit per job is 10 rows — exceeding it marks that job `FAILED`. Before any job runs, startup waits for `pe-sub-api` to answer `GET /api/ping` (it only serves once its Flyway migrations are done); if it never does within the timeout, the startup feeds are skipped with a warning.

`cls-conc-limits-ingest` runs at startup **only when** `ingest.cls-conc-limits-file`
(env `CLS_CONC_LIMITS_FILE`) is set — the API's `V1_5` migration already seeds defaults,
so an unconfigured feed is skipped with a log line rather than re-fed on every boot.

Startup ingest is controlled by `ingest.run-on-startup` (default `true`; env `INGEST_RUN_ON_STARTUP`).
Set it to `false` to skip the seed jobs on boot — for example when `pe-sub-api` is not running,
or in tests (the integration base sets it `false`; job tests launch jobs explicitly against a
mocked `PeSubApiClient`). Disabling startup ingest does **not** remove the on-demand REST
triggers below.

Jobs can also be triggered on demand via REST (see [REST API](#rest-api)).

## Mock data

Development seed files are in `data/mock/`:

| File | Contents |
|---|---|
| `data/mock/facilities.csv` | 65 facilities derived from the Agent Bank Summary workbook |
| `data/mock/lp_master.csv` | 30 LP Master records across all classification tiers |
| `data/mock/lp_facility_seeds.csv` | 42 LP-to-facility assignments across 5 active facilities |
| `data/mock/cls_conc_limit_defaults.csv` | Reference feed for `cls-conc-limits-ingest` — 6 classification labels seeded to the upper bound of each `Concentration_Limits.xls` range (Excluded = 0) |

`lp_facility_seeds.csv` links LP Master records to specific facilities, producing `lp_records` rows the same way the ingestion wizard would. Since the D2 revision (see `pe-sub-docs/LP_DB_EXTRACT_DESIGN.md`) each seed row carries the **full per-LP column set** of the LP DB Export (31 columns: the legacy 7 first, then parent, spv, high_qty, investor_type, inst_vs_hnw, region_location, investment_grade, ubs_cls, sp, mdy, fitch, aum, nav, pension, pension_funded, pct_cap_commit, called_cap, pct_uncalled, pct_called, ubs_conc, ubs_rate, agent_bb, ubs_bb, notes) — `ubs_cls` is derived per row from that row's attributes (ratings, pension assets, NAV, AUM, HNW/SPV flags) via the Borrowing Base Criteria Matrix (`data/reference/bb_criteria_matrix.csv`, transcribed from `pe-sub-docs/BB_CRITERIA_DESIGN.md`), and the row's UBSAR/AgentAR advance rates are slotted into the discrete 90/75/65/50/0 rate groups by the Floor Map (`data/reference/rate_floor_map.csv`). Row values win server-side; the LP Master profile only fills blanks, and a legacy 7-column file still parses (the reader is non-strict and pads blanks). The API inserts a row only when that (facility, investor) pair has none yet — `lp_records` intentionally has **no** unique constraint on the pair (multi-sleeve) — so re-running is a safe no-op that never overwrites records committed through the Shadow BB flow. The startup defaults point at the extract output in `data/out/` (see `ingest.*` properties below).

The extract's input is the simulated, date-stamped LP DB Export produced by `scripts/lp_db_generate.py` into `data/import/`. The generator's built-in **chaos monkey** (`CHAOS_ENABLED`/`CHAOS_SEED` tunables; see `pe-sub-docs/"AI Chaos Monkey for Data Quality.md"`) degrades the values written to the XLSX to realistic manual-entry quality — name drift, `A minus` ratings, unit mix-ups, NAV range strings, categorical drift — while leaving cash/identity columns pristine, and logs every mutation to `data/import/<export name>.chaos_log.csv` as ground truth. `lp_db_extract.py` reads whatever file it is given **as-is** (no cleaning toggle), the same posture it needs for a real export; its unmatched/variance reports plus the chaos log show what the normalizers absorbed.

## Getting started

```bash
# pe-sub-api must be running at PE_SUB_API_URL (default http://localhost:3001)

mvn spring-boot:run
```

On startup the app waits for `pe-sub-api` to answer `/api/ping`, then runs the ingest jobs against the mock data files. No database is required.

## REST API

Both jobs can be triggered independently after startup:

```
POST /jobs/facility-ingest?filePath=<absolute-or-relative-path>
POST /jobs/lp-master-ingest?filePath=<absolute-or-relative-path>
POST /jobs/lp-records-seed?filePath=<absolute-or-relative-path>
POST /jobs/cls-conc-limits-ingest?filePath=<absolute-or-relative-path>
```

**Response:**

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
| `PORT` | `3003` (local profile) | HTTP port |
| `LOG_PATH` | `logs` | Log output directory |
| `FACILITY_INGEST_FILE` | `data/out/facilities.csv` | Path to facilities CSV for startup ingest |
| `LP_MASTER_INGEST_FILE` | `data/out/lp_master.csv` | Path to LP master CSV for startup ingest |
| `LP_FACILITY_SEEDS_FILE` | `data/out/lp_facility_seeds.csv` | Path to LP-facility seed CSV for startup ingest (full 31-column D2-revised format; legacy 7-column files still parse) |
| `CLS_CONC_LIMITS_FILE` | *(unset)* | Path to classification conc-limit defaults CSV; unset → feed skipped at startup |
| `INGEST_RUN_ON_STARTUP` | `true` | Run the seed jobs on startup; set `false` to skip them |
| `INGEST_SCHEMA_WAIT_TIMEOUT` | `30s` | How long startup ingest waits for pe-sub-api to answer `/api/ping` |
| `INGEST_SCHEMA_WAIT_INTERVAL` | `2s` | Poll interval while waiting for pe-sub-api |
| `BB_TEMPLATE_IMPORT_ENABLED` | `true` | Import BB template workbooks from the watched directory |
| `BB_TEMPLATE_IMPORT_DIR` | `data/bb-templates` | Directory scanned for `BB-Template-Import-*.xlsx` workbooks |
| `PE_SUB_API_URL` | `http://localhost:3001` (local profile) | `pe-sub-api` base URL — target of every feed (ingest/seed endpoints) and template upserts |
| `BB_TEMPLATE_SCAN_INTERVAL` | `30s` | How often to rescan the template directory while running |
| `BB_TEMPLATE_STABLE_AGE` | `2s` | Minimum file age before import, to avoid partially copied files |

## BB Template Imports

`pe-sub-jobs` imports structured BB template workbooks from `data/bb-templates/` on startup and
rescans the directory while running. Files are sent to `pe-sub-api` through
`POST /api/bb-templates/import?mode=upsert`, so imports are idempotent by `template_slug`:
existing templates are replaced, and missing templates are created.

Use `.partial` or `.tmp` while copying large files, then rename to `.xlsx` when complete. Temporary
Excel files beginning with `~$` are ignored.

### Generate a template: `scripts/parse_excel_templates.py`

Analyzes a raw Agent BB or investor-list workbook and generates a `BB-Template-Import-<slug>.xlsx`
meta-workbook ready for pe-sub-jobs to pick up:

```bash
python scripts/parse_excel_templates.py data/import/some-agent-bb.xlsx
```

Output lands in `data/bb-templates/` and is imported automatically within 30 seconds.

## Logging

Logs are written to `$LOG_PATH/pe-sub-jobs.log` and rotated daily to `$LOG_PATH/archived/pe-sub-jobs.YYYY-MM-DD.log.gz` (30 days / 2 GB cap). Console output mirrors the file.

## Build

```bash
mvn package              # fat JAR → target/pe-sub-jobs-1.0.0.jar
mvn package -DskipTests
java -jar target/pe-sub-jobs-1.0.0.jar
```

## Testing

**No database.** The app is DB-less, so tests are too: any test that boots the Spring context extends `IntegrationTestBase`, which replaces `PeSubApiClient` with a Mockito mock (`@MockitoBean`). Job tests run the real reader/processor pipeline against temp CSV files and assert the payloads handed to the API client; the write semantics themselves (upsert/skip/merge) are pe-sub-api's responsibility and are covered by its `SeedIngestEndpointsIntegrationTest`.

## Project structure

```
src/main/java/com/ubs/pesubjobs/
  PeSubJobsApplication.java        @SpringBootApplication + @ConfigurationPropertiesScan
  JobStartupRunner.java            ApplicationRunner — waits for /api/ping, runs the feeds
  client/
    PeSubApiClient.java            RestClient wrapper — all pe-sub-api ingest/seed calls
  config/
    IngestProperties.java          @ConfigurationProperties(prefix = "ingest")
    ResourcelessBatchConfig.java   In-memory JobRepository/JobOperator — no BATCH_* tables
    FacilityIngestJobConfig.java   Job + Step + @StepScope reader + API-posting writer
    LpMasterIngestJobConfig.java
    LpRecordsSeedJobConfig.java    Posts raw seed rows (full 31-column set); the API resolves names, row values win, LP Master fills blanks
    ClsConcLimitIngestJobConfig.java
  controller/
    JobController.java             POST /jobs/{jobName}
  exception/
    GlobalExceptionHandler.java    ProblemDetail error responses
  model/
    FacilityRow.java               Raw CSV record (all Strings)
    ProcessedFacility.java         Type-safe record (BigDecimal, LocalDate)
    LpMasterRow.java
    ProcessedLpMaster.java
    LpFacilitySeedRow.java         Raw seed CSV row (full per-LP column set) — posted to the API verbatim
    ClsConcLimitRow.java / ProcessedClsConcLimit.java
  processor/
    FacilityRowProcessor.java      Parses dates/decimals; returns null to skip invalid rows
    LpMasterRowProcessor.java      Parses booleans; normalises blank ratings to ""
    ClsConcLimitRowProcessor.java  Percent parsing + range check + dash normalization
src/main/resources/
  application.yml (+ application-{local,dev,qa,prod}.yml)
  logback-spring.xml
data/mock/
  facilities.csv
  lp_master.csv
```
