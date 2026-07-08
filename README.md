# pe-sub-jobs

Spring Boot 4.1 / Java 25 / Spring Batch 6 data ingestion service for the PE Sub Borrowing Base Platform. Reads flat-file CSV exports and upserts records into the shared PostgreSQL database. Runs at **`http://localhost:3003`**.

## Stack

- Java 25, Spring Boot 4.1.0, Maven 3.9
- Spring Batch 6 — chunk-oriented processing, fault-tolerant skip policy
- Spring JDBC (`JdbcBatchItemWriter`) — direct UPSERT SQL; no JPA
- PostgreSQL 16 (shared with `pe-sub-api`) — Spring Batch meta-tables auto-created on first startup
- Logback — daily rolling log to `pe-sub-jobs.log`, gzip-archived, 30-day retention

## Jobs

### `facility-ingest`

Reads a CSV of facility records and upserts them into the `facilities` table.

**CSV columns (header row required):**

```
agent_bank, name, account_number, loan_amount, maturity_date, bank_status, bank_status_date
```

- `loan_amount` — plain decimal (no `$` or commas)
- `maturity_date`, `bank_status_date` — ISO format `YYYY-MM-DD`; blank = `NULL`
- `name` is the upsert key (`ON CONFLICT (name) DO UPDATE`)
- New rows get `status = 'Not Started'` and `conc_limit_m = 25.00`; existing rows preserve their platform `status`

### `lp-master-ingest`

Reads a CSV of LP Master records and upserts them into the `lp_master` table.

**CSV columns (header row required):**

```
investor_name, parent, spv, high_qty, investor_type, inst_vs_hnw, region_location,
investment_grade, sp, mdy, fitch, aum, nav, pension, pension_funded,
ubs_classification, ubs_default_adv_rate, ubs_default_conc_limit, notes
```

- `spv`, `high_qty`, `investment_grade` — `true` / `false` strings
- `sp`, `mdy`, `fitch` — rating strings; blank stored as `""`
- `investor_name` is the upsert key (`ON CONFLICT (investor_name) DO UPDATE`)

### `cls-conc-limits-ingest`

Reads a CSV of per-classification concentration-limit defaults and merges it into the
`cls_conc_limit_defaults` config row (jsonb map) in the shared `config` table — the same
map edited on the UI's Config screen (Per-LP Concentration Limit Defaults card) and used
by the BB engine's fallback chain (per-LP limit → class default → facility limit).

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
- After a successful run the job calls `POST {PE_SUB_API_URL}/api/config/reload` so
  `pe-sub-api`'s in-memory config cache picks the values up immediately; if the API is
  down this logs a warning and the values apply on its next restart

## Startup behaviour

All three jobs run automatically once the application is up, in sequence:

1. `facility-ingest` against `ingest.facility-file`
2. `lp-master-ingest` against `ingest.lp-master-file`
3. `lp-records-seed` against `ingest.lp-facility-seeds-file`

The seed job runs **after** facilities and LP Master because it depends on both tables being populated. Each job logs `status / readCount / writeCount / skipCount` on completion. A failure on one job is caught and logged; the remaining jobs still run. The skip limit per job is 10 rows — exceeding it marks that job `FAILED`.

`cls-conc-limits-ingest` runs at startup **only when** `ingest.cls-conc-limits-file`
(env `CLS_CONC_LIMITS_FILE`) is set — the API's `V1_5` migration already seeds defaults,
so an unconfigured feed is skipped with a log line rather than re-fed on every boot.

Startup ingest is controlled by `ingest.run-on-startup` (default `true`; env `INGEST_RUN_ON_STARTUP`).
Set it to `false` to skip the seed jobs on boot — for example when the shared schema has not been
migrated yet, or in tests (the integration base sets it `false` so jobs don't run against the empty
embedded database, whose business tables are owned by `pe-sub-api`'s migrations). Disabling startup
ingest does **not** remove the on-demand REST triggers below.

Jobs can also be triggered on demand via REST (see [REST API](#rest-api)).

## Mock data

Development seed files are in `data/mock/`:

| File | Contents |
|---|---|
| `data/mock/facilities.csv` | 65 facilities derived from the Agent Bank Summary workbook |
| `data/mock/lp_master.csv` | 30 LP Master records across all classification tiers |
| `data/mock/lp_facility_seeds.csv` | 42 LP-to-facility assignments across 5 active facilities |
| `data/mock/cls_conc_limit_defaults.csv` | Reference feed for `cls-conc-limits-ingest` — 6 classification labels seeded to the upper bound of each `Concentration_Limits.xls` range (Excluded = 0) |

`lp_facility_seeds.csv` links LP Master records to specific facilities, producing `lp_records` rows the same way the ingestion wizard would. It upserts on `(facility_id, investor_name)`, so re-running refreshes seeded values instead of failing on duplicates. These are the default files used on startup (see `ingest.*` properties below).

## Getting started

```bash
# PostgreSQL must be running (shared with pe-sub-api)

mvn spring-boot:run
```

On startup the app creates the Spring Batch meta-tables in the shared database (if they don't exist) and immediately runs both ingest jobs against the mock data files.

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
| `SPRING_DATASOURCE_URL` | `jdbc:postgresql://localhost:5432/pesub` | JDBC connection URL |
| `SPRING_DATASOURCE_USERNAME` | `pesub` | DB username |
| `SPRING_DATASOURCE_PASSWORD` | `password` | DB password |
| `PORT` | `3003` | HTTP port |
| `LOG_PATH` | `logs` | Log output directory |
| `FACILITY_INGEST_FILE` | `data/mock/facilities.csv` | Path to facilities CSV for startup ingest |
| `LP_MASTER_INGEST_FILE` | `data/mock/lp_master.csv` | Path to LP master CSV for startup ingest |
| `LP_FACILITY_SEEDS_FILE` | `data/mock/lp_facility_seeds.csv` | Path to LP-facility seed CSV for startup ingest |
| `CLS_CONC_LIMITS_FILE` | *(unset)* | Path to classification conc-limit defaults CSV; unset → feed skipped at startup |
| `INGEST_RUN_ON_STARTUP` | `true` | Run the seed jobs on startup; set `false` to skip them |
| `INGEST_SCHEMA_WAIT_TIMEOUT` | `30s` | How long startup ingest waits for API-owned business tables |
| `INGEST_SCHEMA_WAIT_INTERVAL` | `2s` | Poll interval while waiting for business tables |
| `BB_TEMPLATE_IMPORT_ENABLED` | `true` | Import BB template workbooks from the watched directory |
| `BB_TEMPLATE_IMPORT_DIR` | `data/bb-templates` | Directory scanned for `BB-Template-Import-*.xlsx` workbooks |
| `PE_SUB_API_URL` | `http://localhost:3001` | `pe-sub-api` base URL used for idempotent template upserts and post-feed config cache reloads |
| `BB_TEMPLATE_SCAN_INTERVAL` | `30s` | How often to rescan the template directory while running |
| `BB_TEMPLATE_STABLE_AGE` | `2s` | Minimum file age before import, to avoid partially copied files |

## BB Template Imports

`pe-sub-jobs` imports structured BB template workbooks from `data/bb-templates/` on startup and
rescans the directory while running. Files are sent to `pe-sub-api` through
`POST /api/bb-templates/import?mode=upsert`, so imports are idempotent by `template_slug`:
existing templates are replaced, and missing templates are created.

Use `.partial` or `.tmp` while copying large files, then rename to `.xlsx` when complete. Temporary
Excel files beginning with `~$` are ignored.

## Logging

Logs are written to `$LOG_PATH/pe-sub-jobs.log` and rotated daily to `$LOG_PATH/archived/pe-sub-jobs.YYYY-MM-DD.log.gz` (30 days / 2 GB cap). Console output mirrors the file.

## Build

```bash
mvn package              # fat JAR → target/pe-sub-jobs-1.0.0.jar
mvn package -DskipTests
java -jar target/pe-sub-jobs-1.0.0.jar
```

## Testing

**Zonky embedded Postgres only.** Any test that boots the Spring context must extend `IntegrationTestBase`, which starts Zonky's in-process PostgreSQL (`@AutoConfigureEmbeddedDatabase(provider = ZONKY)`) — no Docker, no external `localhost:5432`. Never use H2 or Testcontainers; test behaviour must match production PostgreSQL 16 exactly.

## Project structure

```
src/main/java/com/ubs/pesubjobs/
  PeSubJobsApplication.java        @SpringBootApplication + @ConfigurationPropertiesScan
  JobStartupRunner.java            ApplicationRunner — runs both jobs on startup
  config/
    IngestProperties.java          @ConfigurationProperties(prefix = "ingest")
    FacilityIngestJobConfig.java   Job + Step + @StepScope reader + UPSERT writer
    LpMasterIngestJobConfig.java
    LpRecordsSeedJobConfig.java    Seeds lp_records by linking LP Master → facilities
  controller/
    JobController.java             POST /jobs/{jobName}
  exception/
    GlobalExceptionHandler.java    ProblemDetail error responses
  model/
    FacilityRow.java               Raw CSV record (all Strings)
    ProcessedFacility.java         Type-safe record (BigDecimal, LocalDate)
    LpMasterRow.java
    ProcessedLpMaster.java
    LpFacilitySeedRow.java         Raw seed CSV row
    ProcessedLpFacilitySeed.java   Fully resolved seed row (facility_id, lp_master_id + all LP fields)
  processor/
    FacilityRowProcessor.java      Parses dates/decimals; returns null to skip invalid rows
    LpMasterRowProcessor.java      Parses booleans; normalises blank ratings to ""
    LpFacilitySeedRowProcessor.java  JDBC lookup of facility + LP Master; resolves all FK fields
src/main/resources/
  application.yml
  logback-spring.xml
data/mock/
  facilities.csv
  lp_master.csv
```
