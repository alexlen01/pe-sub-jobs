# pe-sub-jobs

Spring Boot 3.5 / Java 21 / Spring Batch 5 data ingestion service for the PE Sub Borrowing Base Platform. Reads flat-file CSV exports and upserts records into the shared PostgreSQL database. Runs at **`http://localhost:3003`**.

## Stack

- Java 21, Spring Boot 3.5.15, Maven 3.9
- Spring Batch 5 — chunk-oriented processing, fault-tolerant skip policy
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

## Startup behaviour

All three jobs run automatically once the application is up, in sequence:

1. `facility-ingest` against `ingest.facility-file`
2. `lp-master-ingest` against `ingest.lp-master-file`
3. `lp-records-seed` against `ingest.lp-facility-seeds-file`

The seed job runs **after** facilities and LP Master because it depends on both tables being populated. Each job logs `status / readCount / writeCount / skipCount` on completion. A failure on one job is caught and logged; the remaining jobs still run. The skip limit per job is 10 rows — exceeding it marks that job `FAILED`.

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

`lp_facility_seeds.csv` links LP Master records to specific facilities, producing `lp_records` rows the same way the ingestion wizard would. Uses `ON CONFLICT DO NOTHING` — safe to re-run. These are the default files used on startup (see `ingest.*` properties below).

## Getting started

```bash
# PostgreSQL must be running (shared with pe-sub-api)

mvnw spring-boot:run
```

On startup the app creates the Spring Batch meta-tables in the shared database (if they don't exist) and immediately runs both ingest jobs against the mock data files.

## REST API

Both jobs can be triggered independently after startup:

```
POST /jobs/facility-ingest?filePath=<absolute-or-relative-path>
POST /jobs/lp-master-ingest?filePath=<absolute-or-relative-path>
POST /jobs/lp-records-seed?filePath=<absolute-or-relative-path>
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
| `LOG_PATH` | `C:/Users/alexl/apps/pe-sub/logs` | Log output directory |
| `FACILITY_INGEST_FILE` | `data/mock/facilities.csv` | Path to facilities CSV for startup ingest |
| `LP_MASTER_INGEST_FILE` | `data/mock/lp_master.csv` | Path to LP master CSV for startup ingest |
| `LP_FACILITY_SEEDS_FILE` | `data/mock/lp_facility_seeds.csv` | Path to LP-facility seed CSV for startup ingest |
| `INGEST_RUN_ON_STARTUP` | `true` | Run the seed jobs on startup; set `false` to skip them |

## Logging

Logs are written to `$LOG_PATH/pe-sub-jobs.log` and rotated daily to `$LOG_PATH/archived/pe-sub-jobs.YYYY-MM-DD.log.gz` (30 days / 2 GB cap). Console output mirrors the file.

## Build

```bash
mvnw package              # fat JAR → target/pe-sub-jobs-0.1.0.jar
mvnw package -DskipTests
java -jar target/pe-sub-jobs-0.1.0.jar
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
