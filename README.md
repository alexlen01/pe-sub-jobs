# pe-sub-jobs

Spring Batch service for loading data into the PE Sub API.

## Requirements

- Java 21
- Maven
- pe-sub-api running at `http://localhost:3001`
- Data files in `data/out/` or configured input directories

## Local startup

```bash
mvn spring-boot:run
```

- Service URL: `http://localhost:3003`
- Startup waits for `GET /api/ping` on the API.

## Security

All `/jobs` calls require service auth.

Required headers:

- `X-Auth-User`
- `X-Auth-Roles: SERVICE`

Requests without auth fail with `401`.
Requests without the `SERVICE` role fail with `403`.

## Jobs

### `facility-ingest`

- Endpoint: `POST /jobs/facility-ingest`
- Feed target: `POST /api/facilities/ingest`
- Upserts by facility name.

### `lp-master-ingest`

- Endpoint: `POST /jobs/lp-master-ingest`
- Feed target: `POST /api/lp-master/ingest`
- Upserts by investor name.

### `lp-records-seed`

- Endpoint: `POST /jobs/lp-records-seed`
- Feed target: `POST /api/lpRecords/seed`
- Seeds facility LP records.

### `cls-conc-limits-ingest`

- Endpoint: `POST /jobs/cls-conc-limits-ingest`
- Feed target: `PATCH /api/config/cls-conc-limit-defaults`

## Startup behavior

- The app waits for the API to respond on `/api/ping`.
- It runs startup feeds only when the API is confirmed empty.
- `INGEST_RUN_ON_STARTUP=false` skips startup ingestion.
- Jobs run in this order:
  1. `facility-ingest`
  2. `lp-master-ingest`
  3. `lp-records-seed`

## File selection

Use `?file=<name>` on a job endpoint to load a file from `INGEST_IMPORT_ROOT`.

- File must exist in the configured import directory.
- Only files under that root are allowed.
- Invalid or missing files return `400`.

## Environment variables

- `PORT` — default `3003`
- `LOG_PATH` — log directory
- `PE_SUB_API_URL` — API base URL, default `http://localhost:3001`
- `FACILITY_INGEST_FILE` — startup facility input
- `LP_MASTER_INGEST_FILE` — startup LP Master input
- `LP_FACILITY_SEEDS_FILE` — startup LP seed input
- `INGEST_IMPORT_ROOT` — allowed job file directory
- `INGEST_MAX_ROWS_NOT_LANDED` — fail threshold for job runs
- `INGEST_RUN_ON_STARTUP` — startup feed toggle
- `INGEST_SCHEMA_WAIT_TIMEOUT` — API availability wait timeout
- `INGEST_SCHEMA_WAIT_INTERVAL` — API availability poll interval
- `BB_TEMPLATE_IMPORT_ENABLED` — template watch toggle
- `BB_TEMPLATE_IMPORT_DIR` — workbook import directory
- `BB_TEMPLATE_SCAN_INTERVAL` — rescan interval
- `BB_TEMPLATE_STABLE_AGE` — file stability wait before import

## BB template import

- Watches `data/bb-templates` for `*.xlsx` files.
- Imports new or changed templates with `POST /api/bb-templates/import?mode=upsert`.
- Skips files still changing during a scan.

## Logging

- Logs write to `logs/pe-sub-jobs.log`.
- Rotated files are stored under `logs/archived/`.

## Build and run

```bash
mvn clean compile
mvn clean test
mvn package
java -jar target/pe-sub-jobs-1.0.0.jar
```

## Operational warnings

- `docker compose down -v` is not part of this service; reset database state in the API layer.
- Startup ingest is skipped when the API is not ready.
- A job is failed when `rowsNotLanded` exceeds the configured threshold.
- Use `INGEST_RUN_ON_STARTUP=false` when you do not want startup loads to run.

## Troubleshooting

### API not ready

- Confirm the API is running.
- Confirm the configured API URL is correct.
- Check startup wait timeout and poll settings.

### Job fails

- Check `rowsNotLanded` in the job result.
- Confirm the target file exists in the configured import directory.
- Confirm the file contents match the expected feed format.
- Confirm the API is accepting the feed request.

### Auth failure

- Confirm `X-Auth-User` is present.
- Confirm `X-Auth-Roles` includes `SERVICE`.
- Confirm the request is sent to the correct server.

### Template import not happening

- Confirm `BB_TEMPLATE_IMPORT_ENABLED=true`.
- Confirm the directory contains valid `.xlsx` files.
- Confirm files are stable long enough to pass the age threshold.

