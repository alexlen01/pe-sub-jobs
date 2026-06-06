# pe-sub-jobs

Spring Boot 3.5 / Java 21 background jobs service for the PE Sub Borrowing Base Platform. Runs scheduled and triggered jobs: BB recalculation, LP data refresh, notification dispatch, and other async processing tasks.

## Stack

- Java 21 (Eclipse Temurin), Spring Boot 3.5, Maven 3.9
- Spring Data JPA (Hibernate 6), PostgreSQL 16
- Flyway — migrations applied automatically on startup from `src/main/resources/db/migration/`
- Spring Scheduling (`@EnableScheduling`) — cron and fixed-rate job execution

## Prerequisites

- Java 21
- Maven 3.9+
- Docker (for PostgreSQL)

## Getting started

```bash
# 1. Start PostgreSQL (shared with pe-sub-api)
docker compose up -d

# 2. Start the service
mvn spring-boot:run
```

Service runs at `http://localhost:3003`. Health check: `GET /health`.

## Other commands

```bash
mvn package              # build fat JAR → target/pe-sub-jobs-0.1.0.jar
mvn package -DskipTests  # skip tests during build
java -jar target/pe-sub-jobs-0.1.0.jar
```

## Environment variables

| Variable | Default | Description |
|---|---|---|
| `SPRING_DATASOURCE_URL` | `jdbc:postgresql://localhost:5432/pesub` | JDBC connection URL |
| `SPRING_DATASOURCE_USERNAME` | `pesub` | DB username |
| `SPRING_DATASOURCE_PASSWORD` | `password` | DB password |
| `PORT` | `3003` | HTTP port |
| `LOG_PATH` | `C:/Users/alexl/apps/pe-sub/logs` | Log output directory |

## Project structure

```
src/main/java/com/ubs/pesubjobs/
  PeSubJobsApplication.java
src/main/resources/
  application.yml
  db/migration/
```
