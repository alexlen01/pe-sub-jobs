# syntax=docker/dockerfile:1.7
#
# This image ships the Spring Boot service only. data/import (raw LP DB Export drop) and data/out
# (generated seed CSVs) are gitignored, environment-specific, and confidential — they are never
# baked in; mount them at runtime (see pe-sub-jobs-helm-chart's `data` volume). data/reference and
# scripts/ are offline tooling inputs/outputs consumed by the Python scripts that produce data/out,
# not by this service, so they are not copied either.

## ---- Build ----
FROM maven:3.9.9-eclipse-temurin-21-alpine AS build
WORKDIR /build

# Resolve dependencies in their own layer so a source-only change doesn't re-download the world.
COPY pom.xml .
RUN --mount=type=cache,target=/root/.m2 mvn -B -ntp dependency:go-offline

COPY src ./src
RUN --mount=type=cache,target=/root/.m2 mvn -B -ntp -DskipTests clean package

## ---- Runtime ----
FROM eclipse-temurin:21-jre-alpine
RUN addgroup -S pesub && adduser -S pesub -G pesub
# WORKDIR must stay /app: ingest.*-file and bb-template-import.directory in application.yml default
# to relative paths (data/out/..., data/import/...) resolved against the process's working
# directory, not an absolute location.
WORKDIR /app
COPY --from=build /build/target/*.jar app.jar
RUN mkdir -p /app/data && chown -R pesub:pesub /app
USER pesub:pesub

# The app's actual listen port comes from a profile-specific env var set at deploy time
# (DEV_JOBS_PORT / QA_JOBS_PORT / PROD_JOBS_PORT in a cluster, PORT for a bare `docker run`) — see
# application.yml. 8080 here only documents the image's expected default.
EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
  CMD wget -qO- "http://localhost:${PORT:-8080}/pe-sub-jobs/actuator/health" | grep -q '"status":"UP"' || exit 1

ENTRYPOINT ["java", "-jar", "/app/app.jar"]
