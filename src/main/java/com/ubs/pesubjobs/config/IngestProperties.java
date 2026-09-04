package com.ubs.pesubjobs.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.context.properties.bind.DefaultValue;

import java.time.Duration;

@ConfigurationProperties(prefix = "ingest")
public record IngestProperties(
        String facilityFile,
        // The group layer above the facilities. Loaded before the facility feed so a fund naming
        // its group finds it already onboarded under the name the agent printed. Blank/absent →
        // the startup runner skips it, and facilities fall back to grouping by account number.
        String umbrellaFile,
        String lpMasterFile,
        String lpFacilitySeedsFile,
        // Optional classification concentration-limit defaults feed. Blank/absent → the
        // startup runner skips it (the API seeds defaults on schema creation); the on-demand
        // /jobs/cls-conc-limits-ingest endpoint works regardless.
        String clsConcLimitsFile,
        // pe-sub-api base URL, used to reload its in-memory config cache after a config feed.
        String apiBaseUrl,
        // The only directory an on-demand /jobs trigger may read from. Caller-supplied names are
        // bare file names resolved against this root; the configured feeds above are trusted
        // configuration and are read as given, so a mock or alternate feed set still works.
        @DefaultValue("data/out") String importRoot,
        // How many of a feed's rows may fail to land before the run is a failure rather than a
        // success with a footnote. Counts rows pe-sub-api refused and rows dropped before being
        // sent. A handful is a data-quality problem for the next feed; more than that means the
        // reference data the platform now holds is not the reference data the operator loaded.
        @DefaultValue("10") long maxRowsNotLanded,
        @DefaultValue("30s") Duration schemaWaitTimeout,
        @DefaultValue("2s") Duration schemaWaitInterval,
        // When false, JobStartupRunner does not launch the seed jobs on boot. Defaults true to
        // preserve production behaviour; tests set it false so jobs don't run against the empty
        // embedded schema (the business tables are owned by pe-sub-api's migrations).
        @DefaultValue("true") boolean runOnStartup) {}
