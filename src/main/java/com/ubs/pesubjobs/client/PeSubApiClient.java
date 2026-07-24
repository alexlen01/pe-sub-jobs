package com.ubs.pesubjobs.client;

import com.ubs.pesubjobs.config.IngestProperties;
import com.ubs.pesubjobs.model.LpFacilitySeedRow;
import com.ubs.pesubjobs.model.ProcessedFacility;
import com.ubs.pesubjobs.model.ProcessedLpMaster;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.ParameterizedTypeReference;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Component;
import org.springframework.web.client.RestClient;

import java.util.List;
import java.util.Map;

/**
 * All pe-sub-jobs writes go through pe-sub-api's SERVICE-gated bulk endpoints — this app holds
 * no database connection and issues no SQL. pe-sub-api owns the schema; the jobs own only the
 * CSV reading/parsing. Each method posts one chunk; a non-2xx response throws, which surfaces
 * to Spring Batch's fault-tolerant skip handling exactly as a failed JDBC batch write used to.
 */
@Component
public class PeSubApiClient {

    private static final Logger log = LoggerFactory.getLogger(PeSubApiClient.class);

    /** Mirror of pe-sub-api's IngestSummary response body. */
    public record ApiIngestSummary(int created, int updated, int skipped) {}

    private final RestClient rest;

    public PeSubApiClient(IngestProperties props) {
        this.rest = RestClient.builder()
                .baseUrl(props.apiBaseUrl().replaceAll("/+$", ""))
                .build();
    }

    public ApiIngestSummary ingestFacilities(List<? extends ProcessedFacility> rows) {
        return post("/api/facilities/ingest", rows);
    }

    public ApiIngestSummary ingestLpMaster(List<? extends ProcessedLpMaster> rows) {
        return post("/api/lp-master/ingest", rows);
    }

    /**
     * Wipes the LP Master table ahead of a full repopulate ("override, do not preserve") from the
     * one-off LP DB extract feed. Runs once as the lp-master ingest job's pre-step, before the
     * chunked upsert, so re-running the feed does not leave stale master rows. SERVICE-gated;
     * facility LP records are detached (not deleted) by the API.
     */
    public void clearLpMaster() {
        rest.post()
                .uri("/api/lp-master/clear")
                .retrieve()
                .toBodilessEntity();
        log.info("POST /api/lp-master/clear -> LP Master cleared for repopulate");
    }

    public ApiIngestSummary seedLpRecords(List<? extends LpFacilitySeedRow> rows) {
        return post("/api/lpRecords/seed", rows);
    }

    /**
     * Merges per-classification default conc limits into the API's cls_conc_limit_defaults
     * config map. The API persists and refreshes its in-memory cache in the same call, so no
     * follow-up /api/config/reload is needed.
     */
    public void mergeClsConcLimitDefaults(Map<String, Double> limits) {
        rest.patch()
                .uri("/api/config/cls-conc-limit-defaults")
                .contentType(MediaType.APPLICATION_JSON)
                .body(limits)
                .retrieve()
                .toBodilessEntity();
    }

    /** Readiness probe — the API up means its schema is migrated and the feeds can run. */
    public boolean isApiReady() {
        try {
            rest.get().uri("/api/ping").retrieve().toBodilessEntity();
            return true;
        } catch (Exception e) {
            log.debug("pe-sub-api readiness check failed: {}", e.getMessage());
            return false;
        }
    }

    public long getFacilityCount() {
        try {
            Map<String, Long> response = rest.get()
                    .uri("/api/facilities/count")
                    .retrieve()
                    .body(new ParameterizedTypeReference<Map<String, Long>>() {});
            return response != null ? response.getOrDefault("count", 0L) : 0L;
        } catch (Exception e) {
            log.debug("Failed to query facility count: {}", e.getMessage());
            return 0L;
        }
    }

    public long getLpMasterCount() {
        try {
            Map<String, Long> response = rest.get()
                    .uri("/api/lp-master/count")
                    .retrieve()
                    .body(new ParameterizedTypeReference<Map<String, Long>>() {});
            return response != null ? response.getOrDefault("count", 0L) : 0L;
        } catch (Exception e) {
            log.debug("Failed to query LP Master count: {}", e.getMessage());
            return 0L;
        }
    }

    public long getLpRecordCount() {
        try {
            Map<String, Long> response = rest.get()
                    .uri("/api/lpRecords/count")
                    .retrieve()
                    .body(new ParameterizedTypeReference<Map<String, Long>>() {});
            return response != null ? response.getOrDefault("count", 0L) : 0L;
        } catch (Exception e) {
            log.debug("Failed to query LP Record count: {}", e.getMessage());
            return 0L;
        }
    }

    private ApiIngestSummary post(String uri, Object body) {
        ApiIngestSummary summary = rest.post()
                .uri(uri)
                .contentType(MediaType.APPLICATION_JSON)
                .body(body)
                .retrieve()
                .body(ApiIngestSummary.class);
        log.info("POST {} -> created={} updated={} skipped={}",
                uri,
                summary != null ? summary.created() : 0,
                summary != null ? summary.updated() : 0,
                summary != null ? summary.skipped() : 0);
        return summary;
    }
}
