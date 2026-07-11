package com.ubs.pesubjobs.client;

import com.ubs.pesubjobs.config.IngestProperties;
import com.ubs.pesubjobs.model.LpFacilitySeedRow;
import com.ubs.pesubjobs.model.ProcessedFacility;
import com.ubs.pesubjobs.model.ProcessedLpMaster;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
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
