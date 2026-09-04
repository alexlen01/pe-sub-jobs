package com.ubs.pesubjobs.client;

import com.ubs.pesubjobs.config.IngestProperties;
import com.ubs.pesubjobs.model.LpFacilitySeedRow;
import com.ubs.pesubjobs.model.ProcessedFacility;
import com.ubs.pesubjobs.model.ProcessedLpMaster;
import com.ubs.pesubjobs.model.ProcessedUmbrella;
import com.ubs.pesubjobs.security.JobsSecurityProperties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.ParameterizedTypeReference;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Component;
import org.springframework.web.client.RestClient;

import java.util.List;
import java.util.Map;
import java.util.OptionalLong;

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

    public PeSubApiClient(IngestProperties props, JobsSecurityProperties security) {
        this.rest = RestClient.builder()
                .baseUrl(props.apiBaseUrl().replaceAll("/+$", ""))
                // The ingest, seed and clear routes are SERVICE-gated on pe-sub-api. In gateway
                // mode a header-less feed is rejected 401, so this service asserts its own
                // identity on every call — which is also what pe-sub-api's audit trail attributes.
                .defaultHeader(security.getUserHeader(), security.getServiceUser())
                .defaultHeader(security.getRolesHeader(), security.getRequiredRole())
                .build();
    }

    public ApiIngestSummary ingestFacilities(List<? extends ProcessedFacility> rows) {
        return post("/api/facilities/ingest", rows);
    }

    /**
     * The group layer, posted ahead of the facilities that name it — see the umbrella ingest job.
     *
     * <p>Posted to the master-agreement route, which is what the group has always been: a credit
     * agreement several funds borrow under. The API still answers on {@code /api/umbrellas} as a
     * deprecated alias, so a jobs build deployed before this change keeps loading — the two paths
     * are one handler, so neither can drift from the other.
     */
    public ApiIngestSummary ingestUmbrellas(List<? extends ProcessedUmbrella> rows) {
        return post("/api/master-agreements/ingest", rows);
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

    public OptionalLong getFacilityCount() {
        return count("/api/facilities/count");
    }

    public OptionalLong getLpMasterCount() {
        return count("/api/lp-master/count");
    }

    public OptionalLong getLpRecordCount() {
        return count("/api/lpRecords/count");
    }

    /**
     * A row count from the API, or empty when it could not be established.
     *
     * <p>The distinction is the point: these counts decide whether startup treats the platform as
     * unseeded, and a reload that follows that decision replaces LP Master wholesale. Answering a
     * failed query with {@code 0} makes an API outage indistinguishable from an empty database and
     * hands the caller a destructive decision on evidence it does not have. Empty says so instead,
     * and the failure is logged at WARN because it is an operational event, not a detail.
     */
    private OptionalLong count(String uri) {
        try {
            Map<String, Long> response = rest.get()
                    .uri(uri)
                    .retrieve()
                    .body(new ParameterizedTypeReference<Map<String, Long>>() {});
            Long value = response != null ? response.get("count") : null;
            if (value == null) {
                log.warn("GET {} returned no count - treating the row count as unknown", uri);
                return OptionalLong.empty();
            }
            return OptionalLong.of(value);
        } catch (Exception e) {
            log.warn("GET {} failed: {} - treating the row count as unknown", uri, e.getMessage());
            return OptionalLong.empty();
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
