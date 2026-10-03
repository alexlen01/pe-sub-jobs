package com.ubs.pesubjobs;

import com.ubs.pesubjobs.config.BbTemplateImportProperties;
import com.ubs.pesubjobs.config.InternalTlsProperties;
import com.ubs.pesubjobs.security.JobsSecurityProperties;
import com.ubs.pesubjobs.storage.BbTemplateObject;
import com.ubs.pesubjobs.storage.BbTemplateStore;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.http.MediaType;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.ResourceAccessException;
import org.springframework.web.client.RestClient;

import java.io.IOException;
import java.time.Instant;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicBoolean;

@Component
public class BbTemplateDirectoryImporter implements ApplicationRunner {

    private static final Logger log = LoggerFactory.getLogger(BbTemplateDirectoryImporter.class);

    private final BbTemplateImportProperties props;
    private final BbTemplateStore store;
    private final RestClient restClient;
    // Keyed off blob/file identity rather than a filesystem Path — the same map shape works
    // whether the active profile is reading a local directory or an Azure Blob Storage container.
    // Single replica: no cross-pod locking needed for this in-memory dedupe.
    private final Map<String, Fingerprint> imported = new ConcurrentHashMap<>();
    private final AtomicBoolean scanRunning = new AtomicBoolean(false);

    // Two constructors, so the injection point has to be named explicitly.
    @Autowired
    public BbTemplateDirectoryImporter(BbTemplateImportProperties props, BbTemplateStore store,
                                       JobsSecurityProperties security, InternalTlsProperties internalTls) {
        this(props, store, security,
                internalTls.restClientBuilder("bb-template-import.api-base-url", props.apiBaseUrl()));
    }

    /**
     * Takes the builder rather than making one, so the scan's decisions — which files it posts, and
     * what it does when the API refuses one — can be exercised against a stubbed transport instead
     * of only against a live API. Not the injection point: the constructor above is.
     */
    BbTemplateDirectoryImporter(BbTemplateImportProperties props, BbTemplateStore store,
                                JobsSecurityProperties security, RestClient.Builder restClientBuilder) {
        this.props = props;
        this.store = store;
        // Template import is SERVICE-gated on pe-sub-api alongside the ANALYST screen path, so
        // this caller asserts only what it is. In gateway mode a header-less post is 401.
        this.restClient = restClientBuilder
                .baseUrl(trimTrailingSlash(props.apiBaseUrl()))
                .defaultHeader(security.getUserHeader(), security.getServiceUser())
                .defaultHeader(security.getRolesHeader(), security.getRequiredRole())
                .build();
    }

    @Override
    public void run(ApplicationArguments args) {
        scan("startup");
    }

    @Scheduled(fixedDelayString = "${bb-template-import.scan-interval:30s}")
    public void scheduledScan() {
        scan("scheduled");
    }

    private void scan(String reason) {
        if (!props.enabled()) return;
        if (!scanRunning.compareAndSet(false, true)) {
            log.debug("BB template import scan skipped reason={} cause=scan-already-running", reason);
            return;
        }
        try {
            List<BbTemplateObject> objects;
            try {
                objects = store.list();
            } catch (IOException e) {
                log.warn("BB template import scan skipped reason={} location={} error={}",
                        reason, store.describeLocation(), e.getMessage(), e);
                return;
            }
            if (!apiAvailable(reason)) {
                return;
            }

            objects.stream()
                .filter(this::isImportWorkbook)
                .takeWhile(object -> importIfChanged(object, reason))
                .forEach(object -> {});
        } finally {
            scanRunning.set(false);
        }
    }

    private boolean importIfChanged(BbTemplateObject object, String reason) {
        try {
            Fingerprint fp = new Fingerprint(object.size(), object.lastModified().toEpochMilli());
            if (!isStable(fp)) return true;
            if (fp.equals(imported.get(object.identifier()))) return true;

            MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();
            body.add("file", store.open(object.identifier()));
            restClient.post()
                .uri("/api/bb-templates/import?mode=upsert")
                .contentType(MediaType.MULTIPART_FORM_DATA)
                .body(body)
                .retrieve()
                .toBodilessEntity();

            imported.put(object.identifier(), fp);
            log.info("BB template imported reason={} file={}", reason, object.name());
            return true;
        } catch (ResourceAccessException e) {
            log.warn("BB template import scan paused reason={} file={} apiBaseUrl={} error={}",
                reason, object.name(), trimTrailingSlash(props.apiBaseUrl()), e.getMessage());
            return false;
        } catch (Exception e) {
            log.warn("BB template import failed reason={} file={} error={}", reason, object.name(), e.getMessage(), e);
            return true;
        }
    }

    private boolean apiAvailable(String reason) {
        try {
            restClient.get()
                .uri("/api/ping")
                .retrieve()
                .toBodilessEntity();
            return true;
        } catch (ResourceAccessException e) {
            log.warn("BB template import scan skipped reason={} apiBaseUrl={} error={}",
                reason, trimTrailingSlash(props.apiBaseUrl()), e.getMessage());
            return false;
        }
    }

    /**
     * Case folds against ROOT, not the JVM default: a Turkish default locale folds the "I" in
     * ".PARTIAL.XLSX" to a dotless "ı", the guard below stops matching, and a workbook still
     * being written gets posted to the API half-formed.
     */
    private boolean isImportWorkbook(BbTemplateObject object) {
        String name = object.name().toLowerCase(Locale.ROOT);
        return name.endsWith(".xlsx")
            && !name.startsWith("~$")
            && !name.endsWith(".tmp.xlsx")
            && !name.endsWith(".partial.xlsx");
    }

    private boolean isStable(Fingerprint fp) {
        return Instant.now().minus(props.stableAge()).toEpochMilli() >= fp.lastModifiedMillis();
    }

    private static String trimTrailingSlash(String raw) {
        String value = java.util.Objects.requireNonNull(raw, "bb-template-import.api-base-url must be set").trim();
        while (value.endsWith("/")) value = value.substring(0, value.length() - 1);
        return value;
    }

    private record Fingerprint(long size, long lastModifiedMillis) {}
}
