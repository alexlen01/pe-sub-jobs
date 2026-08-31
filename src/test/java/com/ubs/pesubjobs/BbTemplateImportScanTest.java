package com.ubs.pesubjobs;

import com.ubs.pesubjobs.config.BbTemplateImportProperties;
import com.ubs.pesubjobs.security.JobsSecurityProperties;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.springframework.http.HttpMethod;
import org.springframework.http.HttpStatus;
import org.springframework.test.web.client.MockRestServiceServer;
import org.springframework.web.client.RestClient;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;

import static org.springframework.test.web.client.match.MockRestRequestMatchers.header;
import static org.springframework.test.web.client.match.MockRestRequestMatchers.method;
import static org.springframework.test.web.client.match.MockRestRequestMatchers.requestTo;
import static org.springframework.test.web.client.response.MockRestResponseCreators.withServerError;
import static org.springframework.test.web.client.response.MockRestResponseCreators.withStatus;
import static org.springframework.test.web.client.response.MockRestResponseCreators.withSuccess;

/**
 * The directory watcher decides two things on its own: which files in the drop directory are a
 * finished template worth posting, and what to do when one of them is refused. Both decisions are
 * invisible in production — a wrongly skipped file just never appears, and a refused one is a log
 * line — so they are pinned here.
 *
 * <p>The importer holds no parser: pe-sub-api reads the workbook. What this service must get right
 * is not sending a file that is still being written, and not letting one bad file stop the rest.
 */
class BbTemplateImportScanTest {

    private static final String API = "http://pe-sub-api.test";

    @TempDir Path dropDirectory;

    private MockRestServiceServer server;
    private BbTemplateDirectoryImporter importer;

    @BeforeEach
    void setUp() {
        RestClient.Builder builder = RestClient.builder();
        server = MockRestServiceServer.bindTo(builder).build();
        // stableAge zero: files written a moment ago count as finished, so the test does not
        // have to wait out the production settling window to exercise anything else.
        BbTemplateImportProperties props = new BbTemplateImportProperties(
                true, dropDirectory.toString(), API, Duration.ofSeconds(30), Duration.ZERO);
        importer = new BbTemplateDirectoryImporter(props, new JobsSecurityProperties(), builder);
    }

    @Test
    void postsFinishedWorkbooksAndSkipsEverythingElseInTheDirectory() throws IOException {
        // Everything a real drop directory accumulates alongside the templates.
        write("citibank.xlsx");
        write("legacy.xls");            // the API's import path takes OOXML only
        write("~$citibank.xlsx");       // Excel's lock file for an open workbook
        write("upload.tmp.xlsx");       // a copy still in progress
        write("upload.partial.xlsx");
        write("OTHER.PARTIAL.XLSX");    // same guard, as Windows hands the name back
        write("notes.txt");

        expectPing();
        // Only the one finished workbook, and it carries the service identity pe-sub-api gates on.
        server.expect(requestTo(API + "/api/bb-templates/import?mode=upsert"))
                .andExpect(method(HttpMethod.POST))
                .andExpect(header("X-Auth-User", "pe-sub-jobs"))
                .andExpect(header("X-Auth-Roles", "SERVICE"))
                .andRespond(withSuccess());

        importer.scheduledScan();

        server.verify();
    }

    @Test
    void aFileTheApiRefuses_doesNotStopTheFilesAfterIt() throws IOException {
        write("a-broken.xlsx");
        write("b-good.xlsx");

        expectPing();
        // Scanned in name order, so the refused file is reached first.
        server.expect(requestTo(API + "/api/bb-templates/import?mode=upsert"))
                .andRespond(withStatus(HttpStatus.UNPROCESSABLE_CONTENT));
        server.expect(requestTo(API + "/api/bb-templates/import?mode=upsert"))
                .andRespond(withSuccess());

        importer.scheduledScan();

        server.verify();
    }

    /**
     * A file already posted is not posted again on the next sweep — the scan runs every few
     * seconds, and re-importing an unchanged template would rewrite the registry on a timer.
     */
    @Test
    void anUnchangedFileIsNotReposted() throws IOException {
        write("citibank.xlsx");

        expectPing();
        server.expect(requestTo(API + "/api/bb-templates/import?mode=upsert")).andRespond(withSuccess());
        expectPing();

        importer.scheduledScan();
        importer.scheduledScan();

        server.verify();
    }

    /**
     * A refused file is not remembered as imported, so the next sweep tries it again — the API
     * being unable to read it now does not mean it will still be unreadable after a redeploy.
     */
    @Test
    void aRefusedFileIsRetriedOnTheNextSweep() throws IOException {
        write("citibank.xlsx");

        expectPing();
        server.expect(requestTo(API + "/api/bb-templates/import?mode=upsert")).andRespond(withServerError());
        expectPing();
        server.expect(requestTo(API + "/api/bb-templates/import?mode=upsert")).andRespond(withSuccess());

        importer.scheduledScan();
        importer.scheduledScan();

        server.verify();
    }

    private void expectPing() {
        server.expect(requestTo(API + "/api/ping"))
                .andExpect(method(HttpMethod.GET))
                .andRespond(withSuccess());
    }

    private void write(String fileName) throws IOException {
        Files.writeString(dropDirectory.resolve(fileName), "not a real workbook — the API parses, not this service");
    }
}
