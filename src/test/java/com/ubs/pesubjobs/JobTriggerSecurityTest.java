package com.ubs.pesubjobs;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.test.web.servlet.MockMvc;

import static org.springframework.security.test.web.servlet.request.SecurityMockMvcRequestPostProcessors.csrf;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.jsonPath;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

/**
 * The job trigger runs feeds that replace facility and LP Master data wholesale, so it is
 * unreachable without a gateway-asserted SERVICE identity, and it reads only from the configured
 * import directory.
 */
class JobTriggerSecurityTest extends IntegrationTestBase {

    @Autowired MockMvc mockMvc;

    @Test
    void withoutIdentityHeaders_isUnauthorized() throws Exception {
        mockMvc.perform(post("/jobs/facility-ingest").with(csrf()))
                .andExpect(status().isUnauthorized());
    }

    @Test
    void withNonServiceRole_isForbidden() throws Exception {
        mockMvc.perform(post("/jobs/facility-ingest")
                        .header("X-Auth-User", "js25029")
                        .header("X-Auth-Roles", "ANALYST,MANAGER"))
                .andExpect(status().isForbidden());
    }

    @Test
    void withServiceRole_runsTheConfiguredFeed() throws Exception {
        mockMvc.perform(post("/jobs/facility-ingest")
                        .header("X-Auth-User", "pe-sub-scheduler")
                        .header("X-Auth-Roles", "SERVICE")
                        .param("file", "facilities.csv"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.status").value("COMPLETED"))
                .andExpect(jsonPath("$.readCount").value(1));
    }

    @Test
    void unknownJobName_isNotFound() throws Exception {
        mockMvc.perform(post("/jobs/definitely-not-a-job")
                        .header("X-Auth-User", "pe-sub-scheduler")
                        .header("X-Auth-Roles", "SERVICE"))
                .andExpect(status().isNotFound());
    }

    /**
     * The pre-fix contract took an absolute {@code filePath}. A traversal attempt must now be
     * rejected before any file is opened, and the response must not disclose the import root.
     */
    @Test
    void traversalOutsideTheImportRoot_isRejected() throws Exception {
        mockMvc.perform(post("/jobs/facility-ingest")
                        .header("X-Auth-User", "pe-sub-scheduler")
                        .header("X-Auth-Roles", "SERVICE")
                        .param("file", "../../../../Windows/win.ini"))
                .andExpect(status().isBadRequest())
                .andExpect(jsonPath("$.title").value("Feed File Not Accepted"));
    }

    @Test
    void absolutePathAsFileName_isRejected() throws Exception {
        mockMvc.perform(post("/jobs/facility-ingest")
                        .header("X-Auth-User", "pe-sub-scheduler")
                        .header("X-Auth-Roles", "SERVICE")
                        .param("file", "C:\\Windows\\win.ini"))
                .andExpect(status().isBadRequest());
    }

    @Test
    void nonCsvFileName_isRejected() throws Exception {
        mockMvc.perform(post("/jobs/facility-ingest")
                        .header("X-Auth-User", "pe-sub-scheduler")
                        .header("X-Auth-Roles", "SERVICE")
                        .param("file", "application.yml"))
                .andExpect(status().isBadRequest());
    }

    @Test
    void unknownFileInsideTheImportRoot_isRejected() throws Exception {
        mockMvc.perform(post("/jobs/facility-ingest")
                        .header("X-Auth-User", "pe-sub-scheduler")
                        .header("X-Auth-Roles", "SERVICE")
                        .param("file", "no-such-feed.csv"))
                .andExpect(status().isBadRequest());
    }
}
