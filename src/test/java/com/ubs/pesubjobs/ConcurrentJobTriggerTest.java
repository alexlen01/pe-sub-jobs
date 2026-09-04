package com.ubs.pesubjobs;

import com.ubs.pesubjobs.client.PeSubApiClient.ApiIngestSummary;
import com.ubs.pesubjobs.controller.JobController;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.security.authentication.UsernamePasswordAuthenticationToken;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.authority.AuthorityUtils;
import org.springframework.web.server.ResponseStatusException;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.ArgumentMatchers.anyList;
import static org.mockito.Mockito.when;

/**
 * Feeds replace reference data wholesale, and the LP Master feed clears the table before it
 * repopulates it. Two overlapping runs of the same feed can therefore have the second clear away
 * what the first has already written — so a feed already in flight refuses a second trigger rather
 * than racing it. Different feeds are independent and are not serialised against each other.
 */
class ConcurrentJobTriggerTest extends IntegrationTestBase {

    @Autowired JobController jobController;

    private static final Authentication CALLER = new UsernamePasswordAuthenticationToken(
            "pe-sub-scheduler", "n/a", AuthorityUtils.createAuthorityList("ROLE_SERVICE"));

    @Test
    void aSecondTriggerOfARunningFeed_isRefusedRatherThanRacedAgainstTheFirst() throws Exception {
        CountDownLatch insideTheRun = new CountDownLatch(1);
        CountDownLatch releaseTheRun = new CountDownLatch(1);
        when(apiClient.ingestFacilities(anyList())).thenAnswer(invocation -> {
            insideTheRun.countDown();
            releaseTheRun.await(10, TimeUnit.SECONDS);
            return new ApiIngestSummary(1, 0, 0);
        });

        AtomicReference<Exception> firstRunFailure = new AtomicReference<>();
        Thread firstRun = new Thread(() -> {
            try {
                jobController.trigger("facility-ingest", "facilities.csv", CALLER);
            } catch (Exception e) {
                firstRunFailure.set(e);
            }
        });
        firstRun.start();
        assertThat(insideTheRun.await(10, TimeUnit.SECONDS))
                .as("the first run reached the API call and is still in flight")
                .isTrue();

        try {
            ResponseStatusException refusal = org.junit.jupiter.api.Assertions.assertThrows(
                    ResponseStatusException.class,
                    () -> jobController.trigger("facility-ingest", "facilities.csv", CALLER));
            assertThat(refusal.getStatusCode()).isEqualTo(HttpStatus.CONFLICT);
            assertThat(refusal.getReason()).contains("already running");
        } finally {
            releaseTheRun.countDown();
            firstRun.join(10_000);
        }

        assertThat(firstRunFailure.get()).as("the first run was left alone to finish").isNull();
    }

    /** A different feed is a different dataset; holding it behind an unrelated run buys nothing. */
    @Test
    void aDifferentFeed_runsWhileAnotherIsInFlight() throws Exception {
        CountDownLatch insideTheRun = new CountDownLatch(1);
        CountDownLatch releaseTheRun = new CountDownLatch(1);
        when(apiClient.ingestFacilities(anyList())).thenAnswer(invocation -> {
            insideTheRun.countDown();
            releaseTheRun.await(10, TimeUnit.SECONDS);
            return new ApiIngestSummary(1, 0, 0);
        });

        Thread firstRun = new Thread(() -> {
            try {
                jobController.trigger("facility-ingest", "facilities.csv", CALLER);
            } catch (Exception ignored) {
                // asserted by the test above; this one is only interested in the second feed
            }
        });
        firstRun.start();
        assertThat(insideTheRun.await(10, TimeUnit.SECONDS)).isTrue();

        try {
            // The classification-limits feed has no configured file, so this gets as far as the
            // feed-path check and is turned back there — which is the point: it was let through
            // the concurrency guard rather than refused as "already running".
            ResponseStatusException outcome = org.junit.jupiter.api.Assertions.assertThrows(
                    ResponseStatusException.class,
                    () -> jobController.trigger("cls-conc-limits-ingest", null, CALLER));
            assertThat(outcome.getStatusCode()).isEqualTo(HttpStatus.BAD_REQUEST);
        } finally {
            releaseTheRun.countDown();
            firstRun.join(10_000);
        }
    }

    /** A refused second trigger must not latch the name: the feed has to be runnable afterwards. */
    @Test
    void theClaimIsReleasedWhenTheRunEnds() throws Exception {
        when(apiClient.ingestFacilities(anyList())).thenReturn(new ApiIngestSummary(1, 0, 0));

        jobController.trigger("facility-ingest", "facilities.csv", CALLER);
        // No wait between the two triggers. The runId is a millisecond clock reading, so a pause
        // here would be guarding against the second run being rejected as a replay of the first —
        // but ResourcelessBatchConfig runs Batch on the in-memory ResourcelessJobRepository, which
        // keeps no instance history to match a replay against. Sleeping for a distinct millisecond
        // guarded nothing and made the test depend on the platform's clock granularity, which on
        // Windows can exceed the pause.
        ResponseEntity<?> second = jobController.trigger("facility-ingest", "facilities.csv", CALLER);

        assertThat(second.getStatusCode().is2xxSuccessful()).isTrue();
    }
}
