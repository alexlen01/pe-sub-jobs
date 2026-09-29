package com.ubs.pesubjobs;

import com.ubs.pesubjobs.client.PeSubApiClient;
import com.ubs.pesubjobs.config.IngestProperties;
import com.ubs.pesubjobs.storage.FeedFileStore;
import com.ubs.pesubjobs.storage.LocalFeedFileStore;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.mockito.InOrder;
import org.springframework.batch.core.job.Job;
import org.springframework.batch.core.job.JobExecution;
import org.springframework.batch.core.job.parameters.JobParameters;
import org.springframework.batch.core.launch.JobOperator;
import org.springframework.batch.infrastructure.item.ExecutionContext;
import org.springframework.boot.DefaultApplicationArguments;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.OptionalLong;

import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.atLeastOnce;
import static org.mockito.Mockito.inOrder;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

/**
 * The startup decision. The seed feeds it launches replace LP Master wholesale, so it may run them
 * only on a platform it has confirmed is unseeded. A count the API could not answer is not a zero:
 * treating it as one turns an API outage into a wipe of a populated platform.
 */
class JobStartupRunnerTest {

    private final JobOperator jobOperator = mock(JobOperator.class);
    private final PeSubApiClient apiClient = mock(PeSubApiClient.class);
    private final Job umbrellaIngestJob = mock(Job.class);
    private final Job facilityIngestJob = mock(Job.class);
    private final Job lpMasterIngestJob = mock(Job.class);
    private final Job lpRecordsSeedJob = mock(Job.class);
    private final Job clsConcLimitIngestJob = mock(Job.class);

    @TempDir Path feedDir;

    private JobStartupRunner runner;
    private Path umbrellaFeed;

    @BeforeEach
    void setUp() throws Exception {
        // Written, because the runner only launches the group feed when the file is actually there:
        // a feed set extracted before umbrellas.csv existed must not turn startup into a failure.
        umbrellaFeed = feedDir.resolve("umbrellas.csv");
        Files.writeString(umbrellaFeed, "key,name\n");
        runner = newRunner(umbrellaFeed.toString());

        when(apiClient.isApiReady()).thenReturn(true);
        JobExecution execution = mock(JobExecution.class);
        // The runner reads the ingest tally the step listener published here; a real execution
        // always carries a context, so the mock has to as well.
        when(execution.getExecutionContext()).thenReturn(new ExecutionContext());
        when(jobOperator.start(any(Job.class), any(JobParameters.class))).thenReturn(execution);
    }

    private JobStartupRunner newRunner(String umbrellaFile) {
        IngestProperties props = new IngestProperties(
                "facilities.csv", umbrellaFile, "lp-master.csv", "lp-facility-seeds.csv", null,
                "http://localhost:9062", "data/out", 10,
                Duration.ofSeconds(1), Duration.ofMillis(250), true);
        // exists() only stats the identifier itself, so it does not care that importRoot here is
        // unrelated to where the umbrella feed file was actually written.
        FeedFileStore feedFileStore = new LocalFeedFileStore(props);
        return new JobStartupRunner(jobOperator, umbrellaIngestJob, facilityIngestJob,
                lpMasterIngestJob, lpRecordsSeedJob, clsConcLimitIngestJob, props, apiClient, feedFileStore);
    }

    private void counts(OptionalLong facilities, OptionalLong lpMaster, OptionalLong lpRecords) {
        when(apiClient.getFacilityCount()).thenReturn(facilities);
        when(apiClient.getLpMasterCount()).thenReturn(lpMaster);
        when(apiClient.getLpRecordCount()).thenReturn(lpRecords);
    }

    private void run() {
        runner.run(new DefaultApplicationArguments());
    }

    @Test
    void confirmedEmptyPlatform_runsTheSeedFeeds() throws Exception {
        counts(OptionalLong.of(0), OptionalLong.of(0), OptionalLong.of(0));

        run();

        verify(jobOperator, atLeastOnce()).start(any(Job.class), any(JobParameters.class));
    }

    @Test
    void groupFeedRunsBeforeTheFacilitiesThatNameIt() throws Exception {
        // A facility names its umbrella on its own row. Loaded the other way round, the first
        // member through creates the group under a name minted from an account number, and the
        // name the agent printed on the credit agreement is lost for the rest of the run.
        counts(OptionalLong.of(0), OptionalLong.of(0), OptionalLong.of(0));

        run();

        InOrder order = inOrder(jobOperator);
        order.verify(jobOperator).start(eq(umbrellaIngestJob), any(JobParameters.class));
        order.verify(jobOperator).start(eq(facilityIngestJob), any(JobParameters.class));
    }

    @Test
    void absentGroupFeed_isSkippedAndTheOtherFeedsStillRun() throws Exception {
        // A feed set extracted before umbrellas.csv existed. The group layer is simply not stated;
        // that is not a reason to leave the platform unseeded.
        runner = newRunner(feedDir.resolve("not-written.csv").toString());
        counts(OptionalLong.of(0), OptionalLong.of(0), OptionalLong.of(0));

        run();

        verify(jobOperator, never()).start(eq(umbrellaIngestJob), any(JobParameters.class));
        verify(jobOperator).start(eq(facilityIngestJob), any(JobParameters.class));
    }

    @Test
    void populatedPlatform_runsNothing() throws Exception {
        counts(OptionalLong.of(12), OptionalLong.of(0), OptionalLong.of(0));

        run();

        verify(jobOperator, never()).start(any(Job.class), any(JobParameters.class));
    }

    @Test
    void unreadableCount_runsNothing() throws Exception {
        // The LP Master count is the one that came back unknown, and LP Master is the table the
        // reload clears. Proceeding here is the destructive reading of missing evidence.
        counts(OptionalLong.of(0), OptionalLong.empty(), OptionalLong.of(0));

        run();

        verify(jobOperator, never()).start(any(Job.class), any(JobParameters.class));
    }

    @Test
    void allCountsUnreadable_runsNothing() throws Exception {
        counts(OptionalLong.empty(), OptionalLong.empty(), OptionalLong.empty());

        run();

        verify(jobOperator, never()).start(any(Job.class), any(JobParameters.class));
    }

    @Test
    void unreachableApi_runsNothingAndAsksForNoCounts() throws Exception {
        when(apiClient.isApiReady()).thenReturn(false);

        run();

        verify(jobOperator, never()).start(any(Job.class), any(JobParameters.class));
        verify(apiClient, never()).getLpMasterCount();
    }
}
