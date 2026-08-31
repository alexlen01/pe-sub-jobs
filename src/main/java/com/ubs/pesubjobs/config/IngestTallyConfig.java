package com.ubs.pesubjobs.config;

import org.springframework.batch.core.configuration.annotation.JobScope;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

/**
 * Wiring for the per-run ingest tally and the threshold it is judged against. Kept in one place so
 * every feed job is counted the same way.
 */
@Configuration
public class IngestTallyConfig {

    /** One tally per run — job-scoped, so counts never carry from one run into the next. */
    @Bean
    @JobScope
    public IngestTally ingestTally() {
        return new IngestTally();
    }

    @Bean
    public IngestTallyListener ingestTallyListener(IngestTally ingestTally, IngestProperties props) {
        return new IngestTallyListener(ingestTally, props.maxRowsNotLanded());
    }
}
