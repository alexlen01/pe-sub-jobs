package com.ubs.pesubjobs.config;

import org.springframework.batch.infrastructure.support.transaction.ResourcelessTransactionManager;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.transaction.PlatformTransactionManager;

/**
 * DB-less Spring Batch support. pe-sub-jobs writes exclusively through pe-sub-api's bulk
 * endpoints and holds no DataSource, and Batch 6 already defaults to an in-memory
 * {@code ResourcelessJobRepository} in that case (Boot's {@code BatchAutoConfiguration}
 * registers it — do NOT declare a jobRepository/jobOperator bean here; that collides).
 *
 * <p>The only gap: the auto-configuration keeps its transaction manager as a protected
 * getter, not a bean, while the step definitions inject a {@link PlatformTransactionManager}.
 * This no-op implementation fills that hole — appropriate because chunk "transactions" here
 * are just buffered REST posts to pe-sub-api with nothing transactional to roll back.
 *
 * <p>Trade-off, accepted deliberately: no persisted run history and no restart-from-failure.
 * Every job is an idempotent full-file re-feed (the API upserts/skips server-side), so
 * replaying a feed is always safe and restart bookkeeping buys nothing.
 */
@Configuration
public class ResourcelessBatchConfig {

    @Bean
    public PlatformTransactionManager transactionManager() {
        return new ResourcelessTransactionManager();
    }
}
