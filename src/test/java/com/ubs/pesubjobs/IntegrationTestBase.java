package com.ubs.pesubjobs;

import io.zonky.test.db.AutoConfigureEmbeddedDatabase;
import org.springframework.boot.webmvc.test.autoconfigure.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.TestPropertySource;

// Zonky embedded PostgreSQL: a real Postgres binary started in-process — no Docker daemon
// required — so JPA auto-configuration and Flyway connect to a genuine Postgres instead of
// requiring a running localhost:5432 database. Mirrors pe-sub-api's IntegrationTestBase.
//
// provider = EMBEDDED -> spin up the bundled embedded-postgres binary for this platform.
// refresh  = NEVER (default) -> the database is bound to the shared Spring test context and
//            lives for the whole run. Test classes remain responsible for their own teardown.
//
// IMPORTANT: every test that boots the Spring context (@SpringBootTest) or otherwise needs a
// database MUST extend this base. Do NOT point tests at an external PostgreSQL instance and do
// NOT introduce H2 or any other database engine — only Zonky's embedded Postgres is permitted,
// so test behaviour matches the production Postgres exactly.
// Startup ingest is disabled here: the embedded DB has only Spring Batch's meta-tables, not the
// business tables (facilities / lp_master / lp_records) that pe-sub-api's migrations own, so
// running the seed jobs on boot would fail and log noise. Context-load tests validate wiring only;
// job behaviour should be covered by dedicated tests that first provision a schema.
@SpringBootTest(webEnvironment = SpringBootTest.WebEnvironment.RANDOM_PORT)
@AutoConfigureMockMvc
@AutoConfigureEmbeddedDatabase(provider = AutoConfigureEmbeddedDatabase.DatabaseProvider.EMBEDDED)
@TestPropertySource(properties = {
    "ingest.run-on-startup=false",
    "bb-template-import.enabled=false"
})
public abstract class IntegrationTestBase {
}
