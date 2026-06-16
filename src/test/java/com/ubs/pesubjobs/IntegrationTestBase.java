package com.ubs.pesubjobs;

import io.zonky.test.db.AutoConfigureEmbeddedDatabase;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;

// Zonky embedded PostgreSQL: a real Postgres binary started in-process — no Docker daemon
// required — so JPA auto-configuration and Flyway connect to a genuine Postgres instead of
// requiring a running localhost:5432 database. Mirrors pe-sub-api's IntegrationTestBase.
//
// provider = ZONKY  -> spin up the bundled embedded-postgres binary for this platform.
// refresh  = NEVER (default) -> the database is bound to the shared Spring test context and
//            lives for the whole run. Test classes remain responsible for their own teardown.
//
// IMPORTANT: every test that boots the Spring context (@SpringBootTest) or otherwise needs a
// database MUST extend this base. Do NOT point tests at an external PostgreSQL instance and do
// NOT introduce H2 or any other database engine — only Zonky's embedded Postgres is permitted,
// so test behaviour matches the production Postgres exactly.
@SpringBootTest(webEnvironment = SpringBootTest.WebEnvironment.RANDOM_PORT)
@AutoConfigureMockMvc
@AutoConfigureEmbeddedDatabase(provider = AutoConfigureEmbeddedDatabase.DatabaseProvider.ZONKY)
public abstract class IntegrationTestBase {
}
