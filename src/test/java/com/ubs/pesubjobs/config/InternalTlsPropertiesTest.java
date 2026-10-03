package com.ubs.pesubjobs.config;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class InternalTlsPropertiesTest {

    private static final String PASSWORD = "test-only-changeit";

    private static InternalTlsProperties tls(boolean enabled) {
        return new InternalTlsProperties(enabled, "classpath:tls/service.p12", PASSWORD, "PKCS12",
                "classpath:tls/truststore.p12", PASSWORD, "PKCS12");
    }

    @Test
    void enabledRefusesPlainHttpApiUrl() {
        assertThatThrownBy(() -> tls(true).restClientBuilder("ingest.api-base-url", "http://pe-sub-api/pe-sub-api"))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("ingest.api-base-url");
    }

    @Test
    void enabledBuildsMutualTlsClientFromStores() {
        assertThat(tls(true).restClientBuilder("ingest.api-base-url", "https://pe-sub-api:9062/pe-sub-api").build())
                .isNotNull();
    }

    @Test
    void disabledLeavesWorkstationHttpAlone() {
        assertThat(tls(false).restClientBuilder("ingest.api-base-url", "http://localhost:9062/pe-sub-api").build())
                .isNotNull();
    }

    @Test
    void toStringNeverCarriesPasswords() {
        assertThat(tls(true).toString()).doesNotContain(PASSWORD);
    }
}
