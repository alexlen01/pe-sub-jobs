package com.ubs.pesubjobs.config;

import org.junit.jupiter.api.Test;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.logging.DeferredLogs;
import org.springframework.core.env.MapPropertySource;
import org.springframework.core.env.StandardEnvironment;
import org.springframework.core.env.SystemEnvironmentPropertySource;

import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class JasyptEnvironmentPostProcessorTest {

    private static final String PASSWORD = "master-test-password";

    private final JasyptEnvironmentPostProcessor processor = new JasyptEnvironmentPostProcessor(new DeferredLogs());

    private static String enc(String plain) {
        return "ENC(" + JasyptEnvironmentPostProcessor.newEncryptor(PASSWORD).encrypt(plain) + ")";
    }

    private static StandardEnvironment environment(Map<String, Object> config, Map<String, Object> env) {
        StandardEnvironment environment = new StandardEnvironment();
        environment.getPropertySources().addFirst(new MapPropertySource("config", config));
        environment.getPropertySources().addFirst(new SystemEnvironmentPropertySource("test-systemEnvironment", env));
        return environment;
    }

    @Test
    void decryptsYamlAndEnvironmentValues() {
        StandardEnvironment environment = environment(
                Map.of("app.internal-tls.key-store-password", enc("ks-secret"), "plain", "unchanged"),
                Map.of("JASYPT_ENCRYPTOR_PASSWORD", PASSWORD, "SPRING_DATASOURCE_PASSWORD", enc("db-secret")));

        processor.postProcessEnvironment(environment, new SpringApplication());

        assertThat(environment.getProperty("app.internal-tls.key-store-password")).isEqualTo("ks-secret");
        assertThat(environment.getProperty("SPRING_DATASOURCE_PASSWORD")).isEqualTo("db-secret");
        // Relaxed name still resolves for an env-sourced value.
        assertThat(environment.getProperty("spring.datasource.password")).isEqualTo("db-secret");
        assertThat(environment.getProperty("plain")).isEqualTo("unchanged");
    }

    @Test
    void higherPrecedencePlainValueWinsOverLowerEncryptedOne() {
        StandardEnvironment environment = environment(Map.of("secret", enc("from-yaml")),
                Map.of("JASYPT_ENCRYPTOR_PASSWORD", PASSWORD, "secret", "from-env"));

        processor.postProcessEnvironment(environment, new SpringApplication());

        assertThat(environment.getProperty("secret")).isEqualTo("from-env");
    }

    @Test
    void noEncryptedValuesNeedsNoPassword() {
        StandardEnvironment environment = environment(Map.of("plain", "value"), Map.of());

        processor.postProcessEnvironment(environment, new SpringApplication());

        assertThat(environment.getProperty("plain")).isEqualTo("value");
    }

    @Test
    void encryptedValueWithoutPasswordFailsStartup() {
        StandardEnvironment environment = environment(Map.of("secret", enc("x")), Map.of());

        assertThatThrownBy(() -> processor.postProcessEnvironment(environment, new SpringApplication()))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("'secret'")
                .hasMessageContaining("JASYPT_ENCRYPTOR_PASSWORD");
    }

    @Test
    void wrongPasswordFailsWithoutEchoingTheValue() {
        String cipher = enc("do-not-leak");
        StandardEnvironment environment = environment(Map.of("secret", cipher),
                Map.of("JASYPT_ENCRYPTOR_PASSWORD", "wrong"));

        assertThatThrownBy(() -> processor.postProcessEnvironment(environment, new SpringApplication()))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("'secret' could not be decrypted")
                .hasMessageNotContaining(cipher)
                .hasMessageNotContaining("do-not-leak");
    }
}
