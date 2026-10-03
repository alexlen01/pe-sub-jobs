package com.ubs.pesubjobs.config;

import org.apache.commons.logging.Log;
import org.jasypt.encryption.pbe.StandardPBEStringEncryptor;
import org.jasypt.exceptions.EncryptionOperationNotPossibleException;
import org.jasypt.iv.RandomIvGenerator;
import org.springframework.boot.EnvironmentPostProcessor;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.context.config.ConfigDataEnvironmentPostProcessor;
import org.springframework.boot.logging.DeferredLogFactory;
import org.springframework.core.Ordered;
import org.springframework.core.env.ConfigurableEnvironment;
import org.springframework.core.env.EnumerablePropertySource;
import org.springframework.core.env.MapPropertySource;
import org.springframework.core.env.MutablePropertySources;
import org.springframework.core.env.PropertySource;
import org.springframework.core.env.SystemEnvironmentPropertySource;

import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Decrypts Jasypt {@code ENC(...)} property values (YAML or environment) with the master password
 * in {@code jasypt.encryptor.password}, ahead of every other source. Values are never logged.
 */
public class JasyptEnvironmentPostProcessor implements EnvironmentPostProcessor, Ordered {

    static final String PASSWORD_PROPERTY = "jasypt.encryptor.password";
    // Same defaults as the Jasypt CLI / jasypt-spring-boot, so their tooling produces usable values.
    static final String ALGORITHM = "PBEWITHHMACSHA512ANDAES_256";
    static final int KEY_OBTENTION_ITERATIONS = 1000;

    private static final Pattern ENC = Pattern.compile("^ENC\\((.+)\\)$");
    private static final String SOURCE = "jasyptDecrypted";
    private static final String ENV_SOURCE = "jasyptDecrypted-systemEnvironment";

    private final Log log;

    public JasyptEnvironmentPostProcessor(DeferredLogFactory logFactory) {
        this.log = logFactory.getLog(JasyptEnvironmentPostProcessor.class);
    }

    @Override
    public int getOrder() {
        // After application*.yml is loaded, so its ENC(...) values are visible.
        return ConfigDataEnvironmentPostProcessor.ORDER + 1;
    }

    @Override
    public void postProcessEnvironment(ConfigurableEnvironment environment, SpringApplication application) {
        Map<String, Object> decrypted = new LinkedHashMap<>();
        Map<String, Object> decryptedEnv = new LinkedHashMap<>();
        Set<String> seen = new HashSet<>();
        StandardPBEStringEncryptor encryptor = null;

        for (PropertySource<?> source : environment.getPropertySources()) {
            if (!(source instanceof EnumerablePropertySource<?> enumerable)) {
                continue;
            }
            for (String name : enumerable.getPropertyNames()) {
                // Sources iterate highest precedence first; a lower one never overrides the winner.
                if (!seen.add(name)) {
                    continue;
                }
                if (!(enumerable.getProperty(name) instanceof CharSequence raw)) {
                    continue;
                }
                Matcher m = ENC.matcher(raw.toString().trim());
                if (!m.matches()) {
                    continue;
                }
                if (encryptor == null) {
                    encryptor = encryptor(environment, name);
                }
                String plain = decrypt(encryptor, name, m.group(1));
                (source instanceof SystemEnvironmentPropertySource ? decryptedEnv : decrypted).put(name, plain);
            }
        }

        MutablePropertySources sources = environment.getPropertySources();
        if (!decrypted.isEmpty()) {
            sources.addFirst(new MapPropertySource(SOURCE, decrypted));
        }
        if (!decryptedEnv.isEmpty()) {
            // Kept as an environment source so relaxed binding (FOO_BAR -> foo.bar) still applies.
            sources.addFirst(new SystemEnvironmentPropertySource(ENV_SOURCE, decryptedEnv));
        }
        if (!decrypted.isEmpty() || !decryptedEnv.isEmpty()) {
            Set<String> names = new HashSet<>(decrypted.keySet());
            names.addAll(decryptedEnv.keySet());
            log.info("Decrypted " + names.size() + " ENC(...) properties: " + names);
        }
    }

    static StandardPBEStringEncryptor newEncryptor(String password) {
        StandardPBEStringEncryptor encryptor = new StandardPBEStringEncryptor();
        encryptor.setPassword(password);
        encryptor.setAlgorithm(ALGORITHM);
        encryptor.setKeyObtentionIterations(KEY_OBTENTION_ITERATIONS);
        encryptor.setIvGenerator(new RandomIvGenerator());
        encryptor.setStringOutputType("base64");
        return encryptor;
    }

    private static StandardPBEStringEncryptor encryptor(ConfigurableEnvironment environment, String firstEncrypted) {
        String password = environment.getProperty(PASSWORD_PROPERTY);
        if (password == null || password.isBlank()) {
            throw new IllegalStateException("Property '" + firstEncrypted + "' is ENC(...) but "
                    + PASSWORD_PROPERTY + " (JASYPT_ENCRYPTOR_PASSWORD) is not set");
        }
        return newEncryptor(password);
    }

    private static String decrypt(StandardPBEStringEncryptor encryptor, String name, String cipherText) {
        try {
            return encryptor.decrypt(cipherText);
        } catch (EncryptionOperationNotPossibleException e) {
            // No cause chained: Jasypt's messages carry nothing useful and nothing sensitive is echoed.
            throw new IllegalStateException("Property '" + name
                    + "' could not be decrypted — wrong Jasypt master password or corrupt ENC(...) value");
        }
    }
}
