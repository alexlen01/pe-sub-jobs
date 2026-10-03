package com.ubs.pesubjobs.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.context.properties.bind.DefaultValue;
import org.springframework.boot.ssl.SslBundle;
import org.springframework.boot.ssl.SslBundleKey;
import org.springframework.boot.ssl.SslOptions;
import org.springframework.boot.ssl.jks.JksSslStoreBundle;
import org.springframework.boot.ssl.jks.JksSslStoreDetails;
import org.springframework.http.client.JdkClientHttpRequestFactory;
import org.springframework.web.client.RestClient;

import javax.net.ssl.SSLParameters;
import java.net.http.HttpClient;
import java.util.Set;

/**
 * Mutual TLS for service-to-service calls. One key store (this service's certificate) and one
 * trust store (the internal CA) serve both this service's own port and its outbound clients.
 * Passwords are expected as Jasypt ENC(...) values.
 */
@ConfigurationProperties(prefix = "app.internal-tls")
public record InternalTlsProperties(
        boolean enabled,
        String keyStore,
        String keyStorePassword,
        @DefaultValue("PKCS12") String keyStoreType,
        String trustStore,
        String trustStorePassword,
        @DefaultValue("PKCS12") String trustStoreType) {

    static final String PROTOCOL = "TLSv1.3";

    /** Builds the bundle; fails fast on missing stores rather than on the first call. */
    public SslBundle sslBundle() {
        require(keyStore, "key-store");
        require(keyStorePassword, "key-store-password");
        require(trustStore, "trust-store");
        require(trustStorePassword, "trust-store-password");
        JksSslStoreBundle stores = new JksSslStoreBundle(
                new JksSslStoreDetails(keyStoreType, null, keyStore, keyStorePassword),
                new JksSslStoreDetails(trustStoreType, null, trustStore, trustStorePassword));
        return SslBundle.of(stores, SslBundleKey.of(keyStorePassword), SslOptions.of(null, Set.of(PROTOCOL)));
    }

    /** Puts this service's certificate and the internal CA on an outbound client. */
    public HttpClient.Builder applyTo(HttpClient.Builder builder) {
        SSLParameters params = new SSLParameters();
        params.setProtocols(new String[] {PROTOCOL});
        return builder.sslContext(sslBundle().createSslContext()).sslParameters(params);
    }

    /** A RestClient builder for pe-sub-api: mTLS when enabled, the default transport otherwise. */
    public RestClient.Builder restClientBuilder(String property, String url) {
        requireHttps(property, url);
        RestClient.Builder builder = RestClient.builder();
        return enabled
                ? builder.requestFactory(new JdkClientHttpRequestFactory(applyTo(HttpClient.newBuilder()).build()))
                : builder;
    }

    /** Refuses a plain-HTTP peer URL once internal TLS is on. */
    public void requireHttps(String property, String url) {
        if (enabled && (url == null || !url.regionMatches(true, 0, "https://", 0, 8))) {
            throw new IllegalStateException(property + " must be an https:// URL when app.internal-tls.enabled=true");
        }
    }

    private static void require(String value, String name) {
        if (value == null || value.isBlank()) {
            throw new IllegalStateException("app.internal-tls." + name + " must be set when internal TLS is enabled");
        }
    }

    // Passwords stay out of logs and actuator output.
    @Override
    public String toString() {
        return "InternalTlsProperties[enabled=" + enabled + ", keyStore=" + keyStore
                + ", trustStore=" + trustStore + "]";
    }
}
