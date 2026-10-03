package com.ubs.pesubjobs.config;

import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.boot.ssl.SslBundle;
import org.springframework.boot.ssl.SslBundleKey;
import org.springframework.boot.ssl.jks.JksSslStoreBundle;
import org.springframework.boot.ssl.jks.JksSslStoreDetails;
import org.springframework.boot.tomcat.TomcatWebServer;
import org.springframework.boot.tomcat.servlet.TomcatServletWebServerFactory;

import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/** Real handshakes against the server port, no Spring context. */
class InternalTlsServerConfigTest {

    private static final String PASSWORD = "test-only-changeit";
    private static final String CLIENT_CERT_ATTRIBUTE = "jakarta.servlet.request.X509Certificate";

    private final InternalTlsProperties tls = new InternalTlsProperties(true,
            "classpath:tls/service.p12", PASSWORD, "PKCS12",
            "classpath:tls/truststore.p12", PASSWORD, "PKCS12");

    private TomcatWebServer server;
    private int port;

    /** Answers "cert" or "none", so a test sees what the security filters would see. */
    private static final class ClientCertProbe extends HttpServlet {
        @Override
        protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws IOException {
            resp.getWriter().write(req.getAttribute(CLIENT_CERT_ATTRIBUTE) == null ? "none" : "cert");
        }
    }

    @BeforeEach
    void start() {
        TomcatServletWebServerFactory factory = new TomcatServletWebServerFactory(0);
        new InternalTlsServerConfig().internalTlsOnServerPort(tls).customize(factory);
        server = (TomcatWebServer) factory.getWebServer(
                ctx -> ctx.addServlet("probe", new ClientCertProbe()).addMapping("/"));
        server.start();
        port = server.getPort();
    }

    @AfterEach
    void stop() {
        server.stop();
    }

    private HttpResponse<String> get(HttpClient client, String scheme) throws Exception {
        return client.send(HttpRequest.newBuilder(URI.create(scheme + "://localhost:" + port + "/")).build(),
                HttpResponse.BodyHandlers.ofString());
    }

    @Test
    void onlyOneConnectorIsOpened() {
        assertThat(server.getTomcat().getService().findConnectors()).hasSize(1);
    }

    @Test
    void serviceCallerWithCaSignedCertificateIsIdentifiedByIt() throws Exception {
        HttpResponse<String> response = get(tls.applyTo(HttpClient.newBuilder()).build(), "https");

        assertThat(response.statusCode()).isEqualTo(200);
        assertThat(response.body()).isEqualTo("cert");
    }

    @Test
    void gatewayOrProbeWithoutClientCertificateIsServedWithoutOne() throws Exception {
        // Trusts the server, presents nothing.
        SslBundle trustOnly = SslBundle.of(new JksSslStoreBundle(null,
                new JksSslStoreDetails("PKCS12", null, "classpath:tls/truststore.p12", PASSWORD)), SslBundleKey.NONE);
        HttpResponse<String> response = get(HttpClient.newBuilder().sslContext(trustOnly.createSslContext()).build(),
                "https");

        assertThat(response.statusCode()).isEqualTo(200);
        assertThat(response.body()).isEqualTo("none");
    }

    @Test
    void plainHttpIsRefusedBeforeReachingTheApp() throws Exception {
        // Tomcat answers a plain request on a TLS port with 400 "requires TLS"; the servlet never runs.
        HttpResponse<String> response = get(HttpClient.newHttpClient(), "http");

        assertThat(response.statusCode()).isEqualTo(400);
        assertThat(response.body()).isNotIn("cert", "none");
    }

    @Test
    void plainHttpPeerUrlIsRejectedWhenEnabled() {
        assertThatThrownBy(() -> tls.requireHttps("ingest.api-base-url", "http://peer/x"))
                .isInstanceOf(IllegalStateException.class)
                .hasMessageContaining("https://");
        tls.requireHttps("ingest.api-base-url", "https://peer/x");
    }

    @Test
    void missingStoreFailsFast() {
        InternalTlsProperties incomplete = new InternalTlsProperties(true, null, null, "PKCS12", null, null, "PKCS12");

        assertThatThrownBy(incomplete::sslBundle).hasMessageContaining("app.internal-tls.key-store");
    }

    @Test
    void toStringNeverCarriesPasswords() {
        assertThat(tls.toString()).doesNotContain(PASSWORD);
    }
}
