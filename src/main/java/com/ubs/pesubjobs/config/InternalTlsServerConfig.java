package com.ubs.pesubjobs.config;

import org.apache.commons.logging.LogFactory;
import org.springframework.boot.autoconfigure.condition.ConditionalOnBooleanProperty;
import org.springframework.boot.ssl.SslBundle;
import org.springframework.boot.tomcat.SslConnectorCustomizer;
import org.springframework.boot.tomcat.servlet.TomcatServletWebServerFactory;
import org.springframework.boot.web.server.Ssl;
import org.springframework.boot.web.server.WebServerFactoryCustomizer;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import java.util.Map;

/**
 * Serves this service's own port (server.port) over TLS 1.3; no extra port is opened. A client
 * certificate is requested but optional: the gateway ingress and kubelet probes connect without
 * one, service callers present theirs, and only those requests may carry the SERVICE role.
 */
@Configuration(proxyBeanMethods = false)
@ConditionalOnBooleanProperty("app.internal-tls.enabled")
public class InternalTlsServerConfig {

    @Bean
    WebServerFactoryCustomizer<TomcatServletWebServerFactory> internalTlsOnServerPort(InternalTlsProperties tls) {
        return factory -> {
            SslBundle bundle = tls.sslBundle();
            factory.addConnectorCustomizers(connector -> new SslConnectorCustomizer(
                    LogFactory.getLog(InternalTlsServerConfig.class), connector, Ssl.ClientAuth.WANT)
                    .customize(bundle, Map.of()));
        };
    }
}
