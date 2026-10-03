package com.ubs.pesubjobs.security;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockFilterChain;
import org.springframework.mock.web.MockHttpServletRequest;
import org.springframework.mock.web.MockHttpServletResponse;
import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.context.SecurityContextHolder;

import java.security.cert.X509Certificate;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.mock;

/** The required role asserted by header counts only over the mutual-TLS connector once internal TLS is on. */
class ServiceIdentityFilterClientCertTest {

    @AfterEach
    void clearContext() {
        SecurityContextHolder.clearContext();
    }

    private static List<String> authoritiesFor(boolean tlsEnabled, boolean clientCert) throws Exception {
        ServiceIdentityFilter filter = new ServiceIdentityFilter(new JobsSecurityProperties(), tlsEnabled);
        MockHttpServletRequest request = new MockHttpServletRequest("POST", "/jobs/lpMasterIngest");
        request.addHeader("X-Auth-User", "ops-scheduler");
        request.addHeader("X-Auth-Roles", "service,AUDITOR");
        if (clientCert) {
            request.setAttribute(ServiceIdentityFilter.CLIENT_CERT_ATTRIBUTE,
                    new X509Certificate[] {mock(X509Certificate.class)});
        }
        filter.doFilter(request, new MockHttpServletResponse(), new MockFilterChain());
        return SecurityContextHolder.getContext().getAuthentication().getAuthorities().stream()
                .map(GrantedAuthority::getAuthority).toList();
    }

    @Test
    void requiredRoleKeptOverMutualTls() throws Exception {
        assertThat(authoritiesFor(true, true)).containsExactly("ROLE_SERVICE", "ROLE_AUDITOR");
    }

    @Test
    void requiredRoleDroppedWithoutClientCertButOtherRolesKept() throws Exception {
        assertThat(authoritiesFor(true, false)).containsExactly("ROLE_AUDITOR");
    }

    @Test
    void requiredRoleKeptWhenInternalTlsIsOff() throws Exception {
        assertThat(authoritiesFor(false, false)).containsExactly("ROLE_SERVICE", "ROLE_AUDITOR");
    }
}
