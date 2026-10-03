package com.ubs.pesubjobs.security;

import jakarta.servlet.FilterChain;
import jakarta.servlet.ServletException;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.security.authentication.UsernamePasswordAuthenticationToken;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.authority.SimpleGrantedAuthority;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.web.filter.OncePerRequestFilter;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;

/**
 * Establishes the request's {@link Authentication} from the gateway-asserted identity headers.
 *
 * <p>Mirrors pe-sub-api's header contract so one proxy configuration serves both services. There
 * is deliberately no fallback identity: a request without a user header leaves the security
 * context empty and the authorization layer answers 401.
 *
 * <p>If the context is already populated — e.g. by {@code @WithMockUser} in a test — this filter
 * defers to it, so authorization tests can assert arbitrary roles.
 *
 * <p>With internal mTLS on, the required role counts only on a request that carried a verified
 * client certificate; anywhere else it is dropped, so {@code /jobs/**} answers 403.
 */
public class ServiceIdentityFilter extends OncePerRequestFilter {

    private static final Logger log = LoggerFactory.getLogger(ServiceIdentityFilter.class);

    /** Set by the container only when the caller presented a certificate the trust store accepts. */
    static final String CLIENT_CERT_ATTRIBUTE = "jakarta.servlet.request.X509Certificate";

    private final JobsSecurityProperties props;
    private final boolean requiredRoleNeedsClientCert;

    public ServiceIdentityFilter(JobsSecurityProperties props, boolean requiredRoleNeedsClientCert) {
        this.props = props;
        this.requiredRoleNeedsClientCert = requiredRoleNeedsClientCert;
    }

    @Override
    protected void doFilterInternal(HttpServletRequest request,
                                    HttpServletResponse response,
                                    FilterChain filterChain) throws ServletException, IOException {
        Authentication existing = SecurityContextHolder.getContext().getAuthentication();
        if (existing != null && existing.isAuthenticated()) {
            filterChain.doFilter(request, response);
            return;
        }

        Authentication auth = fromHeaders(request);
        if (auth != null) {
            SecurityContextHolder.getContext().setAuthentication(auth);
        }
        filterChain.doFilter(request, response);
    }

    private Authentication fromHeaders(HttpServletRequest request) {
        String user = request.getHeader(props.getUserHeader());
        if (user == null || user.isBlank()) return null;

        String rolesRaw = request.getHeader(props.getRolesHeader());
        List<SimpleGrantedAuthority> authorities = (rolesRaw == null || rolesRaw.isBlank())
                ? List.of()
                : Arrays.stream(rolesRaw.split(","))
                        .map(role -> role.trim())
                        .filter(role -> !role.isEmpty())
                        .map(role -> new SimpleGrantedAuthority("ROLE_" + role.toUpperCase(Locale.ROOT)))
                        .toList();

        if (requiredRoleNeedsClientCert && request.getAttribute(CLIENT_CERT_ATTRIBUTE) == null) {
            String required = "ROLE_" + props.getRequiredRole().toUpperCase(Locale.ROOT);
            List<SimpleGrantedAuthority> kept = authorities.stream()
                    .filter(a -> !a.getAuthority().equals(required)).toList();
            if (kept.size() != authorities.size()) {
                log.warn("{} role dropped for user={}: asserted without a client certificate on port {} {} {}",
                        props.getRequiredRole(), user.trim(), request.getLocalPort(), request.getMethod(),
                        request.getRequestURI());
            }
            authorities = kept;
        }

        return UsernamePasswordAuthenticationToken.authenticated(user.trim(), null, authorities);
    }
}
