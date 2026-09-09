package com.ubs.pesubjobs.security;

import jakarta.servlet.FilterChain;
import jakarta.servlet.ServletException;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
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
 */
public class ServiceIdentityFilter extends OncePerRequestFilter {

    private final JobsSecurityProperties props;

    public ServiceIdentityFilter(JobsSecurityProperties props) {
        this.props = props;
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

        return UsernamePasswordAuthenticationToken.authenticated(user.trim(), null, authorities);
    }
}
