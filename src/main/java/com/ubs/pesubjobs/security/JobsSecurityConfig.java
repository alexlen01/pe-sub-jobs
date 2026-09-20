package com.ubs.pesubjobs.security;

import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.http.HttpStatus;
import org.springframework.http.HttpMethod;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.web.authentication.HttpStatusEntryPoint;
import org.springframework.security.config.http.SessionCreationPolicy;
import org.springframework.security.web.SecurityFilterChain;
import org.springframework.security.web.authentication.UsernamePasswordAuthenticationFilter;

/**
 * Stateless authorization for the job trigger surface.
 *
 * <p>The service exposes exactly one capability — running a feed — and every feed rewrites
 * reference data owned by pe-sub-api. There is therefore no permitted anonymous route: the
 * default is {@code denyAll} and {@code /jobs/**} requires the configured service role.
 *
 * <p>CSRF is disabled because identity travels in a proxy-injected header rather than a cookie,
 * so there is no ambient credential a foreign origin could ride on.
 */
@Configuration
@EnableConfigurationProperties(JobsSecurityProperties.class)
public class JobsSecurityConfig {

    @Bean
    public SecurityFilterChain filterChain(HttpSecurity http, JobsSecurityProperties props) throws Exception {
        return http
                // Safe because this stateless API authenticates only through a proxy-injected
                // header; it neither creates sessions nor accepts cookie-based credentials.
                .csrf(csrf -> csrf.disable())
                .cors(cors -> cors.disable())
                .sessionManagement(session -> session.sessionCreationPolicy(SessionCreationPolicy.STATELESS))
                .addFilterBefore(new ServiceIdentityFilter(props), UsernamePasswordAuthenticationFilter.class)
                // Without an entry point an anonymous denial answers 403, which tells a caller its
                // credentials were rejected rather than that it presented none. 401 is the honest
                // answer for a missing identity header; an authenticated caller still gets 403.
                .exceptionHandling(ex -> ex.authenticationEntryPoint(
                        new HttpStatusEntryPoint(HttpStatus.UNAUTHORIZED)))
                .authorizeHttpRequests(auth -> auth
                        // Container error dispatch must render the ProblemDetail body rather than
                        // being re-evaluated as an unauthenticated request.
                        .requestMatchers("/error").permitAll()
                        .requestMatchers(HttpMethod.GET, "/manage/health").permitAll()
                        .requestMatchers(HttpMethod.GET, "/manage/health/liveness").permitAll()
                        .requestMatchers(HttpMethod.GET, "/manage/health/readiness").permitAll()
                        .requestMatchers(HttpMethod.GET, "/manage/info").permitAll()
                        .requestMatchers("/jobs/**").hasRole(props.getRequiredRole())
                        .anyRequest().denyAll())
                .build();
    }
}
