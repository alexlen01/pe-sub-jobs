package com.ubs.pesubjobs.security;

import org.springframework.boot.context.properties.ConfigurationProperties;

/**
 * Identity configuration for the {@code /jobs} trigger surface.
 *
 * <p>Unlike pe-sub-api this service has no developer mode: there is no browser flow to keep
 * working and no anonymous surface worth preserving, so the gateway-asserted headers are the only
 * accepted identity in every environment. A caller without them is rejected 401, and one without
 * {@link #requiredRole} is rejected 403.
 */
@ConfigurationProperties(prefix = "app.security")
public class JobsSecurityProperties {

    /** Header carrying the authenticated caller id, injected by the trusted reverse proxy. */
    private String userHeader = "X-Auth-User";

    /** Header carrying the caller's comma-separated role list, e.g. {@code "SERVICE"}. */
    private String rolesHeader = "X-Auth-Roles";

    /**
     * Role a caller must hold to trigger a job. Feeds replace reference data wholesale, so this
     * is a machine-to-machine capability rather than an end-user one.
     */
    private String requiredRole = "SERVICE";

    /**
     * Identity this service asserts on its own calls to pe-sub-api. Those routes are SERVICE-gated
     * there, so once pe-sub-api runs in gateway mode a header-less feed would be rejected 401.
     * Carrying a named identity also gives pe-sub-api's audit trail a caller to attribute to.
     */
    private String serviceUser = "pe-sub-jobs";

    public String getUserHeader() { return userHeader; }
    public void setUserHeader(String userHeader) { this.userHeader = userHeader; }

    public String getRolesHeader() { return rolesHeader; }
    public void setRolesHeader(String rolesHeader) { this.rolesHeader = rolesHeader; }

    public String getRequiredRole() { return requiredRole; }
    public void setRequiredRole(String requiredRole) { this.requiredRole = requiredRole; }

    public String getServiceUser() { return serviceUser; }
    public void setServiceUser(String serviceUser) { this.serviceUser = serviceUser; }
}
