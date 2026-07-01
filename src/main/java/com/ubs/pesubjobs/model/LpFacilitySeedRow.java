package com.ubs.pesubjobs.model;

public record LpFacilitySeedRow(
        String facilityName,
        String investorName,
        String capCommit,
        String uncalled,
        String agentCls,
        String agentRate,
        String agentConc) {}