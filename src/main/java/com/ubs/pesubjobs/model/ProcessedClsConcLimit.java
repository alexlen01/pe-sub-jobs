package com.ubs.pesubjobs.model;

/** Validated feed row: classification label + default per-LP concentration limit
 *  as a percent of total uncalled capital (e.g. 7.5). */
public record ProcessedClsConcLimit(
        String classification,
        double limitPct) {}
