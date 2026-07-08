package com.ubs.pesubjobs.model;

/** Raw CSV row from the classification concentration-limit defaults feed. */
public record ClsConcLimitRow(
        String classification,
        String limitPct) {}
