package com.ubs.pesubjobs.processor;

import com.ubs.pesubjobs.model.ClsConcLimitRow;
import com.ubs.pesubjobs.model.ProcessedClsConcLimit;
import org.springframework.batch.infrastructure.item.ItemProcessor;

public class ClsConcLimitRowProcessor implements ItemProcessor<ClsConcLimitRow, ProcessedClsConcLimit> {

    @Override
    public ProcessedClsConcLimit process(ClsConcLimitRow item) {
        // Invalid config fails the job rather than dropping out of it unnoticed.
        if (item.classification() == null || item.classification().isBlank()) {
            throw new IllegalArgumentException("Concentration-limit row names no classification.");
        }
        double pct = parsePct(item.limitPct());
        if (pct < 0 || pct > 100) {
            throw new IllegalArgumentException("Concentration limit for "
                    + item.classification().trim() + " is not a percent between 0 and 100: "
                    + item.limitPct());
        }
        return new ProcessedClsConcLimit(normalizeDashes(item.classification().trim()), pct);
    }

    private double parsePct(String s) {
        if (s == null || s.isBlank()) return -1;
        try {
            return Double.parseDouble(s.trim().replace("%", ""));
        } catch (NumberFormatException e) {
            return -1;
        }
    }

    private String normalizeDashes(String s) {
        return s.replace('\u2013', '-').replace('\u2014', '-').replace('\u2010', '-');
    }
}
