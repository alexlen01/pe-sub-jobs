package com.ubs.pesubjobs.processor;

import com.ubs.pesubjobs.model.LpFacilitySeedRow;
import com.ubs.pesubjobs.model.ProcessedLpFacilitySeed;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.batch.infrastructure.item.ItemProcessor;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

@Component
public class LpFacilitySeedRowProcessor implements ItemProcessor<LpFacilitySeedRow, ProcessedLpFacilitySeed> {

    private static final Logger log = LoggerFactory.getLogger(LpFacilitySeedRowProcessor.class);

    private final NamedParameterJdbcTemplate jdbc;
    private final AtomicInteger seq = new AtomicInteger(0);

    public LpFacilitySeedRowProcessor(NamedParameterJdbcTemplate jdbc) {
        this.jdbc = jdbc;
    }

    @Override
    public ProcessedLpFacilitySeed process(LpFacilitySeedRow row) {
        if (row.facilityName() == null || row.facilityName().isBlank()
                || row.investorName() == null || row.investorName().isBlank()) {
            return null;
        }

        List<Integer> facilityIds = jdbc.queryForList(
                "SELECT id FROM facilities WHERE name = :name",
                Map.of("name", row.facilityName().trim()),
                Integer.class);
        if (facilityIds.isEmpty()) {
            log.warn("Facility not found — skipping row: {}", row.facilityName());
            return null;
        }
        int facilityId = facilityIds.getFirst();

        List<Map<String, Object>> lpRows = jdbc.queryForList("""
                SELECT id, parent, spv, high_qty, investor_type, inst_vs_hnw,
                       region_location, investment_grade, ubs_classification,
                       sp, mdy, fitch, aum, nav, pension, pension_funded
                FROM lp_master WHERE investor_name = :name
                """, Map.of("name", row.investorName().trim()));
        if (lpRows.isEmpty()) {
            log.warn("LP Master record not found — skipping row: {}", row.investorName());
            return null;
        }
        Map<String, Object> lp = lpRows.getFirst();

        String agentCls = normalizeAgentClassification(row.agentCls());
        String classification = normalizeUbsClassification(str(lp.get("ubs_classification")), agentCls);

        return new ProcessedLpFacilitySeed(
                facilityId,
                (Integer) lp.get("id"),
                seq.incrementAndGet(),
                row.investorName().trim(),
                str(lp.get("parent")),
                bool(lp.get("spv")),
                bool(lp.get("high_qty")),
                coalesce(str(lp.get("investor_type")), ""),
                coalesce(str(lp.get("inst_vs_hnw")), "Institutional"),
                coalesce(str(lp.get("region_location")), ""),
                bool(lp.get("investment_grade")),
                classification,
                agentCls,
                coalesce(str(lp.get("sp")),    ""),
                coalesce(str(lp.get("mdy")),   ""),
                coalesce(str(lp.get("fitch")), ""),
                str(lp.get("aum")),
                str(lp.get("nav")),
                str(lp.get("pension")),
                str(lp.get("pension_funded")),
                row.capCommit(),
                row.uncalled(),
                row.agentRate(),
                row.agentConc());
    }

    private static String str(Object o)                       { return o == null ? null : o.toString(); }
    private static boolean bool(Object o)                     { return o instanceof Boolean b && b; }
    private static String coalesce(String v, String fallback) { return (v != null && !v.isBlank()) ? v : fallback; }

    private static String normalizeAgentClassification(String raw) {
        if (raw == null || raw.isBlank()) return "Non-Rated Included";
        return switch (raw.trim()) {
            case "Rated", "Rated Included" -> "Rated Included";
            case "Non-Rated", "Non-Rated Included" -> "Non-Rated Included";
            case "Designated", "Designated Institutional" -> "Designated Institutional";
            case "Designated PWM" -> "Designated PWM";
            case "Ineligible", "Ineligible Investor", "Ineligible Investors" -> "Ineligible Investor";
            default -> raw.trim();
        };
    }

    private static String normalizeUbsClassification(String raw, String agentCls) {
        String value = raw == null ? "" : raw.trim();
        return switch (value) {
            case "Rated Investor" -> "Rated Investor";
            case "Unrated NAV > $1Bn" -> "Unrated NAV > $1Bn";
            case "FoF & Other > $10Bn AUM" -> "FoF & Other > $10Bn AUM";
            case "Corp Pension > $5Bn Assets" -> "Corp Pension > $5Bn Assets";
            case "Other Institutional" -> "Other Institutional";
            case "Excluded" -> "Excluded";
            case "Rated", "Rated Included" -> "Rated Investor";
            case "Unrated >2bn", "Unrated AUM >$2bn", "Unrated AUM $1-2bn", "Eligible <$1bn",
                 "Non-Rated", "Non-Rated Included" -> "Unrated NAV > $1Bn";
            case "Designated Institutional" -> "Corp Pension > $5Bn Assets";
            case "Eligible", "Designated PWM", "Included (PWM)" -> "Other Institutional";
            case "Ineligible Investor", "Ineligible Investors" -> "Excluded";
            default -> switch (agentCls) {
                case "Rated Included" -> "Rated Investor";
                case "Designated Institutional" -> "Corp Pension > $5Bn Assets";
                case "Designated PWM" -> "Other Institutional";
                case "Ineligible Investor" -> "Excluded";
                default -> "Unrated NAV > $1Bn";
            };
        };
    }
}
