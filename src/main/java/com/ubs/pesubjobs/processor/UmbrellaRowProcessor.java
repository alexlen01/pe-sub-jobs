package com.ubs.pesubjobs.processor;

import com.ubs.pesubjobs.model.ProcessedUmbrella;
import com.ubs.pesubjobs.model.UmbrellaRow;
import org.springframework.batch.infrastructure.item.ItemProcessor;

import java.math.BigDecimal;

public class UmbrellaRowProcessor implements ItemProcessor<UmbrellaRow, ProcessedUmbrella> {

    @Override
    public ProcessedUmbrella process(UmbrellaRow item) {
        // Nothing is dropped or defaulted here. A group with no obligor, no account or no line is a
        // group the report stated nothing about, and the API distinguishes that silence from a
        // correction — it leaves what an analyst recorded alone rather than blanking it.
        return new ProcessedUmbrella(
                blankToNull(item.key()),
                blankToNull(item.name()),
                blankToNull(item.obligorName()),
                blankToNull(item.agentBank()),
                blankToNull(item.accountNumber()),
                parseDecimal(item.loanAmount()),
                parseBoolean(item.crossCollateralized()),
                blankToNull(item.agreementRef()),
                blankToNull(item.borrowerEntity()),
                parseDecimal(item.subLimit()),
                // Passed through as written, upper-cased only. The API refuses a reading it does not
                // hold and names the row; folding an unfamiliar spelling onto one it does hold would
                // record a legal position nobody stated.
                item.liabilityType() == null || item.liabilityType().isBlank()
                        ? null : item.liabilityType().trim().toUpperCase(java.util.Locale.ROOT)
        );
    }

    /** Null, not false, for a blank cell: cross-collateralization is a term of the credit agreement,
     *  and silence is not a statement that the members hold separate collateral. */
    private Boolean parseBoolean(String s) {
        String v = blankToNull(s);
        return v == null ? null : Boolean.valueOf("true".equalsIgnoreCase(v) || "1".equals(v));
    }

    private String blankToNull(String s) {
        return (s == null || s.isBlank()) ? null : s.trim();
    }

    private BigDecimal parseDecimal(String s) {
        if (s == null || s.isBlank()) return null;
        return new BigDecimal(s.trim().replace("$", "").replace(",", ""));
    }
}
