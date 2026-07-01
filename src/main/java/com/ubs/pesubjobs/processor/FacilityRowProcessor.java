package com.ubs.pesubjobs.processor;

import com.ubs.pesubjobs.model.FacilityRow;
import com.ubs.pesubjobs.model.ProcessedFacility;
import org.springframework.batch.item.ItemProcessor;
import org.springframework.lang.NonNull;
import org.springframework.lang.Nullable;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class FacilityRowProcessor implements ItemProcessor<FacilityRow, ProcessedFacility> {

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ISO_LOCAL_DATE;

    @Override
    @Nullable
    public ProcessedFacility process(@NonNull FacilityRow item) {
        if (item.name() == null || item.name().isBlank()) return null;
        if (item.agentBank() == null || item.agentBank().isBlank()) return null;
        return new ProcessedFacility(
                item.agentBank().trim(),
                item.name().trim(),
                blankToNull(item.accountNumber()),
                parseDecimal(item.loanAmount()),
                parseDate(item.maturityDate()),
                blankToNull(item.bankStatus()),
                parseDate(item.bankStatusDate())
        );
    }

    private String blankToNull(String s) {
        return (s == null || s.isBlank()) ? null : s.trim();
    }

    private BigDecimal parseDecimal(String s) {
        if (s == null || s.isBlank()) return null;
        return new BigDecimal(s.trim().replace("$", "").replace(",", ""));
    }

    private LocalDate parseDate(String s) {
        if (s == null || s.isBlank()) return null;
        return LocalDate.parse(s.trim(), DATE_FMT);
    }
}