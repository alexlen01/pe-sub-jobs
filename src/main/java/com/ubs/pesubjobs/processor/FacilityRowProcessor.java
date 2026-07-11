package com.ubs.pesubjobs.processor;

import com.ubs.pesubjobs.model.FacilityRow;
import com.ubs.pesubjobs.model.ProcessedFacility;
import org.springframework.batch.infrastructure.item.ItemProcessor;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoField;
import java.util.List;

public class FacilityRowProcessor implements ItemProcessor<FacilityRow, ProcessedFacility> {

    private static final List<DateTimeFormatter> DATE_FORMATS = List.of(
            DateTimeFormatter.ISO_LOCAL_DATE,
            new DateTimeFormatterBuilder()
                    .appendPattern("M/d/")
                    .appendValue(ChronoField.YEAR, 4)
                    .toFormatter()
    );

    @Override
    public ProcessedFacility process(FacilityRow item) {
        if (item.name() == null || item.name().isBlank()) return null;
        if (item.agentBank() == null || item.agentBank().isBlank()) return null;
        return new ProcessedFacility(
                item.agentBank().trim(),
                item.name().trim(),
                blankToNull(item.accountNumber()),
                parseDecimal(item.loanAmount()),
                parseDate(item.maturityDate()),
                blankToNull(item.bankStatus()),
                parseDate(item.bankStatusDate()),
                parseDecimal(item.ubsParticipation()),
                parseDate(item.collateralDate())
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
        String value = s.trim();
        for (DateTimeFormatter formatter : DATE_FORMATS) {
            try {
                return LocalDate.parse(value, formatter);
            } catch (DateTimeParseException ignored) {
                // Try the next feed-supported date shape.
            }
        }
        throw new DateTimeParseException("Unsupported date format", value, 0);
    }
}
