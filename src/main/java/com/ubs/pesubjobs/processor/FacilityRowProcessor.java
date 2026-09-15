package com.ubs.pesubjobs.processor;

import com.ubs.pesubjobs.model.FacilityRow;
import com.ubs.pesubjobs.model.ProcessedFacility;
import org.springframework.batch.infrastructure.item.ItemProcessor;

import java.math.BigDecimal;
import java.time.LocalDate;

public class FacilityRowProcessor implements ItemProcessor<FacilityRow, ProcessedFacility> {

    @Override
    public ProcessedFacility process(FacilityRow item) {
        // A blank name or agent bank is not a reason to drop a facility; both are filled in
        // where the constraints live, on the API side.
        return new ProcessedFacility(
                blankToNull(item.agentBank()),
                blankToNull(item.name()),
                blankToNull(item.accountNumber()),
                parseDecimal(item.facilitySize()),
                parseDate(item.maturityDate()),
                blankToNull(item.status()),
                parseDecimal(item.ubsParticipation()),
                parseDate(item.collateralDate()),
                blankToNull(item.umbrellaName()),
                blankToNull(item.umbrellaKey()),
                parseBoolean(item.umbrellaCrossCollateralized()),
                blankToNull(item.creditAgreementRef())
        );
    }

    /** Null, not false, for a blank cell: the feed states a shared borrowing base only where it can
     *  know of one, and silence must not read as a statement that there is none. */
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

    private LocalDate parseDate(String s) {
        return FeedDates.parse(s);
    }
}
