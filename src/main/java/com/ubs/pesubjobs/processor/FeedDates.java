package com.ubs.pesubjobs.processor;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoField;
import java.util.List;

/**
 * The date shapes every feed file is read in.
 *
 * <p>Shared rather than repeated per processor: the facility feed and the group feed state the same
 * dates about the same credit agreement — a group hands its maturity and collateral date down to its
 * members while it is Active — and two lists that drifted apart would parse one agreement's terms on
 * one file and reject them on the other.
 */
final class FeedDates {

    private static final List<DateTimeFormatter> FORMATS = List.of(
            DateTimeFormatter.ISO_LOCAL_DATE,
            new DateTimeFormatterBuilder()
                    .appendPattern("M/d/")
                    .appendValue(ChronoField.YEAR, 4)
                    .toFormatter()
    );

    private FeedDates() {}

    /**
     * The date a feed cell states, or null where it states none.
     *
     * @throws DateTimeParseException on a cell that is not blank and is not a date — a run stops
     *                                rather than silently recording an agreement with no maturity.
     */
    static LocalDate parse(String s) {
        if (s == null || s.isBlank()) return null;
        String value = s.trim();
        for (DateTimeFormatter formatter : FORMATS) {
            try {
                return LocalDate.parse(value, formatter);
            } catch (DateTimeParseException ignored) {
                // Try the next feed-supported date shape.
            }
        }
        throw new DateTimeParseException("Unsupported date format", value, 0);
    }
}
