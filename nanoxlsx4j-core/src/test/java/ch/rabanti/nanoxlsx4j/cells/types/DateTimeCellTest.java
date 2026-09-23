package ch.rabanti.nanoxlsx4j.cells.types;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.time.Duration;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.util.Date;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import ch.rabanti.nanoxlsx4j.Cell;

class DateTimeCellTest {
    private final CellTypeUtils utils = new CellTypeUtils();

    @Test
    @DisplayName("Date value cell test: Test of the cell values, as well as proper modification")
    void dateCellTest() {
        // Date is complicated to parametrize, therefore hardcoded
        Date defaultDateTime = date(2020, 11, 1, 11, 22, 13, 99);
        utils.assertCellCreation(
                defaultDateTime, date(1900, 1, 1, 0, 0, 0, 0),
                Cell.CellType.DATE, Date::equals
        );
        utils.assertCellCreation(
                defaultDateTime, date(9999, 12, 31, 23, 59, 59, 0),
                Cell.CellType.DATE, Date::equals
        );
    }

    @Test
    @DisplayName("Duration value cell test: Test of the cell values, as well as proper modification")
    void durationCellTest() {
        // Date is complicated to parametrize, therefore hardcoded
        Duration defaultTime = duration(0, 22, 11, 7, 135);
        utils.assertCellCreation(
                defaultTime, duration(0, 0, 0, 0, 0),
                Cell.CellType.TIME, Duration::equals
        );
        utils.assertCellCreation(
                defaultTime, duration(2958465, 23, 59, 59, 999),
                Cell.CellType.TIME, Duration::equals
        );
    }

    @Test
    @DisplayName("Test of the Date comparison method on cells")
    void dateCellComparisonTest() {
        // Complicated to parametrize, thus hardcoded
        Date baseDate = date(2020, 11, 5, 12, 23, 7, 157);
        Date nearBelowBase = date(2020, 11, 5, 12, 23, 7, 156);
        Date belowBase = date(2020, 11, 5, 8, 23, 7, 157);
        Date nearAboveBase = date(2020, 11, 5, 12, 23, 7, 158);
        Date aboveBase = date(2020, 12, 5, 12, 23, 7, 156);

        Cell baseCell = utils.createVariantCell(baseDate, utils.getCellAddress());
        Cell equalCell = utils.createVariantCell(baseDate, utils.getCellAddress());
        Cell nearBelowCell = utils.createVariantCell(nearBelowBase, utils.getCellAddress());
        Cell nearAboveCell = utils.createVariantCell(nearAboveBase, utils.getCellAddress());
        Cell belowCell = utils.createVariantCell(belowBase, utils.getCellAddress());
        Cell aboveCell = utils.createVariantCell(aboveBase, utils.getCellAddress());

        assertEquals(0, ((Date) baseCell.getValue()).compareTo((Date) equalCell.getValue()));
        assertEquals(1, ((Date) baseCell.getValue()).compareTo((Date) nearBelowCell.getValue()));
        assertEquals(-1, ((Date) baseCell.getValue()).compareTo((Date) nearAboveCell.getValue()));
        assertEquals(1, Integer.signum(((Date) baseCell.getValue()).compareTo((Date) belowCell.getValue())));
        assertEquals(-1, Integer.signum(((Date) baseCell.getValue()).compareTo((Date) aboveCell.getValue())));
    }

    @Test
    @DisplayName("Test of the Duration comparison method on cells")
    void durationCellComparisonTest() {
        Duration baseTime = duration(1, 5, 7, 22, 113);
        Duration nearBelowBase = duration(1, 5, 7, 22, 112);
        Duration belowBase = duration(0, 5, 7, 22, 113);
        Duration nearAboveBase = duration(1, 5, 7, 22, 114);
        Duration aboveBase = duration(1, 5, 17, 22, 113);

        Cell baseCell = utils.createVariantCell(baseTime, utils.getCellAddress());
        Cell equalCell = utils.createVariantCell(baseTime, utils.getCellAddress());
        Cell nearBelowCell = utils.createVariantCell(nearBelowBase, utils.getCellAddress());
        Cell nearAboveCell = utils.createVariantCell(nearAboveBase, utils.getCellAddress());
        Cell belowCell = utils.createVariantCell(belowBase, utils.getCellAddress());
        Cell aboveCell = utils.createVariantCell(aboveBase, utils.getCellAddress());

        assertEquals(0, ((Duration) baseCell.getValue()).compareTo((Duration) equalCell.getValue()));
        assertEquals(
                1, Integer.signum(((Duration) baseCell.getValue()).compareTo((Duration) nearBelowCell.getValue())));
        assertEquals(
                -1, Integer.signum(((Duration) baseCell.getValue()).compareTo((Duration) nearAboveCell.getValue())));
        assertEquals(1, Integer.signum(((Duration) baseCell.getValue()).compareTo((Duration) belowCell.getValue())));
        assertEquals(-1, Integer.signum(((Duration) baseCell.getValue()).compareTo((Duration) aboveCell.getValue())));
    }

    private static Date date(int year, int month, int day, int hour, int minute, int second, int millis) {
        return Date.from(LocalDateTime.of(year, month, day, hour, minute, second, millis * 1_000_000)
                .toInstant(ZoneOffset.UTC));
    }

    private static Duration duration(long days, long hours, long minutes, long seconds, long millis) {
        return Duration.ofDays(days).plusHours(hours).plusMinutes(minutes).plusSeconds(seconds).plusMillis(millis);
    }
}
