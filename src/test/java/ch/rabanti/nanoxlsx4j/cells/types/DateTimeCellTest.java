package ch.rabanti.nanoxlsx4j.cells.types;

import ch.rabanti.nanoxlsx4j.Cell;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.time.Duration;
import java.util.Date;

import static ch.rabanti.nanoxlsx4j.TestUtils.buildDate;
import static ch.rabanti.nanoxlsx4j.TestUtils.buildTime;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class DateTimeCellTest {

    CellTypeUtils dateUtils;
    CellTypeUtils timeUtils;

    public DateTimeCellTest() {
        dateUtils = new CellTypeUtils(Date.class);
        timeUtils = new CellTypeUtils(Duration.class);
    }

    @DisplayName("Date value cell test: Test of the cell values, as well as proper modification")
    @Test()
    void dateCellTest() {
        // Date is hard to parametrize, therefore hardcoded
        Date defaultDateTime = buildDate(2020, 11, 1, 11, 22, 13, 99);
        dateUtils.assertCellCreation(defaultDateTime, buildDate(1900, 1, 1), Cell.CellType.DATE, Date::equals);
        dateUtils.assertCellCreation(defaultDateTime, buildDate(9999, 12, 31, 23, 59, 59), Cell.CellType.DATE, Date::equals);
    }

    @DisplayName("Duration value cell test: Test of the cell values, as well as proper modification")
    @Test()
    void DurationCellTest() {
        // LocalTime is hard to parametrize, therefore hardcoded
        Duration defaultTime = buildTime(22, 11, 7, 135);
        timeUtils.assertCellCreation(defaultTime, buildTime(0, 0, 0), Cell.CellType.TIME, Duration::equals);
        timeUtils.assertCellCreation(defaultTime, buildTime(23, 59, 59, 999), Cell.CellType.TIME, Duration::equals);
    }

    @DisplayName("Test of the Date comparison method on cells")
    @Test()
    void dateCellComparisonTest() {
        // Hard to parametrize, thus hardcoded
        Date baseDate = buildDate(2020, 11, 5, 12, 23, 7, 157);
        Date nearBelowBase = buildDate(2020, 11, 5, 12, 23, 7, 156);
        Date belowBase = buildDate(2020, 11, 5, 8, 23, 7, 157);
        Date nearAboveBase = buildDate(2020, 11, 5, 12, 23, 7, 158);
        Date aboveBase = buildDate(2020, 12, 5, 12, 23, 7, 156);

        Cell baseCell = dateUtils.createVariantCell(baseDate, dateUtils.getCellAddress());
        Cell equalCell = dateUtils.createVariantCell(baseDate, dateUtils.getCellAddress());
        Cell nearBelowCell = dateUtils.createVariantCell(nearBelowBase, dateUtils.getCellAddress());
        Cell nearAboveCell = dateUtils.createVariantCell(nearAboveBase, dateUtils.getCellAddress());
        Cell belowCell = dateUtils.createVariantCell(belowBase, dateUtils.getCellAddress());
        Cell aboveCell = dateUtils.createVariantCell(aboveBase, dateUtils.getCellAddress());

        assertEquals(0, ((Date) baseCell.getValue()).compareTo((Date) equalCell.getValue()));
        assertEquals(1, ((Date) baseCell.getValue()).compareTo((Date) nearBelowCell.getValue()));
        assertEquals(-1, ((Date) baseCell.getValue()).compareTo((Date) nearAboveCell.getValue()));
        assertEquals(1, ((Date) baseCell.getValue()).compareTo((Date) belowCell.getValue()));
        assertEquals(-1, ((Date) baseCell.getValue()).compareTo((Date) aboveCell.getValue()));
    }

    @DisplayName("Test of the LocalTime comparison method on cells")
    @Test()
    void durationCellComparisonTest() {
        // Hard to parametrize, thus hardcoded
        Duration baseTime = buildTime(5, 7, 22, 113);
        Duration nearBelowBase = buildTime(5, 7, 22, 112);
        Duration belowBase = buildTime(4, 7, 22, 113);
        Duration nearAboveBase = buildTime(5, 7, 22, 114);
        Duration aboveBase = buildTime(5, 17, 22, 113);

        Cell baseCell = timeUtils.createVariantCell(baseTime, timeUtils.getCellAddress());
        Cell equalCell = timeUtils.createVariantCell(baseTime, timeUtils.getCellAddress());
        Cell nearBelowCell = timeUtils.createVariantCell(nearBelowBase, timeUtils.getCellAddress());
        Cell nearAboveCell = timeUtils.createVariantCell(nearAboveBase, timeUtils.getCellAddress());
        Cell belowCell = timeUtils.createVariantCell(belowBase, timeUtils.getCellAddress());
        Cell aboveCell = timeUtils.createVariantCell(aboveBase, timeUtils.getCellAddress());

        assertCompareTo(0, (Duration) baseCell.getValue(), (Duration) equalCell.getValue());
        assertCompareTo(1, (Duration) baseCell.getValue(), (Duration) nearBelowCell.getValue());
        assertCompareTo(-1, (Duration) baseCell.getValue(), (Duration) nearAboveCell.getValue());
        assertCompareTo(1, (Duration) baseCell.getValue(), (Duration) belowCell.getValue());
        assertCompareTo(-1, (Duration) baseCell.getValue(), (Duration) aboveCell.getValue());
    }

    private static <T extends Comparable> void assertCompareTo(int threshold, T v1, T v2) {
        if (threshold == 0) {
            assertEquals(0, v1.compareTo(v2));
        } else if (threshold > 0) {
            int result = v1.compareTo(v2);
            assertTrue(result >= threshold);
        } else {
            int result = v1.compareTo(v2);
            assertTrue(result <= threshold);
        }
    }

}
