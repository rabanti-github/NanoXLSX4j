package ch.rabanti.nanoxlsx4j.cells.types;

import ch.rabanti.nanoxlsx4j.Cell;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.time.LocalTime;
import java.util.Calendar;
import java.util.Date;

import static org.junit.jupiter.api.Assertions.assertEquals;

class DateTimeCellTest {

    CellTypeUtils dateUtils;
    CellTypeUtils timeUtils;
    Calendar calendarInstance;

    public DateTimeCellTest() {
        dateUtils = new CellTypeUtils(Date.class);
        timeUtils = new CellTypeUtils(LocalTime.class);
        calendarInstance = Calendar.getInstance();
    }

    @DisplayName("Date value cell test: Test of the cell values, as well as proper modification")
    @Test()
    void dateCellTest() {
        // Date is hard to parametrize, therefore hardcoded
        Date defaultDateTime = buildDate(2020, 11, 1, 11, 22, 13, 99);
        dateUtils.assertCellCreation(defaultDateTime, buildDate(1900, 1, 1), Cell.CellType.DATE, Date::equals);
        dateUtils.assertCellCreation(defaultDateTime, buildDate(9999, 12, 31, 23, 59, 59), Cell.CellType.DATE, Date::equals);
    }

    @DisplayName("LocalTime value cell test: Test of the cell values, as well as proper modification")
    @Test()
    void LocalTimeCellTest() {
        // LocalTime is hard to parametrize, therefore hardcoded
        LocalTime defaultTime = buildTime(22, 11, 7, 135);
        timeUtils.assertCellCreation(defaultTime, buildTime(0, 0, 0), Cell.CellType.TIME, LocalTime::equals);
        timeUtils.assertCellCreation(defaultTime, buildTime(23, 59, 59, 999), Cell.CellType.TIME, LocalTime::equals);
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
    void LocalTimeCellComparisonTest() {
        // Hard to parametrize, thus hardcoded
        LocalTime baseTime = buildTime(5, 7, 22, 113);
        LocalTime nearBelowBase = buildTime(5, 7, 22, 112);
        LocalTime belowBase = buildTime(4, 7, 22, 113);
        LocalTime nearAboveBase = buildTime(5, 7, 22, 114);
        LocalTime aboveBase = buildTime(5, 17, 22, 113);

        Cell baseCell = timeUtils.createVariantCell(baseTime, timeUtils.getCellAddress());
        Cell equalCell = timeUtils.createVariantCell(baseTime, timeUtils.getCellAddress());
        Cell nearBelowCell = timeUtils.createVariantCell(nearBelowBase, timeUtils.getCellAddress());
        Cell nearAboveCell = timeUtils.createVariantCell(nearAboveBase, timeUtils.getCellAddress());
        Cell belowCell = timeUtils.createVariantCell(belowBase, timeUtils.getCellAddress());
        Cell aboveCell = timeUtils.createVariantCell(aboveBase, timeUtils.getCellAddress());

        assertEquals(0, ((LocalTime) baseCell.getValue()).compareTo((LocalTime) equalCell.getValue()));
        assertEquals(1, ((LocalTime) baseCell.getValue()).compareTo((LocalTime) nearBelowCell.getValue()));
        assertEquals(-1, ((LocalTime) baseCell.getValue()).compareTo((LocalTime) nearAboveCell.getValue()));
        assertEquals(1, ((LocalTime) baseCell.getValue()).compareTo((LocalTime) belowCell.getValue()));
        assertEquals(-1, ((LocalTime) baseCell.getValue()).compareTo((LocalTime) aboveCell.getValue()));
    }


    private Date buildDate(int year, int month, int day) {
        return buildDate(year, month, day, 0, 0, 0, 0);
    }

    private Date buildDate(int year, int month, int day, int hour, int minute, int second) {
        return buildDate(year, month, day, hour, minute, second, 0);
    }

    private Date buildDate(int year, int month, int day, int hour, int minute, int second, int millisSecond) {
        calendarInstance.set(year, month, day, hour, minute, second);
        calendarInstance.set(Calendar.MILLISECOND, millisSecond);
        return calendarInstance.getTime();
    }

    private LocalTime buildTime(int hour, int minute, int second) {
        return buildTime(hour, minute, second, 0);
    }

    private LocalTime buildTime(int hour, int minute, int second, int milliSecond) {
        return LocalTime.of(hour, minute, second, milliSecond * 1000000);
    }

}
