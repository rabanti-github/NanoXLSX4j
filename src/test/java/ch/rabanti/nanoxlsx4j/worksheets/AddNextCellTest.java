package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.time.Duration;
import java.time.LocalTime;
import java.util.Calendar;
import java.util.Date;

import static ch.rabanti.nanoxlsx4j.TestUtils.buildDate;
import static ch.rabanti.nanoxlsx4j.TestUtils.buildTime;
import static org.junit.jupiter.api.Assertions.assertEquals;

public class AddNextCellTest {
    @DisplayName("Test of the addNextCell function with only the value")
    @ParameterizedTest(name = "Given value {1} ({0}) should lead to cells of the type {2}")
    @CsvSource({
            "NULL,'', EMPTY",
            "STRING,'', STRING",
            "STRING, test, STRING",
            "LONG, '17', NUMBER",
            "DOUBLE, '1.02', NUMBER",
            "FLOAT, '-22.3', NUMBER",
            "INTEGER, '0', NUMBER",
            "BYTE, '127', NUMBER",
            "BOOLEAN, 'true', BOOL",
            "BOOLEAN, 'false', BOOL",
    })
    void addNextCellTest(String sourceType, String sourceValue, Cell.CellType expectedType) {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentCellAddress("D2");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        assertEquals(0, worksheet.getCells().size());
        Object value = TestUtils.createInstance(sourceType, sourceValue);
        worksheet.addNextCell(value);
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", expectedType, null, value, 3, 2);
        worksheet.setCurrentCellAddress("E3");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.ColumnToColumn);
        worksheet.addNextCell(value);
        WorksheetTest.assertAddedCell(worksheet, 2, "E3", expectedType, null, value, 5, 2);
    }

    @DisplayName("Test of the addNextCell function with value and style")
    @ParameterizedTest(name = "Given value {1} ({0}) should lead to cells of the type {2}")
    @CsvSource({
            "NULL,'', EMPTY",
            "STRING,'', STRING",
            "STRING, test, STRING",
            "LONG, '17', NUMBER",
            "DOUBLE, '1.02', NUMBER",
            "FLOAT, '-22.3', NUMBER",
            "INTEGER, '0', NUMBER",
            "BYTE, '127', NUMBER",
            "BOOLEAN, 'true', BOOL",
            "BOOLEAN, 'false', BOOL",
    })
    void addNextCellTest2(String sourceType, String sourceValue, Cell.CellType expectedType) {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentCellAddress("D2");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        assertEquals(0, worksheet.getCells().size());
        Object value = TestUtils.createInstance(sourceType, sourceValue);
        worksheet.addNextCell(value, BasicStyles.BoldItalic());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", expectedType, BasicStyles.BoldItalic(), value, 3, 2);
        worksheet.setCurrentCellAddress("E3");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.ColumnToColumn);
        worksheet.addNextCell(value, BasicStyles.Bold());
        WorksheetTest.assertAddedCell(worksheet, 2, "E3", expectedType, BasicStyles.Bold(), value, 5, 2);
    }

    @DisplayName("Test of the addNextCell function for Date and Duration")
    @Test()
    void addNextCellTest3() {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentCellAddress("D2");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        assertEquals(0, worksheet.getCells().size());
        Date date = buildDate(2020, 6, 10, 11, 12, 22);
        worksheet.addNextCell(date);
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.DATE, BasicStyles.DateFormat(), date, 3, 2);
        worksheet.setCurrentCellAddress("E3");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.ColumnToColumn);
        Duration time = buildTime(6, 22, 13);
        worksheet.addNextCell(time);
        WorksheetTest.assertAddedCell(worksheet, 2, "E3", Cell.CellType.TIME, BasicStyles.TimeFormat(), time, 5, 2);
    }

    @DisplayName("Test of the addNextCell function for Date and Duration with styles")
    @Test()
    void addNextCellTest4() {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentCellAddress("D2");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        assertEquals(0, worksheet.getCells().size());
        Date date = buildDate(2020, 6, 10, 11, 12, 22);
        worksheet.addNextCell(date, BasicStyles.Bold());
        Style mixedStyle = BasicStyles.DateFormat();
        mixedStyle.append(BasicStyles.Bold());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.DATE, mixedStyle, date, 3, 2);
        worksheet.setCurrentCellAddress("E3");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.ColumnToColumn);
        Duration time = buildTime(6, 22, 13);
        worksheet.addNextCell(time, BasicStyles.Underline());
        mixedStyle = BasicStyles.TimeFormat();
        mixedStyle.append(BasicStyles.Underline());
        WorksheetTest.assertAddedCell(worksheet, 2, "E3", Cell.CellType.TIME, mixedStyle, time, 5, 2);
    }

    @DisplayName("Test of the addNextCell function with value and active worksheet style")
    @ParameterizedTest(name = "Given value {1} ({0}) should lead to cells of the type {2}")
    @CsvSource({
            "NULL,'', EMPTY",
            "STRING,'', STRING",
            "STRING, test, STRING",
            "LONG, '17', NUMBER",
            "DOUBLE, '1.02', NUMBER",
            "FLOAT, '-22.3', NUMBER",
            "INTEGER, '0', NUMBER",
            "BYTE, '127', NUMBER",
            "BOOLEAN, 'true', BOOL",
            "BOOLEAN, 'false', BOOL",
    })
    void addNextCellTest5(String sourceType, String sourceValue, Cell.CellType expectedType) {
        Worksheet worksheet = new Worksheet();
        worksheet.setActiveStyle(BasicStyles.BorderFrameHeader());
        worksheet.setCurrentCellAddress("D2");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        assertEquals(0, worksheet.getCells().size());
        Object value = TestUtils.createInstance(sourceType, sourceValue);
        worksheet.addNextCell(value);
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", expectedType, BasicStyles.BorderFrameHeader(), value, 3, 2);
    }

    @DisplayName("Test of the addNextCell function for Date and LocalTime with active worksheet style")
    @Test()
    void addNextCellTest6() {
        Worksheet worksheet = new Worksheet();
        worksheet.setActiveStyle(BasicStyles.BorderFrameHeader());
        worksheet.setCurrentCellAddress("D2");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        assertEquals(0, worksheet.getCells().size());
        Date date = buildDate(2020, 6, 10, 11, 12, 22);
        worksheet.addNextCell(date);
        Style mixedStyle = BasicStyles.DateFormat();
        mixedStyle.append(BasicStyles.BorderFrameHeader());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.DATE, mixedStyle, date, 3, 2);
        worksheet.setCurrentCellAddress("E3");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.ColumnToColumn);
        Duration time = buildTime(6, 22, 13);
        worksheet.addNextCell(time);
        mixedStyle = BasicStyles.TimeFormat();
        mixedStyle.append(BasicStyles.BorderFrameHeader());
        WorksheetTest.assertAddedCell(worksheet, 2, "E3", Cell.CellType.TIME, mixedStyle, time, 5, 2);
    }

    @DisplayName("Test of the addNextCell function for a nested cell object")
    @Test()
    void addNextCellTest7() {
        Worksheet worksheet = new Worksheet();
        Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
        worksheet.setCurrentCellAddress("D2");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        worksheet.addNextCell(cell);
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, null, 33.3d, 3, 2);
    }

    @DisplayName("Test of the addNextCell function for a nested cell object and style")
    @Test()
    void addNextCellTest8() {
        Worksheet worksheet = new Worksheet();
        Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
        worksheet.setCurrentCellAddress("D2");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        worksheet.addNextCell(cell, BasicStyles.Bold());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.Bold(), 33.3d, 3, 2);
        cell = new Cell("test", Cell.CellType.STRING, "R2");
        cell.setStyle(BasicStyles.BorderFrame());
        Style mixedStyle = BasicStyles.BorderFrame();
        mixedStyle.append(BasicStyles.Bold());
        worksheet.addNextCell(cell, BasicStyles.Bold());
        WorksheetTest.assertAddedCell(worksheet, 2, "D3", Cell.CellType.STRING, mixedStyle, "test", 3, 3);
    }

    @DisplayName("Test of the addNextCell function for a nested cell object and active worksheet style")
    @Test()
    void addNextCellTest9() {
        Worksheet worksheet = new Worksheet();
        worksheet.setActiveStyle(BasicStyles.Bold());
        Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
        worksheet.setCurrentCellAddress("D2");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        worksheet.addNextCell(cell);
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.Bold(), 33.3d, 3, 2);
        cell = new Cell("test", Cell.CellType.STRING, "R2");
        cell.setStyle(BasicStyles.BorderFrame());
        Style mixedStyle = BasicStyles.BorderFrame();
        mixedStyle.append(BasicStyles.Bold());
        worksheet.addNextCell(cell);
        WorksheetTest.assertAddedCell(worksheet, 2, "D3", Cell.CellType.STRING, mixedStyle, "test", 3, 3);
    }

    @DisplayName("Test of the addNextCell function with when changing the current cell direction")
    @Test()
    void addNextCellTest10() {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentCellAddress("D2");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        worksheet.addNextCell("test");
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.STRING, null, "test", 3, 2);
        worksheet.setCurrentCellAddress("E3");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.ColumnToColumn);
        worksheet.addNextCell("test");
        WorksheetTest.assertAddedCell(worksheet, 2, "E3", Cell.CellType.STRING, null, "test", 5, 2);
        worksheet.setCurrentCellAddress("F5");
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.Disabled);
        worksheet.addNextCell("test");
        WorksheetTest.assertAddedCell(worksheet, 3, "F5", Cell.CellType.STRING, null, "test", 5, 4);
    }
}
