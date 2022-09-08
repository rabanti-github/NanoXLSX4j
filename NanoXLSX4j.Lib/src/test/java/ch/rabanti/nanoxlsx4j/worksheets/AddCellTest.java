package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Address;
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
import java.util.Date;
import java.util.function.BiConsumer;

import static ch.rabanti.nanoxlsx4j.TestUtils.buildDate;
import static ch.rabanti.nanoxlsx4j.TestUtils.buildTime;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class AddCellTest {

    private Worksheet worksheet;

    @DisplayName("Test of the addCell function with the only the value (with address and column/row invocation)")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource({
            "NULL, '', 0, 0,EMPTY, A1",
            "STRING, '', 2, 2,STRING, C3",
            "STRING, 'test', 5, 1,STRING, F2",
            "LONG, '17', 16383, 0,NUMBER, XFD1",
            "DOUBLE, '1.02', 0, 1048575,NUMBER, A1048576",
            "FLOAT, '-22.3', 16383, 1048575,NUMBER, XFD1048576",
            "INTEGER, '0', 0, 0,NUMBER, A1",
            "BYTE, '127', 2, 2,NUMBER, C3",
            "BOOLEAN, 'true', 5, 1,BOOL, F2",
            "BOOLEAN, 'false', 16383, 0,BOOL, XFD1",
    })
    void addCellTest1(String sourceType, String sourceValue, int column, int row, Cell.CellType expectedType, String expectedAddress) {
        Object value = TestUtils.createInstance(sourceType, sourceValue);
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
        invokeAddCellTest(value, column, row, worksheet::addCell, expectedType, expectedAddress, column, row + 1);
        Address address = new Address(column, row);
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
        invokeAddCellTest(value, address.getAddress(), worksheet::addCell, expectedType, expectedAddress, column + 1, row);
    }

    @DisplayName("Test of the addCell function with value and Style (with address and column/row invocation)")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource({
            "NULL, '', 0, 0,EMPTY, A1",
            "STRING, '', 2, 2,STRING, C3",
            "STRING, 'test', 5, 1,STRING, F2",
            "LONG, '17', 16383, 0,NUMBER, XFD1",
            "DOUBLE, '1.02', 0, 1048575,NUMBER, A1048576",
            "FLOAT, '-22.3', 16383, 1048575,NUMBER, XFD1048576",
            "INTEGER, '0', 0, 0,NUMBER, A1",
            "BYTE, '127', 2, 2,NUMBER, C3",
            "BOOLEAN, 'true', 5, 1,BOOL, F2",
            "BOOLEAN, 'false', 16383, 0,BOOL, XFD1",
    })
    void addCellTest2(String sourceType, String sourceValue, int column, int row, Cell.CellType expectedType, String expectedAddress) {
        Object value = TestUtils.createInstance(sourceType, sourceValue);
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
        invokeAddCellTest(value, column, row, BasicStyles.BoldItalic(), worksheet::addCell, expectedType, expectedAddress, column, row + 1, BasicStyles.BoldItalic());
        Address address = new Address(column, row);
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
        invokeAddCellTest(value, address.getAddress(), BasicStyles.Bold(), worksheet::addCell, expectedType, expectedAddress, column + 1, row, BasicStyles.Bold());
    }

    @DisplayName("Test of the addCell function for DateTime and TimeSpan (with address and column/row invocation)")
    @Test()
    void addCellTest3() {
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
        Date date = buildDate(2020, 6, 10, 11, 12, 22);
        invokeAddCellTest(date, 5, 1, worksheet::addCell, Cell.CellType.DATE, "F2", 5, 2, BasicStyles.DateFormat());
        Address address = new Address(5, 1);
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
        invokeAddCellTest(date, address.getAddress(), worksheet::addCell, Cell.CellType.DATE, "F2", 6, 1, BasicStyles.DateFormat());

        worksheet = WorksheetTest.initWorksheet(worksheet, "S9", Worksheet.CellDirection.RowToRow);
        Duration time = buildTime(6, 22, 13);
        invokeAddCellTest(time, 5, 1, worksheet::addCell, Cell.CellType.TIME, "F2", 5, 2, BasicStyles.TimeFormat());
        worksheet = WorksheetTest.initWorksheet(worksheet, "V6", Worksheet.CellDirection.ColumnToColumn);
        invokeAddCellTest(time, address.getAddress(), worksheet::addCell, Cell.CellType.TIME, "F2", 6, 1, BasicStyles.TimeFormat());
    }

    @DisplayName("Test of the addCell function for DateTime and TimeSpan with styles (with address and column/row invocation)")
    @Test()
    void addCellTest4() {
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
        Date date = buildDate(2020, 6, 10, 11, 12, 22);
        Style mixedStyle = BasicStyles.DateFormat();
        mixedStyle.append(BasicStyles.Bold());
        invokeAddCellTest(date, 5, 1, BasicStyles.Bold(), worksheet::addCell, Cell.CellType.DATE, "F2", 5, 2, mixedStyle);
        Address address = new Address(5, 1);
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
        invokeAddCellTest(date, address.getAddress(), BasicStyles.Bold(), worksheet::addCell, Cell.CellType.DATE, "F2", 6, 1, mixedStyle);

        worksheet = WorksheetTest.initWorksheet(worksheet, "S9", Worksheet.CellDirection.RowToRow);
        Duration time = buildTime(6, 22, 13);
        mixedStyle = BasicStyles.TimeFormat();
        mixedStyle.append(BasicStyles.Underline());
        invokeAddCellTest(time, 5, 1, BasicStyles.Underline(), worksheet::addCell, Cell.CellType.TIME, "F2", 5, 2, mixedStyle);
        worksheet = WorksheetTest.initWorksheet(worksheet, "V6", Worksheet.CellDirection.ColumnToColumn);
        invokeAddCellTest(time, address.getAddress(), BasicStyles.Underline(), worksheet::addCell, Cell.CellType.TIME, "F2", 6, 1, mixedStyle);
    }


    @DisplayName("Test of the addCell function with value and active worksheet style (with address and column/row invocation)")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource({
            "NULL, '', 0, 0,EMPTY, A1",
            "STRING, '', 2, 2,STRING, C3",
            "STRING, 'test', 5, 1,STRING, F2",
            "LONG, '17', 16383, 0,NUMBER, XFD1",
            "DOUBLE, '1.02', 0, 1048575,NUMBER, A1048576",
            "FLOAT, '-22.3', 16383, 1048575,NUMBER, XFD1048576",
            "INTEGER, '0', 0, 0,NUMBER, A1",
            "BYTE, '127', 2, 2,NUMBER, C3",
            "BOOLEAN, 'true', 5, 1,BOOL, F2",
            "BOOLEAN, 'false', 16383, 0,BOOL, XFD1",
    })
    void addCellTest5(String sourceType, String sourceValue, int column, int row, Cell.CellType expectedType, String expectedAddress) {
        Object value = TestUtils.createInstance(sourceType, sourceValue);
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrameHeader());
        invokeAddCellTest(value, column, row, worksheet::addCell, expectedType, expectedAddress, column, row + 1, BasicStyles.BorderFrameHeader());
        Address address = new Address(column, row);
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn, BasicStyles.BorderFrameHeader());
        invokeAddCellTest(value, address.getAddress(), worksheet::addCell, expectedType, expectedAddress, column + 1, row, BasicStyles.BorderFrameHeader());
    }

    @DisplayName("Test of the addCell function for DateTime and TimeSpan with active worksheet style (with address and column/row invocation)")
    @Test()
    void addCellTest6() {
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrameHeader());
        Date date = buildDate(2020, 6, 10, 11, 12, 22);
        Style mixedStyle = BasicStyles.DateFormat();
        mixedStyle.append(BasicStyles.BorderFrameHeader());
        invokeAddCellTest(date, 5, 1, worksheet::addCell, Cell.CellType.DATE, "F2", 5, 2, mixedStyle);
        Address address = new Address(5, 1);
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn, BasicStyles.BorderFrameHeader());
        invokeAddCellTest(date, address.getAddress(), worksheet::addCell, Cell.CellType.DATE, "F2", 6, 1, mixedStyle);

        worksheet = WorksheetTest.initWorksheet(worksheet, "S9", Worksheet.CellDirection.RowToRow, BasicStyles.Underline());
        Duration time = buildTime(6, 22, 13);
        mixedStyle = BasicStyles.TimeFormat();
        mixedStyle.append(BasicStyles.Underline());
        invokeAddCellTest(time, 5, 1, worksheet::addCell, Cell.CellType.TIME, "F2", 5, 2, mixedStyle);
        worksheet = WorksheetTest.initWorksheet(worksheet, "V6", Worksheet.CellDirection.ColumnToColumn, BasicStyles.Underline());
        invokeAddCellTest(time, address.getAddress(), worksheet::addCell, Cell.CellType.TIME, "F2", 6, 1, mixedStyle);
    }


    @DisplayName("Test of the addCell function for a nested cell object (with address and column/row invocation)")
    @Test()
    void addCellTest7() {
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
        Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
        worksheet.addCell(cell, 3, 1);
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, null, 33.3d, 3, 2);
        worksheet = new Worksheet();
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
        Address address = new Address(3, 1);
        worksheet.addCell(cell, address.getAddress());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, null, 33.3d, 4, 1);
    }


    @DisplayName("Test of the addCell function for a nested cell object and style (with address and column/row invocation)")
    @Test()
    void addCellTest8() {
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
        Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
        worksheet.addCell(cell, 3, 1, BasicStyles.Bold());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.Bold(), 33.3d, 3, 2);
        worksheet = new Worksheet();
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
        Address address = new Address(3, 1);
        worksheet.addCell(cell, address.getAddress());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.Bold(), 33.3d, 4, 1);

        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
        cell = new Cell("test", Cell.CellType.STRING, "R2");
        cell.setStyle(BasicStyles.BorderFrame());
        Style mixedStyle = BasicStyles.BorderFrame();
        mixedStyle.append(BasicStyles.Bold());
        worksheet.addCell(cell, 3, 1, BasicStyles.Bold());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.STRING, mixedStyle, "test", 3, 2);
        worksheet = new Worksheet();
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.ColumnToColumn);
        worksheet.addCell(cell, address.getAddress(), BasicStyles.Bold());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.STRING, mixedStyle, "test", 4, 1);
    }


    @DisplayName("Test of the addCell function for a nested cell object and active worksheet style (with address and column/row invocation)")
    @Test()
    void addCellTest9() {
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrame());
        Cell cell = new Cell(33.3d, Cell.CellType.NUMBER, "R1"); // Address should be replaced
        worksheet.addCell(cell, 3, 1);
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.BorderFrame(), 33.3d, 3, 2);
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.ColumnToColumn, BasicStyles.BorderFrame());
        Address address = new Address(3, 1);
        worksheet.addCell(cell, address.getAddress());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.NUMBER, BasicStyles.BorderFrame(), 33.3d, 4, 1);

        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.BorderFrame());
        cell = new Cell("test", Cell.CellType.STRING, "R2");
        cell.setStyle(BasicStyles.Bold());
        Style mixedStyle = BasicStyles.BorderFrame();
        mixedStyle.append(BasicStyles.Bold());
        worksheet.addCell(cell, 3, 1);
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.STRING, mixedStyle, "test", 3, 2);
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.ColumnToColumn, BasicStyles.BorderFrame());
        worksheet.addCell(cell, address.getAddress());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.STRING, mixedStyle, "test", 4, 1);
    }


    @DisplayName("Test of the addCell function with when changing the current cell direction (with address and column/row invocation)")
    @ParameterizedTest(name = "Given column {1} and row {2} with {3} should lead to the next column {4} and row {5}")
    @CsvSource({
            "D2, 3, 1, RowToRow, 3, 2",
            "E7, 7, 2, ColumnToColumn, 8, 2",
            "C9, 10, 5, Disabled, 2, 8",
    })
    void addCellTest10(String worksheetAddress, int initialColumn, int initialRow, Worksheet.CellDirection cellDirection, int expectedNextColumn, int expectedNextRow) {
        Address initialAddress = new Address(initialColumn, initialRow);
        worksheet = WorksheetTest.initWorksheet(worksheet, worksheetAddress, cellDirection);
        invokeAddCellTest("test", initialColumn, initialRow, worksheet::addCell, Cell.CellType.STRING, initialAddress.getAddress(), expectedNextColumn, expectedNextRow);
        worksheet = WorksheetTest.initWorksheet(worksheet, worksheetAddress, cellDirection);
        invokeAddCellTest("test", initialAddress.getAddress(), worksheet::addCell, Cell.CellType.STRING, initialAddress.getAddress(), expectedNextColumn, expectedNextRow);
    }

    @DisplayName("Test of the addCell function where an existing cell is overwritten")
    @Test()
    void addCellOverwriteTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell("test", "C2");
        assertEquals(Cell.CellType.STRING, worksheet.getCells().get("C2").getDataType());
        assertEquals("test", worksheet.getCells().get("C2").getValue());
        worksheet.addCell(22, "C2");
        assertEquals(Cell.CellType.NUMBER, worksheet.getCells().get("C2").getDataType());
        assertEquals(22, worksheet.getCells().get("C2").getValue());
        assertEquals(1, worksheet.getCells().size());
    }

    @DisplayName("Test of the addCell function where existing cells are overwritten and the old cells where dates and times")
    @Test()
    void addCellOverwriteTest2() {
        Worksheet worksheet = new Worksheet();
        Date date = buildDate(2020, 10, 5, 4, 11, 12);
        Duration time = buildTime(11, 12, 13);
        worksheet.addCell(date, "C2");
        worksheet.addCell(time, "C3");
        assertEquals(Cell.CellType.DATE, worksheet.getCells().get("C2").getDataType());
        assertEquals(date, worksheet.getCells().get("C2").getValue());
        assertTrue(BasicStyles.DateFormat().equals(worksheet.getCells().get("C2").getCellStyle()));
        assertEquals(Cell.CellType.TIME, worksheet.getCells().get("C3").getDataType());
        assertEquals(time, worksheet.getCells().get("C3").getValue());
        assertTrue(BasicStyles.TimeFormat().equals(worksheet.getCells().get("C3").getCellStyle()));
        worksheet.addCell(22, "C2");
        worksheet.addCell("test", "C3");
        assertEquals(Cell.CellType.NUMBER, worksheet.getCells().get("C2").getDataType());
        assertEquals(22, worksheet.getCells().get("C2").getValue());
        assertNull(worksheet.getCells().get("C2").getCellStyle());
        assertEquals(Cell.CellType.STRING, worksheet.getCells().get("C3").getDataType());
        assertEquals("test", worksheet.getCells().get("C3").getValue());
        assertNull(worksheet.getCells().get("C3").getCellStyle());
        assertEquals(2, worksheet.getCells().size());
    }


    private <T1> void invokeAddCellTest(Object value, T1 parameter1, BiConsumer<Object, T1> action, Cell.CellType expectedType, String expectedAddress, int expectedNextColumn, int expectedNextRow) {
        invokeAddCellTest(value, parameter1, action, expectedType, expectedAddress, expectedNextColumn, expectedNextRow, null);
    }

    private <T1> void invokeAddCellTest(Object value, T1 parameter1, BiConsumer<Object, T1> action, Cell.CellType expectedType, String expectedAddress, int expectedNextColumn, int expectedNextRow, Style expectedStyle) {
        assertEquals(0, worksheet.getCells().size());
        action.accept(value, parameter1);
        WorksheetTest.assertAddedCell(worksheet, 1, expectedAddress, expectedType, expectedStyle, value, expectedNextColumn, expectedNextRow);
        worksheet = new Worksheet(); // Auto-reset
    }

    private <T1, T2> void invokeAddCellTest(Object value, T1 parameter1, T2 parameter2, TestUtils.TriConsumer<Object, T1, T2> action, Cell.CellType expectedType, String expectedAddress, int expectedNextColumn, int expectedNextRow) {
        invokeAddCellTest(value, parameter1, parameter2, action, expectedType, expectedAddress, expectedNextColumn, expectedNextRow, null);
    }

    private <T1, T2> void invokeAddCellTest(Object value, T1 parameter1, T2 parameter2, TestUtils.TriConsumer<Object, T1, T2> action, Cell.CellType expectedType, String expectedAddress, int expectedNextColumn, int expectedNextRow, Style expectedStyle) {
        assertEquals(0, worksheet.getCells().size());
        action.accept(value, parameter1, parameter2);
        WorksheetTest.assertAddedCell(worksheet, 1, expectedAddress, expectedType, expectedStyle, value, expectedNextColumn, expectedNextRow);
        worksheet = new Worksheet(); // Auto-reset
    }

    private <T1, T2, T3> void invokeAddCellTest(Object value, T1 parameter1, T2 parameter2, T3 parameter3, TestUtils.QuadConsumer<Object, T1, T2, T3> action, Cell.CellType expectedType, String expectedAddress, int expectedNextColumn, int expectedNextRow) {
        invokeAddCellTest(value, parameter1, parameter2, parameter3, action, expectedType, expectedAddress, expectedNextColumn, expectedNextRow, null);
    }

    private <T1, T2, T3> void invokeAddCellTest(Object value, T1 parameter1, T2 parameter2, T3 parameter3, TestUtils.QuadConsumer<Object, T1, T2, T3> action, Cell.CellType expectedType, String expectedAddress, int expectedNextColumn, int expectedNextRow, Style expectedStyle) {
        assertEquals(0, worksheet.getCells().size());
        action.accept(value, parameter1, parameter2, parameter3);
        WorksheetTest.assertAddedCell(worksheet, 1, expectedAddress, expectedType, expectedStyle, value, expectedNextColumn, expectedNextRow);
        worksheet = new Worksheet(); // Auto-reset
    }

}
