package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.ValueSource;
import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;

class RowTest {

    enum RowProperty {
        HIDDEN,
        HEIGHT
    }

    @Test
    @DisplayName("Test of the addHiddenRow function with a row number")
    void addHiddenRowTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getHiddenRows().isEmpty());
        worksheet.addHiddenRow(2);
        assertEquals(1, worksheet.getHiddenRows().size());
        assertTrue(worksheet.getHiddenRows().containsKey(2));
        assertTrue(worksheet.getHiddenRows().get(2));
        worksheet.addHiddenRow(2);  // Should not add an additional entry
        assertEquals(1, worksheet.getHiddenRows().size());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing addHiddenRow function with an invalid row number")
    @ValueSource(
            ints = {
                    -1,
                    -100,
                    1048576}
    )
    void addHiddenRowFailTest(int value) {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.addHiddenRow(value));
    }

    @Test
    @DisplayName("Test of the getCurrentRowNumber function")
    void getCurrentRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCurrentRowNumber());
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.ROW_TO_ROW);
        worksheet.addNextCell("test");
        worksheet.addNextCell("test");
        assertEquals(2, worksheet.getCurrentRowNumber());
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.COLUMN_TO_COLUMN);
        worksheet.addNextCell("test");
        worksheet.addNextCell("test");
        assertEquals(2, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow();
        assertEquals(3, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow(2);
        assertEquals(5, worksheet.getCurrentRowNumber());
        worksheet.goToNextColumn(2);
        assertEquals(0, worksheet.getCurrentRowNumber());
    }

    @ParameterizedTest
    @DisplayName("Test of the goToNextRow function")
    @CsvSource(
            {
                    "0,0,0",
                    "0,1,1",
                    "1,1,2",
                    "3,10,13",
                    "3,-1,2",
                    "3,-3,0"}
    )
    void goToNextRowTest(int initialRowNumber, int number, int expectedRowNumber) {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentRowNumber(initialRowNumber);
        worksheet.goToNextRow(number);
        assertEquals(expectedRowNumber, worksheet.getCurrentRowNumber());
    }

    @ParameterizedTest
    @DisplayName("Test of the goToNextRow function with the option to keep the column")
    @CsvSource(
            {
                    "A1,0,false,A1",
                    "A1,0,true,A1",
                    "A1,1,false,A2",
                    "A1,1,true,A2",
                    "C10,1,false,A11",
                    "C10,1,true,C11",
                    "R5,5,false,A10",
                    "R5,5,true,R10",
                    "F5,-3,false,A2",
                    "F5,-3,true,F2",
                    "F5,-4,false,A1",
                    "F5,-4,true,F1"
            }
    )
    void goToNextRowTest2(String initialAddress, int number, boolean keepColumnPosition, String expectedAddress) {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentCellAddress(initialAddress);
        worksheet.goToNextRow(number, keepColumnPosition);
        Address expected = new Address(expectedAddress);
        assertEquals(expected.column(), worksheet.getCurrentColumnNumber());
        assertEquals(expected.row(), worksheet.getCurrentRowNumber());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing goToNextRow function on invalid values")
    @CsvSource(
            {
                    "0,-1",
                    "10,-12",
                    "0,1048576",
                    "0,1248575"}
    )
    void goToNextRowFailTest(int initialValue, int value) {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentRowNumber(initialValue);
        assertEquals(initialValue, worksheet.getCurrentRowNumber());
        assertThrows(RangeException.class, () -> worksheet.goToNextRow(value));
    }

    @Test
    @DisplayName("Test of the removeRowHeight function")
    void removeRowHeightTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.setRowHeight(2, 22.2f);
        worksheet.setRowHeight(4, 33.3f);
        assertEquals(2, worksheet.getRowHeights().size());
        worksheet.removeRowHeight(2);
        assertEquals(1, worksheet.getRowHeights().size());
        worksheet.removeRowHeight(3);
        worksheet.removeRowHeight(-1);
        assertEquals(1, worksheet.getRowHeights().size());
    }

    @ParameterizedTest
    @DisplayName("Test of the setCurrentRowNumber function")
    @ValueSource(
            ints = {
                    0,
                    3,
                    1048575}
    )
    void setCurrentRowNumberTest(int row) {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow();
        worksheet.setCurrentRowNumber(row);
        assertEquals(row, worksheet.getCurrentRowNumber());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing setCurrentRowNumber function")
    @ValueSource(
            ints = {
                    -1,
                    -10,
                    1048576}
    )
    void setCurrentRowNumberFailTest(int row) {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.setCurrentRowNumber(row));
    }

    @ParameterizedTest
    @DisplayName("Test of the setRowHeight function")
    @ValueSource(
            floats = {
                    0f,
                    0.1f,
                    10f,
                    255f}
    )
    void setRowHeightTest(float height) {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getRowHeights().isEmpty());
        worksheet.setRowHeight(0, height);
        assertEquals(1, worksheet.getRowHeights().size());
        assertEquals(height, worksheet.getRowHeights().get(0));
        worksheet.setRowHeight(0, Worksheet.DEFAULT_WORKSHEET_ROW_HEIGHT);
        assertEquals(1, worksheet.getRowHeights().size());
        assertEquals(Worksheet.DEFAULT_WORKSHEET_ROW_HEIGHT, worksheet.getRowHeights().get(0));
    }

    @ParameterizedTest
    @DisplayName("Test of the failing setRowHeight function")
    @CsvSource(
            {
                    "-1,0",
                    "1048576,0.0",
                    "0,-10",
                    "0,409.51",
                    "0,500"}
    )
    void setRowHeightFailTest(int rowNumber, float height) {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.setRowHeight(rowNumber, height));
    }

    @Test
    @DisplayName("Test of the getRow function")
    void getRowTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(22, "B1");
        worksheet.addCell(23, "B2");
        worksheet.addCell("test", "C2");
        worksheet.addCell(true, "D2");
        worksheet.addCell(false, "B3");
        List<Cell> row = worksheet.getRow(1);
        assertEquals(3, row.size());
        assertEquals(23, row.get(0).getValue());
        assertEquals("test", row.get(1).getValue());
        assertEquals(true, row.get(2).getValue());
    }

    @Test
    @DisplayName("Test of the getRow function when no values are applying")
    void getRowTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(22, "B1");
        worksheet.addCell(false, "B3");
        List<Cell> row = worksheet.getRow(1);
        assertTrue(row.isEmpty());
    }
}
