package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.ValueSource;
import ch.rabanti.nanoxlsx4j.Worksheet;

class GetColumnBoundariesTest {
    @Test
    @DisplayName("Test of the getLastColumnNumber function with an empty worksheet")
    void getLastColumnNumberTest() {
        Worksheet worksheet = new Worksheet();
        int column = worksheet.getLastColumnNumber();
        assertEquals(-1, column);
    }

    @Test
    @DisplayName("Test of the getLastColumnNumber function with defined columns on an empty worksheet")
    void getLastColumnNumberTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(2);
        int column = worksheet.getLastColumnNumber();
        assertEquals(2, column);
    }

    @Test
    @DisplayName("Test of the getLastColumnNumber function with defined columns on an empty worksheet, where the " +
            "column definition has gaps")
    void getLastColumnNumberTest3() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(10);
        int column = worksheet.getLastColumnNumber();
        assertEquals(10, column);
    }

    @Test
    @DisplayName("Test of the getLastColumnNumber function with defined columns where cells are defined below the " +
            "last column")
    void getLastColumnNumberTest4() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(10);
        worksheet.addCell("test", "E5");
        int column = worksheet.getLastColumnNumber();
        assertEquals(10, column);
    }

    @Test
    @DisplayName("Test of the getLastColumnNumber function with defined columns where cells are defined above the " +
            "last column")
    void getLastColumnNumberTest5() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(2);
        worksheet.addCell("test", "F5");
        int column = worksheet.getLastColumnNumber();
        assertEquals(5, column);
    }

    @ParameterizedTest
    @CsvSource(
            {
                    "F5,5",
                    "A1,4"}
    )
    @DisplayName("Test of the getLastColumnNumber function with an explicitly defined, empty cell besides other " +
            "column definitions")
    void getLastColumnNumberTest6(String emptyCellAddress, int expectedLastColumn) {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(4);
        worksheet.addCell(null, emptyCellAddress);
        int column = worksheet.getLastColumnNumber();
        assertEquals(expectedLastColumn, column);
    }

    @Test
    @DisplayName("Test of the getFirstColumnNumber function with an empty worksheet")
    void getFirstColumnNumberTest() {
        Worksheet worksheet = new Worksheet();
        int column = worksheet.getFirstColumnNumber();
        assertEquals(-1, column);
    }

    @Test
    @DisplayName("Test of the getFirstColumnNumber function with defined columns on an empty worksheet")
    void getFirstColumnNumberTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(3);
        int column = worksheet.getFirstColumnNumber();
        assertEquals(1, column);
    }

    @Test
    @DisplayName("Test of the getFirstColumnNumber function with defined columns on an empty worksheet, where the " +
            "column definition has gaps")
    void getFirstColumnNumberTest3() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(10);
        int column = worksheet.getFirstColumnNumber();
        assertEquals(1, column);
    }

    @Test
    @DisplayName("Test of the getFirstColumnNumber function with defined columns where cells are defined above the " +
            "first column")
    void getFirstColumnNumberTest4() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(8);
        worksheet.addHiddenColumn(10);
        worksheet.addCell("test", "E5");
        int column = worksheet.getFirstColumnNumber();
        assertEquals(3, column);
    }

    @Test
    @DisplayName("Test of the getFirstColumnNumber function with defined columns where cells are defined below the " +
            "first column")
    void getFirstColumnNumberTest5() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(7);
        worksheet.addHiddenColumn(8);
        worksheet.addHiddenColumn(9);
        worksheet.addCell("test", "F5");
        int column = worksheet.getFirstColumnNumber();
        assertEquals(5, column);
    }

    @ParameterizedTest
    @CsvSource(
            {
                    "F5,3",
                    "A1,0"}
    )
    @DisplayName("Test of the getFirstColumnNumber function with an explicitly defined, empty cell besides other " +
            "column definitions")
    void getFirstColumnNumberTest6(String emptyCellAddress, int expectedFirstRow) {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(4);
        worksheet.addCell(null, emptyCellAddress);
        int column = worksheet.getFirstColumnNumber();
        assertEquals(expectedFirstRow, column);
    }

    @Test
    @DisplayName("Test of the getLastDataColumnNumber function with an empty worksheet")
    void getLastDataColumnNumberTest() {
        Worksheet worksheet = new Worksheet();
        int column = worksheet.getLastDataColumnNumber();
        assertEquals(-1, column);
    }

    @Test
    @DisplayName("Test of the getLastDataColumnNumber function with defined columns on an empty worksheet")
    void getLastDataColumnNumberTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(2);
        int column = worksheet.getLastDataColumnNumber();
        assertEquals(-1, column);
    }

    @Test
    @DisplayName("Test of the getLastDataColumnNumber function with defined columns where cells are defined below the" +
            " last column")
    void getLastDataColumnNumberTest3() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(10);
        worksheet.addCell("test", "E5");
        int column = worksheet.getLastDataColumnNumber();
        assertEquals(4, column);
    }

    @Test
    @DisplayName("Test of the getLastDataColumnNumber function with defined columns where cells are defined above the" +
            " last column")
    void getLastDataColumnNumberTest4() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(2);
        worksheet.addCell("test", "F5");
        int column = worksheet.getLastDataColumnNumber();
        assertEquals(5, column);
    }

    @Test
    @DisplayName("Test of the getLastDataColumnNumber function with two defined columns")
    void getLastDataColumnNumberTest5() {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell("test", "A1");
        worksheet.addCell("test", "B1");
        int column = worksheet.getLastDataColumnNumber();
        assertEquals(1, column);
    }

    @Test
    @DisplayName("Test of the getFirstDataColumnNumber function with an empty worksheet")
    void getFirstDataColumnNumberTest() {
        Worksheet worksheet = new Worksheet();
        int column = worksheet.getFirstDataColumnNumber();
        assertEquals(-1, column);
    }

    @Test
    @DisplayName("Test of the getFirstDataColumnNumber function with defined columns on an empty worksheet")
    void getFirstDataColumnNumberTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(2);
        int column = worksheet.getFirstDataColumnNumber();
        assertEquals(-1, column);
    }

    @Test
    @DisplayName("Test of the getFirstDataColumnNumber function with defined columns where cells are defined above " +
            "the first column")
    void getFirstDataColumnNumberTest3() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(10);
        worksheet.addCell("test", "E5");
        int column = worksheet.getFirstDataColumnNumber();
        assertEquals(4, column);
    }

    @Test
    @DisplayName("Test of the getFirstDataColumnNumber function with defined columns where cells are defined below " +
            "the first column")
    void getFirstDataColumnNumberTest4() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(10);
        worksheet.addCell("test", "F5");
        int column = worksheet.getFirstDataColumnNumber();
        assertEquals(5, column);
    }

    @Test
    @DisplayName("Test of the getFirstDataColumnNumber function with two defined columns")
    void getFirstDataColumnNumberTest5() {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell("test", "A1");
        worksheet.addCell("test", "B1");
        int column = worksheet.getFirstDataColumnNumber();
        assertEquals(0, column);
    }

    @ParameterizedTest
    @ValueSource(
            strings = {
                    "F5",
                    "A1"}
    )
    @DisplayName("Test of the getFirstDataColumnNumber and getLastDataColumnNumber functions with an explicitly " +
            "defined, empty cell besides other column definitions")
    void getFirstOrLastDataColumnNumberTest(String emptyCellAddress) {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(4);
        worksheet.addCell(null, emptyCellAddress);
        int minColumn = worksheet.getFirstDataColumnNumber();
        int maxColumn = worksheet.getLastDataColumnNumber();
        assertEquals(-1, minColumn);
        assertEquals(-1, maxColumn);
    }

    @Test
    @DisplayName("Test of the getFirstDataColumnNumber and getLastDataColumnNumber functions with exactly one defined" +
            " cell")
    void getFirstOrLastDataColumnNumberTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(10);
        worksheet.addCell("test", "F5");
        int minColumn = worksheet.getFirstDataColumnNumber();
        int maxColumn = worksheet.getLastDataColumnNumber();
        assertEquals(5, minColumn);
        assertEquals(5, maxColumn);
    }

    @ParameterizedTest
    @ValueSource(
            strings = {
                    "F5",
                    "A1"}
    )
    @DisplayName("Test of the getFirstDataColumnNumber and getLastDataColumnNumber functions with an explicitly " +
            "defined, empty cell with empty string besides other column definitions")
    void getFirstOrLastDataColumnNumberTest3(String emptyCellAddress) {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(4);
        worksheet.addCell("", emptyCellAddress);
        int minColumn = worksheet.getFirstDataColumnNumber();
        int maxColumn = worksheet.getLastDataColumnNumber();
        assertEquals(-1, minColumn);
        assertEquals(-1, maxColumn);
    }

}
