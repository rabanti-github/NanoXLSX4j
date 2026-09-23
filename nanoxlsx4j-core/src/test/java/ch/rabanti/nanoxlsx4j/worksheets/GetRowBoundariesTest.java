package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.ValueSource;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.worksheets.RowTest.RowProperty;

class GetRowBoundariesTest {
    @Test
    @DisplayName("Test of the getLastRowNumber function with an empty worksheet")
    void getLastRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getLastRowNumber();
        assertEquals(-1, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getLastRowNumber function with defined rows on an empty worksheet")
    void getLastRowNumberTest2(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        int row = worksheet.getLastRowNumber();
        assertEquals(2, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getLastRowNumber function with defined rows on an empty worksheet, where the row " +
            "definition has gaps")
    void getLastRowNumberTest3(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        int row = worksheet.getLastRowNumber();
        assertEquals(10, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getLastRowNumber function with defined rows where cells are defined below the last row")
    void getLastRowNumberTest4(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        worksheet.addCell("test", "E5");
        int row = worksheet.getLastRowNumber();
        assertEquals(10, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getLastRowNumber function with defined rows where cells are defined above the last row")
    void getLastRowNumberTest5(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        worksheet.addCell("test", "F5");
        int row = worksheet.getLastRowNumber();
        assertEquals(4, row);
    }

    @ParameterizedTest
    @CsvSource(
            {
                    "F7,6",
                    "A1,4"}
    )
    @DisplayName("Test of the getLastRowNumber function with an explicitly defined, empty cell besides other row " +
            "definitions")
    void getLastRowNumberTest6(String emptyCellAddress, int expectedFirstRow) {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenRow(3);
        worksheet.addHiddenRow(4);
        worksheet.addCell(null, emptyCellAddress);
        int row = worksheet.getLastRowNumber();
        assertEquals(expectedFirstRow, row);
    }

    @Test
    @DisplayName("Test of the getLastDataRowNumber function with an empty worksheet")
    void getLastDataRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getLastDataRowNumber();
        assertEquals(-1, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getLastDataRowNumber function with defined rows on an empty worksheet")
    void getLastDataRowNumberTest2(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        int row = worksheet.getLastDataRowNumber();
        assertEquals(-1, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getLastDataRowNumber function with defined rows where cells are defined below the last " +
            "row")
    void getLastDataRowNumberTest3(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }

        worksheet.addCell("test", "E5");
        int row = worksheet.getLastDataRowNumber();
        assertEquals(4, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getLastDataRowNumber function with defined rows where cells are defined above the last " +
            "row")
    void getLastDataRowNumberTest4(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        } else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(3, 44.4f);
        }

        worksheet.addCell("test", "F5");
        int row = worksheet.getLastDataRowNumber();
        assertEquals(4, row);
    }

    @Test
    @DisplayName("Test of the getFirstRowNumber function with an empty worksheet")
    void getFirstRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getFirstRowNumber();
        assertEquals(-1, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getFirstRowNumber function with defined rows on an empty worksheet")
    void getFirstRowNumberTest2(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(3);
        } else {
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 33.3f);
            worksheet.setRowHeight(3, 44.4f);
        }
        int row = worksheet.getFirstRowNumber();
        assertEquals(1, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getFirstRowNumber function with defined rows on an empty worksheet, where the row " +
            "definition has gaps")
    void getFirstRowNumberTest3(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(10);
        } else {
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        int row = worksheet.getFirstRowNumber();
        assertEquals(1, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getFirstRowNumber function with defined rows where cells are defined above the first row")
    void getFirstRowNumberTest4(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(3);
            worksheet.addHiddenRow(10);
        } else {
            worksheet.setRowHeight(2, 22.2f);
            worksheet.setRowHeight(3, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        worksheet.addCell("test", "E5");
        int row = worksheet.getFirstRowNumber();
        assertEquals(2, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getFirstRowNumber function with defined rows where cells are defined below the first row")
    void getFirstRowNumberTest5(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(6);
            worksheet.addHiddenRow(7);
            worksheet.addHiddenRow(8);
        } else {
            worksheet.setRowHeight(6, 22.2f);
            worksheet.setRowHeight(7, 33.3f);
            worksheet.setRowHeight(8, 44.4f);
        }
        worksheet.addCell("test", "F5");
        int row = worksheet.getFirstRowNumber();
        assertEquals(4, row);
    }

    @ParameterizedTest
    @CsvSource(
            {
                    "F5,4",
                    "A1,0"}
    )
    @DisplayName("Test of the getFirstRowNumber function with an explicitly defined, empty cell besides other row " +
            "definitions")
    void getFirstRowNumberTest6(String emptyCellAddress, int expectedFirstRow) {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(4);
        worksheet.addCell(null, emptyCellAddress);
        int row = worksheet.getFirstRowNumber();
        assertEquals(expectedFirstRow, row);
    }

    @Test
    @DisplayName("Test of the getFirstDataRowNumber function with an empty worksheet")
    void getFirstDataRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getFirstDataRowNumber();
        assertEquals(-1, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getFirstDataRowNumber function with defined rows on an empty worksheet")
    void getFirstDataRowNumberTest2(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(3);
        } else {
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 33.3f);
            worksheet.setRowHeight(3, 44.4f);
        }
        int row = worksheet.getFirstDataRowNumber();
        assertEquals(-1, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getFirstDataRowNumber function with defined rows where cells are defined below the last" +
            " row")
    void getFirstDataRowNumberTest3(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(3);
            worksheet.addHiddenRow(10);
        } else {
            worksheet.setRowHeight(2, 22.2f);
            worksheet.setRowHeight(3, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }

        worksheet.addCell("test", "E5");
        int row = worksheet.getFirstDataRowNumber();
        assertEquals(4, row);
    }

    @ParameterizedTest
    @EnumSource(
            value = RowProperty.class,
            names = {
                    "HEIGHT",
                    "HIDDEN"}
    )
    @DisplayName("Test of the getFirstDataRowNumber function with defined rows where cells are defined above the last" +
            " row")
    void getFirstDataRowNumberTest4(RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.HIDDEN) {
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(3);
        } else {
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 33.3f);
            worksheet.setRowHeight(3, 44.4f);
        }

        worksheet.addCell("test", "F6");
        int row = worksheet.getFirstDataRowNumber();
        assertEquals(5, row);
    }

    @ParameterizedTest
    @ValueSource(
            strings = {
                    "F5",
                    "A1"}
    )
    @DisplayName("Test of the getFirstDataRowNumber and getLastDataRowNumber functions with an explicitly defined, " +
            "empty cell besides other row definitions")
    void getFirstOrLastDataRowNumberTest(String emptyCellAddress) {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenRow(3);
        worksheet.addHiddenRow(4);
        worksheet.addCell(null, emptyCellAddress);
        int minRow = worksheet.getFirstDataRowNumber();
        int maxRow = worksheet.getLastDataRowNumber();
        assertEquals(-1, minRow);
        assertEquals(-1, maxRow);
    }

    @Test
    @DisplayName("Test of the getFirstDataRowNumber and getLastDataRowNumber functions with exactly one defined cell")
    void getFirstOrLastDataRowNumberTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenRow(2);
        worksheet.addHiddenRow(3);
        worksheet.addHiddenRow(10);
        worksheet.addCell("test", "C5");
        int minRow = worksheet.getFirstDataRowNumber();
        int maxRow = worksheet.getLastDataRowNumber();
        assertEquals(4, minRow);
        assertEquals(4, maxRow);
    }

    @ParameterizedTest
    @ValueSource(
            strings = {
                    "F5",
                    "A1"}
    )
    @DisplayName("Test of the getFirstDataRowNumber and getLastDataRowNumber functions with an explicitly defined, " +
            "cell with empty string besides other row definitions")
    void getFirstOrLastDataRowNumberTest3(String emptyCellAddress) {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenRow(3);
        worksheet.addHiddenRow(4);
        worksheet.addCell("", emptyCellAddress);
        int minRow = worksheet.getFirstDataRowNumber();
        int maxRow = worksheet.getLastDataRowNumber();
        assertEquals(-1, minRow);
        assertEquals(-1, maxRow);
    }

}
