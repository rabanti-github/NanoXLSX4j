package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Worksheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class GetRowBoundariesTest {

    @DisplayName("Test of the getLastRowNumber function with an empty worksheet")
    @Test()
    void getLastRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getLastRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows on an empty worksheet")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getLastRowNumberTest2(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        }
        else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        int row = worksheet.getLastRowNumber();
        assertEquals(2, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows on an empty worksheet, where the row definition has gaps")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getLastRowNumberTest3(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        }
        else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        int row = worksheet.getLastRowNumber();
        assertEquals(10, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows where cells are defined below the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getLastRowNumberTest4(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        }
        else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        worksheet.addCell(
                "test",
                "E5"
        );
        int row = worksheet.getLastRowNumber();
        assertEquals(10, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows where cells are defined above the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getLastRowNumberTest5(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        }
        else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        worksheet.addCell(
                "test",
                "F5"
        );
        int row = worksheet.getLastRowNumber();
        assertEquals(4, row);
    }

    @DisplayName("Test of the getLastRowNumber function with an explicitly defined, empty cell besides other row definitions")
    @ParameterizedTest(name = "Given empty cell address {0} should lead to the appropriate last row {1}")
    @CsvSource(
            {
                    "'F7', 6",
                    "'A1', 4",}
    )
    void getLastRowNumberTest6(String emptyCellAddress, int expectedLastRow) {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenRow(3);
        worksheet.addHiddenRow(4);
        worksheet.addCell(null, emptyCellAddress);
        int row = worksheet.getLastRowNumber();
        assertEquals(expectedLastRow, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with an empty worksheet")
    @Test()
    void getLastDataRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getLastDataRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with defined rows on an empty worksheet")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getLastDataRowNumberTest2(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        }
        else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        int row = worksheet.getLastDataRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with defined rows where cells are defined below the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getLastDataRowNumberTest3(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        }
        else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        worksheet.addCell(
                "test",
                "E5"
        );
        int row = worksheet.getLastDataRowNumber();
        assertEquals(4, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with defined rows where cells are defined above the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getLastDataRowNumberTest4(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        }
        else {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(3, 44.4f);
        }
        worksheet.addCell(
                "test",
                "F5"
        );
        int row = worksheet.getLastDataRowNumber();
        assertEquals(4, row);
    }

    @DisplayName("Test of the getFirstRowNumber function with an empty worksheet")
    @Test()
    void getFirstRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getFirstRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the getFirstRowNumber function with defined rows on an empty worksheet")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate first row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getFirstRowNumberTest2(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(3);
        }
        else {
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 33.3f);
            worksheet.setRowHeight(3, 44.4f);
        }
        int row = worksheet.getFirstRowNumber();
        assertEquals(1, row);
    }

    @DisplayName("Test of the getFirstRowNumber function with defined rows on an empty worksheet, where the row definition has gaps")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate first row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getFirstRowNumberTest3(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(10);
        }
        else {
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        int row = worksheet.getFirstRowNumber();
        assertEquals(1, row);
    }

    @DisplayName("Test of the getFirstRowNumber function with defined rows where cells are defined above the first row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate first row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getFirstRowNumberTest4(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(3);
            worksheet.addHiddenRow(10);
        }
        else {
            worksheet.setRowHeight(2, 22.2f);
            worksheet.setRowHeight(3, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        worksheet.addCell(
                "test",
                "E5"
        );
        int row = worksheet.getFirstRowNumber();
        assertEquals(2, row);
    }

    @DisplayName("Test of the getFirstRowNumber function with defined rows where cells are defined below the first row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate first row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getFirstRowNumberTest5(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(6);
            worksheet.addHiddenRow(7);
            worksheet.addHiddenRow(8);
        }
        else {
            worksheet.setRowHeight(6, 22.2f);
            worksheet.setRowHeight(7, 33.3f);
            worksheet.setRowHeight(8, 44.4f);
        }
        worksheet.addCell(
                "test",
                "F5"
        );
        int row = worksheet.getFirstRowNumber();
        assertEquals(4, row);
    }

    @DisplayName("Test of the getFirstRowNumber function with an explicitly defined, empty cell besides other row definitions")
    @ParameterizedTest(name = "Given empty cell address {0} should lead to the appropriate last row {1}")
    @CsvSource(
            {
                    "'F5', 4",
                    "'A1', 0",}
    )
    void getFirstRowNumberTest6(String emptyCellAddress, int expectedFirstRow) {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(4);
        worksheet.addCell(null, emptyCellAddress);
        int row = worksheet.getFirstRowNumber();
        assertEquals(expectedFirstRow, row);
    }

    @DisplayName("Test of the getFirstDataRowNumber function with an empty worksheet")
    @Test()
    void getFirstDataRowNumberTest() {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getFirstDataRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the getFirstDataRowNumber function with defined rows on an empty worksheet")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate first row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getFirstDataRowNumberTest2(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(3);
        }
        else {
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 33.3f);
            worksheet.setRowHeight(3, 44.4f);
        }
        int row = worksheet.getFirstDataRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the getFirstDataRowNumber function with defined rows where cells are defined below the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate first row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getFirstDataRowNumberTest3(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(3);
            worksheet.addHiddenRow(10);
        }
        else {
            worksheet.setRowHeight(2, 22.2f);
            worksheet.setRowHeight(3, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        worksheet.addCell(
                "test",
                "E5"
        );
        int row = worksheet.getFirstDataRowNumber();
        assertEquals(4, row);
    }

    @DisplayName("Test of the getFirstDataRowNumber function with defined rows where cells are defined above the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate first row")
    @CsvSource(
            {
                    "Height",
                    "Hidden",}
    )
    void getFirstDataRowNumberTest4(RowTest.RowProperty rowProperty) {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowTest.RowProperty.Hidden) {
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(3);
        }
        else {
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 33.3f);
            worksheet.setRowHeight(3, 44.4f);
        }
        worksheet.addCell(
                "test",
                "F6"
        );
        int row = worksheet.getFirstDataRowNumber();
        assertEquals(5, row);
    }

    @DisplayName("Test of the getFirstDataRowNumber and getLastDataRowNumber functions with an explicitly defined, empty cell besides other row definitions")
    @ParameterizedTest(name = "Given empty cell address {0} should lead to -1 as  first and last row")
    @CsvSource(
            {
                    "'F5'",
                    "'A1'"}
    )
    void getFirstOrLastDataRowNumberTest(String emptyCellAddress) {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenRow(3);
        worksheet.addHiddenRow(4);
        worksheet.addCell(null, emptyCellAddress);
        int minCRow = worksheet.getFirstDataRowNumber();
        int maxRow = worksheet.getLastDataRowNumber();
        assertEquals(-1, minCRow);
        assertEquals(-1, maxRow);
    }

    @DisplayName("Test of the getFirstDataRowNumber and GetLastDataRowNumber functions with exactly one defined cell")
    @Test()
    void getFirstOrLastDataColumnNumberTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenRow(2);
        worksheet.addHiddenRow(3);
        worksheet.addHiddenRow(10);
        worksheet.addCell(
                "test",
                "C5"
        );
        int minRow = worksheet.getFirstDataRowNumber();
        int maxRow = worksheet.getLastDataRowNumber();
        assertEquals(4, minRow);
        assertEquals(4, maxRow);
    }

    @DisplayName("Test of the getFirstDataColumnNumber and getLastDataColumnNumber functions with an explicitly defined, empty cell with empty string besides other column definitions")
    @ParameterizedTest(name = "Given empty cell address {0} should lead to -1 as  first and last row")
    @CsvSource(
            {
                    "'F5'",
                    "'A1'"}
    )
    public void getFirstOrLastDataColumnNumberTest3(String emptyCellAddress)
    {
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
