package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.*;

public class RowTest {

    public enum RowProperty
    {
        Hidden,
        Height
    }

    @DisplayName("Test of the addHiddenRow function with a row number")
    @Test()
    void addHiddenRowTest()
    {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getHiddenRows().size());
        worksheet.addHiddenRow(2);
        assertEquals(1, worksheet.getHiddenRows().size());
        assertTrue(worksheet.getHiddenRows().containsKey(2));
        assertTrue(worksheet.getHiddenRows().get(2));
        worksheet.addHiddenRow(2); // Should not add an additional entry
        assertEquals(1, worksheet.getHiddenRows().size());
    }

    @DisplayName("Test of the failing addHiddenRow function with an invalid column number")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "-1",
            "-100",
            "1048576",
    })
    void addHiddenRowFailTest(int value)
    {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.addHiddenRow(value));
    }

    @DisplayName("Test of the getLastRowNumber function with an empty worksheet")
    @Test()
    void getLastRowNumberTest()
    {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getLastRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows on an empty worksheet")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastRowNumberTest2(RowProperty rowProperty)
    {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden)
        {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        }
        else
        {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        int row = worksheet.getLastRowNumber();
        assertEquals(2, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows on an empty worksheet, where the row definition has gaps")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastRowNumberTest3(RowProperty rowProperty)
    {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden)
        {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        }
        else
        {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        int row = worksheet.getLastRowNumber();
        assertEquals(10, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows where cells are defined below the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastRowNumberTest4(RowProperty rowProperty)
    {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden)
        {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        }
        else
        {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        worksheet.addCell("test", "E5");
        int row = worksheet.getLastRowNumber();
        assertEquals(10, row);
    }

    @DisplayName("Test of the getLastRowNumber function with defined rows where cells are defined above the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastRowNumberTest5(RowProperty rowProperty)
    {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden)
        {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        }
        else
        {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        worksheet.addCell("test", "F5");
        int row = worksheet.getLastRowNumber();
        assertEquals(4, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with an empty worksheet")
    @Test()
    void getLastDataRowNumberTest()
    {
        Worksheet worksheet = new Worksheet();
        int row = worksheet.getLastDataRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with defined rows on an empty worksheet")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastDataRowNumberTest2(RowProperty rowProperty)
    {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden)
        {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        }
        else
        {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(2, 44.4f);
        }
        int row = worksheet.getLastDataRowNumber();
        assertEquals(-1, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with defined rows where cells are defined below the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastDataRowNumberTest3(RowProperty rowProperty)
    {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden)
        {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(10);
        }
        else
        {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(10, 44.4f);
        }
        worksheet.addCell("test", "E5");
        int row = worksheet.getLastDataRowNumber();
        assertEquals(4, row);
    }

    @DisplayName("Test of the GetLastDataRowNumber function with defined rows where cells are defined above the last row")
    @ParameterizedTest(name = "Given property {0} should lead to the appropriate last row")
    @CsvSource({
            "Height",
            "Hidden",
    })
    void getLastDataRowNumberTest4(RowProperty rowProperty)
    {
        Worksheet worksheet = new Worksheet();
        if (rowProperty == RowProperty.Hidden)
        {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        }
        else
        {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 33.3f);
            worksheet.setRowHeight(3, 44.4f);
        }
        worksheet.addCell("test", "F5");
        int row = worksheet.getLastDataRowNumber();
        assertEquals(4, row);
    }

    @DisplayName("Test of the getCurrentRowNumber function")
    @Test()
    void getCurrentRowNumberTest()
    {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCurrentRowNumber());
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
        worksheet.addNextCell("test");
        worksheet.addNextCell("test");
        assertEquals(2, worksheet.getCurrentRowNumber());
        worksheet.setCurrentCellDirection(Worksheet.CellDirection.ColumnToColumn);
        worksheet.addNextCell("test");
        worksheet.addNextCell("test");
        assertEquals(2, worksheet.getCurrentRowNumber()); // should not change
        worksheet.goToNextRow();
        assertEquals(3, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow(2);
        assertEquals(5, worksheet.getCurrentRowNumber());
        worksheet.goToNextColumn(2);
        assertEquals(0, worksheet.getCurrentRowNumber()); // should reset
    }

    @DisplayName("Test of the goToNextRow function")
    @Test()
    void goToNextColumnTest()
    {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow();
        assertEquals(1, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow(5);
        assertEquals(6, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow(-2);
        assertEquals(4, worksheet.getCurrentRowNumber());
    }

    @DisplayName("Test of the failing goToNextRow function on invalid values")
    @ParameterizedTest(name = "Given incremental value {1} an basis {0} should lead to an exception")
    @CsvSource({
            "0, -1",
            "10, -12",
            "0, 1048576",
            "0, 1248575",
    })
    void goToNextRowTest2(int initialValue, int value)
    {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentRowNumber(initialValue);
        assertEquals(initialValue, worksheet.getCurrentRowNumber());
        assertThrows(RangeException.class, () -> worksheet.goToNextRow(value));
    }

    @DisplayName("Test of the removeRowHeight function")
    @Test()
    void removeRowHeightTest()
    {
        Worksheet worksheet = new Worksheet();
        worksheet.setRowHeight(2, 22.2f);
        worksheet.setRowHeight(4, 33.3f);
        assertEquals(2, worksheet.getRowHeights().size());
        worksheet.removeRowHeight(2);
        assertEquals(1, worksheet.getRowHeights().size());
        worksheet.removeRowHeight(3); // Should not cause anything
        worksheet.removeRowHeight(-1); // Should not cause anything
        assertEquals(1, worksheet.getRowHeights().size());
    }

    @DisplayName("Test of the setCurrentRowNumber function")
    @ParameterizedTest(name = "Given row number {0} should lead to the same position on the worksheet")
    @CsvSource({
            "0",
            "3",
            "1048575",
    })
    void setCurrentRowNumberTest(int row)
    {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow();
        worksheet.setCurrentRowNumber(row);
        assertEquals(row, worksheet.getCurrentRowNumber());
    }

    @DisplayName("Test of the failing setCurrentRowNumber function")
    @ParameterizedTest(name = "Given column number {0} should lead to an exception")
    @CsvSource({
            "-1",
            "-10",
            "1048576",
    })
    void setCurrentRowNumberFailTest(int row)
    {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.setCurrentRowNumber(row));
    }

    @DisplayName("Test of the setRowHeight function")
    @ParameterizedTest(name = "Given height {0} should lead to do a valid rowHeight definition")
    @CsvSource({
            "0f",
            "0.1f",
            "10f",
            "255f",
    })
    void setRowHeightTest(float height)
    {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getRowHeights().size());
        worksheet.setRowHeight(0, height);
        assertEquals(1, worksheet.getRowHeights().size());
        assertEquals(height, worksheet.getRowHeights().get(0));
        worksheet.setRowHeight(0, Worksheet.DEFAULT_ROW_HEIGHT);
        assertEquals(1, worksheet.getRowHeights().size()); // No removal so far
        assertEquals(Worksheet.DEFAULT_ROW_HEIGHT, worksheet.getRowHeights().get(0));
    }

    @DisplayName("Test of the failing setRowHeight function")
    @ParameterizedTest(name = "Given row number {0} or width {1} should lead to an exception")
    @CsvSource({
            "-1, 0f",
            "1048576, 0.0f",
            "0, -10f",
            "0, 409.51f",
            "0, 500f",
    })
    void setRowHeightFailTest(int rowNumber, float height)
    {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.setRowHeight(rowNumber, height));
    }


}
