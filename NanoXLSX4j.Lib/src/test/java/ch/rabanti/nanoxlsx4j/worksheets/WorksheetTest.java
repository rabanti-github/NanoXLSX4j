package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Column;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Objects;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class WorksheetTest {

    public enum RangeRepresentation {
        StringExpression,
        RangeObject,
        Addresses
    }

    @DisplayName("Test of the default constructor")
    @Test()
    void constructorTest() {
        Worksheet worksheet = new Worksheet();
        assertConstructorBasics(worksheet);
        assertNull(worksheet.getWorkbookReference());
        assertEquals(0, worksheet.getSheetID());
    }

    @DisplayName("Test of the constructor with the worksheet name")
    @ParameterizedTest(name = "Given name {0} should lead to a valid worksheet")
    @CsvSource(
            {
                    "'.'",
                    "' '",
                    "'Test'",
                    "'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'",}
    )
    void constructorTest2(String name) {
        Worksheet worksheet = new Worksheet(name);
        assertConstructorBasics(worksheet);
        assertNull(worksheet.getWorkbookReference());
        assertEquals(name, worksheet.getSheetName());
    }

    @DisplayName("Test of the constructor with all parameters")
    @ParameterizedTest(name = "Given sheet name {0} and id {1} should lead to a valid worksheet")
    @CsvSource(
            {
                    "., 1",
                    "' ', 2",
                    "Test, 10",
                    "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx, 255",}
    )
    void constructorTest2(String name, int id) {
        Workbook workbook = new Workbook(
                "test.xlsx",
                "sheet2"
        );
        Worksheet worksheet = new Worksheet(name, id, workbook);
        assertConstructorBasics(worksheet);
        assertNotNull(worksheet.getWorkbookReference());
        assertEquals("test.xlsx", worksheet.getWorkbookReference().getFilename());
        assertEquals(id, worksheet.getSheetID());
    }

    @DisplayName("Test of the failing constructor if provided with invalid worksheet names")
    @ParameterizedTest(name = "Given worksheet name {1} should lead to an exception")
    @CsvSource(
            {
                    "STRING, ''",
                    "NULL, ''",
                    "STRING, '['",
                    "STRING, '................................'",}
    )
    void constructorFailingTest(String sourceType, String sourceValue) {
        String name = (String) TestUtils.createInstance(sourceType, sourceValue);
        assertThrows(FormatException.class, () -> new Worksheet(name));
    }

    @DisplayName("Test of the failing constructor if provided with invalid values")
    @ParameterizedTest(name = "Given worksheet name {1} or id {2} should lead to an exception")
    @CsvSource(
            {
                    "STRING, '', 1",
                    "NULL, '', 1",
                    "STRING, '[', 1",
                    "STRING, '................................', 0",
                    "STRING, 'Test', 0",
                    "STRING, 'Test', -1",}
    )
    void constructorFailingTest2(String sourceType, String sourceValue, int id) {
        String name = (String) TestUtils.createInstance(sourceType, sourceValue);
        Workbook workbook = new Workbook(
                "test.xlsx",
                "sheet2"
        );
        assertThrows(FormatException.class, () -> new Worksheet(name, id, workbook));
    }

    @DisplayName("Test of the autoFilterRange field getter")
    @Test()
    void autoFilterRangeTest() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getAutoFilterRange());
        worksheet.setAutoFilter("B2:D4");
        Range ExpectedRange = new Range("B1:D1"); // Function reduces range to row 1
        assertEquals(ExpectedRange, worksheet.getAutoFilterRange());
        worksheet.removeAutoFilter();
        assertNull(worksheet.getAutoFilterRange());
    }

    @DisplayName("Test of the get function of the cells field")
    @Test()
    void cellsTest() {
        Worksheet worksheet = new Worksheet();
        assertNotNull(worksheet.getCells());
        assertEquals(0, worksheet.getCells().size());
        worksheet.addCell(
                "test",
                "C3"
        );
        worksheet.addCell(22, "D4");
        assertEquals(2, worksheet.getCells().size());
        TestUtils.assertMapEntry(
                "C3",
                "test",
                worksheet.getCells(),
                Cell::getValue
        );
        TestUtils.assertMapEntry("D4", 22, worksheet.getCells(), Cell::getValue);
        worksheet.removeCell("C3");
        assertEquals(1, worksheet.getCells().size());
        TestUtils.assertMapEntry("D4", 22, worksheet.getCells(), Cell::getValue);
    }

    @DisplayName("Test of the get function of the columns field")
    @Test()
    void columnsTest() {
        Worksheet worksheet = new Worksheet();
        assertNotNull(worksheet.getColumns());
        assertEquals(0, worksheet.getColumns().size());
        worksheet.setColumnWidth("B", 11f);
        worksheet.setColumnWidth("C", 0.7f);
        assertEquals(2, worksheet.getColumns().size());
        TestUtils.assertMapEntry(1, 11f, worksheet.getColumns(), Column::getWidth);
        TestUtils.assertMapEntry(2, 0.7f, worksheet.getColumns(), Column::getWidth);
        worksheet.resetColumn(1);
        assertEquals(1, worksheet.getColumns().size());
        TestUtils.assertMapEntry(2, 0.7f, worksheet.getColumns(), Column::getWidth);
    }

    @DisplayName("Test of the currentCellDirection field")
    @ParameterizedTest(name = "Given direction {0} with initial column {1} and row {2} should lead to the next column {3} and row {4}")
    @CsvSource(
            {
                    "ColumnToColumn, 2, 7, 3, 7",
                    "RowToRow, 2, 7, 2, 8",
                    "Disabled, 2, 7, 2, 7",}
    )
    void currentCellDirectionTest(Worksheet.CellDirection direction, int givenInitialColumn, int givenInitialRow, int expectedColumn, int expectedRow) {
        Worksheet worksheet = new Worksheet();
        worksheet.setCurrentCellDirection(direction);
        worksheet.setCurrentCellAddress(givenInitialColumn, givenInitialRow);
        assertEquals(givenInitialRow, worksheet.getCurrentRowNumber());
        assertEquals(givenInitialColumn, worksheet.getCurrentColumnNumber());
        worksheet.addNextCell("test");
        assertEquals(expectedRow, worksheet.getCurrentRowNumber());
        assertEquals(expectedColumn, worksheet.getCurrentColumnNumber());
    }

    @DisplayName("Test of the defaultColumnWidth filed")
    @ParameterizedTest(name = "Given value {0} should lead to the same result")
    @CsvSource(
            {
                    "1f",
                    "15.5f",
                    "0f",
                    "255f",}
    )
    void defaultColumnWidthTest(float value) {
        Worksheet worksheet = new Worksheet();
        assertEquals(Worksheet.DEFAULT_COLUMN_WIDTH, worksheet.getDefaultColumnWidth());
        worksheet.setDefaultColumnWidth(value);
        assertEquals(value, worksheet.getDefaultColumnWidth());
    }

    @DisplayName("Test of the failing defaultColumnWidth field, using the setter")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource(
            {
                    "-1f",
                    "255.1f",}
    )
    void defaultColumnWidthTest2(float value) {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.setDefaultColumnWidth(value));
    }

    @DisplayName("Test of the defaultRowHeight filed")
    @ParameterizedTest(name = "Given value {0} should lead to the same result")
    @CsvSource(
            {
                    "1f",
                    "15.5f",
                    "0f",
                    "409.5",}
    )
    void defaultRowHeightTest(float value) {
        Worksheet worksheet = new Worksheet();
        assertEquals(Worksheet.DEFAULT_ROW_HEIGHT, worksheet.getDefaultRowHeight());
        worksheet.setDefaultRowHeight(value);
        assertEquals(value, worksheet.getDefaultRowHeight());
    }

    @DisplayName("Test of the failing defaultRowHeight field, using the setter")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource(
            {
                    "-1f",
                    "410f",}
    )
    void defaultRowHeightTest2(float value) {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.setDefaultRowHeight(value));
    }

    @DisplayName("Test of the get function of the hiddenRows field")
    @Test()
    void hiddenRowsTest() {
        Worksheet worksheet = new Worksheet();
        assertNotNull(worksheet.getHiddenRows());
        assertEquals(0, worksheet.getHiddenRows().size());
        worksheet.addHiddenRow(2);
        worksheet.addHiddenRow(5);
        assertEquals(2, worksheet.getHiddenRows().size());
        TestUtils.assertMapEntry(2, true, worksheet.getHiddenRows());
        TestUtils.assertMapEntry(5, true, worksheet.getHiddenRows());
        worksheet.removeHiddenRow(2);
        assertEquals(1, worksheet.getHiddenRows().size());
        TestUtils.assertMapEntry(5, true, worksheet.getHiddenRows());
    }

    @DisplayName("Test of the get function of the mergedCells filed")
    @Test()
    void mergedCellsTest() {
        Worksheet worksheet = new Worksheet();
        assertNotNull(worksheet.getMergedCells());
        assertEquals(0, worksheet.getMergedCells().size());
        Range range1 = new Range("A2:C3");
        Range range2 = new Range("S3:R2");
        worksheet.mergeCells(range1);
        worksheet.mergeCells(range2);
        assertEquals(2, worksheet.getMergedCells().size());
        TestUtils.assertMapEntry("A2:C3", range1, worksheet.getMergedCells());
        TestUtils.assertMapEntry("R2:S3", range2, worksheet.getMergedCells());
        worksheet.removeMergedCells(range1.toString());
        assertEquals(1, worksheet.getMergedCells().size());
        TestUtils.assertMapEntry("R2:S3", range2, worksheet.getMergedCells());
    }

    @DisplayName("Test of the get function of the rowHeights field")
    @Test()
    void rowHeightsTest() {
        Worksheet worksheet = new Worksheet();
        assertNotNull(worksheet.getRowHeights());
        assertEquals(0, worksheet.getRowHeights().size());
        worksheet.setRowHeight(2, 15.3f);
        worksheet.setRowHeight(5, 100f);
        assertEquals(2, worksheet.getRowHeights().size());
        TestUtils.assertMapEntry(2, 15.3f, worksheet.getRowHeights());
        TestUtils.assertMapEntry(5, 100f, worksheet.getRowHeights());
        worksheet.removeRowHeight(2);
        assertEquals(1, worksheet.getRowHeights().size());
        TestUtils.assertMapEntry(5, 100f, worksheet.getRowHeights());
    }

    @DisplayName("Test of the selectedCells field getter")
    @Test()
    void selectedCellsTest() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getSelectedCells());
        worksheet.setSelectedCells("B2:D4");
        Range ExpectedRange = new Range("B2:D4");
        assertEquals(ExpectedRange, worksheet.getSelectedCells());
        worksheet.removeSelectedCells();
        assertNull(worksheet.getSelectedCells());
    }

    @DisplayName("Test of the sheetID field, as well as failing if invalid")
    @Test()
    void sheetIDTest() {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getSheetID());
        worksheet.setSheetID(12);
        assertEquals(12, worksheet.getSheetID());
        assertThrows(FormatException.class, () -> worksheet.setSheetID(0));
        assertThrows(FormatException.class, () -> worksheet.setSheetID(-1));
    }

    @DisplayName("Test of the  sheetName filed")
    @ParameterizedTest(name = "Given value {0} should lead to the name {1}")
    @CsvSource(
            {
                    ".",
                    "' '",
                    "Test",
                    "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",}
    )
    void nameTest(String name) {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getSheetName());
        worksheet.setSheetName(name);
        assertEquals(name, worksheet.getSheetName());
    }

    @DisplayName("Test failing of the set function of the sheetName filed if a worksheet name is invalid")
    @ParameterizedTest(name = "Given value {1} should lead to an exception")
    @CsvSource(
            {
                    "NULL, ",
                    "STRING, ''",
                    "STRING, xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
                    "STRING, A[B",
                    "STRING, A]B",
                    "STRING, A*B",
                    "STRING, A?B",
                    "STRING, A/B",
                    "STRING, A\\B",}
    )
    void nameFailTest(String sourceType, String sourceValue) {
        String name = (String) TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        assertThrows(Exception.class, () -> worksheet.setSheetName(name));
    }

    @DisplayName("Test of the sheetProtectionValues field")
    @Test()
    void sheetProtectionValuesTest() {
        Worksheet worksheet = new Worksheet();
        assertNotNull(worksheet.getSheetProtectionValues());
        assertEquals(0, worksheet.getSheetProtectionValues().size());
        worksheet.addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.deleteRows);
        worksheet.addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.formatRows);
        assertEquals(2, worksheet.getSheetProtectionValues().size());
        TestUtils.assertListEntry(Worksheet.SheetProtectionValue.deleteRows, worksheet.getSheetProtectionValues());
        TestUtils.assertListEntry(Worksheet.SheetProtectionValue.formatRows, worksheet.getSheetProtectionValues());
        worksheet.removeAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.deleteRows);
        assertEquals(1, worksheet.getSheetProtectionValues().size());
        TestUtils.assertListEntry(Worksheet.SheetProtectionValue.formatRows, worksheet.getSheetProtectionValues());
    }

    @DisplayName("Test of the useSheetProtection field")
    @Test()
    void useSheetProtectionTest() {
        Worksheet worksheet = new Worksheet();
        assertFalse(worksheet.isUseSheetProtection());
        worksheet.setUseSheetProtection(true);
        assertTrue(worksheet.isUseSheetProtection());
    }

    @DisplayName("Test of the workbookReference field")
    @Test()
    void workbookReferenceTest() {
        Workbook workbook = new Workbook(
                "test.xlsx",
                "test"
        );
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getWorkbookReference());
        worksheet.setWorkbookReference(workbook);
        assertNotNull(worksheet.getWorkbookReference());
        assertEquals("test.xlsx", worksheet.getWorkbookReference().getFilename());
    }

    @DisplayName("Test of the hidden field")
    @Test()
    void hiddenTest() {
        Worksheet worksheet = new Worksheet();
        assertFalse(worksheet.isHidden());
        worksheet.setHidden(true);
        assertTrue(worksheet.isHidden());
    }

    @DisplayName("Test of the failing set function of the hidden field when trying to hide all worksheets")
    @Test()
    void hiddenFailTest() {
        Workbook workbook = new Workbook("test1");
        workbook.addWorksheet("test2");
        workbook.getWorksheets().get(1).setHidden(true);
        assertFalse(workbook.getWorksheets().get(0).isHidden());
        assertTrue(workbook.getWorksheets().get(1).isHidden());
        assertThrows(WorksheetException.class, () -> workbook.getWorksheets().get(0).setHidden(true));
    }

    @DisplayName("Test of the failing set function of the hidden field when trying to hide all worksheets (scenario with 3 worksheets)")
    @Test()
    void hiddenFailTest2() {
        Workbook workbook = new Workbook("test1");
        workbook.addWorksheet("test2");
        workbook.addWorksheet("test3");
        workbook.setSelectedWorksheet(1);
        workbook.getWorksheets().get(0).setHidden(true);
        workbook.getWorksheets().get(2).setHidden(true);
        assertTrue(workbook.getWorksheets().get(0).isHidden());
        assertFalse(workbook.getWorksheets().get(1).isHidden());
        assertTrue(workbook.getWorksheets().get(2).isHidden());
        assertThrows(WorksheetException.class, () -> workbook.getWorksheets().get(1).setHidden(true));
    }

    @DisplayName("Test of the failing set function of the hidden field when trying to hide all worksheets by adding hidden worksheets to a workbook")
    @Test()
    void hiddenFailTest3() {
        Worksheet hidden = new Worksheet("test1");
        hidden.setHidden(true);
        Workbook workbook = new Workbook();
        assertEquals(0, workbook.getWorksheets().size());
        assertThrows(WorksheetException.class, () -> workbook.addWorksheet(hidden));
    }

    @DisplayName("Test of the get function of the activeStyle field")
    @Test()
    void activeStyleTest() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getActiveStyle());
        worksheet.setActiveStyle(BasicStyles.DottedFill_0_125());
        assertNotNull(worksheet.getActiveStyle());
        assertEquals(BasicStyles.DottedFill_0_125(), worksheet.getActiveStyle());
        worksheet.clearActiveStyle();
        assertNull(worksheet.getActiveStyle());
    }

    @DisplayName("Test of the removeCell function with column and row")
    @Test()
    void removeCellTest() {
        Worksheet worksheet = new Worksheet();
        String[] array = new String[]{
                "test1",
                "test2",
                "test3"};
        worksheet.addCellRange(Arrays.asList(array), "A1:A3");
        assertEquals(3, worksheet.getCells().size());
        boolean result = worksheet.removeCell(0, 1);
        assertTrue(result);
        assertEquals(2, worksheet.getCells().size());
        assertFalse(worksheet.getCells().containsKey("A2"));
        result = worksheet.removeCell(0, 1); // re-test
        assertFalse(result);
        assertEquals(2, worksheet.getCells().size());
    }

    @DisplayName("Test of the removeCell function with address")
    @Test()
    void removeCellTest2() {
        Worksheet worksheet = new Worksheet();
        String[] array = new String[]{
                "test1",
                "test2",
                "test3"};
        worksheet.addCellRange(Arrays.asList(array), "A1:A3");
        assertEquals(3, worksheet.getCells().size());
        boolean result = worksheet.removeCell("A3");
        assertTrue(result);
        assertEquals(2, worksheet.getCells().size());
        assertFalse(worksheet.getCells().containsKey("A3"));
        result = worksheet.removeCell("A3"); // re-test
        assertFalse(result);
        assertEquals(2, worksheet.getCells().size());
    }

    @DisplayName("Test of the removeCell function when no cells are defined")
    @Test()
    void removeCellTest3() {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCells().size());
        boolean result = worksheet.removeCell(2, 5);
        assertFalse(result);
        result = worksheet.removeCell("A3");
        assertFalse(result);
    }

    @DisplayName("Test of the addAllowedActionOnSheetProtection function")
    @ParameterizedTest(name = "Given value {0} should lead to {1} entries (with an possible additional value of {2})")
    @CsvSource(
            {
                    "deleteRows, 1, ",
                    "formatRows, 1, ",
                    "selectLockedCells, 2, selectUnlockedCells",
                    "selectUnlockedCells, 1,",
                    "autoFilter, 1,",
                    "sort, 1,",
                    "insertRows, 1, ",
                    "deleteColumns, 1, ",
                    "formatCells, 1, ",
                    "formatColumns, 1, ",
                    "insertHyperlinks, 1, ",
                    "insertColumns, 1, ",
                    "objects, 1, ",
                    "pivotTables, 1, ",
                    "scenarios, 1, ",}
    )
    void addAllowedActionOnSheetProtectionTest(Worksheet.SheetProtectionValue typeOfProtection, int expectedSize, Worksheet.SheetProtectionValue additionalExpectedValue) {
        Worksheet worksheet = new Worksheet();
        assertFalse(worksheet.isUseSheetProtection());
        assertEquals(0, worksheet.getSheetProtectionValues().size());
        worksheet.addAllowedActionOnSheetProtection(typeOfProtection);
        TestUtils.assertListEntry(typeOfProtection, worksheet.getSheetProtectionValues());
        if (additionalExpectedValue != null) {
            TestUtils.assertListEntry(additionalExpectedValue, worksheet.getSheetProtectionValues());
        }
        assertEquals(expectedSize, worksheet.getSheetProtectionValues().size());
        worksheet.addAllowedActionOnSheetProtection(typeOfProtection); // Should not lead to an additional value
        assertEquals(expectedSize, worksheet.getSheetProtectionValues().size());
        Worksheet.SheetProtectionValue additionalValue;
        if (typeOfProtection == Worksheet.SheetProtectionValue.objects) {
            additionalValue = Worksheet.SheetProtectionValue.sort;
        }
        else {
            additionalValue = Worksheet.SheetProtectionValue.objects;
        }
        worksheet.addAllowedActionOnSheetProtection(additionalValue);
        TestUtils.assertListEntry(additionalValue, worksheet.getSheetProtectionValues());
        assertEquals(expectedSize + 1, worksheet.getSheetProtectionValues().size());
        assertTrue(worksheet.isUseSheetProtection());
    }

    @DisplayName("Test of the addAllowedActionOnSheetProtection is nothing adding if null is passed")
    @Test()
    void addAllowedActionOnSheetProtectionTest2() {
        Worksheet worksheet = new Worksheet();
        assertFalse(worksheet.isUseSheetProtection());
        assertEquals(0, worksheet.getSheetProtectionValues().size());
        worksheet.addAllowedActionOnSheetProtection(null);
        assertFalse(worksheet.isUseSheetProtection());
        assertEquals(0, worksheet.getSheetProtectionValues().size());
    }

    @DisplayName("Test of the getCell function with an Address object")
    @ParameterizedTest(name = "Given cells {0} should return a cell {3}")
    @CsvSource(
            {
                    "'C2', STRING, test, C2",
                    "'C1,C2,C3', INTEGER, 22, C2",
                    "'A1,B1,C1,D1', BOOLEAN, true, C1",}
    )
    void getCellTest(String definedCells, String sourceType, String sourceValue, String expectedAddress) {
        Object definedSample = TestUtils.createInstance(sourceType, sourceValue);
        List<String> addresses = TestUtils.splitValuesAsList(definedCells);
        Worksheet worksheet = new Worksheet();
        for (String address : addresses) {
            worksheet.addCell(definedSample, address);
        }
        Cell cell = worksheet.getCell(new Address(expectedAddress));
        assertNotNull(cell);
        assertEquals(definedSample, cell.getValue());
        assertEquals(expectedAddress, cell.getCellAddress());
    }

    @DisplayName("Test of the getCell function with a column and row")
    @ParameterizedTest(name = "Given cells {0} should return a cell with column {3} and row {4}")
    @CsvSource(
            {
                    "'C2', STRING, test, 2,1",
                    "'C1,C2,C3', INTEGER, 22, 2,1",
                    "'A1,B1,C1,D1', BOOLEAN, true, 2,0",}
    )
    void getCellTest2(String definedCells, String sourceType, String sourceValue, int expectedColumn, int expectedRow) {
        Object definedSample = TestUtils.createInstance(sourceType, sourceValue);
        List<String> addresses = TestUtils.splitValuesAsList(definedCells);
        Worksheet worksheet = new Worksheet();
        for (String address : addresses) {
            worksheet.addCell(definedSample, address);
        }
        Cell cell = worksheet.getCell(expectedColumn, expectedRow);
        assertNotNull(cell);
        assertEquals(definedSample, cell.getValue());
        assertEquals(new Address(expectedColumn, expectedRow), cell.getCellAddress2());
    }

    @DisplayName("Test of the failing getCell function with an Address object")
    @ParameterizedTest(name = "Given cells {0} should lead to a WorksheetException applied to {3}")
    @CsvSource(
            {
                    "'', NULL, '', C2",
                    "'C1,C2,C3', INTEGER, 22, D2",}
    )
    void getCellFailTest(String definedCells, String sourceType, String sourceValue, String expectedAddress) {
        Object definedSample = TestUtils.createInstance(sourceType, sourceValue);
        List<String> addresses = TestUtils.splitValuesAsList(definedCells);
        Worksheet worksheet = new Worksheet();
        for (String address : addresses) {
            worksheet.addCell(definedSample, address);
        }
        assertThrows(WorksheetException.class, () -> worksheet.getCell(new Address(expectedAddress)));
    }

    @DisplayName("Test of the failing getCell function with a column and row")
    @ParameterizedTest(name = "Given cells {0} should lead to a WorksheetException applied to column {3} and row {4}")
    @CsvSource(
            {
                    "'', NULL, '', 2,1, WorksheetException",
                    "'C1,C2,C3', INTEGER, 22, 3,1, WorksheetException",
                    "'C1,C2,C3', INTEGER, 22, -1,2, RangeException",
                    "'C1,C2,C3', INTEGER, 22, 2,-2, RangeException",
                    "'C1,C2,C3', INTEGER, 22, 16384,2, RangeException",
                    "'C1,C2,C3', INTEGER, 22, 2,1048576, RangeException",}
    )
    void getCellFailTest2(String definedCells, String sourceType, String sourceValue, int expectedColumn, int expectedRow, String exceptionName) {
        Object definedSample = TestUtils.createInstance(sourceType, sourceValue);
        List<String> addresses = TestUtils.splitValuesAsList(definedCells);
        Worksheet worksheet = new Worksheet();
        for (String address : addresses) {
            worksheet.addCell(definedSample, address);
        }
        Exception exception = assertThrows(Exception.class, () -> worksheet.getCell(expectedColumn, expectedRow));
        assertEquals(exceptionName, exception.getClass().getSimpleName());
    }

    @DisplayName("Test of the failing getCell function with null as address object")
    @Test
    void getCellFailTest3() {
        Worksheet worksheet = new Worksheet();
        Address address = null;
        Exception exception = assertThrows(Exception.class, () -> worksheet.getCell(address));
        assertEquals("WorksheetException", exception.getClass().getSimpleName());
    }

    @DisplayName("Test of the hasCell function with an Address object")
    @ParameterizedTest(name = "Given addresses {0} should lead to {2} with address {1}")
    @CsvSource(
            {
                    "C2, C2, true",
                    "C2, C3, false",
                    ", C2, false",
                    "C2,C3,C4, C2, true",
                    "C2,C3,C4, D2, false",}
    )
    void hasCellTest(String definedCells, String givenAddress, boolean expectedResult) {
        List<String> addresses = TestUtils.splitValuesAsList(definedCells);
        Worksheet worksheet = new Worksheet();
        for (String address : addresses) {
            worksheet.addCell("test", address);
        }
        assertEquals(expectedResult, worksheet.hasCell(new Address(givenAddress)));
    }

    @DisplayName("Test of the hasCell function with a column and row")
    @ParameterizedTest(name = "Given addresses {0} should lead to {3} on column {1} and row {2}")
    @CsvSource(
            {
                    "'C2', 2,1, true",
                    "'C2', 2,2, false",
                    "'', 2,1, false",
                    "'C2,C3,C4', 2,1, true",
                    "'C2,C3,C4', 3,1, false",}
    )
    void hasCellTest2(String definedCells, int givenColumn, int givenRow, boolean expectedResult) {
        List<String> addresses = TestUtils.splitValuesAsList(definedCells);
        Worksheet worksheet = new Worksheet();
        for (String address : addresses) {
            worksheet.addCell("test", address);
        }
        assertEquals(expectedResult, worksheet.hasCell(givenColumn, givenRow));
    }

    @DisplayName("Test of the failing hasCell function with a column and row")
    @ParameterizedTest(name = "Given column {0} and row {1} should lead to an exception")
    @CsvSource(
            {
                    "-1, 2",
                    "2, -1",
                    "16384, 2",
                    "2, 1048576",}
    )
    void hasCellFailTest(int givenColumn, int givenRow) {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(
                "test",
                "C3"
        );
        assertThrows(RangeException.class, () -> worksheet.getCell(givenColumn, givenRow));
    }

    @DisplayName("Test of the getLastCellAddress function with an empty worksheet")
    @ParameterizedTest(name = "Column definitions: {0}, and row definitions of hidden states: {1} and heights: {2} should lead to a null address")
    @CsvSource(
            {
                    "false, false, false",
                    "false, false, true",
                    "false, true, true",
                    "false, true, false",
                    "true, false, false"}
    )
    void getLastCellAddressTest(boolean hasColumns, boolean hasHiddenRows, boolean hasRowHeights) {
        Worksheet worksheet = new Worksheet();
        if (hasColumns) {
            worksheet.addHiddenColumn(0);
            worksheet.addHiddenColumn(1);
            worksheet.addHiddenColumn(2);
        }
        if (hasHiddenRows) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        }
        if (hasRowHeights) {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 22.2f);
        }
        Address address = worksheet.getLastCellAddress();
        assertNull(address);
    }

    @DisplayName("Test of the getLastCellAddress function with an empty worksheet but defined columns and rows")
    @Test()
    void getLastCellAddressTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenRow(0);
        worksheet.addHiddenRow(1);
        Address address = worksheet.getLastCellAddress();
        assertNotNull(address);
        assertEquals("C2", address.getAddress());
    }

    @DisplayName("Test of the getLastCellAddress function with an empty worksheet but defined columns and rows with gaps")
    @Test()
    void getLastCellAddressTest3() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(10);
        worksheet.addHiddenRow(0);
        worksheet.addHiddenRow(1);
        worksheet.setRowHeight(10, 22.2f);
        Address address = worksheet.getLastCellAddress();
        assertNotNull(address);
        assertEquals("K11", address.getAddress());
    }

    @DisplayName("Test of the getLastCellAddress function with defined columns and rows where cells are defined below the last column and row")
    @Test()
    void getLastCellAddressTest4() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(10);
        worksheet.addHiddenRow(0);
        worksheet.addHiddenRow(1);
        worksheet.setRowHeight(10, 22.2f);
        worksheet.addCell(
                "test",
                "E5"
        );
        Address address = worksheet.getLastCellAddress();
        assertNotNull(address);
        assertEquals("K11", address.getAddress());
    }

    @DisplayName("Test of the getLastCellAddress function with defined columns and rows where cells are defined above the last column and row")
    @Test()
    void getLastCellAddressTest5() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(10);
        worksheet.addHiddenRow(0);
        worksheet.addHiddenRow(1);
        worksheet.setRowHeight(10, 22.2f);
        worksheet.addCell(
                "test",
                "L12"
        );
        Address address = worksheet.getLastCellAddress();
        assertNotNull(address);
        assertEquals("L12", address.getAddress());
    }

    @DisplayName("Test of the getLastDataCellAddress function with an empty worksheet")
    @ParameterizedTest(name = "Column definitions: {0} and row definitions: {1} should lead to a null address")
    @CsvSource(
            {
                    "false, false, false",
                    "false, false, true",
                    "false, true, true",
                    "false, true, false",
                    "true, false, false",
                    "true, false, true",
                    "true, true, false",
                    "true, true, true",}
    )
    void getLastDataCellAddressTest(boolean hasColumns, boolean hasHiddenRows, boolean hasRowHeights) {
        Worksheet worksheet = new Worksheet();
        if (hasColumns) {
            worksheet.addHiddenColumn(0);
            worksheet.addHiddenColumn(1);
            worksheet.addHiddenColumn(2);
        }
        if (hasHiddenRows) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        }
        if (hasRowHeights) {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 22.2f);
        }
        Address address = worksheet.getLastDataCellAddress();
        assertNull(address);
    }

    @DisplayName("Test of the getLastCellAddress function with defined columns and rows where cells are defined below the last column and row")
    @Test()
    void getLastDataCellAddressTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(10);
        worksheet.addHiddenRow(0);
        worksheet.addHiddenRow(1);
        worksheet.setRowHeight(10, 22.2f);
        worksheet.addCell(
                "test",
                "E7"
        );
        Address address = worksheet.getLastDataCellAddress();
        assertNotNull(address);
        assertEquals("E7", address.getAddress());
    }

    @DisplayName("Test of the getLastDataCellAddress function with defined columns and rows where cells are defined above the last column and row")
    @Test()
    void getLastDataCellAddressTest3() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(10);
        worksheet.addHiddenRow(0);
        worksheet.addHiddenRow(1);
        worksheet.setRowHeight(10, 22.2f);
        worksheet.addCell(
                "test",
                "L12"
        );
        Address address = worksheet.getLastDataCellAddress();
        assertNotNull(address);
        assertEquals("L12", address.getAddress());
    }

    @DisplayName("Test of the getFirstCellAddress function with an empty worksheet")
    @ParameterizedTest(name = "Column definitions: {0}, and row definitions of hidden states: {1} and heights: {2} should lead to a null address")
    @CsvSource(
            {
                    "false, false, false",
                    "false, false, true",
                    "false, true, true",
                    "false, true, false",
                    "true, false, false"}
    )
    void getFirstCellAddressTest(boolean hasColumns, boolean hasHiddenRows, boolean hasRowHeights) {
        Worksheet worksheet = new Worksheet();
        if (hasColumns) {
            worksheet.addHiddenColumn(1);
            worksheet.addHiddenColumn(2);
            worksheet.addHiddenColumn(3);
        }
        if (hasHiddenRows) {
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
            worksheet.addHiddenRow(3);
        }
        if (hasRowHeights) {
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 22.2f);
            worksheet.setRowHeight(3, 22.2f);
        }
        Address address = worksheet.getFirstCellAddress();
        assertNull(address);
    }

    @DisplayName("Test of the getFirstCellAddress function with an empty worksheet but defined columns and rows")
    @Test()
    void getFirstCellAddressTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(4);
        worksheet.addHiddenRow(1);
        worksheet.addHiddenRow(2);
        Address address = worksheet.getFirstCellAddress();
        assertNotNull(address);
        assertEquals("C2", address.getAddress());
    }

    @DisplayName("Test of the getFirstCellAddress function with an empty worksheet but defined columns and rows with gaps")
    @Test()
    void getFirstCellAddressTest3() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(10);
        worksheet.addHiddenRow(3);
        worksheet.addHiddenRow(4);
        worksheet.setRowHeight(10, 22.2f);
        Address address = worksheet.getFirstCellAddress();
        assertNotNull(address);
        assertEquals("C4", address.getAddress());
    }

    @DisplayName("Test of the getFirstCellAddress function with defined columns and rows where cells are defined above the first column and row")
    @Test()
    void getFirstCellAddressTest4() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(4);
        worksheet.addHiddenColumn(10);
        worksheet.addHiddenRow(4);
        worksheet.addHiddenRow(5);
        worksheet.setRowHeight(10, 22.2f);
        worksheet.addCell(
                "test",
                "R5"
        );
        worksheet.addCell(
                "test",
                "F11"
        );
        Address address = worksheet.getFirstCellAddress();
        assertNotNull(address);
        assertEquals("D5", address.getAddress());
    }

    @DisplayName("Test of the getFirstCellAddress function with defined columns and rows where cells are defined below the first column and row")
    @Test()
    void getFirstCellAddressTest5() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(4);
        worksheet.addHiddenColumn(10);
        worksheet.addHiddenRow(3);
        worksheet.addHiddenRow(4);
        worksheet.setRowHeight(100, 22.2f);
        worksheet.addCell(
                "test",
                "E5"
        );
        Address address = worksheet.getFirstCellAddress();
        assertNotNull(address);
        assertEquals("C4", address.getAddress());
    }

    @DisplayName("Test of the getFirstDataCellAddress function with an empty worksheet")
    @ParameterizedTest(name = "Column definitions: {0} and row definitions: {1} should lead to a null address")
    @CsvSource(
            {
                    "false, false, false",
                    "false, false, true",
                    "false, true, true",
                    "false, true, false",
                    "true, false, false",
                    "true, false, true",
                    "true, true, false",
                    "true, true, true",}
    )
    void getFirstDataCellAddressTest(boolean hasColumns, boolean hasHiddenRows, boolean hasRowHeights) {
        Worksheet worksheet = new Worksheet();
        if (hasColumns) {
            worksheet.addHiddenColumn(0);
            worksheet.addHiddenColumn(1);
            worksheet.addHiddenColumn(2);
        }
        if (hasHiddenRows) {
            worksheet.addHiddenRow(0);
            worksheet.addHiddenRow(1);
            worksheet.addHiddenRow(2);
        }
        if (hasRowHeights) {
            worksheet.setRowHeight(0, 22.2f);
            worksheet.setRowHeight(1, 22.2f);
            worksheet.setRowHeight(2, 22.2f);
        }
        Address address = worksheet.getFirstDataCellAddress();
        assertNull(address);
    }

    @DisplayName("Test of the getFirstDataCellAddress function with defined columns and rows where cells are defined above the first column and row")
    @Test()
    void getFirstDataCellAddressTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(4);
        worksheet.addHiddenRow(2);
        worksheet.addHiddenRow(3);
        worksheet.setRowHeight(4, 22.2f);
        worksheet.addCell(
                "test",
                "E5"
        );
        worksheet.addCell(
                "test",
                "H9"
        );
        Address address = worksheet.getFirstDataCellAddress();
        assertNotNull(address);
        assertEquals("E5", address.getAddress());
    }

    @DisplayName("Test of the getFirstDataCellAddress function with defined columns and rows where cells are defined below the first column and row")
    @Test()
    void getFirstDataCellAddressTest3() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(10);
        worksheet.addHiddenRow(1);
        worksheet.addHiddenRow(2);
        worksheet.setRowHeight(10, 22.2f);
        worksheet.addCell(
                "test",
                "C5"
        );
        worksheet.addCell(
                "test",
                "D7"
        );
        Address address = worksheet.getFirstDataCellAddress();
        assertNotNull(address);
        assertEquals("C5", address.getAddress());
    }

    @DisplayName("Test of the MergeCells function")
    @ParameterizedTest(name = "Given representation {0} an column {1}, row {2} to column {3}, and row {4} should lead to a range {5}")
    @CsvSource(
            {
                    "Addresses, 0,0,0,0, A1:A1, 1",
                    "RangeObject, 1, 1, 1, 1, B2:B2, 1",
                    "StringExpression, 2, 2, 2, 2, C3:C3, 1",
                    "Addresses, 0, 0, 2, 2, A1:C3, 9",
                    "RangeObject, 1, 1, 3, 1, B2:D2, 3",
                    "StringExpression, 2, 2, 2, 4, C3:C5, 3",
                    "Addresses, 2, 2, 0, 0, C3:A1, 9",
                    "StringExpression, 2, 4, 2, 2, C3:C5, 3",
                    // String expression is reordered by the test method
            }
    )
    void mergeCellsTest(RangeRepresentation representation, int givenStartColumn, int givenStartRow, int givenEndColumn, int givenEndRow, String expectedMergedCells, int expectedCount) {
        Worksheet worksheet = new Worksheet();
        Address startAddress = new Address(givenStartColumn, givenStartRow);
        Address endAddress = new Address(givenEndColumn, givenEndRow);
        Range range = new Range(startAddress, endAddress);
        assertEquals(0, worksheet.getMergedCells().size());
        String returnedAddress;
        if (representation == RangeRepresentation.Addresses) {
            returnedAddress = worksheet.mergeCells(startAddress, endAddress);
        }
        else if (representation == RangeRepresentation.StringExpression) {
            returnedAddress = worksheet.mergeCells(range.toString());
        }
        else {
            returnedAddress = worksheet.mergeCells(range);
        }

        assertEquals(1, worksheet.getMergedCells().size());
        assertEquals(expectedMergedCells, returnedAddress);
        assertTrue(worksheet.getMergedCells().containsKey(expectedMergedCells));
        assertEquals(
                expectedCount,
                worksheet.getMergedCells().get(expectedMergedCells).resolveEnclosedAddresses().size()
        );
    }

    @DisplayName("Test of the mergeCells function with more than one range")
    @ParameterizedTest(name = "Given representation {0} an column {1}, row {2} to column {3}, and row {4} should lead to a range {5}")
    @CsvSource(
            {
                    "Addresses, 0,0,0,0, A1:A1, 1",
                    "RangeObject, 1, 1, 1, 1, B2:B2, 1",
                    "StringExpression, 2, 2, 2, 2, C3:C3, 1",
                    "Addresses, 0, 0, 2, 2, A1:C3, 9",
                    "RangeObject, 1, 1, 3, 1, B2:D2, 3",
                    "StringExpression, 2, 2, 2, 4, C3:C5, 3",
                    "Addresses, 2, 2, 0, 0, C3:A1, 9",
                    "StringExpression, 2, 4, 2, 2, C3:C5, 3",
                    // String expression is reordered by the test method
            }
    )
    void mergeCellsTest2(RangeRepresentation representation, int givenStartColumn, int givenStartRow, int givenEndColumn, int givenEndRow, String expectedMergedCells, int expectedCount) {
        Worksheet worksheet = new Worksheet();
        Address startAddress = new Address(givenStartColumn, givenStartRow);
        Address endAddress = new Address(givenEndColumn, givenEndRow);
        Range range = new Range(startAddress, endAddress);
        assertEquals(0, worksheet.getMergedCells().size());
        String returnedAddress;
        if (representation == RangeRepresentation.Addresses) {
            returnedAddress = worksheet.mergeCells(startAddress, endAddress);
        }
        else if (representation == RangeRepresentation.StringExpression) {
            returnedAddress = worksheet.mergeCells(range.toString());
        }
        else {
            returnedAddress = worksheet.mergeCells(range);
        }
        String returnedAddress2 = worksheet.mergeCells("X1:X2");
        assertEquals(2, worksheet.getMergedCells().size());
        assertEquals(expectedMergedCells, returnedAddress);
        assertTrue(worksheet.getMergedCells().containsKey(expectedMergedCells));
        assertEquals(
                expectedCount,
                worksheet.getMergedCells().get(expectedMergedCells).resolveEnclosedAddresses().size()
        );
        assertTrue(worksheet.getMergedCells().containsKey(returnedAddress2));
        assertEquals(2, worksheet.getMergedCells().get(returnedAddress2).resolveEnclosedAddresses().size());
    }

    @DisplayName("Test of the failing mergeCells function if cell addresses are colliding")
    @Test()
    void mergeCellsFailTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.mergeCells("A1:D4");
        assertThrows(RangeException.class, () -> worksheet.mergeCells("B4:E4"));
    }

    @DisplayName("Test of the failing mergeCells function if the merge range already exists (full intersection)")
    @Test()
    void mergeCellsFailTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.mergeCells("A1:D4");
        assertThrows(RangeException.class, () -> worksheet.mergeCells("D4:A1")); // Flip addresses
    }

    @DisplayName("Test of the failing MergeCells function if the merge range is invalid (string)")
    @Test()
    void mergeCellsFailTest3() {
        Worksheet worksheet = new Worksheet();
        assertThrows(FormatException.class, () -> worksheet.mergeCells(""));
        String nullValue = null;
        assertThrows(FormatException.class, () -> worksheet.mergeCells(nullValue));
    }

    @DisplayName("Test of the internal RecalculateAutoFilter function")
    @Test()
    void recalculateAutoFilterTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.recalculateAutoFilter(); // Dummy call
        assertNull(worksheet.getAutoFilterRange());
        worksheet.addCell(
                "test",
                "A100"
        );
        worksheet.addCell(
                "test",
                "D50"
        ); // Will expand the range to row 50
        worksheet.addCell(
                "test",
                "F2"
        );
        worksheet.setAutoFilter("B1:E1");
        worksheet.getColumns().get(2).setAutoFilter(false);
        worksheet.resetColumn(2);
        worksheet.recalculateAutoFilter();
        assertTrue(worksheet.getColumns().get(2).hasAutoFilter());
        assertEquals("B1:E50", worksheet.getAutoFilterRange().toString());
    }

    @DisplayName("Test of the internal recalculateColumns function")
    @Test()
    void recalculateColumnsTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.setColumnWidth(1, 22.5f);
        worksheet.setColumnWidth(2, 22.8f);
        worksheet.addHiddenColumn(3);
        worksheet.setColumnWidth(1, Worksheet.DEFAULT_COLUMN_WIDTH); // should not remove the column
        assertEquals(3, worksheet.getColumns().size());
        worksheet.recalculateColumns();
        assertEquals(2, worksheet.getColumns().size());
        assertFalse(worksheet.getColumns().containsKey(1));
    }

    @DisplayName("Test of the internal resolveMergedCells function")
    @Test()
    void resolveMergedCellsTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(
                "test",
                "B1"
        );
        worksheet.addCell(22.2f, "C1");
        assertEquals(2, worksheet.getCells().size());
        worksheet.mergeCells("B1:D1");
        worksheet.resolveMergedCells();
        assertEquals(3, worksheet.getCells().size());
        assertNull(worksheet.getCells().get("B1").getCellStyle());
        assertEquals(Cell.CellType.EMPTY, worksheet.getCells().get("C1").getDataType());
        assertEquals(BasicStyles.MergeCellStyle(), worksheet.getCells().get("C1").getCellStyle());
        assertEquals(22.2f, worksheet.getCells().get("C1").getValue());
        assertEquals(BasicStyles.MergeCellStyle(), worksheet.getCells().get("D1").getCellStyle());
        assertEquals(Cell.CellType.EMPTY, worksheet.getCells().get("D1").getDataType());
    }

    @DisplayName("Test of the internal recalculateColumns function")
    @Test()
    void removeAutoFilterTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.setAutoFilter(1, 5);
        assertNotNull(worksheet.getAutoFilterRange());
        assertEquals("B1:F1", worksheet.getAutoFilterRange().toString());
        worksheet.removeAutoFilter();
        assertNull(worksheet.getAutoFilterRange());
    }

    @DisplayName("Test of the removeHiddenColumn function")
    @Test()
    void removeHiddenColumnTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(3);
        worksheet.setColumnWidth(2, 22.2f);
        assertEquals(3, worksheet.getColumns().size());
        worksheet.removeHiddenColumn(2);
        worksheet.removeHiddenColumn(3);
        assertEquals(2, worksheet.getColumns().size());
        assertFalse(worksheet.getColumns().get(2).isHidden());
    }

    @DisplayName("Test of the removeHiddenColumn function with a string as column expression")
    @Test()
    void removeHiddenColumnTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn("B");
        worksheet.addHiddenColumn("C");
        worksheet.addHiddenColumn("D");
        worksheet.setColumnWidth(2, 22.2f);
        assertEquals(3, worksheet.getColumns().size());
        worksheet.removeHiddenColumn("C");
        worksheet.removeHiddenColumn("D");
        assertEquals(2, worksheet.getColumns().size());
        assertFalse(worksheet.getColumns().get(2).isHidden());
    }

    @DisplayName("Test of the removeHiddenRow function")
    @Test()
    void removeHiddenRowTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenRow(1);
        worksheet.addHiddenRow(2);
        worksheet.addHiddenRow(3);
        assertEquals(3, worksheet.getHiddenRows().size());
        worksheet.removeHiddenRow(2);
        assertEquals(2, worksheet.getHiddenRows().size());
        assertFalse(worksheet.getHiddenRows().containsKey(2));
    }

    @DisplayName("Test of the removeMergedCells function")
    @Test()
    void removeMergedCellsTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(
                "test",
                "B2"
        );
        worksheet.addCell(22, "B3");
        worksheet.mergeCells("B1:B4");
        assertTrue(worksheet.getMergedCells().containsKey("B1:B4"));
        worksheet.removeMergedCells("B1:B4");
        assertEquals(0, worksheet.getMergedCells().size());
    }

    @DisplayName("Test of the removeMergedCells function after resolution of merged cells (on save)")
    @Test()
    void removeMergedCellsTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(
                "test",
                "B2"
        );
        worksheet.addCell(22, "B3");
        worksheet.mergeCells("B1:B4");
        assertTrue(worksheet.getMergedCells().containsKey("B1:B4"));
        worksheet.resolveMergedCells();
        worksheet.removeMergedCells("B1:B4");
        assertEquals(0, worksheet.getMergedCells().size());
        assertNotEquals(BasicStyles.MergeCellStyle(), worksheet.getCells().get("B2").getCellStyle());
        assertNotEquals(BasicStyles.MergeCellStyle(), worksheet.getCells().get("B3").getCellStyle());
        assertEquals("test", worksheet.getCells().get("B2").getValue());
        assertEquals(22, worksheet.getCells().get("B3").getValue());
    }

    @DisplayName("Test of the failing removeMergedCells function on an invalid range")
    @ParameterizedTest(name = "Given range {1} of type {0} should lead to an exception")
    @CsvSource(
            {
                    "Null, ''",
                    "STRING, ''",
                    "STRING, 'B1'",
                    "STRING, 'B1:B5'",}
    )
    void removeMergedCellsFailTest(String sourceType, String sourceValue) {
        String range = (String) TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(
                "test",
                "B2"
        );
        worksheet.addCell(22, "B3");
        worksheet.mergeCells("B1:B4");
        assertTrue(worksheet.getMergedCells().containsKey("B1:B4"));
        assertThrows(RangeException.class, () -> worksheet.removeMergedCells(range));
    }

    @DisplayName("Test of the removeSelectedCells function")
    @Test()
    void removeSelectedCellsTest() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getSelectedCells());
        worksheet.setSelectedCells("B2:D3");
        assertEquals("B2:D3", worksheet.getSelectedCells().toString());
        worksheet.removeSelectedCells();
        assertNull(worksheet.getSelectedCells());
    }

    @DisplayName("Test of the RemoveAllowedActionOnSheetProtection function")
    @ParameterizedTest(name = "Given value {0} should be not present after removal")
    @CsvSource(
            {
                    "deleteRows, objects, sort",
                    "formatRows, objects, sort",
                    "selectLockedCells, objects, sort",
                    "selectUnlockedCells, objects, sort",
                    "autoFilter, objects, sort",
                    "sort, objects, formatRows",
                    "insertRows, objects, sort",
                    "deleteColumns, objects, sort",
                    "formatCells, objects, sort",
                    "formatColumns, objects, sort",
                    "insertHyperlinks, objects, sort",
                    "insertColumns, objects, sort",
                    "objects, formatColumns, sort",
                    "pivotTables, objects, sort",
                    "scenarios, objects, sort",}
    )
    void removeAllowedActionOnSheetProtectionTest(Worksheet.SheetProtectionValue typeOfProtection, Worksheet.SheetProtectionValue additionalValue, Worksheet.SheetProtectionValue notPresentValue) {
        Worksheet worksheet = new Worksheet();
        worksheet.addAllowedActionOnSheetProtection(typeOfProtection);
        worksheet.addAllowedActionOnSheetProtection(additionalValue);
        int count = worksheet.getSheetProtectionValues().size();
        assertTrue(count >= 2);
        worksheet.removeAllowedActionOnSheetProtection(typeOfProtection);
        assertEquals(count - 1, worksheet.getSheetProtectionValues().size());
        assertFalse(worksheet.getSheetProtectionValues().contains(typeOfProtection));
        worksheet.removeAllowedActionOnSheetProtection(notPresentValue); // should not cause anything
        assertEquals(count - 1, worksheet.getSheetProtectionValues().size());
    }

    @DisplayName("Test of the setActiveStyle function")
    @Test()
    void setActiveStyleTest() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getActiveStyle());
        worksheet.setActiveStyle(BasicStyles.Bold());
        assertEquals(BasicStyles.Bold(), worksheet.getActiveStyle());
    }

    @DisplayName("Test of the setActiveStyle function on null")
    @Test()
    void setActiveStyleTest2() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getActiveStyle());
        worksheet.setActiveStyle(null);
        assertNull(worksheet.getActiveStyle());
    }

    @DisplayName("Test of the setCurrentCellAddress function with column and row numbers")
    @ParameterizedTest(name = "Given column {0} and row {1} should lead to the same position as currentCellAddress")
    @CsvSource(
            {
                    "0, 0",
                    "5, 0",
                    "0, 5",
                    "16383, 1048575",}
    )
    void setCurrentCellAddressTest(int column, int row) {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCurrentColumnNumber());
        assertEquals(0, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow();
        worksheet.goToNextColumn();
        worksheet.setCurrentCellAddress(column, row);
        assertEquals(column, worksheet.getCurrentColumnNumber());
        assertEquals(row, worksheet.getCurrentRowNumber());
    }

    @DisplayName("Test of the setCurrentCellAddress function")
    @ParameterizedTest(name = "Given address {0} should lead to the same position as currentCellAddress")
    @CsvSource(
            {
                    "A1",
                    "$A$1",
                    "C$5",
                    "$XFD1",
                    "A$1048575",
                    "XFD1048575",}
    )
    void setCurrentCellAddressTest2(String address) {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCurrentColumnNumber());
        assertEquals(0, worksheet.getCurrentRowNumber());
        worksheet.goToNextRow();
        worksheet.goToNextColumn();
        worksheet.setCurrentCellAddress(address);
        Address addr = new Address(address);
        assertEquals(addr.Column, worksheet.getCurrentColumnNumber());
        assertEquals(addr.Row, worksheet.getCurrentRowNumber());
    }

    @DisplayName("Test of the failing setCurrentCellAddress function on invalid columns or rows")
    @ParameterizedTest(name = "Given column {0} or row {1} should lead to an exception")
    @CsvSource(
            {
                    "-1, 0",
                    "0, -1",
                    "-10, -10",
                    "16384, 1048575",
                    "16383, 1048576",}
    )
    void setCurrentCellAddressFailTest(int column, int row) {
        Worksheet worksheet = new Worksheet();
        assertThrows(RangeException.class, () -> worksheet.setCurrentCellAddress(column, row));
    }

    @DisplayName("Test of the failing setCurrentCellAddress function on an invalid address as string")
    @ParameterizedTest(name = "Given address {1} (type {0}) should lead to an exception")
    @CsvSource(
            {
                    "NULL, ''",
                    "STRING, ''",
                    "STRING, ':'",
                    "STRING, 'XFE1'",
                    "STRING, 'A1:A1'",
                    "STRING, 'A0'",
                    "STRING, 'A1048577'",}
    )
    void setCurrentCellAddressFailTest2(String sourceType, String sourceValue) {
        String address = (String) TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        assertThrows(Exception.class, () -> worksheet.setCurrentCellAddress(address));
    }

    @DisplayName("Test of the setSelectedCells function with range objects --> OBSOLETE")
    @ParameterizedTest(name = "Given expression {0} should lead to the same range as selected cells")
    @CsvSource(
            {
                    "A1:A1",
                    "B2:C10",
                    "C10:B5",
                    "A1:A10",
                    "A1:R1",
                    "$A$1:$R$1",
                    "A1:XFD1048575",}
    )
    void setSelectedCellsTest(String addressExpression) {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getSelectedCells());
        Range range = new Range(addressExpression);
        worksheet.setSelectedCells(range);
        assertEquals(range.toString(), worksheet.getSelectedCells().toString());
    }

    @DisplayName("Test of the AddSelectedCells function with range objects")
    @ParameterizedTest(name = "Given range expression {0} should lead to the same range as selected cells with a range count of {1}")
    @CsvSource(
            {
                    "'A1:A1', 1",
                    "'B2:C10', 1",
                    "'A1:A10', 1",
                    "'A1:R1', 1",
                    "'$A$1:$R$1', 1",
                    "'A1:XFD1048575', 1",
                    "'A1:A1,B2:D2', 2",
                    "'B2:C10,D1:F6,F8:F10', 3",
                    "'A1:A10,A12', 2",
                    "'A1:R1,S1:S1', 2",
                    "'$A$1:$R$1,$S$1:$S$4,X10', 3",
            }
    )
    void addSelectedCellsTest(String rangeExpressions, int expectedRangeCount) {
        String[] rangeStrings = rangeExpressions.split(",");
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getSelectedCellRanges().size());
        for (String range : rangeStrings) {
            Range r = getRangeFromExpression(range);
            worksheet.addSelectedCells(r);
        }
        assertSelectedCellRanges(expectedRangeCount, rangeStrings, worksheet);
    }

    @DisplayName("Test of the setSelectedCells function with strings --> OBSOLETE")
    @ParameterizedTest(name = "Given string {0} should lead to the same range as selected cells")
    @CsvSource(
            {
                    "A1:A1",
                    "B2:C10",
                    "C10:B5",
                    "A1:A10",
                    "A1:R1",
                    "$A$1:$R$1",
                    "A1:XFD1048575",}
    )
    void setSelectedCellsTest2(String rangeString) {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getSelectedCells());
        worksheet.setSelectedCells(rangeString);
        Range range = new Range(rangeString);
        assertEquals(range.toString(), worksheet.getSelectedCells().toString());
    }

    @DisplayName("Test of the AddSelectedCells function with strings")
    @ParameterizedTest(name = "Given range expression {0} should lead to the same range as selected cells with a range count of {1}")
    @CsvSource(
            {
                    "'A1:A1', 1",
                    "'B2:C10', 1",
                    "'A1:A10', 1",
                    "'A1:R1', 1",
                    "'$A$1:$R$1', 1",
                    "'A1:XFD1048575', 1",
                    "'A1:A1,B2:D2', 2",
                    "'B2:C10,D1:F6,F8:F10', 3",
                    "'A1:A10,A12:A12', 2",
                    "'A1:R1,S1:S1', 2",
                    "'$A$1:$R$1,$S$1:$S$4,X10:X10', 3",
                    "'A1:A1,B2:D2,', 2",
                    // Empty string should be ignored
            }
    )
    void addSelectedCellsTest2(String rangeExpressions, int expectedRangeCount) {
        String[] rangeStrings = rangeExpressions.split(",");
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getSelectedCellRanges().size());
        for (String range : rangeStrings) {
            worksheet.addSelectedCells(range);
        }
        assertSelectedCellRanges(expectedRangeCount, rangeStrings, worksheet);
    }


    @DisplayName("Test of the setSelectedCells function with address objects --> OBSOLETE")
    @ParameterizedTest(name = "Given start address {0} and end address {1} should lead to the same range as selected cells")
    @CsvSource(
            {
                    "A1, A1",
                    "B2, C10",
                    "C10, B5",
                    "A1, A10",
                    "A1, R1",
                    "$A$1, $R$1",
                    "A1, XFD1048575",}
    )
    void setSelectedCellsTest3(String startAddress, String endAddress) {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getSelectedCells());
        Address start = new Address(startAddress);
        Address end = new Address(endAddress);
        Range range = new Range(start, end);
        worksheet.setSelectedCells(start, end);
        assertEquals(range.StartAddress.getAddress(), worksheet.getSelectedCells().StartAddress.getAddress());
        assertEquals(range.EndAddress.getAddress(), worksheet.getSelectedCells().EndAddress.getAddress());
    }

    @DisplayName("Test of the setSelectedCells functions with null values --> OBSOLETE")
    @Test
    void setSelectedCellsTest4() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getSelectedCells());
        worksheet.setSelectedCells("R5:X10");
        assertNotNull(worksheet.getSelectedCells());
        String nullString = null;
        worksheet.setSelectedCells(nullString);
        assertNull(worksheet.getSelectedCells());
        worksheet.setSelectedCells("R5:X10");
        assertNotNull(worksheet.getSelectedCells());
        Range nullRange = null;
        worksheet.setSelectedCells(nullRange);
        assertNull(worksheet.getSelectedCells());
        worksheet.setSelectedCells("R5:X10");
        assertNotNull(worksheet.getSelectedCells());
        worksheet.setSelectedCells(null, null);
        assertNull(worksheet.getSelectedCells());
        worksheet.setSelectedCells("R5:X10");
        assertNotNull(worksheet.getSelectedCells());
    }

    @DisplayName("Test of the AddSelectedCells function with address objects")
    @ParameterizedTest(name = "Given start range expression {0} and end range expressions {1} should lead to the same range as selected cells")
    @CsvSource(
            {
                    "'A1', 'A1'",
                    "'B2', 'C10'",
                    "'C10', 'B5'",
                    "'A1', 'A10'",
                    "'A1', 'R1'",
                    "'$A$1', '$R$1'",
                    "'A1', 'XFD1048575'",
                    "'A1,B1', 'A1,B1'",
                    "'B2,D5', 'C10,E10'",
                    "'C10,C20', 'B5,B20'",
                    "'A1,B2,C3', 'A10,B20,C30'",
                    "'$A$1,$X$1', '$R$1,$Y$2'",
            }
    )
    void addSelectedCellsTest3(String startAddressString, String endAddressString) {
        Worksheet worksheet = new Worksheet();
        String[] startAddresses = startAddressString.split(",");
        String[] endAddresses = endAddressString.split(",");
        assertEquals(0, worksheet.getSelectedCellRanges().size());
        String[] expectedRanges = new String[startAddresses.length];
        for (int i = 0; i < startAddresses.length; i++) {
            expectedRanges[i] = startAddresses[i] + ":" + endAddresses[i];
            Address start = new Address(startAddresses[i]);
            Address end = new Address(endAddresses[i]);
            worksheet.addSelectedCells(start, end);
        }
        assertSelectedCellRanges(startAddresses.length, expectedRanges, worksheet);
    }


    @DisplayName("Test of the failing  setSelectedCells function with one null address instead of two")
    @Test
    void setSelectedCellsTest4b() {
        Worksheet worksheet = new Worksheet();
        Address address = new Address("C5");
        assertThrows(RangeException.class, () -> worksheet.setSelectedCells(null, address));
        assertThrows(RangeException.class, () -> worksheet.setSelectedCells(address, null));
    }

    @DisplayName("Test of the setSheetProtectionPassword function")
    @ParameterizedTest(name = "Given value {1} (type {0}) should lead to the password {3} and sheet protection should be {4}")
    @CsvSource(
            {
                    "NULL, '', NULL, '', false",
                    "STRING, '', NULL, '', false",
                    "STRING, 'x', STRING, 'x', true",
                    "STRING, '***', STRING, '***', true",}
    )
    void setSheetProtectionPasswordTest(String sourceType, String sourceValue, String expectedType, String expectedValue, boolean expectedUsage) {
        String password = (String) TestUtils.createInstance(sourceType, sourceValue);
        String expectedPassword = (String) TestUtils.createInstance(expectedType, expectedValue);

        Worksheet worksheet = new Worksheet();
        assertFalse(worksheet.isUseSheetProtection());
        assertNull(worksheet.getSheetProtectionPassword());
        worksheet.setSheetProtectionPassword(password);
        assertEquals(expectedUsage, worksheet.isUseSheetProtection());
        assertEquals(expectedPassword, worksheet.getSheetProtectionPassword());
    }

    @DisplayName("Test of the setSheetName function")
    @ParameterizedTest(name = "Given value {1} (type {0} should be valid = {2} and to a worksheet name {4} ")
    @CsvSource(
            {
                    "'STRING', '1', true, 'STRING', '1'",
                    "'STRING', 'test', true, 'STRING', 'test'",
                    "'STRING', 'test-test', true, 'STRING', 'test-test'",
                    "'STRING', '$$$', true, 'STRING', '$$$'",
                    "'STRING', 'a b', true, 'STRING', 'a b'",
                    "'STRING', 'a\tb', true, 'STRING', 'a\tb'",
                    "'STRING', '-------------------------------', true, STRING, '-------------------------------'",
                    "'STRING', '', false, 'NULL', ''",
                    "'NULL', '', false, 'NULL', ''",
                    "'STRING', 'a[b', false, 'NULL', ''",
                    "'STRING', 'a]b', false, 'NULL', ''",
                    "'STRING', 'a*b', false, 'NULL', ''",
                    "'STRING', 'a?b', false, 'NULL', ''",
                    "'STRING', 'a/b', false, 'NULL', ''",
                    "'STRING', 'a\\b', false, 'NULL', ''",
                    "'STRING', '--------------------------------', false, NULL, ''",}
    )
    void setSheetNameTest(String sourceType, String sourceValue, boolean expectedValid, String expectedType, String expectedValue) {
        String name = (String) TestUtils.createInstance(sourceType, sourceValue);
        String expectedName = (String) TestUtils.createInstance(expectedType, expectedValue);

        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getSheetName());
        if (expectedValid) {
            worksheet.setSheetName(name);
            assertEquals(expectedName, worksheet.getSheetName());
        }
        else {
            assertThrows(FormatException.class, () -> worksheet.setSheetName(name));
        }
    }

    @DisplayName("Test of the setSheetName function")
    @ParameterizedTest(name = "Given value {1} (type {0} should be valid = {2} and to a worksheet name {4} ")
    @CsvSource(
            {
                    "false, 'STRING', 'test', true, 'STRING', 'test'",
                    "false, 'NULL', '', false, 'NULL', ''",
                    "false, 'STRING', 'a[b', false, 'NULL', ''",
                    "false, 'STRING', 'a]b', false, 'NULL', ''",
                    "false, 'STRING', 'a*b', false, 'NULL', ''",
                    "false, 'STRING', 'a?b', false, 'NULL', ''",
                    "false, 'STRING', 'a/b', false, 'NULL', ''",
                    "false, 'STRING', 'a\\b', false, 'NULL', ''",
                    "false, 'STRING', '--------------------------------', false, NULL, ''",
                    "true, 'STRING', 'test', true, 'STRING', 'test'",
                    "true, 'NULL', '', true, 'STRING', 'Sheet2'",
                    "true, 'STRING', '', true, 'STRING', 'Sheet2'",
                    "true, 'STRING', 'a[b', true, 'STRING', 'a_b'",
                    "true, 'STRING', 'a]b', true, 'STRING', 'a_b'",
                    "true, 'STRING', 'a*b', true, 'STRING', 'a_b'",
                    "true, 'STRING', 'a?b', true, 'STRING', 'a_b'",
                    "true, 'STRING', 'a/b', true, 'STRING', 'a_b'",
                    "true, 'STRING', 'a\\b', true, 'STRING', 'a_b'",
                    "true, 'STRING', '--------------------------------', true, STRING, '-------------------------------'",}
    )
    void setSheetNameTest(boolean useSanitation, String sourceType, String sourceValue, boolean expectedValid, String expectedType, String expectedValue) {
        String name = (String) TestUtils.createInstance(sourceType, sourceValue);
        String expectedName = (String) TestUtils.createInstance(expectedType, expectedValue);

        Workbook workbook = new Workbook("Sheet1");
        workbook.addWorksheet("test");
        Worksheet worksheet = workbook.getCurrentWorksheet();
        assertEquals("test", worksheet.getSheetName());
        if (expectedValid) {
            worksheet.setSheetName(name, useSanitation);
            assertEquals(expectedName, worksheet.getSheetName());
        }
        else {
            assertThrows(FormatException.class, () -> worksheet.setSheetName(name, useSanitation));
        }
    }

    @DisplayName("Test of the failing setSheetName function with sanitizing on a missing Workbook reference")
    @Test()
    void setSheetNameFailingTest() {
        Worksheet worksheet = new Worksheet(); // Worksheet was not created over a workbook
        assertThrows(WorksheetException.class, () -> worksheet.setSheetName("test", true));
    }

    // Tests for Insert-Search-Replace

    @ParameterizedTest(name = "Test replaceCellValue: oldValue={1}, newValue={3}, differentValue={5}")
    @CsvSource({
            "STRING, 'oldValue', STRING,'newValue', STRING,'differentValue'",
            "STRING, 'oldValue', INTEGER, 23, BOOLEAN, true",
            "INTEGER, 25, INTEGER, 23, BOOLEAN, false",
            "FLOAT, 23.7, STRING,'string', STRING,'otherString'",
            "FLOAT, -0.01, INTEGER, 15, NULL, null",
            "BOOLEAN, true, NULL, null, BOOLEAN, false",
            "NULL, null, NULL, null, BOOLEAN, false",
            "INTEGER, 0, NULL, null, NULL, null",
            "NULL, null, STRING,'', INTEGER, 0",
            "NULL, null, BOOLEAN, true, INTEGER, -1",
            "NULL, null, INTEGER, 22, BOOLEAN, BOOLEAN, false",
            "NULL, null, FLOAT, 0.01, FLOAT, 0.001",
            "NULL, null, INTEGER,27, INTEGER, 0",
            "INTEGER, 127, NULL, null, INTEGER, 128",
            "INTEGER, 128, STRING, '128', INTEGER, -128"
    })
    void replaceCellValue_ShouldReplaceAllOccurrences(String oldType, String oldValueString, String newType, String newValueString, String differentType, String differentValueString) {
        Worksheet worksheet = new Worksheet();
        Object oldValue = TestUtils.createInstance(oldType, oldValueString);
        Object newValue = TestUtils.createInstance(newType, newValueString);
        Object differentValue = TestUtils.createInstance(differentType, differentValueString);

        worksheet.addCell(oldValue, 0, 0);
        worksheet.addCell(oldValue, 1, 0);
        worksheet.addCell(oldValue, 2, 0);
        worksheet.addCell(differentValue, 3, 0);

        int replacedCount = worksheet.replaceCellValue(oldValue, newValue);

        // Assertions
        assertEquals(3, replacedCount);
        assertEquals(newValue, worksheet.getCells().get("A1").getValue());
        assertEquals(newValue, worksheet.getCells().get("B1").getValue());
        assertEquals(newValue, worksheet.getCells().get("C1").getValue());
        assertEquals(differentValue, worksheet.getCells().get("D1").getValue());
    }

    @DisplayName("Test replaceCellValue on a Date value")
    @Test()
    void replaceCellValue_ShouldReplaceAllOccurrences2() {
        Worksheet worksheet = new Worksheet();
        String oldValue = "oldValue";
        worksheet.addCell(oldValue, 0, 0);
        worksheet.addCell(oldValue, 1, 0);
        worksheet.addCell(oldValue, 2, 0);
        worksheet.addCell("differentValue", 3, 0);

        Date newValue = TestUtils.buildDate(2025, 2, 10, 5, 6, 7);
        int replacedCount = worksheet.replaceCellValue(oldValue, newValue);

        // Assertions
        assertEquals(3, replacedCount);
        assertEquals(newValue, worksheet.getCells().get("A1").getValue());
        assertEquals(newValue, worksheet.getCells().get("B1").getValue());
        assertEquals(newValue, worksheet.getCells().get("C1").getValue());
        assertEquals("differentValue", worksheet.getCells().get("D1").getValue());
    }

    @ParameterizedTest(name = "Test replaceCellValue: oldType={0}, oldValue={1}, newType={2}, newValue={3}, diffType={4}, diffValue={5}")
    @CsvSource({
            "STRING, 'oldValue', STRING, 'newValue', STRING, 'differentValue'",
            "STRING, 'oldValue', INTEGER, 23, BOOLEAN, true",
            "INTEGER, 25, INTEGER, 23, BOOLEAN, false",
            "DOUBLE, 23.7, STRING, 'string', STRING, 'otherString'",
            "DOUBLE, -0.01, INTEGER, 15, NULL, null",
            "BOOLEAN, true, NULL, null, BOOLEAN, false",
            "NULL, null, NULL, null, BOOLEAN, false",
            "INTEGER, 0, NULL, null, NULL, null",
            "NULL, null, STRING, '', INTEGER, 0",
            "NULL, null, BOOLEAN, true, INTEGER, -1",
            "NULL, null, INTEGER, 22, BOOLEAN, false",
            "NULL, null, DOUBLE, 0.01, DOUBLE, 0.001",
            "NULL, null, INTEGER, 127, INTEGER, 0",
            "INTEGER, 127, NULL, null, INTEGER, 128",
            "INTEGER, 128, STRING, '128', INTEGER, -128"
    })
    void replaceCellValue_ShouldNotReplaceAnyOccurrences(
            String oldType, String oldValueString,
            String newType, String newValueString,
            String diffType, String diffValueString) {

        Object oldValue = TestUtils.createInstance(oldType, oldValueString);
        Object newValue = TestUtils.createInstance(newType, newValueString);
        Object differentValue = TestUtils.createInstance(diffType, diffValueString);

        Worksheet worksheet = new Worksheet();
        String unrelatedValue = "UnrelatedValue";
        worksheet.addCell(unrelatedValue, 0, 0);
        worksheet.addCell(unrelatedValue, 1, 0);
        worksheet.addCell(unrelatedValue, 2, 0);
        worksheet.addCell(differentValue, 3, 0);

        int replacedCount = worksheet.replaceCellValue(oldValue, newValue);

        // Assert
        assertEquals(0, replacedCount, "No replacements should be made");
        assertEquals(unrelatedValue, worksheet.getCells().get("A1").getValue());
        assertEquals(unrelatedValue, worksheet.getCells().get("B1").getValue());
        assertEquals(unrelatedValue, worksheet.getCells().get("C1").getValue());
        assertEquals(differentValue, worksheet.getCells().get("D1").getValue());
    }
    @ParameterizedTest(name = "Test firstOrDefaultCell: type={0}, matchingValue={1}")
    @CsvSource({
            "NULL, null",
            "STRING, ''",       // Empty string
            "INTEGER, 0",
            "STRING, 'null'",   // String "null"
            "INTEGER, -1",
            "FLOAT, 0.05",
            "BOOLEAN, true",
            "BOOLEAN, false",
            "FLOAT, -22.357"
    })
    void firstOrDefaultCell_ShouldReturnCorrectCell(String sourceType, String sourceValue) {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell("value1", 0, 0);
        Object matchingValue = TestUtils.createInstance(sourceType, sourceValue);
        worksheet.addCell(matchingValue, 1, 0);
        worksheet.addCell("value3", 2, 0);

        Cell result = worksheet.firstOrDefaultCell(cell -> cell.getValue().equals(matchingValue));

        // Assertions
        assertNotNull(result);
        assertEquals(matchingValue, result.getValue());
    }

    @ParameterizedTest(name = "Test firstOrDefaultCell when no existing values exists: matchingValue={1}")
    @CsvSource({
            "NULL, null",
            "STRING, ''",       // Empty string
            "INTEGER, 0",
            "STRING, 'null'",   // String "null"
            "INTEGER, -1",
            "FLOAT, 0.05",
            "BOOLEAN, true",
            "BOOLEAN, false",
            "FLOAT, -22.357"
    })
    void firstOrDefaultCell_NotFound(String sourceType, String sourceValue) {
        Worksheet worksheet = new Worksheet();
        Object nonMatchingValue = TestUtils.createInstance(sourceType, sourceValue);
        worksheet.addCell("value1", 0, 0);
        worksheet.addCell("value2", 1, 0);
        worksheet.addCell("value3", 2, 0);

        Cell result = worksheet.firstOrDefaultCell(cell -> cell.getValue().equals(nonMatchingValue));

        // Assertions
        assertNull(result);
    }

    @ParameterizedTest(name = "Test firstCellByValue: matchingValue={0}")
    @CsvSource({
            "NULL, null",
            "STRING, ''",       // Empty string
            "INTEGER, 0",
            "STRING, 'null'",   // String "null"
            "INTEGER, -1",
            "FLOAT, 0.05",
            "BOOLEAN, true",
            "BOOLEAN, false",
            "FLOAT, -22.357"
    })
    void firstOrDefaultCellByValue_ShouldReturnCorrectCell(String sourceType, String sourceValue) {
        Worksheet worksheet = new Worksheet();

        var cell1 = new Cell("Test1", Cell.CellType.STRING, "A1");
        Object matchingValue = TestUtils.createInstance(sourceType, sourceValue);
        var cell2 = new Cell(matchingValue, Cell.CellType.DEFAULT, "B1");
        worksheet.getCells().put("A1", cell1);
        worksheet.getCells().put("B1", cell2);

        Cell result = worksheet.firstCellByValue(matchingValue);

        // Assertions
        assertNotNull(result);
        assertEquals(matchingValue, result.getValue());
        assertEquals("B1", result.getCellAddress());
    }

    @DisplayName("Test firstCellByValue on a Date value")
    @Test()
    void firstOrDefaultCellByValue_ShouldReturnCorrectCell2() {
        Worksheet worksheet = new Worksheet();

        var cell1 = new Cell("Test1", Cell.CellType.STRING, "A1");
        Date matchingValue = TestUtils.buildDate(2025, 2, 10, 5, 6, 7);
        var cell2 = new Cell(matchingValue, Cell.CellType.DATE, "B1");
        worksheet.getCells().put("A1", cell1);
        worksheet.getCells().put("B1", cell2);

        Cell result = worksheet.firstCellByValue(matchingValue);

        // Assertions
        assertNotNull(result);
        assertEquals(matchingValue, result.getValue());
        assertEquals("B1", result.getCellAddress());
    }

    @ParameterizedTest(name = "Test FirstCellByValue when no existing value matches: notMatchingType={0}, notMatchingValue={1}")
    @CsvSource({
            "'NULL', null",
            "'STRING', ''",
            "'INTEGER', 0",
            "'STRING', 'null'",
            "'INTEGER', -1",
            "'FLOAT', 0.05",
            "'BOOLEAN', true",
            "'BOOLEAN', false",
            "'DOUBLE', -22.357"
    })
    void firstCellByValue_ShouldReturnNull_WhenValueNotFound(String notMatchingType, String notMatchingValueString) {

        Object notMatchingValue = TestUtils.createInstance(notMatchingType, notMatchingValueString);
        Worksheet worksheet = new Worksheet();
        Cell cell1 = new Cell("Test1", Cell.CellType.STRING, "A1");
        worksheet.getCells().put("A1", cell1);

        Cell result = worksheet.firstCellByValue(notMatchingValue);

        // Assert
        assertNull(result, "Result should be null when the value is not found");
    }

    @DisplayName("Test of the insertRow function")
    @Test
    void testInsertRow() {
        Worksheet worksheet = new Worksheet();
        Style style1 = new Style();
        style1.getFont().setBold(true);
        Style style2 = new Style();
        style2.getFont().setItalic(true);

        worksheet.addCell("A1", 0, 0);
        worksheet.addCell("A2", 0, 1, style1);
        worksheet.addCell("A3", 0, 2, style2);
        worksheet.addCell("A4", 0, 3);
        worksheet.addCell("B6", 1, 5); // Gap

        worksheet.insertRow(1, 2);

        // Assert
        assertEquals("A1", worksheet.getCells().get("A1").getValue());
        assertEquals("A2", worksheet.getCells().get("A2").getValue());
        assertNull(worksheet.getCells().get("A3").getValue(), "A3 should be null");
        assertNull(worksheet.getCells().get("A4").getValue(), "A4 should be null");
        assertEquals("A3", worksheet.getCells().get("A5").getValue());
        assertEquals("A4", worksheet.getCells().get("A6").getValue());
        assertFalse(worksheet.getCells().containsKey("B7"), "B7 should not exist (gap preserved)");
        assertEquals("B6", worksheet.getCells().get("B8").getValue());

        assertNull(worksheet.getCells().get("A1").getCellStyle(), "A1 should not have a style");
        assertTrue(worksheet.getCells().get("A2").getCellStyle().getFont().isBold(), "A2 should be bold");
        assertTrue(worksheet.getCells().get("A3").getCellStyle().getFont().isBold(), "A3 should be bold");
        assertTrue(worksheet.getCells().get("A4").getCellStyle().getFont().isBold(), "A4 should be bold");
        assertTrue(worksheet.getCells().get("A5").getCellStyle().getFont().isItalic(), "A5 should be italic");
        assertNull(worksheet.getCells().get("A6").getCellStyle(), "A6 should not have a style");
        assertNull(worksheet.getCells().get("B8").getCellStyle(), "B8 should not have a style");
    }

    @DisplayName("Test of the insertColumn function")
    @Test
    void testInsertColumn() {
        Worksheet worksheet = new Worksheet();
        Style style1 = new Style();
        style1.getFont().setBold(true);
        Style style2 = new Style();
        style2.getFont().setItalic(true);

        worksheet.addCell("A2", 0, 1);
        worksheet.addCell("B2", 1, 1, style1);
        worksheet.addCell("C2", 2, 1, style2);
        worksheet.addCell("D2", 3, 1);
        worksheet.addCell("F3", 5, 2); // Gap

        worksheet.insertColumn(1, 2);

        // Assert
        assertEquals("A2", worksheet.getCells().get("A2").getValue());
        assertEquals("B2", worksheet.getCells().get("B2").getValue());
        assertNull(worksheet.getCells().get("C2").getValue(), "C2 should be null");
        assertNull(worksheet.getCells().get("D2").getValue(), "D2 should be null");
        assertEquals("C2", worksheet.getCells().get("E2").getValue());
        assertEquals("D2", worksheet.getCells().get("F2").getValue());
        assertFalse(worksheet.getCells().containsKey("G3"), "G3 should not exist (gap preserved)");
        assertEquals("F3", worksheet.getCells().get("H3").getValue());

        assertNull(worksheet.getCells().get("A2").getCellStyle(), "A2 should not have a style");
        assertTrue(worksheet.getCells().get("B2").getCellStyle().getFont().isBold(), "B2 should be bold");
        assertTrue(worksheet.getCells().get("C2").getCellStyle().getFont().isBold(), "C2 should be bold");
        assertTrue(worksheet.getCells().get("D2").getCellStyle().getFont().isBold(), "D2 should be bold");
        assertTrue(worksheet.getCells().get("E2").getCellStyle().getFont().isItalic(), "E2 should be italic");
        assertNull(worksheet.getCells().get("F2").getCellStyle(), "F2 should not have a style");
        assertNull(worksheet.getCells().get("H3").getCellStyle(), "H3 should not have a style");
    }


    public static void assertAddedCell(Worksheet worksheet, int numberOfEntries, String expectedAddress, Cell.CellType expectedType, Style expectedStyle, Object expectedValue, int nextColumn, int nextRow) {
        assertEquals(numberOfEntries, worksheet.getCells().size());
        TestUtils.assertMapEntry(expectedAddress, expectedValue, worksheet.getCells(), Cell::getValue);
        TestUtils.assertMapEntry(expectedAddress, expectedType, worksheet.getCells(), Cell::getDataType);
        TestUtils.assertMapEntry(expectedAddress, expectedAddress, worksheet.getCells(), Cell::getCellAddress);
        if (expectedStyle == null) {
            assertNull(worksheet.getCells().get(expectedAddress).getCellStyle());
        }
        else {
            assertEquals(expectedStyle, worksheet.getCells().get(expectedAddress).getCellStyle());
        }
        assertEquals(nextColumn, worksheet.getCurrentColumnNumber());
        assertEquals(nextRow, worksheet.getCurrentRowNumber());
    }

    public static Worksheet initWorksheet(Worksheet worksheet, String address, Worksheet.CellDirection direction) {
        return initWorksheet(worksheet, address, direction, null);
    }

    public static Worksheet initWorksheet(Worksheet worksheet, String address, Worksheet.CellDirection direction, Style style) {
        if (worksheet == null) {
            worksheet = new Worksheet();
        }
        worksheet.setCurrentCellAddress(address);
        worksheet.setCurrentCellDirection(direction);
        if (style != null) {
            worksheet.setActiveStyle(style);
        }
        return worksheet;
    }

    private void assertConstructorBasics(Worksheet worksheet) {
        assertNotNull(worksheet);
        assertNotNull(worksheet.getCells());
        assertEquals(0, worksheet.getCells().size());
        assertEquals(0, worksheet.getCurrentRowNumber());
        assertEquals(0, worksheet.getCurrentColumnNumber());
        assertEquals(Worksheet.DEFAULT_COLUMN_WIDTH, worksheet.getDefaultColumnWidth());
        assertEquals(Worksheet.DEFAULT_ROW_HEIGHT, worksheet.getDefaultRowHeight());
        assertNotNull(worksheet.getRowHeights());
        assertEquals(0, worksheet.getRowHeights().size());
        assertNotNull(worksheet.getMergedCells());
        assertEquals(0, worksheet.getMergedCells().size());
        assertNotNull(worksheet.getSheetProtectionValues());
        assertEquals(0, worksheet.getSheetProtectionValues().size());
        assertNotNull(worksheet.getHiddenRows());
        assertEquals(0, worksheet.getHiddenRows().size());
        assertNotNull(worksheet.getColumns());
        assertEquals(0, worksheet.getColumns().size());
        assertNull(worksheet.getActiveStyle());
        assertNull(worksheet.getActivePane());
    }

    private static Range getRangeFromExpression(String range) {
        Range r;
        if (range.contains(":")) {
            r = new Range(range);
        }
        else {
            r = new Range(range + ":" + range);
        }
        return r;
    }

    private static void assertSelectedCellRanges(int expectedRangeCount, String[] ranges, Worksheet worksheet) {
        assertEquals(expectedRangeCount, worksheet.getSelectedCellRanges().size());
        for (String range : ranges) {
            if (range == null || range.isEmpty()) {
                continue;
            }
            Range r = getRangeFromExpression(range);
            assertTrue(worksheet.getSelectedCellRanges().stream().anyMatch(x -> x.toString().equals(r.toString())));
        }
    }
}
