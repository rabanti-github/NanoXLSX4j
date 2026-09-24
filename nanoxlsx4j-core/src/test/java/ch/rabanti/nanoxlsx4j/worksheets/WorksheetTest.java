package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Objects;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Column;
import ch.rabanti.nanoxlsx4j.CorePackageTestAccess;
import ch.rabanti.nanoxlsx4j.LegacyPassword;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;

public class WorksheetTest {

    public enum RangeRepresentation {
        STRING_EXPRESSION,
        RANGE_OBJECT,
        ADDRESSES
    }

    @DisplayName("Test of the default constructor")
    @Test()
    void constructorTest() {
        Worksheet worksheet = new Worksheet();
        assertConstructorBasics(worksheet);
        assertNull(worksheet.getWorkbookReference());
        assertEquals(0, worksheet.getSheetId());
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
    void constructorTest3(String name, int id) {
        Workbook workbook = new Workbook(
                "test.xlsx",
                "sheet2"
        );
        Worksheet worksheet = new Worksheet(name, id, workbook);
        assertConstructorBasics(worksheet);
        assertNotNull(worksheet.getWorkbookReference());
        assertEquals("test.xlsx", worksheet.getWorkbookReference().getFilename());
        assertEquals(id, worksheet.getSheetId());
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
        String name = (String) WorksheetTestUtils.createInstance(sourceType, sourceValue);
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
        String name = (String) WorksheetTestUtils.createInstance(sourceType, sourceValue);
        Workbook workbook = new Workbook(
                "test.xlsx",
                "sheet2"
        );
        assertThrows(FormatException.class, () -> new Worksheet(name, id, workbook));
    }

    @DisplayName("Test of the autoFilterRange field getter")
    @Test()
    void autoFilterRangTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getAutoFilterRange().isEmpty());
        worksheet.setAutoFilter("B2:D4");
        Range ExpectedRange = new Range("B1:D1"); // Function reduces range to row 1
        assertEquals(ExpectedRange, worksheet.getAutoFilterRange().orElseThrow());
        worksheet.removeAutoFilter();
        assertTrue(worksheet.getAutoFilterRange().isEmpty());
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
        WorksheetTestUtils.assertMapEntry(
                "C3",
                "test",
                worksheet.getCells(),
                Cell::getValue
        );
        WorksheetTestUtils.assertMapEntry("D4", 22, worksheet.getCells(), Cell::getValue);
        worksheet.removeCell("C3");
        assertEquals(1, worksheet.getCells().size());
        WorksheetTestUtils.assertMapEntry("D4", 22, worksheet.getCells(), Cell::getValue);
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
        WorksheetTestUtils.assertMapEntry(1, 11f, worksheet.getColumns(), Column::getWidth);
        WorksheetTestUtils.assertMapEntry(2, 0.7f, worksheet.getColumns(), Column::getWidth);
        worksheet.resetColumn(1);
        assertEquals(1, worksheet.getColumns().size());
        WorksheetTestUtils.assertMapEntry(2, 0.7f, worksheet.getColumns(), Column::getWidth);
    }

    @DisplayName("Test of the currentCellDirection field")
    @ParameterizedTest(
            name = "Given direction {0} with initial column {1} and row {2} should lead to the next column" +
                    " {3} and row {4}"
    )
    @CsvSource(
            {
                    "COLUMN_TO_COLUMN, 2, 7, 3, 7",
                    "ROW_TO_ROW, 2, 7, 2, 8",
                    "DISABLED, 2, 7, 2, 7",}
    )
    void currentCellDirectionTest(
            Worksheet.CellDirection direction, int givenInitialColumn, int givenInitialRow, int expectedColumn,
            int expectedRow
    ) {
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
        assertEquals(Worksheet.DEFAULT_WORKSHEET_COLUMN_WIDTH, worksheet.getDefaultColumnWidth());
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
        assertEquals(Worksheet.DEFAULT_WORKSHEET_ROW_HEIGHT, worksheet.getDefaultRowHeight());
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
        WorksheetTestUtils.assertMapEntry(2, true, worksheet.getHiddenRows());
        WorksheetTestUtils.assertMapEntry(5, true, worksheet.getHiddenRows());
        worksheet.removeHiddenRow(2);
        assertEquals(1, worksheet.getHiddenRows().size());
        WorksheetTestUtils.assertMapEntry(5, true, worksheet.getHiddenRows());
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
        WorksheetTestUtils.assertMapEntry("A2:C3", range1, worksheet.getMergedCells());
        WorksheetTestUtils.assertMapEntry("R2:S3", range2, worksheet.getMergedCells());
        worksheet.removeMergedCells(range1.toString());
        assertEquals(1, worksheet.getMergedCells().size());
        WorksheetTestUtils.assertMapEntry("R2:S3", range2, worksheet.getMergedCells());
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
        WorksheetTestUtils.assertMapEntry(2, 15.3f, worksheet.getRowHeights());
        WorksheetTestUtils.assertMapEntry(5, 100f, worksheet.getRowHeights());
        worksheet.removeRowHeight(2);
        assertEquals(1, worksheet.getRowHeights().size());
        WorksheetTestUtils.assertMapEntry(5, 100f, worksheet.getRowHeights());
    }

    @DisplayName("Test of the selectedCells field getter")
    @Test()
    void selectedCellsTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getSelectedCells().isEmpty());
        worksheet.addSelectedCells("B2:D4");
        Range expectedRange = new Range("B2:D4");
        assertTrue(worksheet.getSelectedCells().contains(expectedRange));
        worksheet.clearSelectedCells();
        assertTrue(worksheet.getSelectedCells().isEmpty());
    }

    @DisplayName("Test of the sheetId field, as well as failing if invalid")
    @Test()
    void sheetIDTest() {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getSheetId());
        worksheet.setSheetId(12);
        assertEquals(12, worksheet.getSheetId());
        assertThrows(FormatException.class, () -> worksheet.setSheetId(0));
        assertThrows(FormatException.class, () -> worksheet.setSheetId(-1));
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
        String name = (String) WorksheetTestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        assertThrows(Exception.class, () -> worksheet.setSheetName(name));
    }

    @ParameterizedTest
    @CsvSource(
            value = {
                    "NULL|NULL|false",
                    "''|NULL|false",
                    "' '|' '|true",
                    "test|test|true",
                    "123|123|true",
                    "@é#|@é#|true"
            },
            delimiter = '|',
            nullValues = "NULL"
    )
    @DisplayName("Test of the getSheetProtectionPassword function")
    void sheetProtectionPasswordTest(String givenValue, String expectedValue, boolean expectedPasswordIsSet) {
        Worksheet worksheet = new Worksheet();
        assertFalse(worksheet.getSheetProtectionPassword().passwordIsSet());
        worksheet.setSheetProtectionPassword(givenValue);
        assertEquals(expectedValue, worksheet.getSheetProtectionPassword().getPassword());
        assertEquals(expectedPasswordIsSet, worksheet.getSheetProtectionPassword().passwordIsSet());
        worksheet.setSheetProtectionPassword(null);
        assertFalse(worksheet.getSheetProtectionPassword().passwordIsSet());
    }

    @Test
    @DisplayName("Internal test of the replacement of a password instance in the getSheetProtectionPassword function")
    void sheetProtectionReplacementTest() {
        Worksheet worksheet = new Worksheet();
        assertNotNull(worksheet.getSheetProtectionPassword());
        CorePackageTestAccess.setSheetProtectionPassword(worksheet, null);
        assertNull(worksheet.getSheetProtectionPassword());
        LegacyPassword newInstance = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION);
        CorePackageTestAccess.setSheetProtectionPassword(worksheet, newInstance);
        assertNotNull(worksheet.getSheetProtectionPassword());
    }

    @DisplayName("Test of the sheetProtectionValues field")
    @Test()
    void sheetProtectionValuesTest() {
        Worksheet worksheet = new Worksheet();
        assertNotNull(worksheet.getSheetProtectionValues());
        assertEquals(0, worksheet.getSheetProtectionValues().size());
        worksheet.addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.DELETE_ROWS);
        worksheet.addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.FORMAT_ROWS);
        assertEquals(2, worksheet.getSheetProtectionValues().size());
        WorksheetTestUtils.assertListEntry(
                Worksheet.SheetProtectionValue.DELETE_ROWS, worksheet.getSheetProtectionValues());
        WorksheetTestUtils.assertListEntry(
                Worksheet.SheetProtectionValue.FORMAT_ROWS, worksheet.getSheetProtectionValues());
        worksheet.removeAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.DELETE_ROWS);
        assertEquals(1, worksheet.getSheetProtectionValues().size());
        WorksheetTestUtils.assertListEntry(
                Worksheet.SheetProtectionValue.FORMAT_ROWS, worksheet.getSheetProtectionValues());
    }

    @DisplayName("Test of the useSheetProtection field")
    @Test()
    void useSheetProtectionTest() {
        Worksheet worksheet = new Worksheet();
        assertFalse(worksheet.getSheetProtection());
        worksheet.setSheetProtection(true);
        assertTrue(worksheet.getSheetProtection());
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
        assertThrows(WorksheetException.class, () -> workbook.getWorksheets().getFirst().setHidden(true));
    }

    @DisplayName(
            "Test of the failing set function of the hidden field when trying to hide all worksheets (scenario " +
                    "with 3 worksheets)"
    )
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

    @DisplayName(
            "Test of the failing set function of the hidden field when trying to hide all worksheets by adding " +
                    "hidden worksheets to a workbook"
    )
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
        worksheet.setActiveStyle(BasicStyles.getDottedFill0125());
        assertNotNull(worksheet.getActiveStyle());
        assertEquals(BasicStyles.getDottedFill0125(), worksheet.getActiveStyle());
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
        worksheet.addCellRange(Arrays.asList((Object[]) array), "A1:A3");
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
        worksheet.addCellRange(Arrays.asList((Object[]) array), "A1:A3");
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
                    "DELETE_ROWS, 1, ",
                    "FORMAT_ROWS, 1, ",
                    "SELECT_LOCKED_CELLS, 2, SELECT_UNLOCKED_CELLS",
                    "SELECT_UNLOCKED_CELLS, 1,",
                    "AUTO_FILTER, 1,",
                    "SORT, 1,",
                    "INSERT_ROWS, 1, ",
                    "DELETE_COLUMNS, 1, ",
                    "FORMAT_CELLS, 1, ",
                    "FORMAT_COLUMNS, 1, ",
                    "INSERT_HYPERLINKS, 1, ",
                    "INSERT_COLUMNS, 1, ",
                    "OBJECTS, 1, ",
                    "PIVOT_TABLES, 1, ",
                    "SCENARIOS, 1, ",}
    )
    void addAllowedActionOnSheetProtectionTest(
            Worksheet.SheetProtectionValue typeOfProtection, int expectedSize,
            Worksheet.SheetProtectionValue additionalExpectedValue
    ) {
        Worksheet worksheet = new Worksheet();
        assertFalse(worksheet.getSheetProtection());
        assertEquals(0, worksheet.getSheetProtectionValues().size());
        worksheet.addAllowedActionOnSheetProtection(typeOfProtection);
        WorksheetTestUtils.assertListEntry(typeOfProtection, worksheet.getSheetProtectionValues());
        if (additionalExpectedValue != null) {
            WorksheetTestUtils.assertListEntry(additionalExpectedValue, worksheet.getSheetProtectionValues());
        }
        assertEquals(expectedSize, worksheet.getSheetProtectionValues().size());
        worksheet.addAllowedActionOnSheetProtection(typeOfProtection); // Should not lead to an additional value
        assertEquals(expectedSize, worksheet.getSheetProtectionValues().size());
        Worksheet.SheetProtectionValue additionalValue;
        if (typeOfProtection == Worksheet.SheetProtectionValue.OBJECTS) {
            additionalValue = Worksheet.SheetProtectionValue.SORT;
        } else {
            additionalValue = Worksheet.SheetProtectionValue.OBJECTS;
        }
        worksheet.addAllowedActionOnSheetProtection(additionalValue);
        WorksheetTestUtils.assertListEntry(additionalValue, worksheet.getSheetProtectionValues());
        assertEquals(expectedSize + 1, worksheet.getSheetProtectionValues().size());
        assertTrue(worksheet.getSheetProtection());
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
        Object definedSample = WorksheetTestUtils.createInstance(sourceType, sourceValue);
        List<String> addresses = WorksheetTestUtils.splitValuesAsList(definedCells);
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
        Object definedSample = WorksheetTestUtils.createInstance(sourceType, sourceValue);
        List<String> addresses = WorksheetTestUtils.splitValuesAsList(definedCells);
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
        Object definedSample = WorksheetTestUtils.createInstance(sourceType, sourceValue);
        List<String> addresses = WorksheetTestUtils.splitValuesAsList(definedCells);
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
    void getCellFailTest2(
            String definedCells, String sourceType, String sourceValue, int expectedColumn, int expectedRow,
            String exceptionName
    ) {
        Object definedSample = WorksheetTestUtils.createInstance(sourceType, sourceValue);
        List<String> addresses = WorksheetTestUtils.splitValuesAsList(definedCells);
        Worksheet worksheet = new Worksheet();
        for (String address : addresses) {
            worksheet.addCell(definedSample, address);
        }
        Exception exception = assertThrows(Exception.class, () -> worksheet.getCell(expectedColumn, expectedRow));
        assertEquals(exceptionName, exception.getClass().getSimpleName());
    }

    @DisplayName("Test of the hasCell function with an Address object")
    @ParameterizedTest(name = "Given addresses {0} should lead to {2} with address {1}")
    @CsvSource(
            {
                    "C2, C2, true",
                    "C2, C3, false",
                    ", C2, false",
                    "'C2,C3,C4', C2, true",
                    "'C2,C3,C4', D2, false",}
    )
    void hasCellTest(String definedCells, String givenAddress, boolean expectedResult) {
        List<String> addresses = WorksheetTestUtils.splitValuesAsList(definedCells);
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
        List<String> addresses = WorksheetTestUtils.splitValuesAsList(definedCells);
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
    @ParameterizedTest(
            name = "Column definitions: {0}, and row definitions of hidden states: {1} and heights: {2} " +
                    "should lead to a null address"
    )
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
        Address address = worksheet.getLastCellAddress().orElse(null);
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
        Address address = worksheet.getLastCellAddress().orElse(null);
        assertNotNull(address);
        assertEquals("C2", address.getAddress());
    }

    @DisplayName(
            "Test of the getLastCellAddress function with an empty worksheet but defined columns and rows with " +
                    "gaps"
    )
    @Test()
    void getLastCellAddressTest3() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(0);
        worksheet.addHiddenColumn(1);
        worksheet.addHiddenColumn(10);
        worksheet.addHiddenRow(0);
        worksheet.addHiddenRow(1);
        worksheet.setRowHeight(10, 22.2f);
        Address address = worksheet.getLastCellAddress().orElse(null);
        assertNotNull(address);
        assertEquals("K11", address.getAddress());
    }

    @DisplayName(
            "Test of the getLastCellAddress function with defined columns and rows where cells are defined below" +
                    " the last column and row"
    )
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
        Address address = worksheet.getLastCellAddress().orElse(null);
        assertNotNull(address);
        assertEquals("K11", address.getAddress());
    }

    @DisplayName(
            "Test of the getLastCellAddress function with defined columns and rows where cells are defined above" +
                    " the last column and row"
    )
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
        Address address = worksheet.getLastCellAddress().orElse(null);
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
        Address address = worksheet.getLastDataCellAddress().orElse(null);
        assertNull(address);
    }

    @DisplayName(
            "Test of the getLastCellAddress function with defined columns and rows where cells are defined below" +
                    " the last column and row"
    )
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
        Address address = worksheet.getLastDataCellAddress().orElse(null);
        assertNotNull(address);
        assertEquals("E7", address.getAddress());
    }

    @DisplayName(
            "Test of the getLastDataCellAddress function with defined columns and rows where cells are defined " +
                    "above the last column and row"
    )
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
        Address address = worksheet.getLastDataCellAddress().orElse(null);
        assertNotNull(address);
        assertEquals("L12", address.getAddress());
    }

    @DisplayName("Test of the getFirstCellAddress function with an empty worksheet")
    @ParameterizedTest(
            name = "Column definitions: {0}, and row definitions of hidden states: {1} and heights: {2} " +
                    "should lead to a null address"
    )
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
        Address address = worksheet.getFirstCellAddress().orElse(null);
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
        Address address = worksheet.getFirstCellAddress().orElse(null);
        assertNotNull(address);
        assertEquals("C2", address.getAddress());
    }

    @DisplayName(
            "Test of the getFirstCellAddress function with an empty worksheet but defined columns and rows with " +
                    "gaps"
    )
    @Test()
    void getFirstCellAddressTest3() {
        Worksheet worksheet = new Worksheet();
        worksheet.addHiddenColumn(2);
        worksheet.addHiddenColumn(3);
        worksheet.addHiddenColumn(10);
        worksheet.addHiddenRow(3);
        worksheet.addHiddenRow(4);
        worksheet.setRowHeight(10, 22.2f);
        Address address = worksheet.getFirstCellAddress().orElse(null);
        assertNotNull(address);
        assertEquals("C4", address.getAddress());
    }

    @DisplayName(
            "Test of the getFirstCellAddress function with defined columns and rows where cells are defined " +
                    "above the first column and row"
    )
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
        Address address = worksheet.getFirstCellAddress().orElse(null);
        assertNotNull(address);
        assertEquals("D5", address.getAddress());
    }

    @DisplayName(
            "Test of the getFirstCellAddress function with defined columns and rows where cells are defined " +
                    "below the first column and row"
    )
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
        Address address = worksheet.getFirstCellAddress().orElse(null);
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
        Address address = worksheet.getFirstDataCellAddress().orElse(null);
        assertNull(address);
    }

    @DisplayName(
            "Test of the getFirstDataCellAddress function with defined columns and rows where cells are defined " +
                    "above the first column and row"
    )
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
        Address address = worksheet.getFirstDataCellAddress().orElse(null);
        assertNotNull(address);
        assertEquals("E5", address.getAddress());
    }

    @DisplayName(
            "Test of the getFirstDataCellAddress function with defined columns and rows where cells are defined " +
                    "below the first column and row"
    )
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
        Address address = worksheet.getFirstDataCellAddress().orElse(null);
        assertNotNull(address);
        assertEquals("C5", address.getAddress());
    }

    @DisplayName("Test of the mergeCells function")
    @ParameterizedTest(
            name = "Given representation {0} an column {1}, row {2} to column {3}, and row {4} should lead" +
                    " to a range {5}"
    )
    @CsvSource(
            {
                    "ADDRESSES, 0,0,0,0, A1:A1, 1",
                    "RANGE_OBJECT, 1, 1, 1, 1, B2:B2, 1",
                    "STRING_EXPRESSION, 2, 2, 2, 2, C3:C3, 1",
                    "ADDRESSES, 0, 0, 2, 2, A1:C3, 9",
                    "RANGE_OBJECT, 1, 1, 3, 1, B2:D2, 3",
                    "STRING_EXPRESSION, 2, 2, 2, 4, C3:C5, 3",
                    "ADDRESSES, 2, 2, 0, 0, C3:A1, 9",
                    "STRING_EXPRESSION, 2, 4, 2, 2, C3:C5, 3",
                    // String expression is reordered by the test method
            }
    )
    void mergeCellsTest(
            RangeRepresentation representation, int givenStartColumn, int givenStartRow, int givenEndColumn,
            int givenEndRow, String expectedMergedCells, int expectedCount
    ) {
        Worksheet worksheet = new Worksheet();
        Address startAddress = new Address(givenStartColumn, givenStartRow);
        Address endAddress = new Address(givenEndColumn, givenEndRow);
        Range range = new Range(startAddress, endAddress);
        assertEquals(0, worksheet.getMergedCells().size());
        String returnedAddress;
        if (representation == RangeRepresentation.ADDRESSES) {
            returnedAddress = worksheet.mergeCells(startAddress, endAddress);
        } else if (representation == RangeRepresentation.STRING_EXPRESSION) {
            returnedAddress = worksheet.MergeCells(range.toString());
        } else {
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
    @ParameterizedTest(
            name = "Given representation {0} an column {1}, row {2} to column {3}, and row {4} should lead" +
                    " to a range {5}"
    )
    @CsvSource(
            {
                    "ADDRESSES, 0,0,0,0, A1:A1, 1",
                    "RANGE_OBJECT, 1, 1, 1, 1, B2:B2, 1",
                    "STRING_EXPRESSION, 2, 2, 2, 2, C3:C3, 1",
                    "ADDRESSES, 0, 0, 2, 2, A1:C3, 9",
                    "RANGE_OBJECT, 1, 1, 3, 1, B2:D2, 3",
                    "STRING_EXPRESSION, 2, 2, 2, 4, C3:C5, 3",
                    "ADDRESSES, 2, 2, 0, 0, C3:A1, 9",
                    "STRING_EXPRESSION, 2, 4, 2, 2, C3:C5, 3",
                    // String expression is reordered by the test method
            }
    )
    void mergeCellsTest2(
            RangeRepresentation representation, int givenStartColumn, int givenStartRow, int givenEndColumn,
            int givenEndRow, String expectedMergedCells, int expectedCount
    ) {
        Worksheet worksheet = new Worksheet();
        Address startAddress = new Address(givenStartColumn, givenStartRow);
        Address endAddress = new Address(givenEndColumn, givenEndRow);
        Range range = new Range(startAddress, endAddress);
        assertEquals(0, worksheet.getMergedCells().size());
        String returnedAddress;
        if (representation == RangeRepresentation.ADDRESSES) {
            returnedAddress = worksheet.mergeCells(startAddress, endAddress);
        } else if (representation == RangeRepresentation.STRING_EXPRESSION) {
            returnedAddress = worksheet.MergeCells(range.toString());
        } else {
            returnedAddress = worksheet.mergeCells(range);
        }
        String returnedAddress2 = worksheet.MergeCells("X1:X2");
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
        worksheet.MergeCells("A1:D4");
        assertThrows(RangeException.class, () -> worksheet.MergeCells("B4:E4"));
    }

    @DisplayName("Test of the failing mergeCells function if the merge range already exists (full intersection)")
    @Test()
    void mergeCellsFailTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.MergeCells("A1:D4");
        assertThrows(RangeException.class, () -> worksheet.MergeCells("D4:A1")); // Flip addresses
    }

    @DisplayName("Test of the failing mergeCells function if the merge range is invalid (string)")
    @Test()
    void mergeCellsFailTest3() {
        Worksheet worksheet = new Worksheet();
        assertThrows(FormatException.class, () -> worksheet.MergeCells(""));
        String nullValue = null;
        assertThrows(FormatException.class, () -> worksheet.MergeCells(nullValue));
    }

    @DisplayName("Test of the internal RecalculateAutoFilter function")
    @Test()
    void recalculateAutoFilterTest() {
        Worksheet worksheet = new Worksheet();
        CorePackageTestAccess.recalculateAutoFilter(worksheet); // Dummy call
        assertTrue(worksheet.getAutoFilterRange().isEmpty());
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
        CorePackageTestAccess.recalculateAutoFilter(worksheet);
        assertTrue(worksheet.getColumns().get(2).hasAutoFilter());
        assertEquals("B1:E50", worksheet.getAutoFilterRange().orElseThrow().toString());
    }

    @DisplayName("Test of the internal recalculateColumns function")
    @Test()
    void recalculateColumnsTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.setColumnWidth(1, 22.5f);
        worksheet.setColumnWidth(2, 22.8f);
        worksheet.addHiddenColumn(3);
        worksheet.setColumnWidth(1, Worksheet.DEFAULT_WORKSHEET_COLUMN_WIDTH); // should not remove the column
        assertEquals(3, worksheet.getColumns().size());
        CorePackageTestAccess.recalculateColumns(worksheet);
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
        worksheet.MergeCells("B1:D1");
        CorePackageTestAccess.resolveMergedCells(worksheet);
        assertEquals(3, worksheet.getCells().size());
        assertNull(worksheet.getCells().get("B1").getCellStyle());
        assertEquals(Cell.CellType.EMPTY, worksheet.getCells().get("C1").getDataType());
        assertEquals(BasicStyles.getMergeCellStyle(), worksheet.getCells().get("C1").getCellStyle());
        assertEquals(22.2f, worksheet.getCells().get("C1").getValue());
        assertEquals(BasicStyles.getMergeCellStyle(), worksheet.getCells().get("D1").getCellStyle());
        assertEquals(Cell.CellType.EMPTY, worksheet.getCells().get("D1").getDataType());
    }

    @DisplayName("Test of the internal recalculateColumns function")
    @Test()
    void removeAutoFilterTest() {
        Worksheet worksheet = new Worksheet();
        worksheet.setAutoFilter("B1:F1");
        assertTrue(worksheet.getAutoFilterRange().isPresent());
        assertEquals("B1:F1", worksheet.getAutoFilterRange().orElseThrow().toString());
        worksheet.removeAutoFilter();
        assertTrue(worksheet.getAutoFilterRange().isEmpty());
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
        worksheet.MergeCells("B1:B4");
        assertTrue(worksheet.getMergedCells().containsKey("B1:B4"));
        worksheet.removeMergedCells("B1:B4");
        assertEquals(0, worksheet.getMergedCells().size());
    }

    @DisplayName("Test of the removeMergedCells function after resolution of merged cells (on save)")
    @Disabled
    @Test()
    void removeMergedCellsTest2() {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(
                "test",
                "B2"
        );
        worksheet.addCell(22, "B3");
        worksheet.MergeCells("B1:B4");
        assertTrue(worksheet.getMergedCells().containsKey("B1:B4"));
        CorePackageTestAccess.resolveMergedCells(worksheet);
        worksheet.removeMergedCells("B1:B4");
        // TODO Invoke save
        assertEquals(0, worksheet.getMergedCells().size());
        assertNotEquals(BasicStyles.getMergeCellStyle(), worksheet.getCells().get("B2").getCellStyle());
        assertNotEquals(BasicStyles.getMergeCellStyle(), worksheet.getCells().get("B3").getCellStyle());
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
        String range = (String) WorksheetTestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(
                "test",
                "B2"
        );
        worksheet.addCell(22, "B3");
        worksheet.MergeCells("B1:B4");
        assertTrue(worksheet.getMergedCells().containsKey("B1:B4"));
        assertThrows(RangeException.class, () -> worksheet.removeMergedCells(range));
    }

    @Test
    @DisplayName("Test of the clearSelectedCells function")
    void clearSelectedCellsTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getSelectedCells().isEmpty());
        worksheet.addSelectedCells("B2:D3");
        assertTrue(worksheet.getSelectedCells().contains(new Range("B2:D3")));
        worksheet.clearSelectedCells();
        assertTrue(worksheet.getSelectedCells().isEmpty());
    }

    @DisplayName("Test of the removeAllowedActionOnSheetProtection function")
    @ParameterizedTest(name = "Given value {0} should be not present after removal")
    @CsvSource(
            {
                    "DELETE_ROWS, OBJECTS, SORT",
                    "FORMAT_ROWS, OBJECTS, SORT",
                    "SELECT_LOCKED_CELLS, OBJECTS, SORT",
                    "SELECT_UNLOCKED_CELLS, OBJECTS, SORT",
                    "AUTO_FILTER, OBJECTS, SORT",
                    "SORT, OBJECTS, FORMAT_ROWS",
                    "INSERT_ROWS, OBJECTS, SORT",
                    "DELETE_COLUMNS, OBJECTS, SORT",
                    "FORMAT_CELLS, OBJECTS, SORT",
                    "FORMAT_COLUMNS, OBJECTS, SORT",
                    "INSERT_HYPERLINKS, OBJECTS, SORT",
                    "INSERT_COLUMNS, OBJECTS, SORT",
                    "OBJECTS, FORMAT_COLUMNS, SORT",
                    "PIVOT_TABLES, OBJECTS, SORT",
                    "SCENARIOS, OBJECTS, SORT",}
    )
    void removeAllowedActionOnSheetProtectionTest(
            Worksheet.SheetProtectionValue typeOfProtection, Worksheet.SheetProtectionValue additionalValue,
            Worksheet.SheetProtectionValue notPresentValue
    ) {
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
        worksheet.setActiveStyle(BasicStyles.getBold());
        assertEquals(BasicStyles.getBold(), worksheet.getActiveStyle());
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
        assertEquals(addr.column(), worksheet.getCurrentColumnNumber());
        assertEquals(addr.row(), worksheet.getCurrentRowNumber());
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
        String address = (String) WorksheetTestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        assertThrows(Exception.class, () -> worksheet.setCurrentCellAddress(address));
    }

    @DisplayName("Test of the addSelectedCells function with range objects")
    @ParameterizedTest
    @CsvSource(
            {
                    "A1:A1,A1:A1",
                    "B2:C10,B2:C10",
                    "A1:A10,A1:A10",
                    "A1:R1,A1:R1",
                    "$A$1:$R$1,A1:R1",
                    "A1:XFD1048575,A1:XFD1048575"
            }
    )
    void addSelectedCellsTest(String addressExpression, String expectedRangeExpression) {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getSelectedCells().isEmpty());
        worksheet.addSelectedCells(new Range(addressExpression));
        assertTrue(worksheet.getSelectedCells().contains(new Range(expectedRangeExpression)));
    }

    @DisplayName("Test of the addSelectedCells function with strings")
    @ParameterizedTest
    @CsvSource(
            value = {
                    "A1:A1|A1:A1",
                    "B2:C10|B2:C10",
                    "C10:B5|C10:B5",
                    "A1:A10|A1:A10",
                    "A1:R1|A1:R1",
                    "$A$1:$R$1|A1:R1",
                    "A1:XFD1048575|A1:XFD1048575",
                    "NULL|NULL"
            },
            delimiter = '|',
            nullValues = "NULL"
    )
    void addSelectedCellsTest2(String addressExpression, String expectedRange) {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getSelectedCells().isEmpty());
        worksheet.addSelectedCells(addressExpression);
        if (expectedRange == null) {
            assertTrue(worksheet.getSelectedCells().isEmpty());
        } else {
            assertTrue(worksheet.getSelectedCells().contains(new Range(expectedRange)));
        }
    }

    @DisplayName("Test of the addSelectedCells function with address objects")
    @ParameterizedTest
    @CsvSource(
            {
                    "A1,A1,A1:A1",
                    "B2,C10,B2:C10",
                    "C10,B5,B5:C10",
                    "A1,A10,A1:A10",
                    "A1,R1,A1:R1",
                    "$A$1,$R$1,A1:R1",
                    "A1,XFD1048575,A1:XFD1048575"
            }
    )
    void addSelectedCellsTest3(String startAddress, String endAddress, String expectedRange) {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getSelectedCells().isEmpty());
        worksheet.addSelectedCells(new Address(startAddress), new Address(endAddress));
        assertTrue(worksheet.getSelectedCells().contains(new Range(expectedRange)));
    }

    @DisplayName("Test of the addSelectedCells function with a single address object")
    @ParameterizedTest
    @CsvSource(
            {
                    "A1,A1:A1",
                    "B2,B2:B2",
                    "C10,C10:C10",
                    "A10,A10:A10",
                    "A$1,A1:A1",
                    "$R$1,R1:R1",
                    "XFD1048575,XFD1048575:XFD1048575"
            }
    )
    void addSelectedCellsTest4(String startAddress, String expectedRange) {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getSelectedCells().isEmpty());
        worksheet.addSelectedCells(new Address(startAddress));
        assertTrue(worksheet.getSelectedCells().contains(new Range(expectedRange)));
    }

    @DisplayName("Test of the addSelectedCells function with where an initial range already exists")
    @ParameterizedTest
    @CsvSource(
            {
                    "A1:A1,A1:A1,A1:A1",
                    "B2:B2,A2:A2,A2:B2",
                    "A2:B2,A2:A2,A2:B2",  // merged
                    "A2:C4,F1:F2,'A2:C4,F1:F2'",
                    "A2:C4,B3:D3,'A2:C4,D3:D3'",
                    "B2:C2,A1:D4,A1:D4",  // merged
                    "A2:C4,B3:D5,'A2:A4,B2:C5,D3:D5'" // merged by columns
            }
    )
    void addSelectedCellsTest5(
            String initialRangeExpression, String rangeToAddExpression,
            String expectedRangesExpression
    ) {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getSelectedCells().isEmpty());
        Range initialRange = new Range(initialRangeExpression);
        worksheet.addSelectedCells(initialRange);
        assertTrue(worksheet.getSelectedCells().contains(initialRange));
        worksheet.addSelectedCells(new Range(rangeToAddExpression));
        assertSelectedCellRanges(expectedRangesExpression, worksheet);
    }

    @DisplayName("Test of the removeSelectedCells function with a range object")
    @ParameterizedTest
    @CsvSource(
            value = {
                    "NULL|A1:A1|''",
                    "A1:A1|A1:A1|''",
                    "A1:A1|B1:B1|A1:A1",
                    "A1:C3|A1:C3|''",
                    "B2:C10|B3:C9|B2:C2,B10:C10",
                    "B2:C10|C1:D8|B2:B10,C9:C10",
                    "A1:C3|B2:B2|A1:A3,B1:B1,B3:B3,C1:C3"
            },
            delimiter = '|',
            nullValues = "NULL"
    )
    void removeSelectedCellsTest(String initialRange, String rangeToRemove, String expectedRangesExpression) {
        Worksheet worksheet = new Worksheet();
        addInitialSelectedRange(worksheet, initialRange);
        worksheet.removeSelectedCells(new Range(rangeToRemove));
        assertSelectedCellRanges(expectedRangesExpression, worksheet);
    }

    @DisplayName("Test of the removeSelectedCells function with a string (address or range)")
    @ParameterizedTest
    @CsvSource(
            value = {
                    "NULL|A1:A1|''",
                    "A1:A1|A1:A1|''",
                    "A1:A1|A1|''",
                    "A1:A1|B1:B1|A1:A1",
                    "A1:A1|B1|A1:A1",
                    "A1:C3|A1:C3|''",
                    "A1:C3|C3|A1:B3,C1:C2",
                    "A1:A2|A1|A2:A2",
                    "B2:C10|B3:C9|B2:C2,B10:C10",
                    "B2:C10|C1:D8|B2:B10,C9:C10",
                    "A1:C3|B2:B2|A1:A3,B1:B1,B3:B3,C1:C3",
                    "A1:C3|B2|A1:A3,B1:B1,B3:B3,C1:C3"
            },
            delimiter = '|',
            nullValues = "NULL"
    )
    void removeSelectedCellsTest2(
            String initialRange, String rangeOrAddressToRemove,
            String expectedRangesExpression
    ) {
        Worksheet worksheet = new Worksheet();
        addInitialSelectedRange(worksheet, initialRange);
        worksheet.removeSelectedCells(rangeOrAddressToRemove);
        assertSelectedCellRanges(expectedRangesExpression, worksheet);
    }

    @DisplayName("Test of the removeSelectedCells function with an address object")
    @ParameterizedTest
    @CsvSource(
            value = {
                    "NULL|A1|''",
                    "A1:A1|A1|''",
                    "A1:A1|B1|A1:A1",
                    "A1:C3|C3|A1:B3,C1:C2",
                    "A1:A2|A1|A2:A2",
                    "A1:C3|B2|A1:A3,B1:B1,B3:B3,C1:C3",
                    "NULL|$A1|''",
                    "A1:A1|$A$1|''",
                    "A1:A1|B$1|A1:A1",
                    "A1:C3|$B$2|A1:A3,B1:B1,B3:B3,C1:C3"
            },
            delimiter = '|',
            nullValues = "NULL"
    )
    void removeSelectedCellsTest3(String initialRange, String addressToRemove, String expectedRangesExpression) {
        Worksheet worksheet = new Worksheet();
        addInitialSelectedRange(worksheet, initialRange);
        worksheet.removeSelectedCells(new Address(addressToRemove));
        assertSelectedCellRanges(expectedRangesExpression, worksheet);
    }

    @DisplayName("Test of the removeSelectedCells function with a start and end address objects")
    @ParameterizedTest
    @CsvSource(
            value = {
                    "NULL|A1|A1|''",
                    "A1:A1|A1|A1|''",
                    "A1:A1|B1|B1|A1:A1",
                    "A1:C3|A1|C3|''",
                    "B2:C10|B3|C9|B2:C2,B10:C10",
                    "B2:C10|C1|D8|B2:B10,C9:C10",
                    "A1:C3|B2|B2|A1:A3,B1:B1,B3:B3,C1:C3"
            },
            delimiter = '|',
            nullValues = "NULL"
    )
    void removeSelectedCellsTest4(
            String initialRange, String startAddressExpression,
            String endAddressExpression, String expectedRangesExpression
    ) {
        Worksheet worksheet = new Worksheet();
        addInitialSelectedRange(worksheet, initialRange);
        worksheet.removeSelectedCells(new Address(startAddressExpression), new Address(endAddressExpression));
        assertSelectedCellRanges(expectedRangesExpression, worksheet);
    }

    @DisplayName("Test of the setSheetProtectionPassword function")
    @ParameterizedTest(
            name = "Given value {1} (type {0}) should lead to the password {3} and sheet protection should" +
                    " be {4}"
    )
    @CsvSource(
            {
                    "NULL, '', NULL, '', false",
                    "STRING, '', NULL, '', false",
                    "STRING, 'x', STRING, 'x', true",
                    "STRING, '***', STRING, '***', true",}
    )
    void setSheetProtectionPasswordTest(
            String sourceType, String sourceValue, String expectedType, String expectedValue, boolean expectedUsage) {
        String password = (String) WorksheetTestUtils.createInstance(sourceType, sourceValue);
        String expectedPassword = (String) WorksheetTestUtils.createInstance(expectedType, expectedValue);

        Worksheet worksheet = new Worksheet();
        assertFalse(worksheet.getSheetProtection());
        assertFalse(worksheet.getSheetProtectionPassword().passwordIsSet());
        worksheet.setSheetProtectionPassword(password);
        assertEquals(expectedUsage, worksheet.getSheetProtection());
        assertEquals(expectedPassword, worksheet.getSheetProtectionPassword().getPassword());
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
    void setSheetNameTest(
            String sourceType, String sourceValue, boolean expectedValid, String expectedType, String expectedValue) {
        String name = (String) WorksheetTestUtils.createInstance(sourceType, sourceValue);
        String expectedName = (String) WorksheetTestUtils.createInstance(expectedType, expectedValue);

        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getSheetName());
        if (expectedValid) {
            worksheet.setSheetName(name);
            assertEquals(expectedName, worksheet.getSheetName());
        } else {
            assertThrows(FormatException.class, () -> worksheet.setSheetName(name));
        }
    }

    @DisplayName("Test of the SetSheetName function with sanitation when one worksheet already exists")
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
                    "true, 'STRING', '--------------------------------', true, STRING, " +
                            "'-------------------------------'",}
    )
    void setSheetNameTest2(
            boolean useSanitation, String sourceType, String sourceValue, boolean expectedValid, String expectedType,
            String expectedValue
    ) {
        String name = (String) WorksheetTestUtils.createInstance(sourceType, sourceValue);
        String expectedName = (String) WorksheetTestUtils.createInstance(expectedType, expectedValue);

        Workbook workbook = new Workbook("Sheet1");
        workbook.addWorksheet("test");
        Worksheet worksheet = workbook.getCurrentWorksheet();
        assertEquals("test", worksheet.getSheetName());
        if (expectedValid) {
            worksheet.setSheetName(name, useSanitation);
            assertEquals(expectedName, worksheet.getSheetName());
        } else {
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

    @DisplayName("Test of the replaceCellValue on an existing occurrence")
    @ParameterizedTest(name = "Test replaceCellValue: oldValue={1}, newValue={3}, differentValue={5}")
    @CsvSource(
            {
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
            }
    )
    void replaceCellValue_ShouldReplaceAllOccurrences(
            String oldType, String oldValueString, String newType, String newValueString, String differentType,
            String differentValueString
    ) {
        Worksheet worksheet = new Worksheet();
        Object oldValue = WorksheetTestUtils.createInstance(oldType, oldValueString);
        Object newValue = WorksheetTestUtils.createInstance(newType, newValueString);
        Object differentValue = WorksheetTestUtils.createInstance(differentType, differentValueString);

        worksheet.addCell(oldValue, 0, 0);
        worksheet.addCell(oldValue, 1, 0);
        worksheet.addCell(oldValue, 2, 0);
        worksheet.addCell(differentValue, 3, 0);

        int replacedCount = worksheet.replaceCellValue(oldValue, newValue);

        assertEquals(3, replacedCount);
        assertEquals(newValue, worksheet.getCells().get("A1").getValue());
        assertEquals(newValue, worksheet.getCells().get("B1").getValue());
        assertEquals(newValue, worksheet.getCells().get("C1").getValue());
        assertEquals(differentValue, worksheet.getCells().get("D1").getValue());
    }

    @DisplayName("Test of the replaceCellValue on an existing occurrence of a Date value")
    @Test()
    void replaceCellValue_ShouldReplaceAllOccurrences2() {
        Worksheet worksheet = new Worksheet();
        String oldValue = "oldValue";
        worksheet.addCell(oldValue, 0, 0);
        worksheet.addCell(oldValue, 1, 0);
        worksheet.addCell(oldValue, 2, 0);
        worksheet.addCell("differentValue", 3, 0);

        Date newValue = WorksheetTestUtils.buildDate(2025, 2, 10, 5, 6, 7);
        int replacedCount = worksheet.replaceCellValue(oldValue, newValue);

        assertEquals(3, replacedCount);
        assertEquals(newValue, worksheet.getCells().get("A1").getValue());
        assertEquals(newValue, worksheet.getCells().get("B1").getValue());
        assertEquals(newValue, worksheet.getCells().get("C1").getValue());
        assertEquals("differentValue", worksheet.getCells().get("D1").getValue());
    }

    @DisplayName("Test of the replaceCellValue when one existing values is to replace")
    @ParameterizedTest(
            name = "Test replaceCellValue: oldType={0}, oldValue={1}, newType={2}, newValue={3}, " +
                    "diffType={4}, diffValue={5}"
    )
    @CsvSource(
            {
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
            }
    )
    void replaceCellValue_ShouldNotReplaceAnyOccurrences(
            String oldType, String oldValueString,
            String newType, String newValueString,
            String diffType, String diffValueString
    ) {

        Object oldValue = WorksheetTestUtils.createInstance(oldType, oldValueString);
        Object newValue = WorksheetTestUtils.createInstance(newType, newValueString);
        Object differentValue = WorksheetTestUtils.createInstance(diffType, diffValueString);

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

    @DisplayName("Test of the firstOrDefaultCell when one existing values exists")
    @ParameterizedTest(name = "Test firstOrDefaultCell: type={0}, matchingValue={1}")
    @CsvSource(
            {
                    "NULL, null",
                    "STRING, ''",
                    // Empty string
                    "INTEGER, 0",
                    "STRING, 'null'",
                    // String "null"
                    "INTEGER, -1",
                    "FLOAT, 0.05",
                    "BOOLEAN, true",
                    "BOOLEAN, false",
                    "FLOAT, -22.357"
            }
    )
    void firstOrDefaultCell_ShouldReturnCorrectCell(String sourceType, String sourceValue) {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell("value1", 0, 0);
        Object matchingValue = WorksheetTestUtils.createInstance(sourceType, sourceValue);
        worksheet.addCell(matchingValue, 1, 0);
        worksheet.addCell("value3", 2, 0);

        Cell result = worksheet.firstOrDefaultCell(cell -> Objects.equals(cell.getValue(), matchingValue)).orElse(null);

        assertNotNull(result);
        assertEquals(matchingValue, result.getValue());
    }

    @DisplayName("Test of the firstOrDefaultCell when no existing values exists")
    @ParameterizedTest(name = "Test firstOrDefaultCell when no existing values exists: matchingValue={1}")
    @CsvSource(
            {
                    "NULL, null",
                    "STRING, ''",
                    // Empty string
                    "INTEGER, 0",
                    "STRING, 'null'",
                    // String "null"
                    "INTEGER, -1",
                    "FLOAT, 0.05",
                    "BOOLEAN, true",
                    "BOOLEAN, false",
                    "FLOAT, -22.357"
            }
    )
    void firstOrDefaultCell_NotFound(String sourceType, String sourceValue) {
        Worksheet worksheet = new Worksheet();
        Object nonMatchingValue = WorksheetTestUtils.createInstance(sourceType, sourceValue);
        worksheet.addCell("value1", 0, 0);
        worksheet.addCell("value2", 1, 0);
        worksheet.addCell("value3", 2, 0);

        Cell result = worksheet.firstOrDefaultCell(cell -> Objects.equals(cell.getValue(), nonMatchingValue))
                .orElse(null);

        assertNull(result);
    }

    @DisplayName("Test of the firstCellByValue when one existing values exists")
    @ParameterizedTest(name = "Test firstCellByValue: matchingValue={0}")
    @CsvSource(
            {
                    "NULL, null",
                    "STRING, ''",
                    // Empty string
                    "INTEGER, 0",
                    "STRING, 'null'",
                    // String "null"
                    "INTEGER, -1",
                    "FLOAT, 0.05",
                    "BOOLEAN, true",
                    "BOOLEAN, false",
                    "FLOAT, -22.357"
            }
    )
    void firstCellByValue_ShouldReturnCorrectCell(String sourceType, String sourceValue) {
        Worksheet worksheet = new Worksheet();

        var cell1 = new Cell("Test1", Cell.CellType.STRING, "A1");
        Object matchingValue = WorksheetTestUtils.createInstance(sourceType, sourceValue);
        var cell2 = new Cell(matchingValue, Cell.CellType.DEFAULT, "B1");
        worksheet.addCell(cell1.getValue(), "A1");
        worksheet.addCell(cell2.getValue(), "B1");

        Cell result = worksheet.firstCellByValue(matchingValue).orElse(null);

        // Assertions
        assertNotNull(result);
        assertEquals(matchingValue, result.getValue());
        assertEquals("B1", result.getCellAddress());
    }

    @DisplayName("Test of the FirstCellByValue when one existing values exists, on a Date value")
    @Test()
    void firstCellByValue_ShouldReturnCorrectCell2() {
        Worksheet worksheet = new Worksheet();

        var cell1 = new Cell("Test1", Cell.CellType.STRING, "A1");
        Date matchingValue = WorksheetTestUtils.buildDate(2025, 2, 10, 5, 6, 7);
        var cell2 = new Cell(matchingValue, Cell.CellType.DATE, "B1");
        worksheet.addCell(cell1.getValue(), "A1");
        worksheet.addCell(cell2.getValue(), "B1");

        Cell result = worksheet.firstCellByValue(matchingValue).orElse(null);

        assertNotNull(result);
        assertEquals(matchingValue, result.getValue());
        assertEquals("B1", result.getCellAddress());
    }

    @DisplayName("Test of the FirstCellByValue when no existing values exists")
    @ParameterizedTest(
            name = "Test FirstCellByValue when no existing value matches: notMatchingType={0}, " +
                    "notMatchingValue={1}"
    )
    @CsvSource(
            {
                    "'NULL', null",
                    "'STRING', ''",
                    "'INTEGER', 0",
                    "'STRING', 'null'",
                    "'INTEGER', -1",
                    "'FLOAT', 0.05",
                    "'BOOLEAN', true",
                    "'BOOLEAN', false",
                    "'DOUBLE', -22.357"
            }
    )
    void firstCellByValue_ShouldReturnNull_WhenValueNotFound(String notMatchingType, String notMatchingValueString) {

        Object notMatchingValue = WorksheetTestUtils.createInstance(notMatchingType, notMatchingValueString);
        Worksheet worksheet = new Worksheet();
        Cell cell1 = new Cell("Test1", Cell.CellType.STRING, "A1");
        worksheet.addCell(cell1.getValue(), "A1");

        Cell result = worksheet.firstCellByValue(notMatchingValue).orElse(null);
        
        assertNull(result, "Result should be null when the value is not found");
    }

    @DisplayName("Test of the insertRow function")
    @Test
    void testInsertRow() {
        Worksheet worksheet = new Worksheet();
        Style style1 = new Style();
        style1.getCurrentFont().setBold(true);
        Style style2 = new Style();
        style2.getCurrentFont().setItalic(true);

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
        assertTrue(worksheet.getCells().get("A2").getCellStyle().getCurrentFont().isBold(), "A2 should be bold");
        assertTrue(worksheet.getCells().get("A3").getCellStyle().getCurrentFont().isBold(), "A3 should be bold");
        assertTrue(worksheet.getCells().get("A4").getCellStyle().getCurrentFont().isBold(), "A4 should be bold");
        assertTrue(worksheet.getCells().get("A5").getCellStyle().getCurrentFont().isItalic(), "A5 should be italic");
        assertNull(worksheet.getCells().get("A6").getCellStyle(), "A6 should not have a style");
        assertNull(worksheet.getCells().get("B8").getCellStyle(), "B8 should not have a style");
    }

    @DisplayName("Test of the insertColumn function")
    @Test
    void testInsertColumn() {
        Worksheet worksheet = new Worksheet();
        Style style1 = new Style();
        style1.getCurrentFont().setBold(true);
        Style style2 = new Style();
        style2.getCurrentFont().setItalic(true);

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
        assertTrue(worksheet.getCells().get("B2").getCellStyle().getCurrentFont().isBold(), "B2 should be bold");
        assertTrue(worksheet.getCells().get("C2").getCellStyle().getCurrentFont().isBold(), "C2 should be bold");
        assertTrue(worksheet.getCells().get("D2").getCellStyle().getCurrentFont().isBold(), "D2 should be bold");
        assertTrue(worksheet.getCells().get("E2").getCellStyle().getCurrentFont().isItalic(), "E2 should be italic");
        assertNull(worksheet.getCells().get("F2").getCellStyle(), "F2 should not have a style");
        assertNull(worksheet.getCells().get("H3").getCellStyle(), "H3 should not have a style");
    }

    public static void assertAddedCell(
            Worksheet worksheet, int numberOfEntries, String expectedAddress, Cell.CellType expectedType,
            Style expectedStyle, Object expectedValue, int nextColumn, int nextRow
    ) {
        assertEquals(numberOfEntries, worksheet.getCells().size());
        WorksheetTestUtils.assertMapEntry(expectedAddress, expectedValue, worksheet.getCells(), Cell::getValue);
        WorksheetTestUtils.assertMapEntry(expectedAddress, expectedType, worksheet.getCells(), Cell::getDataType);
        WorksheetTestUtils.assertMapEntry(expectedAddress, expectedAddress, worksheet.getCells(), Cell::getCellAddress);
        if (expectedStyle == null) {
            assertNull(worksheet.getCells().get(expectedAddress).getCellStyle());
        } else {
            assertEquals(expectedStyle, worksheet.getCells().get(expectedAddress).getCellStyle());
        }
        assertEquals(nextColumn, worksheet.getCurrentColumnNumber());
        assertEquals(nextRow, worksheet.getCurrentRowNumber());
    }

    public static Worksheet initWorksheet(Worksheet worksheet, String address, Worksheet.CellDirection direction) {
        return initWorksheet(worksheet, address, direction, null);
    }

    public static Worksheet initWorksheet(
            Worksheet worksheet, String address, Worksheet.CellDirection direction, Style style) {
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
        assertEquals(Worksheet.DEFAULT_WORKSHEET_COLUMN_WIDTH, worksheet.getDefaultColumnWidth());
        assertEquals(Worksheet.DEFAULT_WORKSHEET_ROW_HEIGHT, worksheet.getDefaultRowHeight());
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
        assertTrue(worksheet.getActivePane().isEmpty());
    }

    private static void addInitialSelectedRange(Worksheet worksheet, String initialRange) {
        assertTrue(worksheet.getSelectedCells().isEmpty());
        if (initialRange != null) {
            worksheet.addSelectedCells(new Range(initialRange));
            assertFalse(worksheet.getSelectedCells().isEmpty());
        }
    }

    private static void assertSelectedCellRanges(String expectedRangesExpression, Worksheet worksheet) {
        List<Range> expectedRanges = new ArrayList<>();
        if (expectedRangesExpression != null && !expectedRangesExpression.isEmpty()) {
            for (String range : expectedRangesExpression.split(",")) {
                expectedRanges.add(new Range(range));
            }
        }
        assertEquals(expectedRanges.size(), worksheet.getSelectedCells().size());
        for (Range expectedRange : expectedRanges) {
            assertTrue(worksheet.getSelectedCells().stream()
                    .anyMatch(range -> range.toString().equals(expectedRange.toString())));
        }
    }
}



