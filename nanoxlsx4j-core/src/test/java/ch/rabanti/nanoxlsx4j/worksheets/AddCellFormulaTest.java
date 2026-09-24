package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.function.BiConsumer;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.CorePackageTestAccess;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;

public class AddCellFormulaTest {

    private Worksheet worksheet;

    @DisplayName("Test of the addCellFormula function with the only the value (with address and column/row invocation)")
    @Test()
    void addCellFormulaTest1() {
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.ROW_TO_ROW);
        invokeAddCellFormulaTest("=B2", 2, 3, worksheet::addCellFormula, "C4", 2, 4);
        Address address = new Address(3, 1);
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.COLUMN_TO_COLUMN);
        invokeAddCellFormulaTest("=B2", address.getAddress(), worksheet::addCellFormula, "D2", 4, 1);
    }

    @DisplayName("Test of the addCellFormula function with value and Style (with address and column/row invocation)")
    @Test()
    void addCellFormulaTest2() {
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.ROW_TO_ROW);
        invokeAddCellFormulaTest(
                "=B2", 2, 3, BasicStyles.getBoldItalic(), worksheet::addCellFormula, "C4", 2, 4,
                BasicStyles.getBoldItalic()
        );
        Address address = new Address(3, 1);
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.COLUMN_TO_COLUMN);
        invokeAddCellFormulaTest(
                "=B2", address.getAddress(), BasicStyles.getBold(), worksheet::addCellFormula, "D2", 4,
                1, BasicStyles.getBold()
        );
    }

    @DisplayName(
            "Test of the addCellFormula function with value and active worksheet style (with address and " +
                    "column/row invocation)"
    )
    @Test()
    void addCellFormulaTest3() {
        worksheet = WorksheetTest.initWorksheet(
                worksheet, "D2", Worksheet.CellDirection.ROW_TO_ROW,
                BasicStyles.getBorderFrameHeader()
        );
        invokeAddCellFormulaTest(
                "=B2", 2, 3, worksheet::addCellFormula, "C4", 2, 4, BasicStyles.getBorderFrameHeader());
        Address address = new Address(3, 1);
        worksheet = WorksheetTest.initWorksheet(
                worksheet, "R3", Worksheet.CellDirection.COLUMN_TO_COLUMN,
                BasicStyles.getBorderFrameHeader()
        );
        invokeAddCellFormulaTest(
                "=B2", address.getAddress(), worksheet::addCellFormula, "D2", 4, 1,
                BasicStyles.getBorderFrameHeader()
        );
    }

    @DisplayName(
            "Test of the AddCell function for a nested cell object with a formula (with address and column/row " +
                    "invocation)"
    )
    @Test()
    void addCellFormulaTest4() {
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.ROW_TO_ROW);
        Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
        worksheet.addCell(cell, 3, 1);
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, null, "=B2", 3, 2);
        worksheet = new Worksheet();
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.COLUMN_TO_COLUMN);
        Address address = new Address(3, 1);
        worksheet.addCell(cell, address.getAddress());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, null, "=B2", 4, 1);
    }

    @DisplayName(
            "Test of the AddCell function for a nested cell object with a formula and style (with address and " +
                    "column/row invocation)"
    )
    @Test()
    void addCellFormulaTest5() {
        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.ROW_TO_ROW);
        Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
        worksheet.addCell(cell, 3, 1, BasicStyles.getBold());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.getBold(), "=B2", 3, 2);
        worksheet = new Worksheet();
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.COLUMN_TO_COLUMN);
        Address address = new Address(3, 1);
        worksheet.addCell(cell, address.getAddress());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.getBold(), "=B2", 4, 1);

        worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.ROW_TO_ROW);
        cell = new Cell("=B2", Cell.CellType.FORMULA, "R2");
        cell.setStyle(BasicStyles.getBorderFrame());
        Style mixedStyle = BasicStyles.getBorderFrame();
        mixedStyle.append(BasicStyles.getBold());
        worksheet.addCell(cell, 3, 1, BasicStyles.getBold());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, mixedStyle, "=B2", 3, 2);
        worksheet = new Worksheet();
        worksheet = WorksheetTest.initWorksheet(worksheet, "R3", Worksheet.CellDirection.COLUMN_TO_COLUMN);
        worksheet.addCell(cell, address.getAddress(), BasicStyles.getBold());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, mixedStyle, "=B2", 4, 1);
    }

    @DisplayName(
            "Test of the AddCell function for a nested cell object with a formula and active worksheet style " +
                    "(with address and column/row invocation)"
    )
    @Test()
    void addCellFormulaTest6() {
        worksheet = WorksheetTest.initWorksheet(
                worksheet, "D2", Worksheet.CellDirection.ROW_TO_ROW,
                BasicStyles.getBorderFrame()
        );
        Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
        worksheet.addCell(cell, 3, 1);
        WorksheetTest.assertAddedCell(
                worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.getBorderFrame(), "=B2", 3,
                2
        );
        worksheet = WorksheetTest.initWorksheet(
                worksheet, "D2", Worksheet.CellDirection.COLUMN_TO_COLUMN,
                BasicStyles.getBorderFrame()
        );
        Address address = new Address(3, 1);
        worksheet.addCell(cell, address.getAddress());
        WorksheetTest.assertAddedCell(
                worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.getBorderFrame(), "=B2", 4,
                1
        );

        worksheet = WorksheetTest.initWorksheet(
                worksheet, "D2", Worksheet.CellDirection.ROW_TO_ROW,
                BasicStyles.getBorderFrame()
        );
        cell = new Cell("=B2", Cell.CellType.FORMULA, "R2");
        cell.setStyle(BasicStyles.getBold());
        Style mixedStyle = BasicStyles.getBorderFrame();
        mixedStyle.append(BasicStyles.getBold());
        worksheet.addCell(cell, 3, 1);
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, mixedStyle, "=B2", 3, 2);
        worksheet = WorksheetTest.initWorksheet(
                worksheet, "D2", Worksheet.CellDirection.COLUMN_TO_COLUMN,
                BasicStyles.getBorderFrame()
        );
        worksheet.addCell(cell, address.getAddress());
        WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, mixedStyle, "=B2", 4, 1);
    }

    @DisplayName(
            "Test of the addCellFormula function with when changing the current cell direction (with address and" +
                    " column/row invocation)"
    )
    @ParameterizedTest(
            name = "Given initial column {0} and row {1} should lead to the column {3} and row {4} for " +
                    "direction {2}"
    )
    @CsvSource(
            {
                    "D2, 3, 1, ROW_TO_ROW, 3, 2",
                    "E7, 7, 2, COLUMN_TO_COLUMN, 8, 2",
                    "C9, 10, 5, DISABLED, 2, 8",}
    )
    void addCellFormulaTest7(
            String worksheetAddress, int initialColumn, int initialRow,
            Worksheet.CellDirection cellDirection, int expectedNextColumn, int expectedNextRow
    ) {
        Address initialAddress = new Address(initialColumn, initialRow);
        worksheet = WorksheetTest.initWorksheet(worksheet, worksheetAddress, cellDirection);
        invokeAddCellFormulaTest(
                "=B2", initialColumn, initialRow, worksheet::addCellFormula, initialAddress.getAddress(),
                expectedNextColumn, expectedNextRow
        );
        worksheet = WorksheetTest.initWorksheet(worksheet, worksheetAddress, cellDirection);
        invokeAddCellFormulaTest(
                "=B2", initialAddress.getAddress(), worksheet::addCellFormula, initialAddress.getAddress(),
                expectedNextColumn, expectedNextRow
        );
    }

    @Test
    @DisplayName(
            "Test of the formula handling of attached feature sets. an identical feature set should not be " +
                    "reattached (coverage)"
    )
    void preventFeatureRebindTest() {
        Workbook workbook = new Workbook();
        Worksheet target = new Worksheet("worksheet1");
        workbook.addWorksheet(target);
        Cell cell = new Cell("A1+A2", Cell.CellType.FORMULA);
        target.addCell(cell, "B1");
        target.addCell(cell, "B2");
        assertTrue(
                CorePackageTestAccess.hasSameFeatureSet(
                        workbook.getCurrentWorksheet().getCells().get("B1").getFormula(),
                        workbook.getCurrentWorksheet().getCells().get("B2").getFormula()
                ));
    }

    private <T1> void invokeAddCellFormulaTest(
            String value, T1 parameter1, BiConsumer<String, T1> action, String expectedAddress, int expectedNextColumn,
            int expectedNextRow
    ) {
        invokeAddCellFormulaTest(value, parameter1, action, expectedAddress, expectedNextColumn, expectedNextRow, null);
    }

    private <T1> void invokeAddCellFormulaTest(
            String value, T1 parameter1, BiConsumer<String, T1> action, String expectedAddress, int expectedNextColumn,
            int expectedNextRow, Style expectedStyle
    ) {
        assertEquals(0, worksheet.getCells().size());
        action.accept(value, parameter1);
        assertAddedFormulaCell(
                worksheet, 1, expectedAddress, expectedStyle, value, expectedNextColumn,
                expectedNextRow
        );
        worksheet = new Worksheet(); // Auto-reset
    }

    private <T1, T2> void invokeAddCellFormulaTest(
            String value, T1 parameter1, T2 parameter2,
            WorksheetTestUtils.TriConsumer<String, T1, T2> action, String expectedAddress, int expectedNextColumn,
            int expectedNextRow
    ) {
        invokeAddCellFormulaTest(
                value, parameter1, parameter2, action, expectedAddress, expectedNextColumn,
                expectedNextRow, null
        );
    }

    private <T1, T2> void invokeAddCellFormulaTest(
            String value, T1 parameter1, T2 parameter2,
            WorksheetTestUtils.TriConsumer<String, T1, T2> action, String expectedAddress, int expectedNextColumn,
            int expectedNextRow, Style expectedStyle
    ) {
        assertEquals(0, worksheet.getCells().size());
        action.accept(value, parameter1, parameter2);
        assertAddedFormulaCell(
                worksheet, 1, expectedAddress, expectedStyle, value, expectedNextColumn,
                expectedNextRow
        );
        worksheet = new Worksheet(); // Auto-reset
    }

    private <T1, T2, T3> void invokeAddCellFormulaTest(
            String value, T1 parameter1, T2 parameter2, T3 parameter3,
            WorksheetTestUtils.QuadConsumer<String, T1, T2, T3> action, String expectedAddress, int expectedNextColumn,
            int expectedNextRow
    ) {
        invokeAddCellFormulaTest(
                value, parameter1, parameter2, parameter3, action, expectedAddress, expectedNextColumn, expectedNextRow,
                null
        );
    }

    private <T1, T2, T3> void invokeAddCellFormulaTest(
            String value, T1 parameter1, T2 parameter2, T3 parameter3,
            WorksheetTestUtils.QuadConsumer<String, T1, T2, T3> action, String expectedAddress, int expectedNextColumn,
            int expectedNextRow, Style expectedStyle
    ) {
        assertEquals(0, worksheet.getCells().size());
        action.accept(value, parameter1, parameter2, parameter3);
        assertAddedFormulaCell(
                worksheet, 1, expectedAddress, expectedStyle, value, expectedNextColumn,
                expectedNextRow
        );
        worksheet = new Worksheet(); // Auto-reset
    }

    private void assertAddedFormulaCell(
            Worksheet worksheet, int numberOfEntries, String expectedAddress, Style expectedStyle, String expectedValue,
            int nextColumn, int nextRow
    ) {
        assertEquals(numberOfEntries, worksheet.getCells().size());
        WorksheetTestUtils.assertMapEntry(expectedAddress, expectedValue, worksheet.getCells(), Cell::getValue);
        assertEquals(Cell.CellType.FORMULA, worksheet.getCells().get(expectedAddress).getDataType());
        assertEquals(expectedValue, worksheet.getCells().get(expectedAddress).getValue());
        if (expectedStyle == null) {
            assertNull(worksheet.getCells().get(expectedAddress).getCellStyle());
        } else {
            assertEquals(expectedStyle, worksheet.getCells().get(expectedAddress).getCellStyle());
        }
        assertEquals(nextColumn, worksheet.getCurrentColumnNumber());
        assertEquals(nextRow, worksheet.getCurrentRowNumber());
    }
}
