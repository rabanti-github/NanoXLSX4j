package ch.rabanti.nanoxlsx4j.cells.types;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.EnumSource;
import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.CorePackageTestAccess;
import ch.rabanti.nanoxlsx4j.FormulaData;
import ch.rabanti.nanoxlsx4j.enums.FormulaError;

class FormulaCellTest {

    @ParameterizedTest
    @DisplayName("Formula value cell test: Test creation and value modification")
    @CsvSource(
            value = {
                    "A1|SUM(A1:A3)",
                    "Text|[1]ExternalSheet!A1",
                    "''|B2"},
            delimiter = '|'
    )
    void formulaValueTest(String initialValue, String expectedValue) {
        Cell cell = new Cell(initialValue, Cell.CellType.FORMULA, new Address(0, 0));

        assertEquals(Cell.CellType.FORMULA, cell.getDataType());
        assertEquals(initialValue, cell.getValue());
        assertNotNull(cell.getFormula());
        assertEquals(initialValue, cell.getFormula().getExpression());

        FormulaData formula = cell.getFormula();
        cell.setValue(expectedValue);

        assertEquals(Cell.CellType.FORMULA, cell.getDataType());
        assertSame(formula, cell.getFormula());
        assertEquals(expectedValue, cell.getValue());
        assertEquals(expectedValue, cell.getFormula().getExpression());
    }

    @Test
    @DisplayName("Changing a string cell to Formula creates synchronized formula metadata")
    void stringToFormulaTest() {
        Cell cell = new Cell("Initial Value", Cell.CellType.STRING);

        cell.setDataType(Cell.CellType.FORMULA);

        assertEquals(Cell.CellType.FORMULA, cell.getDataType());
        assertNotNull(cell.getFormula());
        assertEquals("Initial Value", cell.getFormula().getExpression());

        cell.setValue("[1]ExternalSheet!A1");
        assertEquals("[1]ExternalSheet!A1", cell.getFormula().getExpression());
        assertTrue(cell.getFormula().hasExternalReferences());
    }

    @ParameterizedTest
    @DisplayName("Changing a formula cell to a non-formula type clears formula metadata")
    @EnumSource(
            value = Cell.CellType.class,
            names = {
                    "STRING",
                    "NUMBER",
                    "BOOL",
                    "EMPTY",
                    "DATE",
                    "TIME",
                    "ERROR"}
    )
    void formulaToNonFormulaTest(Cell.CellType targetType) {
        Cell cell = new Cell("A1", Cell.CellType.FORMULA);

        cell.setDataType(targetType);

        assertEquals(targetType, cell.getDataType());
        assertEquals("A1", cell.getValue());
        assertNull(cell.getFormula());
    }

    @Test
    @DisplayName("Test of reattaching the feature set to a formula cell, when a new FormulaData object is assigned")
    void formulaReattachFeatureSetTest() {
        Cell cell = new Cell("A1+B1", Cell.CellType.FORMULA);
        cell.setDataType(Cell.CellType.STRING);
        assertNull(cell.getFormula());
        assertEquals(Cell.CellType.STRING, cell.getDataType());
        assertEquals("A1+B1", cell.getValue().toString());

        FormulaData formulaData = new FormulaData("A2+B2");
        CorePackageTestAccess.setFormula(cell, formulaData);
        assertEquals("A2+B2", cell.getValue().toString());

        cell.setDataType(Cell.CellType.FORMULA);
        assertNotNull(cell.getFormula());
        assertEquals(Cell.CellType.FORMULA, cell.getDataType());
        assertEquals("A2+B2", cell.getFormula().getExpression());
    }

    @Test
    @DisplayName("Assigning null to a formula cell resolves Empty and clears formula metadata")
    void formulaToEmptyByValueTest() {
        Cell cell = new Cell("A1", Cell.CellType.FORMULA);

        cell.setValue(null);

        assertEquals(Cell.CellType.EMPTY, cell.getDataType());
        assertNull(cell.getValue());
        assertNull(cell.getFormula());
    }

    @Test
    @DisplayName("A cached formula error does not change the formula cell type or expression")
    void formulaCachedErrorTest() {
        Cell cell = new Cell("A1/0", Cell.CellType.FORMULA);
        FormulaData formula = cell.getFormula();

        formula.setCachedValue(FormulaError.DIVISION_BY_ZERO);
        CorePackageTestAccess.setCachedValueType(formula, Cell.CellType.ERROR);

        assertEquals(Cell.CellType.FORMULA, cell.getDataType());
        assertEquals("A1/0", cell.getValue());
        assertSame(formula, cell.getFormula());
        assertEquals("A1/0", cell.getFormula().getExpression());
        assertEquals(FormulaError.DIVISION_BY_ZERO, cell.getFormula().getCachedValue());
        assertEquals(Cell.CellType.ERROR, cell.getFormula().getCachedValueType());
    }

    @Test
    @DisplayName("Linked formula cells retain cached-value behavior")
    void linkedFormulaValueTest() {
        Cell cell = new Cell(null, Cell.CellType.FORMULA);
        cell.getFormula().setMasterCellAddress("A1");

        cell.setValue("cached value");

        assertEquals(Cell.CellType.FORMULA, cell.getDataType());
        assertEquals("cached value", cell.getValue());
        assertNull(cell.getFormula().getExpression());
    }

    @Test
    @DisplayName("Copying a formula cell creates independent formula metadata")
    void formulaCopyTest() {
        Cell cell = new Cell("[1]ExternalSheet!A1", Cell.CellType.FORMULA, new Address(2, 3));

        Cell copy = CorePackageTestAccess.copyCell(cell);

        assertEquals(cell.getDataType(), copy.getDataType());
        assertEquals(cell.getValue(), copy.getValue());
        assertNotNull(copy.getFormula());
        assertNotSame(cell.getFormula(), copy.getFormula());
        assertEquals(cell.getFormula().getExpression(), copy.getFormula().getExpression());
        assertEquals(cell.getFormula().hasExternalReferences(), copy.getFormula().hasExternalReferences());
    }
}
