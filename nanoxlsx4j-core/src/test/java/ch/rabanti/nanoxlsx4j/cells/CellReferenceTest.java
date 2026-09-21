package ch.rabanti.nanoxlsx4j.cells;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.CorePackageTestAccess;
import ch.rabanti.nanoxlsx4j.DefinedName;
import ch.rabanti.nanoxlsx4j.FormulaData;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.Workbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.ValueSource;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;

class CellReferenceTest {

    @ParameterizedTest
    @DisplayName("Test of all addCellReference overloads")
    @ValueSource(
            ints = {
                    0, // row, column
                    1, // row, column, style
                    2, // address
                    3} // address, style
    )
    void addCellReferenceOverloadTest(int overload) {
        Workbook workbook = new Workbook("Sheet1");
        DefinedName name = workbook.addDefinedNameFormula("FormulaName", "SUM(A1:A2)");
        Style style = (Style) BasicStyles.getBold().copy();
        List<Address> addresses;
        switch (overload) {
            case 0 -> addresses = workbook.getCurrentWorksheet().addCellReference(name, 1, 1, 7);
            case 1 -> addresses = workbook.getCurrentWorksheet().addCellReference(name, 1, 1, style, 7);
            case 2 -> addresses = workbook.getCurrentWorksheet().addCellReference(name, "B2", 7);
            default -> addresses = workbook.getCurrentWorksheet().addCellReference(name, "B2", style, 7);
        }

        assertEquals(1, addresses.size());
        assertEquals(new Address("B2"), addresses.getFirst());
        Cell cell = workbook.getCurrentWorksheet().getCells().get("B2");
        assertEquals(Cell.CellType.FORMULA, cell.getDataType());
        assertEquals("FormulaName", cell.getValue());
        assertEquals("FormulaName", cell.getFormula().getExpression());
        assertSame(name, cell.getFormula().getDefinedNameReference());
        assertEquals("7", cell.getFormula().getCachedValue());
        assertEquals(Cell.CellType.NUMBER, cell.getFormula().getCachedValueType());
        assertNull(cell.getFormula().getFormulaRange());
        if (overload == 1 || overload == 3) {
            assertEquals(style.hashCode(), cell.getCellStyle().hashCode());
        } else {
            assertNull(cell.getCellStyle());
        }
    }

    @Test
    @DisplayName("Test of cell and constant defined name references")
    void cellAndConstantReferenceTest() {
        Workbook workbook = new Workbook("Sheet1");
        DefinedName cellName = workbook.addDefinedNameCell("CellName", workbook.getCurrentWorksheet(), "A1");
        DefinedName constantName = workbook.addDefinedNameConstant("ConstantName", true);

        workbook.getCurrentWorksheet().addCellReference(cellName, "B1");
        workbook.getCurrentWorksheet().addCellReference(constantName, "B2", 999);

        assertNull(workbook.getCurrentWorksheet().getCells().get("B1").getFormula().getFormulaRange());
        assertEquals(
                Cell.CellType.NUMBER,
                workbook.getCurrentWorksheet().getCells().get("B1").getFormula().getCachedValueType()
        );
        assertEquals("TRUE", workbook.getCurrentWorksheet().getCells().get("B2").getFormula().getCachedValue());
        assertEquals(
                Cell.CellType.BOOL,
                workbook.getCurrentWorksheet().getCells().get("B2").getFormula().getCachedValueType()
        );
    }

    @ParameterizedTest
    @DisplayName("Test of range defined name array dimensions")
    @CsvSource(
            value = {
                    "A1:A3|D4|D4,D5,D6",
                    "A1:C1|D4|D4,E4,F4",
                    "A1:B3|D4|D4,D5,D6,E4,E5,E6",
                    "A1:A1|D4|D4"
            },
            delimiter = '|'
    )
    void rangeReferenceTest(String sourceRange, String target, String expectedAddresses) {
        Workbook workbook = new Workbook("Sheet1");
        DefinedName name = workbook.addDefinedNameRange("RangeName", workbook.getCurrentWorksheet(), sourceRange);
        workbook.getCurrentWorksheet().addCell("overwritten", "E5");

        List<Address> addresses = workbook.getCurrentWorksheet().addCellReference(name, target);
        String[] expected = expectedAddresses.split(",");
        assertEquals(expected.length, addresses.size());
        for (int i = 0; i < expected.length; i++) {
            assertEquals(new Address(expected[i]), addresses.get(i));
        }

        Cell master = workbook.getCurrentWorksheet().getCells().get(target);
        assertEquals(FormulaData.FormulaType.ARRAY, master.getFormula().getType());
        assertEquals(
                new Range(new Address(expected[0]), new Address(expected[expected.length - 1])).toString(),
                master.getFormula().getFormulaRange()
        );
        assertNull(master.getFormula().getMasterCellAddress());
        for (int i = 1; i < expected.length; i++) {
            Cell dependent = workbook.getCurrentWorksheet().getCells().get(expected[i]);
            assertEquals(Cell.CellType.FORMULA, dependent.getDataType());
            assertEquals(FormulaData.FormulaType.ARRAY, dependent.getFormula().getType());
            assertEquals(target, dependent.getFormula().getMasterCellAddress());
            assertEquals(master.getFormula().getFormulaRange(), dependent.getFormula().getFormulaRange());
            assertNull(dependent.getFormula().getDefinedNameReference());
        }
    }

    @Test
    @DisplayName("Test of a styled range defined name reference")
    void styledRangeReferenceTest() {
        Workbook workbook = new Workbook("Sheet1");
        DefinedName name = workbook.addDefinedNameRange("RangeName", workbook.getCurrentWorksheet(), "A1:B2");
        Style style = (Style) BasicStyles.getItalic().copy();
        List<Address> addresses = workbook.getCurrentWorksheet().addCellReference(name, "C3", style, "cached");

        assertEquals(4, addresses.size());
        for (Address address : addresses) {
            assertEquals(
                    style.hashCode(), workbook.getCurrentWorksheet().getCells().get(address.toString())
                            .getCellStyle().hashCode()
            );
        }
        assertEquals("cached", workbook.getCurrentWorksheet().getCells().get("C3").getFormula().getCachedValue());
        assertEquals(
                Cell.CellType.STRING,
                workbook.getCurrentWorksheet().getCells().get("C3").getFormula().getCachedValueType()
        );
    }

    @Test
    @DisplayName("Test of invalid addCellReference inputs")
    void invalidReferenceTest() {
        Workbook workbook = new Workbook("Sheet1");
        DefinedName name = workbook.addDefinedNameConstant("Name", 1);
        assertThrows(WorksheetException.class, () -> workbook.getCurrentWorksheet().addCellReference(null, "A1"));
        assertThrows(WorksheetException.class, () -> workbook.getCurrentWorksheet().addCellReference(null, 0, 0));
        assertThrows(
                WorksheetException.class,
                () -> workbook.getCurrentWorksheet().addCellReference(null, 0, 0, BasicStyles.getBold())
        );
        assertThrows(WorksheetException.class, () -> CorePackageTestAccess.setNullReference(new Cell()));
        assertThrows(FormatException.class, () -> workbook.getCurrentWorksheet().addCellReference(name, "invalid"));
        assertThrows(RangeException.class, () -> workbook.getCurrentWorksheet().addCellReference(name, -1, 0));
    }
}
