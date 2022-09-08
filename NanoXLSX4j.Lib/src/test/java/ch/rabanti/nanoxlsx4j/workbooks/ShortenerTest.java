package ch.rabanti.nanoxlsx4j.workbooks;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.util.HashMap;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

public class ShortenerTest {
    @DisplayName("Test of the SetCurrentWorksheet function")
    @Test()
    void setCurrentWorksheetTest() {
        Workbook workbook = new Workbook("Sheet1");
        Worksheet worksheet = new Worksheet("Sheet2");
        workbook.addWorksheet(worksheet);
        workbook.setCurrentWorksheet("Sheet1");
        assertEquals("Sheet1", workbook.getCurrentWorksheet().getSheetName());
        workbook.WS.setCurrentWorksheet(worksheet);
        assertEquals("Sheet2", workbook.getCurrentWorksheet().getSheetName());
    }


    @DisplayName("Test of the failing SetCurrentWorksheet function on a unreferenced worksheet")
    @Test()
    void setCurrentWorksheetFailTest() {
        Workbook workbook = new Workbook("Sheet1");
        Worksheet worksheet = new Worksheet("Sheet2");
        assertThrows(WorksheetException.class, () -> workbook.WS.setCurrentWorksheet(worksheet));
    }


    @DisplayName("Test of the failing SetCurrentWorksheet function on a null object")
    @Test()
    void setCurrentWorksheetFailTest2() {
        Workbook workbook = new Workbook("Sheet1");
        assertThrows(WorksheetException.class, () -> workbook.WS.setCurrentWorksheet(null));
    }


    @DisplayName("Test of the Value function")
    @Test()
    void valueTest() {
        Workbook workbook = new Workbook("Sheet1");
        Map<String, Object> values = new HashMap<String, Object>();
        values.put("A1", "Test");
        values.put("B1", 22);
        assertValue(workbook, workbook.WS::value, null, values, Worksheet.CellDirection.ColumnToColumn, 0, 0, 2, 0, null);

        values.clear();
        values.put("C3", "Test2");
        values.put("C4", 22.2);
        assertValue(workbook, workbook.WS::value, null, values, Worksheet.CellDirection.RowToRow, 2, 2, 2, 4, null);
    }


    @DisplayName("Test of the Value function with a style")
    @Test()
    void valueTest2() {
        Workbook workbook = new Workbook("Sheet1");
        Map<String, Object> values = new HashMap<String, Object>();
        values.put("A1", true);
        values.put("B1", "");
        Style style = BasicStyles.BoldItalic();
        assertValue(workbook, null, workbook.WS::value, values, Worksheet.CellDirection.ColumnToColumn, 0, 0, 2, 0, style);

        values.clear();
        values.put("C3", -22.3);
        values.put("C4", false);
        style = BasicStyles.DoubleUnderline();
        assertValue(workbook, null, workbook.WS::value, values, Worksheet.CellDirection.RowToRow, 2, 2, 2, 4, style);
    }


    @DisplayName("Test of the Formula function")
    @Test()
    void formulaTest() {
        Workbook workbook = new Workbook("Sheet1");
        Map<String, String> values = new HashMap<String, String>();
        values.put("A1", "=A3");
        values.put("B1", "=ROUNDDOWN(22.1)");
        assertValue(workbook, workbook.WS::formula, null, values, Worksheet.CellDirection.ColumnToColumn, 0, 0, 2, 0, null);

        values.clear();
        values.put("C3", "=C3");
        values.put("C4", "=ROUNDDOWN(11.1)");
        assertValue(workbook, workbook.WS::value, null, values, Worksheet.CellDirection.RowToRow, 2, 2, 2, 4, null);
    }


    @DisplayName("Test of the Formula function with a style")
    @Test()
    void formulaTest2() {
        Workbook workbook = new Workbook("Sheet1");
        Map<String, String> values = new HashMap<String, String>();
        values.put("A1", "=A3");
        values.put("B1", "=ROUNDDOWN(22.1)");
        Style style = BasicStyles.BoldItalic();
        assertValue(workbook, null, workbook.WS::formula, values, Worksheet.CellDirection.ColumnToColumn, 0, 0, 2, 0, style);

        values.clear();
        values.put("C3", "=C3");
        values.put("C4", "=ROUNDDOWN(11.1)");
        style = BasicStyles.DoubleUnderline();
        assertValue(workbook, null, workbook.WS::formula, values, Worksheet.CellDirection.RowToRow, 2, 2, 2, 4, style);
    }


    @DisplayName("Test of the Down function")
    @Test()
    void downTest() {
        Workbook workbook = new Workbook("Sheet1");
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        workbook.WS.down();
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(1, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }


    @DisplayName("Test of the Down function with a row number")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource({
            "0,0,0,0,0",
            "0, 0, 1, 0, 1",
            "5, 5, 5, 0, 10",
            "5, 5, -2, 0, 3",
            "5, 5, -5, 0, 0",
    })
    void downTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow) {
        Workbook workbook = new Workbook("Sheet1");
        assertJumpTo(workbook, workbook.WS::down, startColumn, startRow, number, expectedColumn, expectedRow);
    }

    @DisplayName("Test of the Down function with a row number and the option to keep the column position")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource({
            "A1, 0, false, A1",
            "A1, 0, true, A1",
            "A1, 1, false, A2",
            "A1, 1, true, A2",
            "C10, 1, false, A11",
            "C10, 1, true, C11",
            "R5, 5, false, A10",
            "R5, 5, true, R10",
            "F5, -3, false, A2",
            "F5, -3, true, F2",
            "F5, -4, false, A1",
            "F5, -4, true, F1",
    })
    void downTest3(String initialAddress, int number, boolean keepColumn, String expectedAdrress) {
        Workbook workbook = new Workbook("Sheet1");
        assertJumpKeep(workbook, workbook.WS::down, initialAddress, number, keepColumn, expectedAdrress);
    }

    @DisplayName("Test of the failing Down function with a negative row number")
    @Test()
    void downFailingTest() {
        Workbook workbook = new Workbook("Sheet1");
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        assertThrows(RangeException.class, () -> workbook.WS.down(-2));
    }


    @DisplayName("Test of the Up function")
    @Test()
    void upTest() {
        Workbook workbook = new Workbook("Sheet1");
        workbook.getCurrentWorksheet().setCurrentCellAddress("C4");
        assertEquals(2, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(3, workbook.getCurrentWorksheet().getCurrentRowNumber());
        workbook.WS.up();
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(2, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }


    @DisplayName("Test of the Up function with a row number")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource({
            "0, 0, 0, 0, 0",
            "1, 1, 1, 0, 0",
            "5, 10, 5, 0, 5",
            "5, 5, -2, 0, 7",
            "5, 5, 5, 0, 0",
    })
    void upTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow) {
        Workbook workbook = new Workbook("Sheet1");
        assertJumpTo(workbook, workbook.WS::up, startColumn, startRow, number, expectedColumn, expectedRow);
    }

    @DisplayName("Test of the Up function with a row number and the option to keep the column position")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource({
            "A1, 0, false, A1",
            "A1, 0, true, A1",
            "A2, 1, false, A1",
            "A2, 1, true, A1",
            "C10, 1, false, A9",
            "C10, 1, true, C9",
            "R10, 5, false, A5",
            "R10, 5, true, R5",
            "F5, -3, false, A8",
            "F5, -3, true, F8",
            "F5, 4, false, A1",
            "F5, 4, true, F1",
    })
    void upTest3(String initialAddress, int number, boolean keepColumn, String expectedAdrress) {
        Workbook workbook = new Workbook("Sheet1");
        assertJumpKeep(workbook, workbook.WS::up, initialAddress, number, keepColumn, expectedAdrress);
    }

    @DisplayName("Test of the failing Up function with a negative row number")
    @Test()
    void upFailingTest() {
        Workbook workbook = new Workbook("Sheet1");
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        assertThrows(RangeException.class, () -> workbook.WS.up(2));
    }


    @DisplayName("Test of the Right function")
    @Test()
    void rightTest() {
        Workbook workbook = new Workbook("Sheet1");
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        workbook.WS.right();
        assertEquals(1, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }


    @DisplayName("Test of the Right function with a column number")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource({
            "0, 0, 0, 0, 0",
            "0, 0, 1, 1, 0",
            "5, 5, 5, 10, 0",
            "5, 5, -2, 3, 0",
            "5, 5, -5, 0, 0",
    })
    void rightTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow) {
        Workbook workbook = new Workbook("Sheet1");
        assertJumpTo(workbook, workbook.WS::right, startColumn, startRow, number, expectedColumn, expectedRow);
    }

    @DisplayName("Test of the Right function with a column number and the option to keep the row position")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource({
            "A1, 0, false, A1",
            "A1, 0, true, A1",
            "A1, 1, false, B1",
            "A1, 1, true, B1",
            "C10, 1, false, D1",
            "C10, 1, true, D10",
            "R5, 5, false, W1",
            "R5, 5, true, W5",
            "F5, -3, false, C1",
            "F5, -3, true, C5",
            "F5, -5, false, A1",
            "F5, -5, true, A5",
    })
    void rightTest3(String initialAddress, int number, boolean keepColumn, String expectedAdrress) {
        Workbook workbook = new Workbook("Sheet1");
        assertJumpKeep(workbook, workbook.WS::right, initialAddress, number, keepColumn, expectedAdrress);
    }

    @DisplayName("Test of the failing Right function with a negative row number")
    @Test()
    void rightFailingTest() {
        Workbook workbook = new Workbook("Sheet1");
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        assertThrows(RangeException.class, () -> workbook.WS.right(-2));
    }


    @DisplayName("Test of the Left function")
    @Test()
    void leftTest() {
        Workbook workbook = new Workbook("Sheet1");
        workbook.getCurrentWorksheet().setCurrentCellAddress("D4");
        assertEquals(3, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(3, workbook.getCurrentWorksheet().getCurrentRowNumber());
        workbook.WS.left();
        assertEquals(2, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }


    @DisplayName("Test of the Left function with a column number")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource({
            "0, 0, 0, 0, 0",
            "1, 1, 1, 0, 0",
            "5, 5, 2, 3, 0",
            "5, 5, -2, 7, 0",
            "5, 5, 5, 0, 0",
    })
    void leftTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow) {
        Workbook workbook = new Workbook("Sheet1");
        assertJumpTo(workbook, workbook.WS::left, startColumn, startRow, number, expectedColumn, expectedRow);
    }

    @DisplayName("Test of the Left function with a column number and the option to keep the row position")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource({
            "A1, 0, false, A1",
            "A1, 0, true, A1",
            "B1, 1, false, A1",
            "B1, 1, true, A1",
            "C10, 1, false, B1",
            "C10, 1, true, B10",
            "R5, 5, false, M1",
            "R5, 5, true, M5",
            "F5, -3, false, I1",
            "F5, -3, true, I5",
            "F5, 5, false, A1",
            "F5, 5, true, A5",
    })
    void leftTest3(String initialAddress, int number, boolean keepColumn, String expectedAdrress) {
        Workbook workbook = new Workbook("Sheet1");
        assertJumpKeep(workbook, workbook.WS::left, initialAddress, number, keepColumn, expectedAdrress);
    }

    @DisplayName("Test of the failing Left function with a negative row number")
    @Test()
    void leftFailingTest() {
        Workbook workbook = new Workbook("Sheet1");
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        assertThrows(RangeException.class, () -> workbook.WS.left(2));
    }

    // For code coverage
    @DisplayName("Singular Test of the NullCheck method")
    @Test()
    void nullCheckTest() {
        Workbook workbook = new Workbook(); // No worksheet created
        assertThrows(WorksheetException.class, () -> workbook.WS.value(22));
    }

    private void assertJumpTo(Workbook workbook, Consumer<Integer> consumer, int startColumn, int startRow, int number, int expectedColumn, int expectedRow) {
        workbook.getCurrentWorksheet().setCurrentColumnNumber(startColumn);
        workbook.getCurrentWorksheet().setCurrentRowNumber(startRow);
        assertEquals(startColumn, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(startRow, workbook.getCurrentWorksheet().getCurrentRowNumber());
        consumer.accept(number);
        assertEquals(expectedColumn, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(expectedRow, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }

    private void assertJumpKeep(Workbook workbook, BiConsumer<Integer, Boolean> consumer, String initialAddress, int number, boolean keepOther, String expectedAddress) {
        Address initial = new Address(initialAddress);
        workbook.getCurrentWorksheet().setCurrentCellAddress(initial.Column, initial.Row);

        assertEquals(initial.Column, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(initial.Row, workbook.getCurrentWorksheet().getCurrentRowNumber());
        consumer.accept(number, keepOther);
        Address expected = new Address(expectedAddress);
        assertEquals(expected.Column, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(expected.Row, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }

    private <T> void assertValue(Workbook workbook, Consumer<T> consumer, BiConsumer<T, Style> styleConsumer, Map<String, T> values, Worksheet.CellDirection direction, int startColumn, int startRow, int expectedEndColumn, int expectedEndRow, Style style) {
        workbook.getCurrentWorksheet().setCurrentColumnNumber(startColumn);
        workbook.getCurrentWorksheet().setCurrentRowNumber(startRow);
        workbook.getCurrentWorksheet().setCurrentCellDirection(direction);

        for (Map.Entry<String, T> cell : values.entrySet()) {
            if (style == null) {
                consumer.accept(cell.getValue());
            } else {
                styleConsumer.accept(cell.getValue(), style);
            }
        }

        for (Map.Entry<String, T> cell : values.entrySet()) {
            Address address = new Address(cell.getKey());
            T value = (T) workbook.getCurrentWorksheet().getCell(address).getValue();
            assertEquals(cell.getValue(), value);
            if (style != null) {
                assertEquals(style.hashCode(), workbook.getCurrentWorksheet().getCell(address).getCellStyle().hashCode());
            }
        }
        assertEquals(expectedEndColumn, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(expectedEndRow, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }


}
