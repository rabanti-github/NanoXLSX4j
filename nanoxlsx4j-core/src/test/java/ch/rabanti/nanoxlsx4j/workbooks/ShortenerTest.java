package ch.rabanti.nanoxlsx4j.workbooks;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Shortener;
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

public class ShortenerTest {

    @Test
    @DisplayName("Test of the setCurrentWorksheet function")
    void setCurrentWorksheetTest() {
        Workbook workbook = new Workbook("Sheet1");
        workbook.addWorksheet("Sheet2");
        Worksheet worksheet = workbook.getWorksheets().get(1);
        Shortener shortener = new Shortener(workbook);
        workbook.setCurrentWorksheet(workbook.getWorksheets().get(0));
        assertEquals("Sheet1", workbook.getCurrentWorksheet().getSheetName());
        shortener.SetCurrentWorksheet(worksheet);
        assertEquals("Sheet2", workbook.getCurrentWorksheet().getSheetName());
    }

    @Test
    @DisplayName("Test of the failing setCurrentWorksheet function on a unreferenced worksheet")
    void setCurrentWorksheetFailTest() {
        Workbook workbook = new Workbook("Sheet1");
        Worksheet worksheet = new Worksheet("Sheet2");
        Shortener shortener = new Shortener(workbook);
        assertThrows(WorksheetException.class, () -> shortener.SetCurrentWorksheet(worksheet));
    }

    @Test
    @DisplayName("Test of the failing setCurrentWorksheet function on a null object")
    void setCurrentWorksheetFailTest2() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertThrows(WorksheetException.class, () -> shortener.SetCurrentWorksheet(null));
    }

    @Test
    @DisplayName("Test of the value function")
    void valueTest() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        Map<String, Object> values = new LinkedHashMap<>();
        values.put("A1", "Test");
        values.put("B1", 22);
        assertValue(
                workbook,
                shortener::value,
                null,
                values,
                Worksheet.CellDirection.COLUMN_TO_COLUMN,
                0,
                0,
                2,
                0,
                null
        );

        values.clear();
        values.put("C3", "Test2");
        values.put("C4", 22.2);
        assertValue(
                workbook,
                shortener::value,
                null,
                values,
                Worksheet.CellDirection.ROW_TO_ROW,
                2,
                2,
                2,
                4,
                null
        );
    }

    @Test
    @DisplayName("Test of the value function with a style")
    void valueTest2() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        Map<String, Object> values = new LinkedHashMap<>();
        values.put("A1", true);
        values.put("B1", "");
        Style style = BasicStyles.getBoldItalic();
        assertValue(
                workbook,
                null,
                shortener::value,
                values,
                Worksheet.CellDirection.COLUMN_TO_COLUMN,
                0,
                0,
                2,
                0,
                style
        );

        values.clear();
        values.put("C3", -22.3);
        values.put("C4", false);
        style = BasicStyles.getDoubleUnderline();
        assertValue(
                workbook,
                null,
                shortener::value,
                values,
                Worksheet.CellDirection.ROW_TO_ROW,
                2,
                2,
                2,
                4,
                style
        );
    }

    @Test
    @DisplayName("Test of the formula function")
    void formulaTest() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        Map<String, String> values = new LinkedHashMap<>();
        values.put("A1", "=A3");
        values.put("B1", "=ROUNDDOWN(22.1)");
        assertValue(
                workbook,
                shortener::formula,
                null,
                values,
                Worksheet.CellDirection.COLUMN_TO_COLUMN,
                0,
                0,
                2,
                0,
                null
        );

        values.clear();
        values.put("C3", "=C3");
        values.put("C4", "=ROUNDDOWN(11.1)");
        assertValue(
                workbook,
                shortener::value,
                null,
                values,
                Worksheet.CellDirection.ROW_TO_ROW,
                2,
                2,
                2,
                4,
                null
        );
    }

    @Test
    @DisplayName("Test of the formula function with a style")
    void formulaTest2() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        Map<String, String> values = new LinkedHashMap<>();
        values.put("A1", "=A3");
        values.put("B1", "=ROUNDDOWN(22.1)");
        Style style = BasicStyles.getBoldItalic();
        assertValue(
                workbook,
                null,
                shortener::formula,
                values,
                Worksheet.CellDirection.COLUMN_TO_COLUMN,
                0,
                0,
                2,
                0,
                style
        );

        values.clear();
        values.put("C3", "=C3");
        values.put("C4", "=ROUNDDOWN(11.1)");
        style = BasicStyles.getDoubleUnderline();
        assertValue(
                workbook,
                null,
                shortener::formula,
                values,
                Worksheet.CellDirection.ROW_TO_ROW,
                2,
                2,
                2,
                4,
                style
        );
    }

    @Test
    @DisplayName("Test of the down function")
    void downTest() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        shortener.down();
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(1, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }

    @ParameterizedTest
    @DisplayName("Test of the down function with a row number")
    @CsvSource({
        "0, 0, 0, 0, 0",
        "0, 0, 1, 0, 1",
        "5, 5, 5, 0, 10",
        "5, 5, -2, 0, 3",
        "5, 5, -5, 0, 0"
    })
    void downTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow) {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertJumpTo(workbook, shortener::down, startColumn, startRow, number, expectedColumn, expectedRow);
    }

    @ParameterizedTest
    @DisplayName("Test of the down function with a row number and the option to keep the column position")
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
        "F5, -4, true, F1"
    })
    void downTest3(String initialAddress, int number, boolean keepColumn, String expectedAddress) {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertJumpKeep(workbook, shortener::down, initialAddress, number, keepColumn, expectedAddress);
    }

    @Test
    @DisplayName("Test of the failing down function with a negative row number")
    void downFailingTest() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        assertThrows(RangeException.class, () -> shortener.down(-2));
    }

    @Test
    @DisplayName("Test of the up function")
    void upTest() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        workbook.getCurrentWorksheet().setCurrentCellAddress("C4");
        assertEquals(2, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(3, workbook.getCurrentWorksheet().getCurrentRowNumber());
        shortener.up();
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(2, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }

    @ParameterizedTest
    @DisplayName("Test of the up function with a row number")
    @CsvSource({
        "0, 0, 0, 0, 0",
        "1, 1, 1, 0, 0",
        "5, 10, 5, 0, 5",
        "5, 5, -2, 0, 7",
        "5, 5, 5, 0, 0"
    })
    void upTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow) {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertJumpTo(workbook, shortener::up, startColumn, startRow, number, expectedColumn, expectedRow);
    }

    @ParameterizedTest
    @DisplayName("Test of the up function with a row number and the option to keep the column position")
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
        "F5, 4, true, F1"
    })
    void upTest3(String initialAddress, int number, boolean keepColumn, String expectedAddress) {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertJumpKeep(workbook, shortener::up, initialAddress, number, keepColumn, expectedAddress);
    }

    @Test
    @DisplayName("Test of the failing up function with a negative row number")
    void upFailingTest() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        assertThrows(RangeException.class, () -> shortener.up(2));
    }

    @Test
    @DisplayName("Test of the right function")
    void rightTest() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        shortener.right();
        assertEquals(1, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }

    @ParameterizedTest
    @DisplayName("Test of the right function with a column number")
    @CsvSource({
        "0, 0, 0, 0, 0",
        "0, 0, 1, 1, 0",
        "5, 5, 5, 10, 0",
        "5, 5, -2, 3, 0",
        "5, 5, -5, 0, 0"
    })
    void rightTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow) {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertJumpTo(workbook, shortener::right, startColumn, startRow, number, expectedColumn, expectedRow);
    }

    @ParameterizedTest
    @DisplayName("Test of the right function with a column number and the option to keep the row position")
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
        "F5, -5, true, A5"
    })
    void rightTest3(String initialAddress, int number, boolean keepRow, String expectedAddress) {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertJumpKeep(workbook, shortener::right, initialAddress, number, keepRow, expectedAddress);
    }

    @Test
    @DisplayName("Test of the failing right function with a negative row number")
    void rightFailingTest() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        assertThrows(RangeException.class, () -> shortener.right(-2));
    }

    @Test
    @DisplayName("Test of the left function")
    void leftTest() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        workbook.getCurrentWorksheet().setCurrentCellAddress("D4");
        assertEquals(3, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(3, workbook.getCurrentWorksheet().getCurrentRowNumber());
        shortener.left();
        assertEquals(2, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }

    @ParameterizedTest
    @DisplayName("Test of the left function with a column number")
    @CsvSource({
        "0, 0, 0, 0, 0",
        "1, 1, 1, 0, 0",
        "5, 5, 2, 3, 0",
        "5, 5, -2, 7, 0",
        "5, 5, 5, 0, 0"
    })
    void leftTest2(int startColumn, int startRow, int number, int expectedColumn, int expectedRow) {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertJumpTo(workbook, shortener::left, startColumn, startRow, number, expectedColumn, expectedRow);
    }

    @ParameterizedTest
    @DisplayName("Test of the left function with a column number and the option to keep the row position")
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
        "F5, 5, true, A5"
    })
    void leftTest3(String initialAddress, int number, boolean keepRow, String expectedAddress) {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertJumpKeep(workbook, shortener::left, initialAddress, number, keepRow, expectedAddress);
    }

    @Test
    @DisplayName("Test of the failing left function with a negative row number")
    void leftFailingTest() {
        Workbook workbook = new Workbook("Sheet1");
        Shortener shortener = new Shortener(workbook);
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(0, workbook.getCurrentWorksheet().getCurrentRowNumber());
        assertThrows(RangeException.class, () -> shortener.left(2));
    }

    // For code coverage --> TODO Check whether this is still applicable
   // @Test
   // @DisplayName("Singular Test of the nullCheck method")
   // void nullCheckTest() {
   //     Workbook workbook = new Workbook("Sheet1");
   //     workbook.setCurrentWorksheet(null);
   //     Shortener shortener = new Shortener(workbook);
   //     assertThrows(WorksheetException.class, () -> shortener.value(22));
   // }

    private void assertJumpTo(
            Workbook workbook,
            Consumer<Integer> action,
            int startColumn,
            int startRow,
            int number,
            int expectedColumn,
            int expectedRow) {
        workbook.getCurrentWorksheet().setCurrentColumnNumber(startColumn);
        workbook.getCurrentWorksheet().setCurrentRowNumber(startRow);
        assertEquals(startColumn, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(startRow, workbook.getCurrentWorksheet().getCurrentRowNumber());
        action.accept(number);
        assertEquals(expectedColumn, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(expectedRow, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }

    private void assertJumpKeep(
            Workbook workbook,
            BiConsumer<Integer, Boolean> action,
            String initialAddress,
            int number,
            boolean keepOther,
            String expectedAddress) {
        Address initial = new Address(initialAddress);
        workbook.getCurrentWorksheet().setCurrentCellAddress(initial.column(), initial.row());

        assertEquals(initial.column(), workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(initial.row(), workbook.getCurrentWorksheet().getCurrentRowNumber());
        action.accept(number, keepOther);
        Address expected = new Address(expectedAddress);
        assertEquals(expected.column(), workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(expected.row(), workbook.getCurrentWorksheet().getCurrentRowNumber());
    }

    @SuppressWarnings("unchecked")
    private <T> void assertValue(
            Workbook workbook,
            Consumer<T> action,
            BiConsumer<T, Style> styleAction,
            Map<String, T> values,
            Worksheet.CellDirection direction,
            int startColumn,
            int startRow,
            int expectedEndColumn,
            int expectedEndRow,
            Style style) {
        workbook.getCurrentWorksheet().setCurrentColumnNumber(startColumn);
        workbook.getCurrentWorksheet().setCurrentRowNumber(startRow);
        workbook.getCurrentWorksheet().setCurrentCellDirection(direction);

        for (Map.Entry<String, T> cell : values.entrySet()) {
            if (style == null) {
                action.accept(cell.getValue());
            } else {
                styleAction.accept(cell.getValue(), style);
            }
        }

        for (Map.Entry<String, T> cell : values.entrySet()) {
            Address address = new Address(cell.getKey());
            T value = (T) workbook.getCurrentWorksheet().getCell(address).getValue();
            assertEquals(cell.getValue(), value);
            if (style != null) {
                assertEquals(
                        style.hashCode(),
                        workbook.getCurrentWorksheet().getCell(address).getCellStyle().hashCode()
                );
            }
        }
        assertEquals(expectedEndColumn, workbook.getCurrentWorksheet().getCurrentColumnNumber());
        assertEquals(expectedEndRow, workbook.getCurrentWorksheet().getCurrentRowNumber());
    }
}
