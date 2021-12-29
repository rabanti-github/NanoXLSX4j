package ch.rabanti.nanoxlsx4j.styles.writeRead;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.styles.Border;
import ch.rabanti.nanoxlsx4j.styles.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;

import static org.junit.jupiter.api.Assertions.*;

public class BorderWriteReadTest {

    @Test
    void timeTest() {
        String pattern = "n.HH:mm:ss";
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(pattern);

        String time = "23547.12:13:15";
        TemporalAccessor d = null;
        try {
            d = formatter.parse(time);
        } catch (Exception e) {
            e.printStackTrace();
        }
        //Duration d = Duration.parse(time);
        int x = 1;
    }

    @DisplayName("Test of the writing and reading of the diagonal color style value")
    @ParameterizedTest(name = "Given value {0} should lead to a cell with this diagonal color style")
    @CsvSource({
            "'FFAACC00', 'STRING', 'test'",
            "'FFAADD00', 'FLOAT', '0.5f'",
            "'FFDDCC00', 'BOOLEAN', 'true'",
            "'FFAACCDD', 'NULL', ''",
    })
    void diagonalColorTest(String color, String sourceType, String sourceValue) {
        Object value = TestUtils.createInstance(sourceType, sourceValue);
        Style style = new Style();
        style.getBorder().setDiagonalColor(color);
        style.getBorder().setDiagonalStyle(Border.StyleValue.dashDot);
        style.getBorder().setDiagonalUp(true);

        Cell cell = createWorkbook(value, style);

        assertEquals(color, cell.getCellStyle().getBorder().getDiagonalColor());
        assertEquals(Border.StyleValue.dashDot, cell.getCellStyle().getBorder().getDiagonalStyle());
        assertTrue(cell.getCellStyle().getBorder().isDiagonalUp());
        assertFalse(cell.getCellStyle().getBorder().isDiagonalDown());
    }

    private static Cell createWorkbook(Object value, Style style) {
        try {
            Workbook workbook = new Workbook(false);
            workbook.addWorksheet("sheet1");
            workbook.getCurrentWorksheet().addCell(value, "A1", style);

            Workbook givenWorkbook = TestUtils.saveAndLoadWorkbook(workbook, null);

            Cell cell = givenWorkbook.getCurrentWorksheet().getCell(new Address("A1"));
            assertEquals(value, cell.getValue());
            return cell;
        } catch (Exception ex) {
            return null;
        }
    }


}
