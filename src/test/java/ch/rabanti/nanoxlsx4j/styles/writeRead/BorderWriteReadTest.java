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
            "'FFAACC00', 'STRING', 'test', true, true",
            "'FFAADD00', 'FLOAT', '0.5f', true, false",
            "'FFDDCC00', 'BOOLEAN', true, false, true",
            "'FFAACCDD', 'NULL', '', false, false",
    })
    void diagonalColorTest(String color, String sourceType, String sourceValue, boolean diagonalUp, boolean diagonalDown) {
        Object value = TestUtils.createInstance(sourceType, sourceValue);
        Style style = new Style();
        style.getBorder().setDiagonalColor(color);
        style.getBorder().setDiagonalStyle(Border.StyleValue.dashDot);
        style.getBorder().setDiagonalUp(diagonalUp);
        style.getBorder().setDiagonalDown(diagonalDown);

        Cell cell = createWorkbook(value, style);

        assertEquals(color, cell.getCellStyle().getBorder().getDiagonalColor());
        assertEquals(Border.StyleValue.dashDot, cell.getCellStyle().getBorder().getDiagonalStyle());
        assertEquals(diagonalUp, cell.getCellStyle().getBorder().isDiagonalUp());
        assertEquals(diagonalDown, cell.getCellStyle().getBorder().isDiagonalDown());
    }

    @DisplayName("Test of the writing and reading of the top color style value")
    @ParameterizedTest(name = "Given value {0} should lead to a cell with this top style")
    @CsvSource({
            "FFAACC00, test",
            "FFAADD00, 0.5f",
            "FFDDCC00, true",
            "FFAACCDD, null",
    })
    void topColorTest(String color, Object value)
    {
        Style style = new Style();
        style.getBorder().setTopColor(color);
        style.getBorder().setTopStyle(Border.StyleValue.s_double);

        Cell cell = createWorkbook(value, style);

        assertEquals(color, cell.getCellStyle().getBorder().getTopColor());
        assertEquals(Border.StyleValue.s_double, cell.getCellStyle().getBorder().getTopStyle());
    }

    @DisplayName("Test of the writing and reading of the bottom color style value")
    @ParameterizedTest(name = "Given value {0} should lead to a cell with this bottom style")
    @CsvSource({
            "FFAACC00, test",
            "FFAADD00, 0.5f",
            "FFDDCC00, true",
            "FFAACCDD, null",
    })
    void bottomColorTest(String color, Object value)
    {
        Style style = new Style();
        style.getBorder().setBottomColor(color);
        style.getBorder().setBottomStyle(Border.StyleValue.thin);

        Cell cell = createWorkbook(value, style);

        assertEquals(color, cell.getCellStyle().getBorder().getBottomColor());
        assertEquals(Border.StyleValue.thin, cell.getCellStyle().getBorder().getBottomStyle());
    }

    @DisplayName("Test of the writing and reading of the left color style value")
    @ParameterizedTest(name = "Given value {0} should lead to a cell with this left style")
    @CsvSource({
            "FFAACC00, test",
            "FFAADD00, 0.5f",
            "FFDDCC00, true",
            "FFAACCDD, null",
    })
    void leftColorTest(String color, Object value)
    {
        Style style = new Style();
        style.getBorder().setLeftColor(color);
        style.getBorder().setLeftStyle(Border.StyleValue.dashDotDot);

        Cell cell = createWorkbook(value, style);

        assertEquals(color, cell.getCellStyle().getBorder().getLeftColor());
        assertEquals(Border.StyleValue.dashDotDot, cell.getCellStyle().getBorder().getLeftStyle());
    }

    @DisplayName("Test of the writing and reading of the right color style value")
    @ParameterizedTest(name = "Given value {0} should lead to a cell with this right style")
    @CsvSource({
            "FFAACC00, test",
            "FFAADD00, 0.5f",
            "FFDDCC00, true",
            "FFAACCDD, null",
    })
    void rightColorTest(String color, Object value)
    {
        Style style = new Style();
        style.getBorder().setRightColor(color);
        style.getBorder().setRightStyle(Border.StyleValue.dashed);

        Cell cell = createWorkbook(value, style);

        assertEquals(color, cell.getCellStyle().getBorder().getRightColor());
        assertEquals(Border.StyleValue.dashed, cell.getCellStyle().getBorder().getRightStyle());
    }

    @DisplayName("Test of the writing and reading of border style value")
    @ParameterizedTest(name = "Given value {0} should lead to a cell with a border style of this vale")
    @CsvSource({
            "dashDotDot",
            "dashDot",
            "dashed",
            "dotted",
            "hair",
            "medium",
            "mediumDashDot",
            "mediumDashDotDot",
            "mediumDashed",
            "slantDashDot",
            "thin",
            "s_double",
            "thick",
            "none",
    })
    void borderStyleTest(Border.StyleValue styleValue)
    {
        Style style = new Style();
        style.getBorder().setRightStyle(styleValue);

        Cell cell = createWorkbook("test", style);

        assertEquals(styleValue, cell.getCellStyle().getBorder().getRightStyle());
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
