package ch.rabanti.nanoxlsx4j.styles.writeRead;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.styles.Border;
import ch.rabanti.nanoxlsx4j.styles.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class BorderWriteReadTest {

    public enum BorderDirection {
        Diagonal,
        Left,
        Right,
        Top,
        Bottom
    }

    @DisplayName("Test of the 'diagonal' value when writing and reading a Border style")
    @ParameterizedTest(name = "Given value {0} should lead to a cell with this diagonal color style")
    @CsvSource(
            {
                    "'FFAACC00', 'STRING', 'test', true, true",
                    "'FFAADD00', 'FLOAT', '0.5f', true, false",
                    "'FFDDCC00', 'BOOLEAN', true, false, true",
                    "'FFAACCDD', 'NULL', '', false, false",
                    "'', 'INTEGER', '22', true, true"}
    )
    void diagonalColorTest(String color, String sourceType, String sourceValue, boolean diagonalUp, boolean diagonalDown) {
        Object value = TestUtils.createInstance(sourceType, sourceValue);
        Style style = new Style();
        style.getBorder().setDiagonalColor(color);
        style.getBorder().setDiagonalStyle(Border.StyleValue.dashDot);
        style.getBorder().setDiagonalUp(diagonalUp);
        style.getBorder().setDiagonalDown(diagonalDown);

        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");

        assertEquals(color, cell.getCellStyle().getBorder().getDiagonalColor());
        assertEquals(Border.StyleValue.dashDot, cell.getCellStyle().getBorder().getDiagonalStyle());
        assertEquals(diagonalUp, cell.getCellStyle().getBorder().isDiagonalUp());
        assertEquals(diagonalDown, cell.getCellStyle().getBorder().isDiagonalDown());
    }

    @DisplayName("Test of the 'top' value when writing and reading a Border style")
    @ParameterizedTest(name = "Given value {0} should lead to a cell with this top style")
    @CsvSource(
            {
                    "FFAACC00, STRING, test",
                    "FFAADD00, FLOAT, 0.5f",
                    "FFDDCC00, BOOLEAN, true",
                    "FFAACCDD, NULL, null",
                    "'', INTEGER, 22"}
    )
    void topColorTest(String color, String sourceType, String sourceValue) {
        Style style = new Style();
        style.getBorder().setTopColor(color);
        style.getBorder().setTopStyle(Border.StyleValue.s_double);

        Object value = TestUtils.createInstance(sourceType, sourceValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");

        assertEquals(color, cell.getCellStyle().getBorder().getTopColor());
        assertEquals(Border.StyleValue.s_double, cell.getCellStyle().getBorder().getTopStyle());
    }

    @DisplayName("Test of the 'bottom' value when writing and reading a Border style")
    @ParameterizedTest(name = "Given value {0} should lead to a cell with this bottom style")
    @CsvSource(
            {
                    "FFAACC00, STRING, test",
                    "FFAADD00, FLOAT, 0.5f",
                    "FFDDCC00, BOOLEAN, true",
                    "FFAACCDD, NULL, null",
                    "'', INTEGER, 22"}
    )
    void bottomColorTest(String color, String sourceType, String sourceValue) {
        Style style = new Style();
        style.getBorder().setBottomColor(color);
        style.getBorder().setBottomStyle(Border.StyleValue.thin);

        Object value = TestUtils.createInstance(sourceType, sourceValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");

        assertEquals(color, cell.getCellStyle().getBorder().getBottomColor());
        assertEquals(Border.StyleValue.thin, cell.getCellStyle().getBorder().getBottomStyle());
    }

    @DisplayName("Test of the 'left' value when writing and reading a Border style")
    @ParameterizedTest(name = "Given value {0} should lead to a cell with this left style")
    @CsvSource(
            {
                    "FFAACC00, STRING, test",
                    "FFAADD00, FLOAT, 0.5f",
                    "FFDDCC00, BOOLEAN, true",
                    "FFAACCDD, NULL, null",
                    "'', INTEGER, 22"}
    )
    void leftColorTest(String color, String sourceType, String sourceValue) {
        Style style = new Style();
        style.getBorder().setLeftColor(color);
        style.getBorder().setLeftStyle(Border.StyleValue.dashDotDot);

        Object value = TestUtils.createInstance(sourceType, sourceValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");

        assertEquals(color, cell.getCellStyle().getBorder().getLeftColor());
        assertEquals(Border.StyleValue.dashDotDot, cell.getCellStyle().getBorder().getLeftStyle());
    }

    @DisplayName("Test of the 'right' value when writing and reading a Border style")
    @ParameterizedTest(name = "Given value {0} should lead to a cell with this right style")
    @CsvSource(
            {
                    "FFAACC00, STRING, test",
                    "FFAADD00, FLOAT, 0.5f",
                    "FFDDCC00, BOOLEAN, true",
                    "FFAACCDD, NULL, null",
                    "'', INTEGER, 22"}
    )
    void rightColorTest(String color, String sourceType, String sourceValue) {
        Style style = new Style();
        style.getBorder().setRightColor(color);
        style.getBorder().setRightStyle(Border.StyleValue.dashed);

        Object value = TestUtils.createInstance(sourceType, sourceValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");

        assertEquals(color, cell.getCellStyle().getBorder().getRightColor());
        assertEquals(Border.StyleValue.dashed, cell.getCellStyle().getBorder().getRightStyle());
    }

    @DisplayName("Test of the 'styleValue' property when writing and reading a Font style")
    @ParameterizedTest(name = "Given value {0} should lead on the to a cell with a border style (direction {1}) of this vale")
    @CsvSource(
            {

                    "dashDotDot, Bottom",
                    "dashDot, Top",
                    "dashed, Left",
                    "dotted, Right",
                    "hair, Diagonal",
                    "medium, Bottom",
                    "mediumDashDot, Top",
                    "mediumDashDotDot, Left",
                    "mediumDashed, Right",
                    "slantDashDot, Diagonal",
                    "thin, Bottom",
                    "s_double, Top",
                    "thick, Left",
                    "none, Right",}
    )
    void borderStyleTest(Border.StyleValue styleValue, BorderDirection direction) {
        Style style = new Style();
        switch (direction) {
            case Diagonal:
                style.getBorder().setDiagonalStyle(styleValue);
                break;
            case Left:
                style.getBorder().setLeftStyle(styleValue);
                break;
            case Right:
                style.getBorder().setRightStyle(styleValue);
                break;
            case Top:
                style.getBorder().setTopStyle(styleValue);
                break;
            case Bottom:
                style.getBorder().setBottomStyle(styleValue);
                break;
        }
        Cell cell = TestUtils.saveAndReadStyledCell("test", style, "A1");
        switch (direction) {

            case Diagonal:
                assertEquals(styleValue, cell.getCellStyle().getBorder().getDiagonalStyle());
                break;
            case Left:
                assertEquals(styleValue, cell.getCellStyle().getBorder().getLeftStyle());
                break;
            case Right:
                assertEquals(styleValue, cell.getCellStyle().getBorder().getRightStyle());
                break;
            case Top:
                assertEquals(styleValue, cell.getCellStyle().getBorder().getTopStyle());
                break;
            case Bottom:
                assertEquals(styleValue, cell.getCellStyle().getBorder().getBottomStyle());
                break;
        }
    }

}
