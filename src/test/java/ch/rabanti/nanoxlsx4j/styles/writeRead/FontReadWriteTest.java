package ch.rabanti.nanoxlsx4j.styles.writeRead;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.styles.Font;
import ch.rabanti.nanoxlsx4j.styles.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class FontReadWriteTest {

    @DisplayName("Test of the writing and reading of the bold font style value")
    @ParameterizedTest(name = "Given font state for bold '{0}' with value {1} should lead to a cell with this bold state")
    @CsvSource({
            "true, test",
            "false, 0.5f",
    })
    void boldFontTest(boolean styleValue,Object value)
    {
        Style style = new Style();
        style.getFont().setBold(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().isBold());
    }

    @DisplayName("Test of the writing and reading of the italic font style value")
    @ParameterizedTest(name = "Given font state for italic '{0}' with value {1} should lead to a cell with this italic state")
    @CsvSource({
            "true, test",
            "false, 0.5f",
    })
    void italicFontTest(boolean styleValue,Object value)
    {
        Style style = new Style();
        style.getFont().setItalic(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().isItalic());
    }

    @DisplayName("Test of the writing and reading of the strike font style value")
    @ParameterizedTest(name = "Given font state for strike '{0}' with value {1} should lead to a cell with this strike state")
    @CsvSource({
            "true, test",
            "false, 0.5f",
    })
    void strikeFontTest(boolean styleValue,Object value)
    {
        Style style = new Style();
        style.getFont().setStrike(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().isStrike());
    }

    @DisplayName("Test of the writing and reading of the underline font style value")
    @ParameterizedTest(name = "Given font state '{0}' for underline with value {1} should lead to a cell with this underline state")
    @CsvSource({
            "u_single, test",
            "u_double, 0.5f",
            "doubleAccounting, true",
            "singleAccounting, 42",
            "none, ''",
    })
    void underlineFontTest(Font.UnderlineValue styleValue, Object value)
    {
        Style style = new Style();
        style.getFont().setUnderline(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getUnderline());
    }

    @DisplayName("Test of the writing and reading of the vertical alignment font style value")
    @ParameterizedTest(name = "Given font state '{0}' for vertical align with value {1} should lead to a cell with this vertical align state")
    @CsvSource({
            "subscript, test",
            "superscript, 0.5f",
            "none, true",
    })
    void verticalAlignFontTest(Font.VerticalAlignValue styleValue, Object value)
    {
        Style style = new Style();
        style.getFont().setVerticalAlign(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getVerticalAlign());
    }

    @DisplayName("Test of the writing and reading of the size font style value")
    @ParameterizedTest(name = "Given font size '{0}' with value {1} should lead to a cell with this size")
    @CsvSource({
            "10.5f, test",
            "11f, 0.5f",
            "50.55f, true",
    })
    void sizeFontTest(float styleValue, Object value)
    {
        Style style = new Style();
        style.getFont().setSize(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getSize());
    }



}
