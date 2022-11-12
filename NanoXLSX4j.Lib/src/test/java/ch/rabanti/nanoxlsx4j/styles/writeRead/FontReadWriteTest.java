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

    @DisplayName("Test of the 'bold' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font state for bold '{0}' with value {1} should lead to a cell with this bold state")
    @CsvSource(
            {
                    "true, test",
                    "false, 0.5f",}
    )
    void boldFontTest(boolean styleValue, Object value) {
        Style style = new Style();
        style.getFont().setBold(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().isBold());
    }

    @DisplayName("Test of the 'italic' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font state for italic '{0}' with value {1} should lead to a cell with this italic state")
    @CsvSource(
            {
                    "true, test",
                    "false, 0.5f",}
    )
    void italicFontTest(boolean styleValue, Object value) {
        Style style = new Style();
        style.getFont().setItalic(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().isItalic());
    }

    @DisplayName("Test of the 'strike' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font state for strike '{0}' with value {1} should lead to a cell with this strike state")
    @CsvSource(
            {
                    "true, test",
                    "false, 0.5f",}
    )
    void strikeFontTest(boolean styleValue, Object value) {
        Style style = new Style();
        style.getFont().setStrike(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().isStrike());
    }

    @DisplayName("Test of the 'underline' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font state '{0}' for underline with value {1} should lead to a cell with this underline state")
    @CsvSource(
            {
                    "u_single, test",
                    "u_double, 0.5f",
                    "doubleAccounting, true",
                    "singleAccounting, 42",
                    "none, ''",}
    )
    void underlineFontTest(Font.UnderlineValue styleValue, Object value) {
        Style style = new Style();
        style.getFont().setUnderline(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getUnderline());
    }

    @DisplayName("Test of the 'vertical alignment' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font state '{0}' for vertical align with value {1} should lead to a cell with this vertical align state")
    @CsvSource(
            {
                    "subscript, test",
                    "superscript, 0.5f",
                    "none, true",}
    )
    void verticalAlignFontTest(Font.VerticalAlignValue styleValue, Object value) {
        Style style = new Style();
        style.getFont().setVerticalAlign(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getVerticalAlign());
    }

    @DisplayName("Test of the 'size' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font size '{0}' with value {1} should lead to a cell with this size")
    @CsvSource(
            {
                    "10.5f, test",
                    "11f, 0.5f",
                    "50.55f, true",}
    )
    void sizeFontTest(float styleValue, Object value) {
        Style style = new Style();
        style.getFont().setSize(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getSize());
    }

    @DisplayName("Test of the 'theme' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font theme '{0}' with value {1} should lead to a cell with this theme")
    @CsvSource(
            {
                    "1, test",
                    "2, 0.5f",
                    "64, true",}
    )
    void sizeFontTest(int styleValue, Object value) {
        Style style = new Style();
        style.getFont().setColorTheme(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getColorTheme());
    }

    @DisplayName("Test of the 'colorValue' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font colorValue '{0}' with value {1} should lead to a cell with this colorValue")
    @CsvSource(
            {
                    "'FFAABBCC', test",
                    "'', 0.5f",}
    )
    void colorValueFontTest(String styleValue, Object value) {
        Style style = new Style();
        style.getFont().setColorValue(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getColorValue());
    }

    @DisplayName("Test of the 'name' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font name '{0}' with value {1} should lead to a cell with this name")
    @CsvSource(
            {
                    "' ', test",
                    "'test', 0.5f",}
    )
    void nameFontTest(String styleValue, Object value) {
        Style style = new Style();
        style.getFont().setName(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getName());
    }

    @DisplayName("Test of the 'family' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font family '{0}' with value {1} should lead to a cell with this family")
    @CsvSource(
            {
                    "' ', test",
                    "'test', 0.5f",
                    "'', true",}
    )
    void familyFontTest(String styleValue, Object value) {
        Style style = new Style();
        style.getFont().setFamily(styleValue);
        style.getFont().setSize(22);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getFamily());
    }

    @DisplayName("Test of the 'scheme' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font scheme '{0}' with value {1} should lead to a cell with this scheme")
    @CsvSource(
            {
                    "minor, test",
                    "major, 0.5f",}
    )
    void schemeFontTest(Font.SchemeValue styleValue, Object value) {
        Style style = new Style();
        style.getFont().setScheme(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getScheme());
    }

    @DisplayName("Test of the 'charset' value when writing and reading a Font style")
    @ParameterizedTest(name = "Given font charset '{0}' with value {1} should lead to a cell with this charset")
    @CsvSource(
            {
                    "' ', test",
                    "'test', 0.5f",
                    "'', true",}
    )
    void charsetFontTest(String styleValue, Object value) {
        Style style = new Style();
        style.getFont().setCharset(styleValue);
        style.getFont().setSize(22);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getFont().getCharset());
    }

}
