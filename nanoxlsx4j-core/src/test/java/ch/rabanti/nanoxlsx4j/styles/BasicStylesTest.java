/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

class BasicStylesTest {

    @Test
    @DisplayName("Test of the static Bold style")
    void boldTest() {
        Style style = BasicStyles.getBold();
        assertNotNull(style);
        assertTrue(style.getCurrentFont().isBold());
    }

    @Test
    @DisplayName("Test of the static Italic style")
    void italicTest() {
        Style style = BasicStyles.getItalic();
        assertNotNull(style);
        assertTrue(style.getCurrentFont().isItalic());
    }

    @Test
    @DisplayName("Test of the static BoldItalic style")
    void boldItalicTest() {
        Style style = BasicStyles.getBoldItalic();
        assertNotNull(style);
        assertTrue(style.getCurrentFont().isItalic());
        assertTrue(style.getCurrentFont().isBold());
    }

    @Test
    @DisplayName("Test of the static Underline style")
    void underlineTest() {
        Style style = BasicStyles.getUnderline();
        assertNotNull(style);
        assertEquals(Font.UnderlineValue.SINGLE, style.getCurrentFont().getUnderline());
    }

    @Test
    @DisplayName("Test of the static DoubleUnderline style")
    void doubleUnderlineTest() {
        Style style = BasicStyles.getDoubleUnderline();
        assertNotNull(style);
        assertEquals(Font.UnderlineValue.DOUBLE, style.getCurrentFont().getUnderline());
    }

    @Test
    @DisplayName("Test of the static Strike style")
    void strikeTest() {
        Style style = BasicStyles.getStrike();
        assertNotNull(style);
        assertTrue(style.getCurrentFont().isStrike());
    }

    @Test
    @DisplayName("Test of the static TimeFormat style")
    void timeFormatTest() {
        Style style = BasicStyles.getTimeFormat();
        assertNotNull(style);
        assertEquals(NumberFormat.FormatNumber.FORMAT_21, style.getCurrentNumberFormat().getNumber());
    }

    @Test
    @DisplayName("Test of the static DateFormat style")
    void dateFormatTest() {
        Style style = BasicStyles.getDateFormat();
        assertNotNull(style);
        assertEquals(NumberFormat.FormatNumber.FORMAT_14, style.getCurrentNumberFormat().getNumber());
    }

    @Test
    @DisplayName("Test of the static RoundFormat style")
    void roundFormatTest() {
        Style style = BasicStyles.getRoundFormat();
        assertNotNull(style);
        assertEquals(NumberFormat.FormatNumber.FORMAT_1, style.getCurrentNumberFormat().getNumber());
    }

    @Test
    @DisplayName("Test of the static MergeCell style")
    void mergeCellStyleTest() {
        Style style = BasicStyles.getMergeCellStyle();
        assertNotNull(style);
        assertTrue(style.getCurrentCellXf().isForceApplyAlignment());
    }

    @Test
    @DisplayName("Test of the static DottedFill_0_125 style")
    void dottedFill_0_125Test() {
        Style style = BasicStyles.getDottedFill0125();
        assertNotNull(style);
        assertEquals(Fill.PatternValue.GRAY_125, style.getCurrentFill().getPatternFill());
    }

    @Test
    @DisplayName("Test of the static BorderFrame style")
    void borderFrameTest() {
        Style style = BasicStyles.getBorderFrame();
        assertNotNull(style);
        assertEquals(Border.StyleValue.THIN, style.getCurrentBorder().getTopStyle());
        assertEquals(Border.StyleValue.THIN, style.getCurrentBorder().getBottomStyle());
        assertEquals(Border.StyleValue.THIN, style.getCurrentBorder().getLeftStyle());
        assertEquals(Border.StyleValue.THIN, style.getCurrentBorder().getRightStyle());
    }

    @Test
    @DisplayName("Test of the static BorderFrameHeader style")
    void borderFrameHeaderTest() {
        Style style = BasicStyles.getBorderFrameHeader();
        assertNotNull(style);
        assertEquals(Border.StyleValue.THIN, style.getCurrentBorder().getTopStyle());
        assertEquals(Border.StyleValue.MEDIUM, style.getCurrentBorder().getBottomStyle());
        assertEquals(Border.StyleValue.THIN, style.getCurrentBorder().getLeftStyle());
        assertEquals(Border.StyleValue.THIN, style.getCurrentBorder().getRightStyle());
        assertTrue(style.getCurrentFont().isBold());
    }

    @ParameterizedTest
    @DisplayName("Test of the ColorizedText function")
    @CsvSource({
            "000000, FF000000",
            "3CDEF0, FF3CDEF0",
            "af3cd1, FFAF3CD1",
            "FFFFFF, FFFFFFFF"
    })
    void colorizedTextTest(String hexCode, String expectedHexCode) {
        Style style = BasicStyles.colorizedText(hexCode);
        assertNotNull(style);
        assertEquals(expectedHexCode, style.getCurrentFont().getColorValue().getArgbValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing ColorizedText function")
    @NullSource
    @ValueSource(strings = {"", " ", "AAFF", "AAFFCC22", "XXXXVV"})
    void colorizedTextFailTest(String hexCode) {
        assertThrows(StyleException.class, () -> BasicStyles.colorizedText(hexCode));
    }

    @ParameterizedTest
    @DisplayName("Test of the ColorizedBackground function")
    @CsvSource({
            "000000, FF000000",
            "3CDEF0, FF3CDEF0",
            "af3cd1, FFAF3CD1",
            "FFFFFF, FFFFFFFF"
    })
    void colorizedBackgroundTest(String hexCode, String expectedHexCode) {
        Style style = BasicStyles.colorizedBackground(hexCode);
        assertNotNull(style);
        assertEquals(expectedHexCode, style.getCurrentFill().getForegroundColor().getArgbValue());
        assertEquals(Fill.DEFAULT_COLOR, style.getCurrentFill().getBackgroundColor());
        assertEquals(Fill.PatternValue.SOLID, style.getCurrentFill().getPatternFill());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing ColorizedBackground function")
    @NullSource
    @ValueSource(strings = {"", " ", "AAFF", "AAFFCC22", "XXXXVV"})
    void colorizedBackgroundFailTest(String hexCode) {
        assertThrows(StyleException.class, () -> BasicStyles.colorizedBackground(hexCode));
    }

    @ParameterizedTest
    @DisplayName("Test of the Font function with name")
    @ValueSource(strings = {"Calibri", "Arial", "Times New Roman", "Sans Serif", "Tahoma"})
    void fontTest(String name) {
        Style style = BasicStyles.font(name);
        assertEquals(name, style.getCurrentFont().getName());
        assertEquals(Font.DEFAULT_FONT_SIZE, style.getCurrentFont().getSize());
        assertFalse(style.getCurrentFont().isBold());
        assertFalse(style.getCurrentFont().isItalic());
    }

    @ParameterizedTest
    @DisplayName("Test of the Font function with name and size")
    @CsvSource({
            "Calibri, 12",
            "Arial, 1",
            "Times New Roman, 409",
            "Sans Serif, 50",
            "Tahoma, 11"
    })
    void fontTest2(String name, float size) {
        Style style = BasicStyles.font(name, size);
        assertEquals(name, style.getCurrentFont().getName());
        assertEquals(size, style.getCurrentFont().getSize());
        assertFalse(style.getCurrentFont().isBold());
        assertFalse(style.getCurrentFont().isItalic());
    }

    @ParameterizedTest
    @DisplayName("Test of the Font function with name, size and bold state")
    @CsvSource({
            "Calibri, 12, false",
            "Arial, 1, false",
            "Times New Roman, 409, true",
            "Sans Serif, 50, false",
            "Tahoma, 11, true"
    })
    void fontTest3(String name, float size, boolean bold) {
        Style style = BasicStyles.font(name, size, bold);
        assertEquals(name, style.getCurrentFont().getName());
        assertEquals(size, style.getCurrentFont().getSize());
        assertEquals(bold, style.getCurrentFont().isBold());
        assertFalse(style.getCurrentFont().isItalic());
    }

    @Test
    @DisplayName("Test of the Font function for the auto adjustment of invalid font sizes")
    void fontTest4() {
        Style style = BasicStyles.font("Arial", -1f);
        assertEquals(Font.MIN_FONT_SIZE, style.getCurrentFont().getSize());
        style = BasicStyles.font("Arial", 0.5f);
        assertEquals(Font.MIN_FONT_SIZE, style.getCurrentFont().getSize());
        style = BasicStyles.font("Arial", 409.1f);
        assertEquals(Font.MAX_FONT_SIZE, style.getCurrentFont().getSize());
        style = BasicStyles.font("Arial", 1000f);
        assertEquals(Font.MAX_FONT_SIZE, style.getCurrentFont().getSize());
    }

    @Test
    @DisplayName("Test of the failing Font function on a invalid font name")
    void fontFailTest() {
        assertThrows(StyleException.class, () -> BasicStyles.font(null));
        assertThrows(StyleException.class, () -> BasicStyles.font(""));
    }
}
