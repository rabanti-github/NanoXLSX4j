package ch.rabanti.nanoxlsx4j.styles;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

public class BasicStylesTest {

	@DisplayName("Test of the static Bold style")
	@Test()
	void boldTest() {
		Style style = BasicStyles.Bold();
		assertNotNull(style);
		assertTrue(style.getFont().isBold());
	}

	@DisplayName("Test of the static Italic style")
	@Test()
	void italicTest() {
		Style style = BasicStyles.Italic();
		assertNotNull(style);
		assertTrue(style.getFont().isItalic());
	}

	@DisplayName("Test of the static BoldItalic style")
	@Test()
	void boldItalicTest() {
		Style style = BasicStyles.BoldItalic();
		assertNotNull(style);
		assertTrue(style.getFont().isItalic());
		assertTrue(style.getFont().isBold());
	}

	@DisplayName("Test of the static Underline style")
	@Test()
	void underlineTest() {
		Style style = BasicStyles.Underline();
		assertNotNull(style);
		assertEquals(Font.UnderlineValue.u_single, style.getFont().getUnderline());
	}

	@DisplayName("Test of the static DoubleUnderline style")
	@Test()
	void doubleUnderlineTest() {
		Style style = BasicStyles.DoubleUnderline();
		assertNotNull(style);
		assertEquals(Font.UnderlineValue.u_double, style.getFont().getUnderline());
	}

	@DisplayName("Test of the static Strike style")
	@Test()
	void strikeTest() {
		Style style = BasicStyles.Strike();
		assertNotNull(style);
		assertTrue(style.getFont().isStrike());
	}

	@DisplayName("Test of the static TimeFormat style")
	@Test()
	void timeFormatTest() {
		Style style = BasicStyles.TimeFormat();
		assertNotNull(style);
		assertEquals(NumberFormat.FormatNumber.format_21, style.getNumberFormat().getNumber());
	}

	@DisplayName("Test of the static DateFormat style")
	@Test()
	void dateFormatTest() {
		Style style = BasicStyles.DateFormat();
		assertNotNull(style);
		assertEquals(NumberFormat.FormatNumber.format_14, style.getNumberFormat().getNumber());
	}

	@DisplayName("Test of the static RoundFormat style")
	@Test()
	void roundFormatTest() {
		Style style = BasicStyles.RoundFormat();
		assertNotNull(style);
		assertEquals(NumberFormat.FormatNumber.format_1, style.getNumberFormat().getNumber());
	}

	@DisplayName("Test of the static MergeCell style")
	@Test()
	void mergeCellStyleTest() {
		Style style = BasicStyles.MergeCellStyle();
		assertNotNull(style);
		assertTrue(style.getCellXf().isForceApplyAlignment());
	}

	@DisplayName("Test of the static DottedFill_0_125 style")
	@Test()
	void dottedFill_0_125Test() {
		Style style = BasicStyles.DottedFill_0_125();
		assertNotNull(style);
		assertEquals(Fill.PatternValue.gray125, style.getFill().getPatternFill());
	}

	@DisplayName("Test of the static BorderFrame style")
	@Test()
	void borderFrameTest() {
		Style style = BasicStyles.BorderFrame();
		assertNotNull(style);
		assertEquals(Border.StyleValue.thin, style.getBorder().getTopStyle());
		assertEquals(Border.StyleValue.thin, style.getBorder().getBottomStyle());
		assertEquals(Border.StyleValue.thin, style.getBorder().getLeftStyle());
		assertEquals(Border.StyleValue.thin, style.getBorder().getRightStyle());
	}

	@DisplayName("Test of the static BorderFrameHeader style")
	@Test()
	void borderFrameHeaderTest() {
		Style style = BasicStyles.BorderFrameHeader();
		assertNotNull(style);
		assertEquals(Border.StyleValue.thin, style.getBorder().getTopStyle());
		assertEquals(Border.StyleValue.medium, style.getBorder().getBottomStyle());
		assertEquals(Border.StyleValue.thin, style.getBorder().getLeftStyle());
		assertEquals(Border.StyleValue.thin, style.getBorder().getRightStyle());
		assertTrue(style.getFont().isBold());
	}

	@DisplayName("Test of the ColorizedText function")
	@ParameterizedTest(name = "Given value {0} should lead to the color {1}")
	@CsvSource({ "000000, FF000000", "3CDEF0, FF3CDEF0", "af3cd1, FFAF3CD1", "FFFFFF, FFFFFFFF", })
	void colorizedTextTest(String hexCode, String expectedHexCode) {
		Style style = BasicStyles.colorizedText(hexCode);
		assertNotNull(style);
		assertEquals(expectedHexCode, style.getFont().getColorValue());
	}

	@DisplayName("Test of the failing colorizedText function")
	@ParameterizedTest(name = "Given value {1} should lead to an exception")
	@CsvSource({ "NULL, ''", "STRING, ''", "STRING, ' '", "STRING, 'AAFF'", "STRING, 'AAFFCC22'", "STRING, 'XXXXVV'", })
	void colorizedTextFailTest(String sourceType, String sourceValue) {
		String hexCode = (String) TestUtils.createInstance(sourceType, sourceValue);
		assertThrows(StyleException.class, () -> BasicStyles.colorizedText(hexCode));
	}

	@DisplayName("Test of the ColorizedBackground function")
	@ParameterizedTest(name = "Given value {0} should lead to the color {1}")
	@CsvSource({ "000000, FF000000", "3CDEF0, FF3CDEF0", "af3cd1, FFAF3CD1", "FFFFFF, FFFFFFFF", })
	void colorizedBackgroundTest(String hexCode, String expectedHexCode) {
		Style style = BasicStyles.colorizedBackground(hexCode);
		assertNotNull(style);
		assertEquals(expectedHexCode, style.getFill().getForegroundColor());
		assertEquals(Fill.DEFAULT_COLOR, style.getFill().getBackgroundColor());
		assertEquals(Fill.PatternValue.solid, style.getFill().getPatternFill());
	}

	@DisplayName("Test of the failing colorizedBackground function")
	@ParameterizedTest(name = "Given value {1} should lead to an exception")
	@CsvSource({ "NULL, ''", "STRING, ''", "STRING, ' '", "STRING, 'AAFF'", "STRING, 'AAFFCC22'", "STRING, 'XXXXVV'", })
	void colorizedBackgroundFailTest(String sourceType, String sourceValue) {
		String hexCode = (String) TestUtils.createInstance(sourceType, sourceValue);
		assertThrows(StyleException.class, () -> BasicStyles.colorizedBackground(hexCode));
	}

	@DisplayName("Test of the Font function with name")
	@ParameterizedTest(name = "Given name {0} should lead to a valid Font style")
	@CsvSource({ "Calibri", "Arial", "Times New Roman", "Sans Serif", "Tahoma", })
	void fontTest3(String name) {
		Style style = BasicStyles.font(name);
		assertEquals(name, style.getFont().getName());
		assertEquals(Font.DEFAULT_FONT_SIZE, style.getFont().getSize());
		assertFalse(style.getFont().isBold());
		assertFalse(style.getFont().isItalic());
	}

	@DisplayName("Test of the Font function with name and size")
	@ParameterizedTest(name = "Given name {0} and size {1} should lead to a valid Font style")
	@CsvSource({ "Calibri, 12f", "Arial, 1f", "Times New Roman, 409f", "Sans Serif, 50f", "Tahoma, 11f", })
	void fontTest3(String name, float size) {
		Style style = BasicStyles.font(name, size);
		assertEquals(name, style.getFont().getName());
		assertEquals(size, style.getFont().getSize());
		assertFalse(style.getFont().isBold());
		assertFalse(style.getFont().isItalic());
	}

	@DisplayName("Test of the Font function with name, size and bold state")
	@ParameterizedTest(name = "Given name {0}, size {1} and bold state {2} should lead to a valid Font style")
	@CsvSource({ "Calibri, 12f, false", "Arial, 1f, false", "Times New Roman, 409f, true", "Sans Serif, 50f, false",
			"Tahoma, 11f, true", })
	void fontTest3(String name, float size, boolean bold) {
		Style style = BasicStyles.font(name, size, bold);
		assertEquals(name, style.getFont().getName());
		assertEquals(size, style.getFont().getSize());
		assertEquals(bold, style.getFont().isBold());
		assertFalse(style.getFont().isItalic());
	}

	@DisplayName("Test of the Font function with all parameters")
	@ParameterizedTest(name = "Given name {0}, size {1}, bold state {2} and italic state {3} should lead to a valid Font style")
	@CsvSource({ "Calibri, 12f, false, false", "Arial, 1f, false, false", "Times New Roman, 409f, true, false",
			"Sans Serif, 50f, false, true", "Tahoma, 11f, true, true", })
	void fontTest3(String name, float size, boolean bold, boolean italic) {
		Style style = BasicStyles.font(name, size, bold, italic);
		assertEquals(name, style.getFont().getName());
		assertEquals(size, style.getFont().getSize());
		assertEquals(bold, style.getFont().isBold());
		assertEquals(italic, style.getFont().isItalic());
	}

	@DisplayName("Test of the Font function for the auto adjustment of invalid font sizes")
	@Test()
	void fontTest4() {
		Style style = BasicStyles.font("Arial", -1f);
		assertEquals(Font.MIN_FONT_SIZE, style.getFont().getSize());
		style = BasicStyles.font("Arial", 0.5f);
		assertEquals(Font.MIN_FONT_SIZE, style.getFont().getSize());
		style = BasicStyles.font("Arial", 409.1f);
		assertEquals(Font.MAX_FONT_SIZE, style.getFont().getSize());
		style = BasicStyles.font("Arial", 1000f);
		assertEquals(Font.MAX_FONT_SIZE, style.getFont().getSize());
	}

	@DisplayName("Test of the failing font function on a invalid font name")
	@Test()
	void fontFailTest() {
		assertThrows(StyleException.class, () -> BasicStyles.font(null));
		assertThrows(StyleException.class, () -> BasicStyles.font(""));
	}
}
