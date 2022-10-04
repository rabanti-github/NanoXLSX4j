package ch.rabanti.nanoxlsx4j.misc;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;

public class XlsxWriterTest {

	@DisplayName("Test of the 'EscapeXmlChars' method on characters that has to be replaced, when writing a workbook")
	@ParameterizedTest(name = "Given char number {1}, pre- and appended with {0} should lead to a text: {2}")
	@CsvSource({
			"test, 0x41, testAtest", // Not printable
			"test, 0x8, test test", // "
			"test, 0xC, test test", // "
			"test, 0x1F, test test", // "
			"test, 0xD800, test test", // Above valid UTF range
			"test, 0x3C, test<test", // internally saved as &lt;
			"test, 0x3E, test>test", // internally saved as &gt;
			"test, 0x26, test&test", // internally saved as &amp;
	})
	void escapeXmlCharsTest(String givenPrePostFix, int charToEscape, String expectedText) throws Exception {
		String givenText = givenPrePostFix + (char) charToEscape + givenPrePostFix;
		Workbook workbook = new Workbook("worksheet1");
		workbook.getCurrentWorksheet().addCell(givenText, "A1");
		Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
		assertEquals(expectedText, givenWorkbook.getCurrentWorksheet().getCell("A1").getValue());
	}

	@DisplayName("Test of the 'EscapeXmlChars' method on characters that has to be replaced, when writing a workbook")
	@ParameterizedTest(name = "Given char number {1}, pre- and appended with {0} should lead to a text: {2}")
	@CsvSource({
			"ws, 0x41, wsAws", // Not printable
			"ws, 0x8, ws ws", // "
			"ws, 0xC, ws ws", // "
			"ws, 0x1F, ws ws", // "
			"ws, 0xD800, ws ws", // Above valid UTF range
			"ws, 0x22, ws\"ws", // internally saved as &quot; "ws, 0x3C, ws<ws",
			"ws, 0x3C, ws<ws", // internally saved as &lt;
			"ws, 0x3E, ws>ws", // internally saved as &gt;
			"ws, 0x26, ws&ws", // internally saved as &amp;
	})
	void escapeXmlAttributeCharsTest(String givenPrePostFix, int charToEscape, String expectedText) throws Exception {
		// To test the function, the worksheet name is used, since defined as workbook
		// attribute
		String givenName = givenPrePostFix + (char) charToEscape + givenPrePostFix;
		Workbook workbook = new Workbook(givenName);
		workbook.getCurrentWorksheet().addCell(42, "A1");
		Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
		assertEquals(expectedText, givenWorkbook.getCurrentWorksheet().getSheetName());
	}

}
