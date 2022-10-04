package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;

public class StaticTest {

	@DisplayName("Test of the SanitizeWorksheetName function")
	@ParameterizedTest(name = "Given name {0} and number of existing worksheets {1} with prefix {2} should lead to a worksheet name {3}")
	@CsvSource({
			"test, 0, '', test",
			"Sheet2, 1, Sheet, Sheet2",
			", 0, '', Sheet1",
			"'', 0, '', Sheet1",
			"a[b, 0, '', a_b",
			"a]b, 0, '', a_b",
			"a*b, 0, '', a_b",
			"a?b, 0, '', a_b",
			"a/b, 0, '', a_b",
			"a\\b,0, '', a_b",
			"--------------------------------, 0, '', -------------------------------",
			"Sheet10, 20, Sheet, Sheet21",
			"*1, 1, _, _2",
			"------------------------------9, 9, ------------------------------, -----------------------------10",
			"'9999999999999999999999999999999', 9, '999999999999999999999999999999', '0'", })
	void sanitizeWorksheetNameTest(String givenName, int numberOfExistingWorksheets, String existingWorksheetPrefix, String expectedName) {
		Workbook workbook = new Workbook(false);
		for (int i = 0; i < numberOfExistingWorksheets; i++) {
			workbook.addWorksheet(existingWorksheetPrefix + (i + 1));
		}
		String name = Worksheet.sanitizeWorksheetName(givenName, workbook);
		assertEquals(expectedName, name);
	}

}
