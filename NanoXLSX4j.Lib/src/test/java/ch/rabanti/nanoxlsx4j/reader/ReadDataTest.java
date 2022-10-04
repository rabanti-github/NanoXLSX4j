package ch.rabanti.nanoxlsx4j.reader;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.InputStream;
import java.time.Duration;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.function.BiConsumer;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.IOException;

public class ReadDataTest {

	@DisplayName("Test of the reader functionality for strings")
	@Test()
	void readStringTest() throws Exception {
		Map<String, String> cells = new HashMap<>();
		cells.put("A1",
				"Test");
		cells.put("B2",
				"22");
		cells.put("C3",
				"");
		cells.put("D4",
				" ");
		cells.put("E4",
				"x ");
		cells.put("F4",
				" X");
		cells.put("G4",
				" x ");
		cells.put("H4",
				"x x");
		cells.put("E5",
				"#@+-\"'?!\\(){}[]<>/|.,;:");
		cells.put("L6",
				"\t");
		cells.put("M6",
				"\tx");
		cells.put("N6",
				"x\t");
		cells.put("E7",
				"日本語");
		cells.put("F7",
				"हिन्दी");
		cells.put("G7",
				"한국어");
		cells.put("H7",
				"官話");
		cells.put("I7",
				"ελληνική γλώσσα");
		cells.put("J7",
				"русский язык");
		cells.put("K7",
				"עברית");
		cells.put("L7",
				"اَلْعَرَبِيَّة");
		assertValues(cells, ReadDataTest::assertEqualsFunction);
	}

	@DisplayName("Test of the reader functionality for new lines in strings")
	@Test()
	void readStringNewLineTest() throws Exception {
		Map<String, String> given = new HashMap<>();
		given.put("A1",
				"\r");
		given.put("A2",
				"\n");
		given.put("A3",
				"\r\n");
		given.put("A4",
				"a\n");
		given.put("A5",
				"\nx");
		given.put("A6",
				"a\r");
		given.put("A7",
				"\rx");
		given.put("A8",
				"a\r\n");
		given.put("A9",
				"\r\nx");
		given.put("A10",
				"\n\n\n");
		given.put("A11",
				"\r\r\r");
		given.put("A12",
				"\n\r"); // irregular use
		Map<String, String> expected = new HashMap<>();
		expected.put("A1",
				"\r\n");
		expected.put("A2",
				"\r\n");
		expected.put("A3",
				"\r\n");
		expected.put("A4",
				"a\r\n");
		expected.put("A5",
				"\r\nx");
		expected.put("A6",
				"a\r\n");
		expected.put("A7",
				"\r\nx");
		expected.put("A8",
				"a\r\n");
		expected.put("A9",
				"\r\nx");
		expected.put("A10",
				"\r\n\r\n\r\n");
		expected.put("A11",
				"\r\n\r\n\r\n");
		expected.put("A12",
				"\r\n\r\n");
		assertValues(given, ReadDataTest::assertEqualsFunction, expected);
	}

	@DisplayName("Test of the reader functionality for null / empty values")
	@Test()
	void readNullTest() throws Exception {
		Map<String, Object> cells = new HashMap<>();
		cells.put("A1", null);
		cells.put("A2", null);
		cells.put("A3", null);
		assertValues(cells, ReadDataTest::assertEqualsFunction);
	}

	@DisplayName("Test of the reader functionality for long values (above int32 range)")
	@Test()
	void readLongTest() throws Exception {
		Map<String, Long> cells = new HashMap<>();
		cells.put("A1", 4294967296L);
		cells.put("A2", -2147483649L);
		cells.put("A3", 21474836480L);
		cells.put("A4", -21474836480L);
		cells.put("A5", Long.MIN_VALUE);
		cells.put("A6", Long.MAX_VALUE);
		assertValues(cells, ReadDataTest::assertEqualsFunction);
	}

	@DisplayName("Test of the reader functionality for int values")
	@Test()
	public void ReadIntTest() throws Exception {
		Map<String, Integer> cells = new HashMap<>();
		cells.put("A1", 0);
		cells.put("A2", 10);
		cells.put("A3", -10);
		cells.put("A4", 999999);
		cells.put("A5", -999999);
		cells.put("A6", Integer.MIN_VALUE);
		cells.put("A7", Integer.MAX_VALUE);
		assertValues(cells, ReadDataTest::assertEqualsFunction);
	}

	@DisplayName("Test of the reader functionality for byte values (cast to int)")
	@Test()
	public void ReadByteTest() throws Exception {
		Map<String, Byte> cells = new HashMap<>();
		cells.put("A1", (byte) 0);
		cells.put("A2", (byte) 10);
		cells.put("A3", (byte) -10);
		cells.put("A4", (byte) 127);
		cells.put("A5", (byte) -127);
		cells.put("A6", Byte.MIN_VALUE);
		cells.put("A7", Byte.MAX_VALUE);
		Map<String, Integer> expected = new HashMap<>();
		expected.put("A1", 0);
		expected.put("A2", 10);
		expected.put("A3", -10);
		expected.put("A4", 127);
		expected.put("A5", -127);
		expected.put("A6", (int) Byte.MIN_VALUE);
		expected.put("A7", (int) Byte.MAX_VALUE);
		assertValuesCast(cells, ReadDataTest::assertEqualsFunction, expected);
	}

	@DisplayName("Test of the reader functionality for short values (cast to int)")
	@Test()
	public void ReadShortTest() throws Exception {
		Map<String, Short> cells = new HashMap<>();
		cells.put("A1", (short) 0);
		cells.put("A2", (short) 10);
		cells.put("A3", (short) -10);
		cells.put("A4", (short) 32767);
		cells.put("A5", (short) -32768);
		cells.put("A6", Short.MIN_VALUE);
		cells.put("A7", Short.MAX_VALUE);
		Map<String, Integer> expected = new HashMap<>();
		expected.put("A1", 0);
		expected.put("A2", 10);
		expected.put("A3", -10);
		expected.put("A4", 32767);
		expected.put("A5", -32768);
		expected.put("A6", (int) Short.MIN_VALUE);
		expected.put("A7", (int) Short.MAX_VALUE);
		assertValuesCast(cells, ReadDataTest::assertEqualsFunction, expected);
	}

	@DisplayName("Test of the reader functionality for float values")
	@Test()
	void readFloatTest() throws Exception {
		// Numbers without fraction elements are always interpreted as float
		Map<String, Float> cells = new HashMap<>();
		cells.put("A1", 0.000001f);
		cells.put("A2", 10.1f);
		cells.put("A3", -10.22f);
		cells.put("A4", 999999.9f);
		cells.put("A5", -999999.9f);
		cells.put("A6", Float.MIN_VALUE);
		cells.put("A7", Float.MAX_VALUE);
		assertValues(cells, ReadDataTest::assertApproximateFloatFunction);
	}

	@DisplayName("Test of the reader functionality for double values (above float range)")
	@Test()
	void readDoubleTest() throws Exception {
		Map<String, Double> cells = new HashMap<>();
		cells.put("A1", 440282346700000000000000000000000000009.1d);
		cells.put("A2", -440282347600000000000000000000000000009.1d);
		cells.put("A3", 21474836480648356436538453467583788456343865.227d);
		cells.put("A4", -21474836480648356436538453467583748856343865.9d);
		cells.put("A5", Double.MIN_VALUE);
		cells.put("A6", Double.MAX_VALUE);
		assertValues(cells, ReadDataTest::assertApproximateDoubleFunction);
	}

	@DisplayName("Test of the reader functionality for boolean values")
	@Test()
	void readBooleanTest() throws Exception {
		Map<String, Boolean> cells = new HashMap<>();
		cells.put("A1", true);
		cells.put("A2", false);
		cells.put("A3", true);
		assertValues(cells, ReadDataTest::assertEqualsFunction);
	}

	@DisplayName("Test of the reader functionality for Date values")
	@Test()
	void readDateTest() throws Exception {
		Map<String, Date> cells = new HashMap<>();
		cells.put("A1", TestUtils.buildDate(2021, 4, 11, 15, 7, 2));
		cells.put("A2", TestUtils.buildDate(1900, 0, 1, 0, 0, 0));
		cells.put("A3", TestUtils.buildDate(1960, 11, 12));
		cells.put("A4", TestUtils.buildDate(9999, 11, 31, 23, 59, 59));
		assertValues(cells, ReadDataTest::assertEqualsFunction);
	}

	@DisplayName("Test of the reader functionality for LocalTime values")
	@Test()
	void readLocalTimeTest() throws Exception {
		Map<String, Duration> cells = new HashMap<>();
		cells.put("A1", Helper.createDuration(0, 0, 0));
		cells.put("A2", Helper.createDuration(13, 18, 22));
		cells.put("A3", Helper.createDuration(12, 0, 0));
		cells.put("A4", Helper.createDuration(23, 59, 59));
		assertValues(cells, ReadDataTest::assertEqualsFunction);
	}

	@DisplayName("Test of the reader functionality for formulas (no formula parsing)")
	@Test()
	void readFormulaTest() throws Exception {
		Map<String, String> cells = new HashMap<>();
		long lmax = Long.MAX_VALUE;
		cells.put("A1",
				"=B2");
		cells.put("A2",
				"MIN(C2:D2)");
		cells.put("A3",
				"MAX(worksheet2!A1:worksheet2:A100");

		Workbook workbook = new Workbook("worksheet1");
		for (Map.Entry<String, String> cell : cells.entrySet()) {
			workbook.getCurrentWorksheet().addCellFormula(cell.getValue(), cell.getKey());
		}
		Workbook givenWorkbook = TestUtils.saveAndLoadWorkbook(workbook, null);

		assertNotNull(givenWorkbook);
		Worksheet givenWorksheet = givenWorkbook.setCurrentWorksheet(0);
		assertEquals("worksheet1", givenWorksheet.getSheetName());
		for (String address : cells.keySet()) {
			Cell givenCell = givenWorksheet.getCell(new Address(address));
			assertEquals(Cell.CellType.FORMULA, givenCell.getDataType());
			assertEquals(cells.get(address), givenCell.getValue());
		}
	}

	@DisplayName("Test of the reader functionality on invalid / unexpected values")
	@ParameterizedTest(name = "Given value {3} should lead to a {1} on address {0}")
	@CsvSource({
			"A1, STRING, STRING, 'Test'",
			"B1, STRING, STRING, 'x'",
			"C1, NUMBER, DOUBLE, '-1.8538541667'",
			"D1, NUMBER, INTEGER, 2",
			"E1, STRING, STRING, 'x'",
			"F1, STRING, STRING, '1'", // Reference 1 is cast to string '1'
			"G1, NUMBER, FLOAT, '-1.5f'",
			"H1, STRING, STRING, 'y'",
			"I1, BOOL, BOOLEAN,'true'",
			"J1, BOOL, BOOLEAN,'false'",
			"K1, STRING, STRING,'z'",
			"L1, STRING, STRING,'z'",
			"M1, STRING, STRING,'a'", })
	void readInvalidDataTest(String cellAddress, Cell.CellType expectedType, String sourceType, String sourceValue)
			throws Exception {
		// Note: Cell A1 is a valid String
		// Cell B1 is declared numerical, but contains a String
		// Cell C1 is defined as date but has a negative number
		// Cell D1 is defined ad boolean but has an invalid value of 2
		// Cell E1 is defined as boolean but has an invalid value of 'x'
		// Cell F1 is defines as shared String value, but the value does not exist
		// Cell G1 is defined as time but has a negative number
		// Cell H1 is defined as the unknown type 'z'
		// Cell I1 is defined as boolean but has 'true' instead of 1 as XML value
		// Cell J1 is defined as boolean but has 'FALSE' instead of 0 as XML value
		// Cell K1 is defined as date but has an invalid value of 'z'
		// Cell L1 is defined as time but has an invalid value of 'z'
		// Cell M1 is defined as shared string but has an invalid value of 'a'

		Object expectedValue = TestUtils.createInstance(sourceType, sourceValue);
		InputStream stream = TestUtils.getResource("tampered.xlsx");
		Workbook workbook = Workbook.load(stream);
		assertEquals(expectedType, workbook.getWorksheets().get(0).getCells().get(cellAddress).getDataType());
		assertEquals(expectedValue, workbook.getWorksheets().get(0).getCells().get(cellAddress).getValue());
	}

	@DisplayName("Test of the failing reader functionality on invalid XML content")
	@ParameterizedTest(name = "Given invalid workbook {0} should lead to an exception")
	@CsvSource({
			"invalid_workbook.xlsx",
			"invalid_workbook_sheet-definition.xlsx",
			"invalid_worksheet.xlsx",
			"invalid_style.xlsx",
			"invalid_metadata_app.xlsx",
			"invalid_metadata_core.xlsx",
			"invalid_sharedStrings.xlsx",
			"invalid_sharedStrings2.xlsx",
			"missing_worksheet.xlsx", })
	void failingReadInvalidDataTest(String invalidFile) {
		// Note: all referenced (embedded) files contains invalid XML documents
		// (malformed, missing start or end tags, missing attributes)
		InputStream stream = TestUtils.getResource(invalidFile);
		assertThrows(IOException.class, () -> Workbook.load(stream));
	}

	@DisplayName("Test of the reader functionality on an invalid stream")
	@Test()
	void readInvalidStreamTest() {
		InputStream nullStream = null;
		assertThrows(IOException.class, () -> Workbook.load(nullStream));
	}

	private static <T> void assertEqualsFunction(T expected, T given) {
		assertEquals(expected, given);
	}

	private static void assertApproximateDoubleFunction(double expected, double given) {
		double threshold = 0.00000001d;
		assertTrue(Math.abs(given - expected) < threshold);
	}

	private static void assertApproximateFloatFunction(float expected, float given) {
		float threshold = 0.00000001f;
		assertTrue(Math.abs(given - expected) < threshold);
	}

	private static <T> void assertValues(Map<String, T> givenCells, BiConsumer<T, T> assertionConsumer)
			throws Exception {
		assertValues(givenCells, assertionConsumer, null);
	}

	private static <T> void assertValues(Map<String, T> givenCells, BiConsumer<T, T> assertionConsumer, Map<String, T> expectedCells) throws Exception {
		Worksheet givenWorksheet = getWorksheet(givenCells);
		for (String address : givenCells.keySet()) {
			Cell givenCell = givenWorksheet.getCell(new Address(address));
			T value;
			if (expectedCells == null) {
				value = givenCells.get(address);
			}
			else {
				value = expectedCells.get(address);
			}

			if (value == null) {
				assertEquals(Cell.CellType.EMPTY, givenCell.getDataType());
			}
			else {
				assertionConsumer.accept(value, (T) givenCell.getValue());
			}
		}
	}

	private static <T, D> void assertValuesCast(Map<String, T> givenCells, BiConsumer<D, D> assertionConsumer, Map<String, D> expectedCells) throws Exception {
		Worksheet givenWorksheet = getWorksheet(givenCells);
		for (String address : givenCells.keySet()) {
			Cell givenCell = givenWorksheet.getCell(new Address(address));
			D givenValue = (D) givenCell.getValue();
			D expectedValue = expectedCells.get(address);
			if (givenValue == null) {
				assertEquals(Cell.CellType.EMPTY, givenCell.getDataType());
			}
			else {
				assertionConsumer.accept(givenValue, expectedValue);
			}
		}
	}

	private static <T> Worksheet getWorksheet(Map<String, T> givenCells) throws IOException, java.io.IOException {
		Workbook workbook = new Workbook("worksheet1");
		for (Map.Entry<String, T> cell : givenCells.entrySet()) {
			workbook.getCurrentWorksheet().addCell(cell.getValue(), cell.getKey());
		}
		Workbook givenWorkbook = TestUtils.saveAndLoadWorkbook(workbook, null);

		assertNotNull(givenWorkbook);
		Worksheet givenWorksheet = givenWorkbook.setCurrentWorksheet(0);
		assertEquals("worksheet1", givenWorksheet.getSheetName());
		return givenWorksheet;
	}
}
