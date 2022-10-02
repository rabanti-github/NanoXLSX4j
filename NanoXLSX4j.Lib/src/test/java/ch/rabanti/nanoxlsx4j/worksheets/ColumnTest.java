package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;

public class ColumnTest {
	@DisplayName("Test of the addHiddenColumn function with a column number")
	@Test()
	void addHiddenColumnTest() {
		Worksheet worksheet = new Worksheet();
		assertEquals(0, worksheet.getColumns().size());
		worksheet.addHiddenColumn(2);
		assertEquals(1, worksheet.getColumns().size());
		assertTrue(worksheet.getColumns().containsKey(2));
		assertTrue(worksheet.getColumns().get(2).isHidden());
		worksheet.addHiddenColumn(2); // Should not add an additional entry
		assertEquals(1, worksheet.getColumns().size());
	}

	@DisplayName("Test of the addHiddenColumn function with a column as String")
	@Test()
	void addHiddenColumnTest2() {
		Worksheet worksheet = new Worksheet();
		assertEquals(0, worksheet.getColumns().size());
		worksheet.addHiddenColumn("C");
		assertEquals(1, worksheet.getColumns().size());
		assertTrue(worksheet.getColumns().containsKey(2));
		assertTrue(worksheet.getColumns().get(2).isHidden());
		worksheet.addHiddenColumn("C"); // Should not add an additional entry
		assertEquals(1, worksheet.getColumns().size());
	}

	@DisplayName("Test of the failing addHiddenColumn function with an invalid column number")
	@ParameterizedTest(name = "Given value {0} should lead to an exception")
	@CsvSource({ "-1", "-100", "16384", })
	void addHiddenColumnFailTest(int value) {
		Worksheet worksheet = new Worksheet();
		assertThrows(RangeException.class, () -> worksheet.addHiddenColumn(value));
	}

	@DisplayName("Test of the failing addHiddenColumn function with an invalid column string")
	@ParameterizedTest(name = "Given value {0} should lead to an exception")
	@CsvSource({ "NULL, ''", "STRING, ''", "STRING,'#'", "STRING,'XFE'", })
	void addHiddenColumnFailTest2(String sourceType, String sourceValue) {
		String value = (String) TestUtils.createInstance(sourceType, sourceValue);
		Worksheet worksheet = new Worksheet();
		assertThrows(RangeException.class, () -> worksheet.addHiddenColumn(value));
	}

	@DisplayName("Test of the resetColumn function with an empty worksheet")
	@Test()
	void resetColumnTest() {
		Worksheet worksheet = new Worksheet();
		assertEquals(0, worksheet.getColumns().size());
		worksheet.resetColumn(0); // Should do nothing and not fail
		assertEquals(0, worksheet.getColumns().size());
	}

	@DisplayName("Test of the resetColumn function with defined columns but a not defined columns")
	@Test()
	void resetColumnTest2() {
		Worksheet worksheet = new Worksheet();
		assertEquals(0, worksheet.getColumns().size());
		worksheet.addHiddenColumn(0);
		worksheet.addHiddenColumn(2);
		worksheet.resetColumn(1); // Should do nothing and not fail
		assertEquals(2, worksheet.getColumns().size());
	}

	@DisplayName("Test of the resetColumn function with defined columns and an existing column")
	@Test()
	void resetColumnTest3() {
		Worksheet worksheet = new Worksheet();
		assertEquals(0, worksheet.getColumns().size());
		worksheet.addHiddenColumn(0);
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(2);
		assertEquals(3, worksheet.getColumns().size());
		worksheet.resetColumn(1);
		assertEquals(2, worksheet.getColumns().size());
		assertFalse(worksheet.getColumns().containsKey(1));
	}

	@DisplayName("Test of the resetColumn function with defined columns and a AutoFilter definition")
	@Test()
	void resetColumnTest4() {
		Worksheet worksheet = new Worksheet();
		assertEquals(0, worksheet.getColumns().size());
		worksheet.addHiddenColumn(0);
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(2);
		worksheet.setAutoFilter("A1:C1");
		assertEquals(3, worksheet.getColumns().size());
		worksheet.setColumnWidth("B", 66.6f);
		worksheet.resetColumn(1); // Should not remove the column, since in a AutoFilter
		assertEquals(3, worksheet.getColumns().size());
		assertTrue(worksheet.getColumns().containsKey(1));
		assertFalse(worksheet.getColumns().get(1).isHidden());
		assertTrue(worksheet.getColumns().get(1).hasAutoFilter());
		assertEquals(Worksheet.DEFAULT_COLUMN_WIDTH, worksheet.getColumns().get(1).getWidth());
	}

	@DisplayName("Test of the getCurrentColumnNumber function")
	@Test()
	void getCurrentColumnNumberTest() {
		Worksheet worksheet = new Worksheet();
		assertEquals(0, worksheet.getCurrentColumnNumber());
		worksheet.setCurrentCellDirection(Worksheet.CellDirection.ColumnToColumn);
		worksheet.addNextCell("test");
		worksheet.addNextCell("test");
		worksheet.setCurrentCellDirection(Worksheet.CellDirection.RowToRow);
		worksheet.addNextCell("test");
		worksheet.addNextCell("test");
		assertEquals(2, worksheet.getCurrentRowNumber()); // should not change
		assertEquals(2, worksheet.getCurrentColumnNumber());
		worksheet.goToNextColumn();
		assertEquals(3, worksheet.getCurrentColumnNumber());
		worksheet.goToNextColumn(2);
		assertEquals(5, worksheet.getCurrentColumnNumber());
		worksheet.goToNextRow(2);
		assertEquals(0, worksheet.getCurrentColumnNumber()); // should reset
	}

	@DisplayName("Test of the goToNextColumn function")
	@ParameterizedTest(name = "Given initial column number {0} and number {1} should lead to the column {2}")
	@CsvSource({ "0, 0, 0", "0, 1, 1", "1, 1, 2", "3, 10, 13", "3, -1, 2", "3, -3, 0", })
	void goToNextColumnTest(int initialColumnNumber, int number, int expectedColumnNumber) {
		Worksheet worksheet = new Worksheet();
		worksheet.setCurrentColumnNumber(initialColumnNumber);
		worksheet.goToNextColumn(number);
		assertEquals(expectedColumnNumber, worksheet.getCurrentColumnNumber());
	}

	@DisplayName("Test of the goToNextColumn function with the option to keep the row")
	@ParameterizedTest(name = "Given start address {0} and number {1} with the option to keep the row: {2} should lead to the address {3}")
	@CsvSource({ "A1, 0, false, A1", "A1, 0, true, A1", "A1, 1, false, B1", "A1, 1, true, B1", "C10, 1, false, D1",
			"C10, 1, true, D10", "R5, 5, false, W1", "R5, 5, true, W5", "F5, -3, false, C1", "F5, -3, true, C5",
			"F5, -5, false, A1", "F5, -5, true, A5", })
	void goToNextColumnTest2(String initialAddress, int number, boolean keepRowPosition, String expectedAddress) {
		Worksheet worksheet = new Worksheet();
		worksheet.setCurrentCellAddress(initialAddress);
		worksheet.goToNextColumn(number, keepRowPosition);
		Address expected = new Address(expectedAddress);
		assertEquals(expected.Column, worksheet.getCurrentColumnNumber());
		assertEquals(expected.Row, worksheet.getCurrentRowNumber());
	}

	@DisplayName("Test of the failing goToNextColumn function on invalid values")
	@ParameterizedTest(name = "Given incremental value {1} an basis {0} should lead to an exception")
	@CsvSource({ "0, -1", "10, -12", "0, 16384", "0, 20383", })
	void goToNextColumnFailTest(int initialValue, int value) {
		Worksheet worksheet = new Worksheet();
		worksheet.setCurrentColumnNumber(initialValue);
		assertEquals(initialValue, worksheet.getCurrentColumnNumber());
		assertThrows(RangeException.class, () -> worksheet.goToNextColumn(value));
	}

	@DisplayName("Test of the setAutoFilter function on a start and end column")
	@ParameterizedTest(name = "Given start column {0} and end column {1} should lead to a range {2}")
	@CsvSource({ "0, 0, A1:A1", "0, 5, A1:F1", "1, 5, B1:F1", "5, 1, B1:F1", })
	void setAutoFilterTest(int startColumn, int endColumn, String expectedRange) {
		Worksheet worksheet = new Worksheet();
		assertNull(worksheet.getAutoFilterRange());
		worksheet.setAutoFilter(startColumn, endColumn);
		assertNotNull(worksheet.getAutoFilterRange());
		assertEquals(expectedRange, worksheet.getAutoFilterRange().toString());
	}

	@DisplayName("Test of the setAutoFilter function on a range as string")
	@ParameterizedTest(name = "Given range {0} should lead to {0}")
	@CsvSource({ "A1:A1, A1:A1", "A1:F1, A1:F1", "B1:F1, B1:F1", "F1:B1, B1:F1", "$B$1:$F$1, B1:F1", })
	void setAutoFilterTest2(String givenRange, String expectedRange) {
		Worksheet worksheet = new Worksheet();
		assertNull(worksheet.getAutoFilterRange());
		worksheet.setAutoFilter(givenRange);
		assertNotNull(worksheet.getAutoFilterRange());
		assertEquals(expectedRange, worksheet.getAutoFilterRange().toString());
	}

	@DisplayName("Test of the failing setAutoFilter function on an invalid start and / or end column")
	@ParameterizedTest(name = "Given start column {0} and end column {1} should lead to an exception")
	@CsvSource({ "-1, 0", "0, -1", "-1, -1", "2, 16384", "16384, 2", "16384, 16384", })
	void setAutoFilterFailingTest(int startColumn, int endColumn) {
		Worksheet worksheet = new Worksheet();
		assertThrows(RangeException.class, () -> worksheet.setAutoFilter(startColumn, endColumn));
	}

	@DisplayName("Test of the failing setAutoFilter function on an invalid string expression")
	@ParameterizedTest(name = "Given range {0} should lead to an exception")
	@CsvSource({ "NULL, ''", "STRING, ''", "STRING, 'A1'", "STRING, ':'", })
	void setAutoFilterFailingTest2(String sourceType, String sourceValue) {
		String range = (String) TestUtils.createInstance(sourceType, sourceValue);
		Worksheet worksheet = new Worksheet();
		assertThrows(FormatException.class, () -> worksheet.setAutoFilter(range));
	}

	@DisplayName("Test of the setColumnWidth function with column number and column address")
	@ParameterizedTest(name = "Given width {0} should lead to a valid column definition with this width")
	@CsvSource({ "0f", "0.1f", "10f", "255f", })
	void setColumnWidthTest(float width) {
		Worksheet worksheet = new Worksheet();
		assertEquals(0, worksheet.getColumns().size());
		worksheet.setColumnWidth(0, width);
		worksheet.setColumnWidth("B", width);
		assertEquals(2, worksheet.getColumns().size());
		assertEquals(width, worksheet.getColumns().get(0).getWidth());
		assertEquals(width, worksheet.getColumns().get(1).getWidth());
		worksheet.setColumnWidth(0, Worksheet.DEFAULT_COLUMN_WIDTH);
		worksheet.setColumnWidth("B", Worksheet.DEFAULT_COLUMN_WIDTH);
		assertEquals(2, worksheet.getColumns().size()); // No removal so far
		assertEquals(Worksheet.DEFAULT_COLUMN_WIDTH, worksheet.getColumns().get(0).getWidth());
		assertEquals(Worksheet.DEFAULT_COLUMN_WIDTH, worksheet.getColumns().get(1).getWidth());
	}

	@DisplayName("Test of the failing setColumnWidth function with column number")
	@ParameterizedTest(name = "Given column number {0} or width {1} should lead to an exception")
	@CsvSource({ "-1, 0f", "16384, 0.0f", "0, -10f", "0, 255.01f", "0, 500f", })
	void setColumnWidthFailTest(int columnNumber, float width) {
		Worksheet worksheet = new Worksheet();
		assertThrows(Exception.class, () -> worksheet.setColumnWidth(columnNumber, width));
	}

	@DisplayName("Test of the failing setColumnWidth function with column address")
	@ParameterizedTest(name = "Given column address {1} (type: {0}) or width {2} should lead to an exception")
	@CsvSource({ "NULL, ''1, 0f", "STRING, '', 0.0f", "STRING, ':', 0.0f", "STRING, 'XFE', 0.0f", "STRING, 'A', -10f",
			"STRING, 'XFD', 255.01f", "STRING, 'A', 500f", })
	void setColumnWidthFailTest2(String sourceType, String sourceValue, float width) {
		String address = (String) TestUtils.createInstance(sourceType, sourceValue);
		Worksheet worksheet = new Worksheet();
		assertThrows(Exception.class, () -> worksheet.setColumnWidth(address, width));
	}

	@DisplayName("Test of the setCurrentColumnNumber function")
	@ParameterizedTest(name = "Given column number {0} should lead to the same currentColumnNumber on the worksheet")
	@CsvSource({ "0", "5", "16383", })
	void setCurrentColumnNumberTest(int column) {
		Worksheet worksheet = new Worksheet();
		assertEquals(0, worksheet.getCurrentColumnNumber());
		worksheet.goToNextColumn();
		worksheet.setCurrentColumnNumber(column);
		assertEquals(column, worksheet.getCurrentColumnNumber());
	}

	@DisplayName("Test of the failing setCurrentColumnNumber function")
	@ParameterizedTest(name = "Given column number {0} should lead to an exception")
	@CsvSource({ "-1", "-10", "16384", })
	void setCurrentColumnNumberFailTest(int column) {
		Worksheet worksheet = new Worksheet();
		assertThrows(RangeException.class, () -> worksheet.setCurrentColumnNumber(column));
	}

	@DisplayName("Test of the getColumn function")
	@Test()
	void getColumnTest() {
		Worksheet worksheet = new Worksheet();
		List<Object> values = getCellValues(worksheet);
		List<Cell> column = worksheet.getColumn(1);
		AssertColumnValues(column, values);
	}

	@DisplayName("Test of the getColumn function with a column address")
	@Test()
	void getColumnTest2() {
		Worksheet worksheet = new Worksheet();
		List<Object> values = getCellValues(worksheet);
		List<Cell> column = worksheet.getColumn("B");
		AssertColumnValues(column, values);
	}

	@DisplayName("Test of the getColumn function when no values are applying")
	@Test()
	void getColumnTest3() {
		Worksheet worksheet = new Worksheet();
		worksheet.addCell(22, "A1");
		worksheet.addCell(false, "C3");
		List<Cell> column = worksheet.getColumn(1);
		assertEquals(0, column.size());
	}

	private void AssertColumnValues(List<Cell> givenList, List<Object> expectedValues) {
		assertEquals(expectedValues.size(), givenList.size());
		for (int i = 0; i < expectedValues.size(); i++) {
			assertEquals(expectedValues.get(i), givenList.get(i).getValue());
		}
	}

	private static List<Object> getCellValues(Worksheet worksheet) {
		List<Object> expectedList = new ArrayList<Object>();
		expectedList.add(23);
		expectedList.add("test");
		expectedList.add(true);
		worksheet.addCell(22, "A1");
		worksheet.addCell(expectedList.get(0), "B1");
		worksheet.addCell(expectedList.get(1), "B2");
		worksheet.addCell(expectedList.get(2), "B3");
		worksheet.addCell(false, "C2");
		return expectedList;
	}

}
