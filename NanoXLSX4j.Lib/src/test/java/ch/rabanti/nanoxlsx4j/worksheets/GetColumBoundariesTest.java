package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.Worksheet;

public class GetColumBoundariesTest {

	@DisplayName("Test of the getLastColumnNumber function with an empty worksheet")
	@Test()
	void getLastColumnNumberTest() {
		Worksheet worksheet = new Worksheet();
		int column = worksheet.getLastColumnNumber();
		assertEquals(-1, column);
	}

	@DisplayName("Test of the getLastColumnNumber function with defined columns on an empty worksheet")
	@Test()
	void getLastColumnNumberTest2() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(0);
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(2);
		int column = worksheet.getLastColumnNumber();
		assertEquals(2, column);
	}

	@DisplayName("Test of the getLastColumnNumber function with defined columns on an empty worksheet, where the column definition has gaps")
	@Test()
	void getLastColumnNumberTest3() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(0);
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(10);
		int column = worksheet.getLastColumnNumber();
		assertEquals(10, column);
	}

	@DisplayName("Test of the getLastColumnNumber function with defined columns where cells are defined below the last column")
	@Test()
	void getLastColumnNumberTest4() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(0);
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(10);
		worksheet.addCell("test",
				"E5");
		int column = worksheet.getLastColumnNumber();
		assertEquals(10, column);
	}

	@DisplayName("Test of the getLastColumnNumber function with defined columns where cells are defined above the last column")
	@Test()
	void getLastColumnNumberTest5() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(0);
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(2);
		worksheet.addCell("test",
				"F5");
		int column = worksheet.getLastColumnNumber();
		assertEquals(5, column);
	}

	@DisplayName("Test of the getLastColumnNumber function with an explicitly defined, empty cell besides other column definitions")
	@ParameterizedTest(name = "Given empty cell address {0} should lead to the appropriate last column {1}")
	@CsvSource({
			"'F5', 5",
			"'A1', 4", })
	void getLastColumnNumberTest6(String emptyCellAddress, int expectedLastColumn) {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(3);
		worksheet.addHiddenColumn(4);
		worksheet.addCell(null, emptyCellAddress);
		int column = worksheet.getLastColumnNumber();
		assertEquals(expectedLastColumn, column);
	}

	@DisplayName("Test of the getFirstColumnNumber function with an empty worksheet")
	@Test()
	void getFirstColumnNumberTest() {
		Worksheet worksheet = new Worksheet();
		int column = worksheet.getFirstColumnNumber();
		assertEquals(-1, column);
	}

	@DisplayName("Test of the getFirstColumnNumber function with defined columns on an empty worksheet")
	@Test()
	void getFirstColumnNumberTest2() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(2);
		worksheet.addHiddenColumn(3);
		int column = worksheet.getFirstColumnNumber();
		assertEquals(1, column);
	}

	@DisplayName("Test of the getFirstColumnNumber function with defined columns on an empty worksheet, where the column definition has gaps")
	@Test()
	void getFirstColumnNumberTest3() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(2);
		worksheet.addHiddenColumn(10);
		int column = worksheet.getFirstColumnNumber();
		assertEquals(1, column);
	}

	@DisplayName("Test of the getFirstColumnNumber function with defined columns where cells are defined above the first column")
	@Test()
	void getFirstColumnNumberTest4() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(3);
		worksheet.addHiddenColumn(8);
		worksheet.addHiddenColumn(10);
		worksheet.addCell("test",
				"E5");
		int column = worksheet.getFirstColumnNumber();
		assertEquals(3, column);
	}

	@DisplayName("Test of the getFirstColumnNumber function with defined columns where cells are defined below the first column")
	@Test()
	void getFirstColumnNumberTest5() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(7);
		worksheet.addHiddenColumn(8);
		worksheet.addHiddenColumn(9);
		worksheet.addCell("test",
				"F5");
		int column = worksheet.getFirstColumnNumber();
		assertEquals(5, column);
	}

	@DisplayName("Test of the getFirstColumnNumber function with an explicitly defined, empty cell besides other column definitions")
	@ParameterizedTest(name = "Given empty cell address {0} should lead to the appropriate last column {1}")
	@CsvSource({
			"'F5', 3",
			"'A1', 0", })
	void getFirstColumnNumberTest6(String emptyCellAddress, int expectedFirstColumn) {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(3);
		worksheet.addHiddenColumn(4);
		worksheet.addCell(null, emptyCellAddress);
		int column = worksheet.getFirstColumnNumber();
		assertEquals(expectedFirstColumn, column);
	}

	@DisplayName("Test of the getLastDataColumnNumber function with an empty worksheet")
	@Test()
	void getLastDataColumnNumberTest() {
		Worksheet worksheet = new Worksheet();
		int column = worksheet.getLastDataColumnNumber();
		assertEquals(-1, column);
	}

	@DisplayName("Test of the getLastDataColumnNumber function with defined columns on an empty worksheet")
	@Test()
	void getLastDataColumnNumberTest2() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(0);
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(2);
		int column = worksheet.getLastDataColumnNumber();
		assertEquals(-1, column);
	}

	@DisplayName("Test of the getLastDataColumnNumber function with defined columns where cells are defined below the last column")
	@Test()
	void getLastDataColumnNumberTest3() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(0);
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(10);
		worksheet.addCell("test",
				"E5");
		int column = worksheet.getLastDataColumnNumber();
		assertEquals(4, column);
	}

	@DisplayName("Test of the getLastDataColumnNumber function with defined columns where cells are defined below the last column")
	@Test()
	void getLastDataColumnNumberTest4() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(0);
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(10);
		worksheet.addCell("test",
				"E5");
		int column = worksheet.getLastDataColumnNumber();
		assertEquals(4, column);
	}

	@DisplayName("Test of the getFirstDataColumnNumber function with an empty worksheet")
	@Test()
	void getFirstDataColumnNumberTest() {
		Worksheet worksheet = new Worksheet();
		int column = worksheet.getFirstDataColumnNumber();
		assertEquals(-1, column);
	}

	@DisplayName("Test of the getFirstDataColumnNumber function with defined columns on an empty worksheet")
	@Test()
	void getFirstDataColumnNumberTest2() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(0);
		worksheet.addHiddenColumn(1);
		worksheet.addHiddenColumn(2);
		int column = worksheet.getFirstDataColumnNumber();
		assertEquals(-1, column);
	}

	@DisplayName("Test of the getFirstDataColumnNumber function with defined columns where cells are defined above the first column")
	@Test()
	void getFirstDataColumnNumberTest3() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(2);
		worksheet.addHiddenColumn(3);
		worksheet.addHiddenColumn(10);
		worksheet.addCell("test",
				"E5");
		int column = worksheet.getFirstDataColumnNumber();
		assertEquals(4, column);
	}

	@DisplayName("Test of the getFirstDataColumnNumber function with defined columns where cells are defined below the last column")
	@Test()
	void getFirstDataColumnNumberTest4() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(2);
		worksheet.addHiddenColumn(3);
		worksheet.addHiddenColumn(10);
		worksheet.addCell("test",
				"F5");
		int column = worksheet.getFirstDataColumnNumber();
		assertEquals(5, column);
	}

	@DisplayName("Test of the getFirstDataColumnNumber and getLastDataColumnNumber functions with an explicitly defined, empty cell besides other column definitions")
	@ParameterizedTest(name = "Given empty cell address {0} should lead to -1 as  first and last column")
	@CsvSource({
			"'F5'",
			"'A1'" })
	void getFirstOrLastDataColumnNumberTest(String emptyCellAddress) {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(3);
		worksheet.addHiddenColumn(4);
		worksheet.addCell(null, emptyCellAddress);
		int minColumn = worksheet.getFirstDataColumnNumber();
		int maxColumn = worksheet.getLastDataColumnNumber();
		assertEquals(-1, minColumn);
		assertEquals(-1, maxColumn);
	}

	@DisplayName("Test of the getFirstDataColumnNumber and getLastDataColumnNumber functions with exactly one defined cell")
	@Test()
	void getFirstOrLastDataColumnNumberTest2() {
		Worksheet worksheet = new Worksheet();
		worksheet.addHiddenColumn(2);
		worksheet.addHiddenColumn(3);
		worksheet.addHiddenColumn(10);
		worksheet.addCell("test",
				"F5");
		int minColumn = worksheet.getFirstDataColumnNumber();
		int maxColumn = worksheet.getLastDataColumnNumber();
		assertEquals(5, minColumn);
		assertEquals(5, maxColumn);
	}

}
