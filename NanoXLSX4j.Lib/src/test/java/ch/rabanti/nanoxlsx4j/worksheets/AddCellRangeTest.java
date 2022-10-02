package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;

public class AddCellRangeTest {

	public enum RangeType {
		OneColumn, OneRow, ThreeColumnsFourRows, FourColumnsThreeRows
	}

	public enum TestType {
		RandomList, CellList
	}

	@DisplayName("Test of the AddCellRange function for a random list or list of nested cells with start and end address")
	@ParameterizedTest(name = "Given start column {0} and row {1} with the range type {2} should lead to a valid cell range of 12 entries ({3})")
	@CsvSource({ "0,0, OneColumn, RandomList", "7, 27, OneRow, RandomList", "5, 13, FourColumnsThreeRows, RandomList",
			"22, 11, ThreeColumnsFourRows, RandomList", "0,0, OneColumn, CellList", "7, 27, OneRow, CellList",
			"5, 13, FourColumnsThreeRows, CellList", "22, 11, ThreeColumnsFourRows, CellList", })
	void addCellRangeTest1(int startColumn, int startRow, RangeType type, TestType testType) {
		ListTuple data = getList(startColumn, startRow, type, testType);
		Worksheet worksheet = new Worksheet();
		Address startAddress = new Address(startColumn, startRow);
		Address endAddress = ListTuple.getEndAddress(startColumn, startRow, type);

		assertEquals(0, worksheet.getCells().size());
		worksheet.addCellRange(data.getValues(), startAddress, endAddress);
		assertRange(worksheet, data);
	}

	@DisplayName("Test of the addCellRange function for a random list or list of nested cells with start and address with a style")
	@ParameterizedTest(name = "Given start column {0} and row {1} with the range type {2} should lead to a valid cell range of 12 entries ({3})")
	@CsvSource({ "0, 0, OneColumn, RandomList", "7, 27, OneRow, RandomList", "5, 13, FourColumnsThreeRows, RandomList",
			"22, 11, ThreeColumnsFourRows, RandomList", "0, 0, OneColumn, CellList", "7, 27, OneRow, CellList",
			"5, 13, FourColumnsThreeRows, CellList", "22, 11, ThreeColumnsFourRows, CellList", })
	void addCellRangeTest2(int startColumn, int startRow, RangeType type, TestType testType) {
		ListTuple data = getList(startColumn, startRow, type, testType);
		Worksheet worksheet = new Worksheet();
		Address startAddress = new Address(startColumn, startRow);
		Address endAddress = ListTuple.getEndAddress(startColumn, startRow, type);

		assertEquals(0, worksheet.getCells().size());
		worksheet.addCellRange(data.getValues(), startAddress, endAddress, BasicStyles.Bold());
		assertRange(worksheet, data);
		assertRangeStyle(worksheet, data, BasicStyles.Bold());
	}

	@DisplayName("Test of the addCellRange function for a random list or list of nested cells with start and end address with a active style on the workbook")
	@ParameterizedTest(name = "Given start column {0} and row {1} with the range type {2} should lead to a valid cell range of 12 entries ({3})")
	@CsvSource({ "0, 0, OneColumn, RandomList", "7, 27, OneRow, RandomList", "5, 13, FourColumnsThreeRows, RandomList",
			"22, 11, ThreeColumnsFourRows, RandomList", "0, 0, OneColumn, CellList", "7, 27, OneRow, CellList",
			"5, 13, FourColumnsThreeRows, CellList", "22, 11, ThreeColumnsFourRows, CellList", })
	void addCellRangeTest3(int startColumn, int startRow, RangeType type, TestType testType) {
		ListTuple data = getList(startColumn, startRow, type, testType);
		Worksheet worksheet = new Worksheet();
		worksheet.setActiveStyle(BasicStyles.BorderFrame());
		Address startAddress = new Address(startColumn, startRow);
		Address endAddress = ListTuple.getEndAddress(startColumn, startRow, type);

		assertEquals(0, worksheet.getCells().size());
		worksheet.addCellRange(data.getValues(), startAddress, endAddress);
		assertRange(worksheet, data);
		assertRangeStyle(worksheet, data, BasicStyles.BorderFrame());
	}

	@DisplayName("Test of the addCellRange function for a random list or list of nested cells with a range as string")
	@ParameterizedTest(name = "Given range {0} with the range type {1} should lead to a valid cell range of 12 entries ({2})")
	@CsvSource({ "A1:A12, OneColumn, RandomList", "H28:S28, OneRow, RandomList",
			"F14:I16, FourColumnsThreeRows, RandomList", "T12:V15, ThreeColumnsFourRows, RandomList",
			"A1:A12, OneColumn, CellList", "H28:S28, OneRow, CellList", "F14:I16, FourColumnsThreeRows, CellList",
			"T12:V15, ThreeColumnsFourRows, CellList", })
	void addCellRangeTest4(String range, RangeType type, TestType testType) {
		Range cellRange = Cell.resolveCellRange(range);
		ListTuple data = getList(cellRange.StartAddress.Column, cellRange.StartAddress.Row, type, testType);
		Worksheet worksheet = new Worksheet();

		assertEquals(0, worksheet.getCells().size());
		worksheet.addCellRange(data.getValues(), range);
		assertRange(worksheet, data);
	}

	@DisplayName("Test of the addCellRange function for a random list or list of nested cells with a range as string and a style")
	@ParameterizedTest(name = "Given range {0} with the range type {1} should lead to a valid cell range of 12 entries ({2})")
	@CsvSource({ "A1:A12, OneColumn, RandomList", "H28:S28, OneRow, RandomList",
			"F14:I16, FourColumnsThreeRows, RandomList", "T12:V15, ThreeColumnsFourRows, RandomList",
			"A1:A12, OneColumn, CellList", "H28:S28, OneRow, CellList", "F14:I16, FourColumnsThreeRows, CellList",
			"T12:V15, ThreeColumnsFourRows, CellList", })
	void addCellRangeTest5(String range, RangeType type, TestType testType) {
		Range cellRange = Cell.resolveCellRange(range);
		ListTuple data = getList(cellRange.StartAddress.Column, cellRange.StartAddress.Row, type, testType);
		Worksheet worksheet = new Worksheet();

		assertEquals(0, worksheet.getCells().size());
		worksheet.addCellRange(data.getValues(), range, BasicStyles.BorderFrame());
		assertRange(worksheet, data);
		assertRangeStyle(worksheet, data, BasicStyles.BorderFrame());
	}

	@DisplayName("Test of the addCellRange function for a random list or list of nested cells with a range as string and an active style on the worksheet")
	@ParameterizedTest(name = "Given range {0} with the range type {1} should lead to a valid cell range of 12 entries ({2})")
	@CsvSource({ "A1:A12, OneColumn, RandomList", "H28:S28, OneRow, RandomList",
			"F14:I16, FourColumnsThreeRows, RandomList", "T12:V15, ThreeColumnsFourRows, RandomList",
			"A1:A12, OneColumn, CellList", "H28:S28, OneRow, CellList", "F14:I16, FourColumnsThreeRows, CellList",
			"T12:V15, ThreeColumnsFourRows, CellList", })
	void addCellRangeTest6(String range, RangeType type, TestType testType) {
		Range cellRange = Cell.resolveCellRange(range);
		ListTuple data = getList(cellRange.StartAddress.Column, cellRange.StartAddress.Row, type, testType);
		Worksheet worksheet = new Worksheet();
		worksheet.setActiveStyle(BasicStyles.BorderFrameHeader());

		assertEquals(0, worksheet.getCells().size());
		worksheet.addCellRange(data.getValues(), range);
		assertRange(worksheet, data);
		assertRangeStyle(worksheet, data, BasicStyles.BorderFrameHeader());
	}

	@DisplayName("Test of the failing addCellRange function with a deviating row and column number")
	@ParameterizedTest(name = "Given start column {0} and row {1} with the range type {2} should lead to a exception")
	@CsvSource({ "0, 0, OneColumn", "7, 27, OneRow", "5, 13, FourColumnsThreeRows", "22, 11, ThreeColumnsFourRows", })
	void addCellRangeFailingTest(int startColumn, int startRow, RangeType type) {
		ListTuple data = getRandomList(0, 0, type);
		Worksheet worksheet = new Worksheet();
		Address startAddress = new Address(startColumn, startRow);
		Address endAddress = ListTuple.getEndAddress(startColumn + 1, startRow + 1, type);

		assertEquals(0, worksheet.getCells().size());
		assertThrows(RangeException.class, () -> worksheet.addCellRange(data.getValues(), startAddress, endAddress));
	}

	@DisplayName("Test of the failing addCellRange function with a deviating range definition (string)")
	@ParameterizedTest(name = "Given range {1} with the type type {2} should lead to a exception, where {0} would be valid")
	@CsvSource({ "A1:A12, A1:A13, OneColumn", "H28:S28,H28:S29, OneRow", "F14:I16,F14:J16, FourColumnsThreeRows",
			"T12:V15, T12:W15, ThreeColumnsFourRows", })
	void addCellRangeFailingTest2(String givenRange, String falseRange, RangeType type) {
		Range cellRange = Cell.resolveCellRange(givenRange);
		ListTuple data = getRandomList(cellRange.StartAddress.Column, cellRange.StartAddress.Row, type);
		Worksheet worksheet = new Worksheet();

		assertEquals(0, worksheet.getCells().size());
		assertThrows(RangeException.class, () -> worksheet.addCellRange(data.getValues(), falseRange));
	}

	private void assertRange(Worksheet worksheet, ListTuple expectedData) {
		assertEquals(expectedData.size(), worksheet.getCells().size());
		for (int i = 0; i < expectedData.size(); i++) {
			String expectedAddress = expectedData.getAddresses().get(i).getAddress();
			assertTrue(worksheet.getCells().containsKey(expectedAddress));
			assertEquals(expectedData.getValues().get(i), worksheet.getCells().get(expectedAddress).getValue());
			assertEquals(expectedData.getTypes().get(i), worksheet.getCells().get(expectedAddress).getDataType());
		}
	}

	private void assertRangeStyle(Worksheet worksheet, ListTuple expectedData, Style expectedSourceStyle) {
		for (int i = 0; i < expectedData.size(); i++) {
			String expectedAddress = expectedData.getAddresses().get(i).getAddress();
			if (expectedData.getStyles().get(i) == null) {
				assertTrue(expectedSourceStyle.equals(worksheet.getCells().get(expectedAddress).getCellStyle()));
			} else {
				Style mergedStyle = (Style) expectedSourceStyle.copy();
				mergedStyle.append(expectedData.getStyles().get(i));
				assertTrue(mergedStyle.equals(worksheet.getCells().get(expectedAddress).getCellStyle()));
			}
		}
	}

	private static ListTuple getList(int startColumn, int startRow, RangeType type, TestType testType) {
		ListTuple data;
		if (testType == TestType.RandomList) {
			data = getRandomList(startColumn, startRow, type);
		} else {
			data = getCellList(startColumn, startRow, type);
		}
		return data;
	}

	private static ListTuple getRandomList(int startColumn, int startRow, RangeType type) {
		return getRandomList(startColumn, startRow, type, true);
	}

	private static ListTuple getRandomList(int startColumn, int startRow, RangeType type, boolean addNull) {
		ListTuple list = new ListTuple(startColumn, startRow, type);
		// IMPORTANT: The list must contain 12 entries
		list.add(22, Cell.CellType.NUMBER);
		list.add(-0.55f, Cell.CellType.NUMBER);
		list.add(22.22d, Cell.CellType.NUMBER);
		list.add(true, Cell.CellType.BOOL);
		list.add(false, Cell.CellType.BOOL);
		list.add("", Cell.CellType.STRING);
		list.add("test", Cell.CellType.STRING);
		list.add((byte) 64, Cell.CellType.NUMBER);
		list.add(77777l, Cell.CellType.NUMBER);
		Calendar calendar = Calendar.getInstance();
		calendar.set(2020, 11, 1, 11, 22, 13);
		list.add(calendar.getTime(), Cell.CellType.DATE);

		list.add(LocalTime.of(13, 16, 22), Cell.CellType.TIME);
		if (addNull) {
			list.add(null, Cell.CellType.EMPTY);
		} else {
			list.add("substitute", Cell.CellType.STRING);
		}
		return list;
	}

	private static ListTuple getCellList(int startColumn, int startRow, RangeType type) {
		return getCellList(startColumn, startRow, type, true);
	}

	private static ListTuple getCellList(int startColumn, int startRow, RangeType type, boolean addNull) {
		Calendar calendar = Calendar.getInstance();
		calendar.set(2020, 10, 1, 11, 22, 13);
		ListTuple list = new ListTuple(startColumn, startRow, type);
		// IMPORTANT: The list must contain 12 entries
		list.add(new Cell(23, Cell.CellType.DEFAULT, "X1"), Cell.CellType.NUMBER);
		list.add(new Cell(-0.65f, Cell.CellType.DEFAULT, "X2"), Cell.CellType.NUMBER);
		list.add(new Cell(42.22d, Cell.CellType.DEFAULT, "X3"), Cell.CellType.NUMBER);
		list.add(new Cell(false, Cell.CellType.DEFAULT, "X4"), Cell.CellType.BOOL);
		list.add(new Cell(true, Cell.CellType.DEFAULT, "X5"), Cell.CellType.BOOL);
		list.add(new Cell("test2", Cell.CellType.DEFAULT, "X6"), Cell.CellType.STRING);
		list.add(new Cell("", Cell.CellType.DEFAULT, "X7"), Cell.CellType.STRING);
		list.add(new Cell((byte) 64, Cell.CellType.DEFAULT, "X8"), Cell.CellType.NUMBER);
		list.add(new Cell(88888l, Cell.CellType.DEFAULT, "X9"), Cell.CellType.NUMBER);
		list.add(new Cell(calendar.getTime(), Cell.CellType.DEFAULT, "X10"), Cell.CellType.DATE);
		list.add(new Cell(LocalTime.of(13, 15, 22), Cell.CellType.DEFAULT, "X11"), Cell.CellType.TIME);
		if (addNull) {
			list.add(new Cell(null, Cell.CellType.DEFAULT, "X12"), Cell.CellType.EMPTY);
		} else {
			list.add(new Cell("substitute2", Cell.CellType.DEFAULT, "X12"), Cell.CellType.STRING);
		}
		return list;
	}

	public static class ListTuple {
		private final List<Address> preparedAddresses;
		private int currentIndex = 0;

		private final List<Object> values;
		private final List<Cell.CellType> types;
		private final List<Address> addresses;
		private final List<Style> styles;
		private final int size;

		public List<Object> getValues() {
			return values;
		}

		public List<Cell.CellType> getTypes() {
			return types;
		}

		public List<Address> getAddresses() {
			return addresses;
		}

		public List<Style> getStyles() {
			return styles;
		}

		public int size() {
			return size;
		}

		public ListTuple(int startColumn, int startRow, RangeType rangeType) {
			values = new ArrayList<Object>();
			types = new ArrayList<Cell.CellType>();
			addresses = new ArrayList<Address>();
			styles = new ArrayList<>();
			size = 12;
			Address endAddress = getEndAddress(startColumn, startRow, rangeType);
			preparedAddresses = Cell.getCellRange(startColumn, startRow, endAddress.Column, endAddress.Row);

		}

		public void add(Object value, Cell.CellType type) {
			if (value instanceof Cell) {
				values.add(((Cell) value).getValue());
			} else {
				values.add(value);
			}
			types.add(type);
			addresses.add(preparedAddresses.get(currentIndex));
			if (type.equals(Cell.CellType.DATE)) {
				styles.add(BasicStyles.DateFormat());
			} else if (type.equals(Cell.CellType.TIME)) {
				styles.add(BasicStyles.TimeFormat());
			} else {
				styles.add(null);
			}
			currentIndex++;
		}

		public static Address getEndAddress(int startColumn, int startRow, RangeType rangeType) {
			switch (rangeType) {
			case OneColumn:
				return new Address(startColumn, startRow + 11);
			case OneRow:
				return new Address(startColumn + 11, startRow);
			case ThreeColumnsFourRows:
				return new Address(startColumn + 2, startRow + 3);
			default:
				return new Address(startColumn + 3, startRow + 2);
			}
		}

	}

}
