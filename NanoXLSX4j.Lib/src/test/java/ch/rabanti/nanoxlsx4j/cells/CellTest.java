package ch.rabanti.nanoxlsx4j.cells;

import static ch.rabanti.nanoxlsx4j.TestUtils.buildTime;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.cells.types.CellTypeUtils;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.CellXf;
import ch.rabanti.nanoxlsx4j.styles.Style;
import ch.rabanti.nanoxlsx4j.styles.StyleRepository;

public class CellTest {

	private final Cell cell;
	private final Class<?> cellType;
	private final Object cellValue;
	private final Address cellAddress;
	private final CellTypeUtils utils;

	public CellTest() {
		this.utils = new CellTypeUtils(String.class);
		this.cellType = String.class;
		this.cellValue = "value";
		this.cellAddress = this.utils.getCellAddress();
		this.cell = this.utils.createVariantCell(this.cellValue, this.cellAddress, BasicStyles.BoldItalic());
	}

	@DisplayName("Test of the default constructor")
	@Test()
	void cellConstructorTest() {
		Cell cell = new Cell();
		assertEquals(Cell.CellType.DEFAULT, cell.getDataType());
		assertNull(cell.getValue());
		assertNull(cell.getCellStyle());
		assertEquals("A1", cell.getCellAddress()); // The address comes from initial row and column of 0
	}

	@DisplayName("Test of the constructor with value and type")
	@ParameterizedTest(name = "Given Object {0} and type {2} should lead to {1} as {3}")
	@CsvSource({
			"string, string, STRING, STRING",
			"true, true, BOOL, BOOL",
			"false, false, BOOL, BOOL",
			"22, 22, NUMBER, NUMBER",
			"22.1f, 22.1f, NUMBER, NUMBER",
			"=B1, =B1, FORMULA, FORMULA",
			"'','' , DEFAULT, STRING",
			",, DEFAULT, EMPTY",
			"'', , EMPTY, EMPTY",
			"11, , EMPTY, EMPTY", })
	void cellConstructorTest2(Object givenValue, Object expectedValue, Cell.CellType givenType, Cell.CellType expectedType) throws Exception {
		Cell cell = new Cell(givenValue, givenType);
		assertEquals(expectedType, cell.getDataType());
		assertEquals(expectedValue, cell.getValue());
		assertNull(cell.getCellStyle());
		assertEquals("A1", cell.getCellAddress()); // The address comes from initial row and column of 0
	}

	@DisplayName("Test of the constructor with value and type with a date as value")
	@Test()
	void cellConstructorTest2b() {
		Date date = new Date();
		Cell cell = new Cell(date, Cell.CellType.DEFAULT);
		assertEquals(Cell.CellType.DATE, cell.getDataType());
		assertEquals(date, cell.getValue());
		assertNotNull(cell.getCellStyle());
		assertEquals(BasicStyles.DateFormat(), cell.getCellStyle());
	}

	@DisplayName("Test of the constructor with value, type and address string")
	@ParameterizedTest(name = "Given Object {0} and type {2} should lead to {1} as {3}")
	@CsvSource({
			"string, string, STRING, STRING, C$7, C$7",
			"true, true, BOOL, BOOL, D100, D100",
			"false, false, BOOL, BOOL, $A$2, $A$2",
			"22, 22, NUMBER, NUMBER, $B5, $B5",
			"22.1f, 22.1f, NUMBER, NUMBER, AA10, AA10",
			"=B1, =B1, FORMULA, FORMULA, $A$15, $A$15",
			"'','' , DEFAULT, STRING, r1, R1",
			",, DEFAULT, EMPTY, $ab$999, $AB$999",
			"'', , EMPTY, EMPTY, c$90000, C$90000",
			"11, , EMPTY, EMPTY, a17, A17", })
	void cellConstructorTest3(Object givenValue, Object expectedValue, Cell.CellType givenType, Cell.CellType expectedType, String givenAddress, String expectedAddress) {
		Cell cell = new Cell(givenValue, givenType, givenAddress);
		assertEquals(expectedType, cell.getDataType());
		assertEquals(expectedValue, cell.getValue());
		assertNull(cell.getCellStyle());
		assertEquals(expectedAddress, cell.getCellAddress());
	}

	@DisplayName("Test of the constructor with value, type and address Object or row and column")
	@ParameterizedTest(name = "Given Object {0} and {2} should lead to {1} as {3}")
	@CsvSource({
			"string, string, STRING, STRING, 2, 6, C7",
			"true, true, BOOL, BOOL, 3, 99, D100",
			"false, false, BOOL, BOOL, 0, 1, A2",
			"22, 22, NUMBER, NUMBER, 1, 4, B5",
			"22.1f, 22.1f, NUMBER, NUMBER, 26, 9, AA10",
			"=B1, =B1, FORMULA, FORMULA, 0, 14, A15",
			"'','' , DEFAULT, STRING, 17, 0, R1",
			", , DEFAULT, EMPTY, 27, 998, AB999",
			"'', , EMPTY, EMPTY, 2, 89999, C90000",
			"11, , EMPTY, EMPTY, 0, 16, A17", })
	void cellConstructorTest4(Object givenValue, Object expectedValue, Cell.CellType givenType, Cell.CellType expectedType, int givenColumn, int givenRow, String expectedAddress) {
		Address address = new Address(givenColumn, givenRow);
		Cell cell = new Cell(givenValue, givenType, address);
		Cell cell2 = new Cell(givenValue, givenType, givenColumn, givenRow);
		assertEquals(expectedType, cell.getDataType());
		assertEquals(expectedValue, cell.getValue());
		assertNull(cell.getCellStyle());
		assertEquals(expectedAddress, cell.getCellAddress());
		assertEquals(expectedType, cell2.getDataType());
		assertEquals(expectedValue, cell2.getValue());
		assertNull(cell2.getCellStyle());
		assertEquals(expectedAddress, cell2.getCellAddress());
	}

	@DisplayName("Test of the set function of the cellAddress field")
	@ParameterizedTest(name = "Given address {0} should lead to column {1}, row {2} and type {3}")
	@CsvSource({
			"A1, 0, 0, Default",
			"J100, 9, 99, Default",
			"$B2, 1, 1, FixedColumn",
			"B$2, 1, 1, FixedRow",
			"$B$2, 1, 1, FixedRowAndColumn", })
	void setAddressStringPropertyTest(String givenAddress, int expectedColumn, int expectedRow, Cell.AddressType expectedType) {
		this.cell.setCellAddress(givenAddress);
		assertEquals(this.cell.getCellAddress2().Column, expectedColumn);
		assertEquals(this.cell.getCellAddress2().Row, expectedRow);
		assertEquals(this.cell.getCellAddress2().Type, expectedType);
	}

	@DisplayName("Test of the get function of the CellAddressString field")
	@ParameterizedTest(name = "Given column {0}, row {1} and type {2} should lead to the address {3}")
	@CsvSource({
			"0, 0, Default, A1",
			"9, 99,Default, J100",
			"1, 1, FixedColumn, $B2",
			"1, 1, FixedRow, B$2",
			"1, 1, FixedRowAndColumn, $B$2", })
	void getAddressStringPropertyTest(int givenColumn, int givenRow, Cell.AddressType givenTyp, String expectedAddress) {
		this.cell.setCellAddressType(givenTyp);
		this.cell.setColumnNumber(givenColumn);
		this.cell.setRowNumber(givenRow);
		assertEquals(this.cell.getCellAddress(), expectedAddress);
	}

	@DisplayName("Test of the set function of the CellAddress field, as well as get functions of columnNumber, RowNumber and AddressType")
	@ParameterizedTest(name = "Given address {0} should lead to column {1}, row {2} and type {3}")
	@CsvSource({
			"A1, 0, 0, Default",
			"XFD1048576, 16383, 1048575, Default",
			"$B2, 1, 1, FixedColumn",
			"B$2, 1, 1, FixedRow",
			"$B$2, 1, 1, FixedRowAndColumn", })
	void setAddressPropertyTest(String givenAddress, int expectedColumn, int expectedRow, Cell.AddressType expectedType) {
		Address given = new Address(givenAddress);
		this.cell.setCellAddress2(given);
		assertEquals(this.cell.getColumnNumber(), expectedColumn);
		assertEquals(this.cell.getRowNumber(), expectedRow);
		assertEquals(this.cell.getCellAddressType(), expectedType);
	}

	@DisplayName("Test of the get function of the CellAddress field, as well as set functions of columnNumber, RowNumber and AddressType")
	@ParameterizedTest(name = "Given column {0}, row {1} and type {2} should lead to the address {3}")
	@CsvSource({
			"0, 0, Default, A1",
			"16383, 1048575, Default, XFD1048576",
			"1, 1, FixedColumn, $B2",
			"1, 1, FixedRow, B$2",
			"1, 1, FixedRowAndColumn, $B$2", })
	void getAddressPropertyTest(int givenColumn, int givenRow, Cell.AddressType givenTyp, String expectedAddress) {
		this.cell.setColumnNumber(givenColumn);
		this.cell.setRowNumber(givenRow);
		this.cell.setCellAddressType(givenTyp);
		Address expected = new Address(expectedAddress);
		assertEquals(this.cell.getCellAddress2(), expected);
	}

	@DisplayName("Test of the get function of the Style field")
	@Test()
	void cellStyleTest() {
		Calendar calendarInstance = Calendar.getInstance();
		calendarInstance.set(2020, 11, 1, 12, 30, 22);
		Date givenDate = calendarInstance.getTime();
		Cell cell = utils.createVariantCell(givenDate, this.cellAddress);
		Style expectedStyle = BasicStyles.DateFormat();
		assertEquals(expectedStyle, cell.getCellStyle()); // Note: assertEquals fails here because of Object
															// reference comparison
	}

	@DisplayName("Test of the get function of the Style field, when no style was assigned")
	@Test()
	void cellStyleTest2() {
		Cell cell = new Cell(42, Cell.CellType.NUMBER, this.cellAddress);
		assertNull(cell.getCellStyle());
	}

	@DisplayName("Test of the set function of the Style filed")
	@Test()
	void cellStyleTest3() {
		Cell cell = utils.createVariantCell(42, this.cellAddress);
		assertNull(cell.getCellStyle());
		Style style = BasicStyles.BoldItalic();
		Style returnedStyle = cell.setStyle(style);
		assertNotNull(cell.getCellStyle());
		// Note: assertEquals fails here because of Object reference comparison
		assertEquals(style, cell.getCellStyle());
		assertEquals(style, returnedStyle);
		assertFalse(StyleRepository.getInstance().getStyles().isEmpty());
	}

	@DisplayName("Test of the set function of the Style field where the style repository is unmanaged")
	@Test()
	void cellStyleTest3b() {
		Cell cell = utils.createVariantCell(42, this.cellAddress);
		assertNull(cell.getCellStyle());
		StyleRepository.getInstance().flushStyles();
		assertEquals(0, StyleRepository.getInstance().getStyles().size());
		Style style = BasicStyles.BoldItalic();
		Style returnedStyle = cell.setStyle(style, true);
		// Note: assertEquals fails here because of Object reference comparison
		assertEquals(style, cell.getCellStyle());
		assertEquals(0, StyleRepository.getInstance().getStyles().size());
	}

	@DisplayName("Test of the failing set function of the Style field, when the style is null")
	@Test()
	void cellStyleFailTest() {
		Cell intCell = new Cell(42, Cell.CellType.NUMBER, this.cellAddress);
		Style style = null;
		assertThrows(StyleException.class, () -> {
			cell.setStyle(style);
		});
	}

	@DisplayName("Test of the set function of the Value field")
	@ParameterizedTest(name = "Initial type {0} with value {1} should lead with value {3} to type {2}")
	@CsvSource({
			"INTEGER, 0, NUMBER, STRING, 'test', STRING",
			"BOOLEAN, true, BOOL, INTEGER, 1, NUMBER",
			"FLOAT, 22.5, NUMBER, INTEGER, 22, NUMBER",
			"STRING, 'True', STRING, BOOLEAN, true, BOOL",
			"NULL, '', EMPTY, INTEGER, 22, NUMBER",
			"STRING, 'test', STRING, NULL, '', EMPTY", })
	void cellValueTest(String initialSourceType, String initialSourceValue, Cell.CellType initialType, String givenSourceType, String givenSourceValue, Cell.CellType expectedType) {
		Object initialValue = TestUtils.createInstance(initialSourceType, initialSourceValue);
		Object givenValue = TestUtils.createInstance(givenSourceType, givenSourceValue);
		Cell cell2 = new Cell(initialValue, initialType);
		assertEquals(cell2.getDataType(), initialType);
		cell2.setValue(givenValue);
		assertEquals(expectedType, cell2.getDataType());
	}

	@DisplayName("Test of the removeStyle method")
	@Test()
	void removeStyleTest() {
		Style style = BasicStyles.Bold();
		Cell floatCell = utils.createVariantCell(11.11f, this.cellAddress, style);
		assertEquals(style, floatCell.getCellStyle()); // Note: assertEquals fails here because of Object reference
														// comparison
		floatCell.removeStyle();
		assertNull(floatCell.getCellStyle());
	}

	@DisplayName("Test of the compareTo method (simplified use cases)")
	@ParameterizedTest(name = "Given cell address {0} compared to {1} should lead to {2}")
	@CsvSource({
			"A1, A1, 0",
			"A1, A2, -1",
			"A1, B1, -1",
			"A2, A1, 1",
			"B1, A1, 1",
			"A1, , -1", })
	void compareToTest(String cell1Address, String cell2Address, int expectedResult) {
		Cell cell1 = new Cell("Test", Cell.CellType.DEFAULT, cell1Address);
		Cell cell2 = null;
		if (cell2Address != null) {
			cell2 = new Cell("Test", Cell.CellType.DEFAULT, cell2Address);
		}
		assertEquals(expectedResult, cell1.compareTo(cell2));
	}

	@DisplayName("Test of the equals method (simplified use cases)")
	@Test()
	void equalsTest() {
		TestUtils.assertCellsEqual(null, null, "Data", this.cellAddress);
		TestUtils.assertCellsEqual(27, 27, 28, this.cellAddress);
		TestUtils.assertCellsEqual(0.27778f, 0.27778f, 0.27777f, this.cellAddress);
		TestUtils.assertCellsEqual("ABC",
				"ABC",
				"abc",
				this.cellAddress);
		TestUtils.assertCellsEqual("",
				"",
				" ",
				this.cellAddress);
		TestUtils.assertCellsEqual(true, true, false, this.cellAddress);
		TestUtils.assertCellsEqual(false, false, true, this.cellAddress);
		Calendar calendarInstance = Calendar.getInstance();
		calendarInstance.set(2020, 10, 9, 8, 7, 6);
		Date d1 = calendarInstance.getTime();
		calendarInstance.set(2020, 10, 9, 8, 7, 6);
		Date d2 = calendarInstance.getTime();
		calendarInstance.set(2020, 10, 9, 8, 7, 5);
		Date d3 = calendarInstance.getTime();
		TestUtils.assertCellsEqual(d1, d2, d3, this.cellAddress);
	}

	@DisplayName("Test failing of the equals method, when other cell is null (simplified use cases)")
	@Test()
	void equalsFailTest() {
		Cell cell1 = utils.createVariantCell("test", this.cellAddress, BasicStyles.BoldItalic());
		Cell cell2 = null;
		assertNotEquals(cell1, cell2);
	}

	@DisplayName("Test failing of the equals method, when the address of the other cell is different (simplified use cases)")
	@Test()
	void EqualsFailTest2() {
		Cell cell1 = utils.createVariantCell("test", this.cellAddress, BasicStyles.BoldItalic());
		Cell cell2 = utils.createVariantCell("test", new Address(99, 99), BasicStyles.BoldItalic());
		assertNotEquals(cell1, cell2);
	}

	@DisplayName("Test failing of the equals method, when the style of the other cell is different (simplified use cases)")
	@Test()
	void EqualsFailTest3() {
		Cell cell1 = utils.createVariantCell("test", this.cellAddress, BasicStyles.BoldItalic());
		Cell cell2 = utils.createVariantCell("test", this.cellAddress, BasicStyles.Italic());
		assertNotEquals(cell1, cell2);
	}

	@DisplayName("Test of the equals method, when two identical cells occur in different workbooks and worksheets (simplified use cases)")
	@Test()
	void EqualsFailTest4() {
		Workbook workbook1 = new Workbook(true);
		Workbook workbook2 = new Workbook(true);
		Cell cell1 = utils.createVariantCell("test", this.cellAddress, BasicStyles.BoldItalic());
		workbook1.getCurrentWorksheet().addCell(cell1, this.cellAddress.getAddress());
		Cell cell2 = utils.createVariantCell("test", this.cellAddress, BasicStyles.BoldItalic());
		workbook2.getCurrentWorksheet().addCell(cell2, this.cellAddress.getAddress());
		Cell cell3 = utils.createVariantCell("test", this.cellAddress, BasicStyles.BoldItalic());
		workbook2.addWorksheet("2nd");
		workbook2.getWorksheets().get(1).addCell(cell3, this.cellAddress.getAddress());
		assertEquals(cell1, cell2);
		assertEquals(cell2, cell3);
	}

	@DisplayName("Test of the compareTo method (simplified use cases)")
	@ParameterizedTest(name = "Given value {0} and type {1} should lead to type {2}")
	@CsvSource({
			"STRING, string, NUMBER, STRING",
			"INTEGER, 12, STRING, NUMBER",
			"DOUBLE, -12.12d, STRING, NUMBER",
			"BOOLEAN, true, STRING, BOOL",
			"BOOLEAN, false, STRING, BOOL",
			"STRING, =A2, FORMULA, FORMULA",
			"NULL, , STRING, EMPTY",
			"STRING, Actually not empty, EMPTY, EMPTY",
			"STRING, string, DEFAULT, STRING",
			"INTEGER, 0, DEFAULT, NUMBER",
			"FLOAT, -12.12f, DEFAULT, NUMBER",
			"BOOLEAN, true, DEFAULT, BOOL",
			"BOOLEAN, false, DEFAULT, BOOL",
			"STRING, =A2, DEFAULT, STRING",
			"STRING, '', DEFAULT, STRING",
			"NULL, , DEFAULT, EMPTY", })
	void resolveCellTypeTest(String sourceType, String givenStringValue, Cell.CellType givenCellType, Cell.CellType expectedCllType) {
		Object givenValue = TestUtils.createInstance(sourceType, givenStringValue);
		Cell cell = new Cell(givenValue, givenCellType, this.cellAddress);
		cell.resolveCellType();
		assertEquals(expectedCllType, cell.getDataType());
	}

	@DisplayName("Test of the compareTo method for dates and times")
	@Test()
	void resolveCellTypeTest2() {
		Cell dateCell = new Cell(new Date(), Cell.CellType.NUMBER, this.cellAddress);
		dateCell.resolveCellType();
		assertEquals(Cell.CellType.DATE, dateCell.getDataType());
		dateCell = new Cell(new Date(), Cell.CellType.DEFAULT, this.cellAddress);
		dateCell.resolveCellType();
		assertEquals(Cell.CellType.DATE, dateCell.getDataType());
		Cell timeCell = new Cell(buildTime(0, 0, 59), Cell.CellType.NUMBER, this.cellAddress);
		timeCell.resolveCellType();
		assertEquals(Cell.CellType.TIME, timeCell.getDataType());
		timeCell = new Cell(buildTime(0, 0, 59), Cell.CellType.DEFAULT, this.cellAddress);
		dateCell.resolveCellType();
		assertEquals(Cell.CellType.TIME, timeCell.getDataType());
	}

	@DisplayName("Test of the setCellLockedState method")
	@ParameterizedTest(name = "Given boolean {0} should lead to {1}")
	@CsvSource({
			"true, true",
			"true, false",
			"false, true",
			"false, false", })
	void setCellLockedState(boolean isLocked, boolean isHidden) {
		Cell cell = utils.createVariantCell("test", this.cellAddress);
		cell.setCellLockedState(isLocked, isHidden);
		assertEquals(isLocked, cell.getCellStyle().getCellXf().isLocked());
		assertEquals(isHidden, cell.getCellStyle().getCellXf().isHidden());
	}

	@DisplayName("Test of the setCellLockedState method when a cell style already exists")
	@ParameterizedTest(name = "Given boolean {0} should lead to {1}")
	@CsvSource({
			"true, true",
			"true, false",
			"false, true",
			"false, false", })
	void setCellLockedState2(boolean isLocked, boolean isHidden) {
		Style style = new Style();
		style.getFont().setItalic(true);
		style.getCellXf().setHorizontalAlign(CellXf.HorizontalAlignValue.justify);
		Cell cell = utils.createVariantCell("test", this.cellAddress, style);
		cell.setCellLockedState(isLocked, isHidden);
		assertEquals(isLocked, cell.getCellStyle().getCellXf().isLocked());
		assertEquals(isHidden, cell.getCellStyle().getCellXf().isHidden());
		assertTrue(cell.getCellStyle().getFont().isItalic());
		assertEquals(CellXf.HorizontalAlignValue.justify, cell.getCellStyle().getCellXf().getHorizontalAlign());
	}

	@DisplayName("Test of the getCellRange method with string as range")
	@ParameterizedTest(name = "Given range expression {0} should lead to the cell addresses {1}")
	@CsvSource({
			"A1:A1, 'A1'",
			"A1:A4, 'A1,A2,A3,A4'",
			"A1:B3, 'A1,A2,A3,B1,B2,B3'",
			"B3:A2, 'A2,A3,B2,B3'", })
	void getCellRangeTest(String range, String expectedAddresses) {
		List<Address> addresses = Cell.getCellRange(range);
		TestUtils.assertCellRange(expectedAddresses, addresses);
	}

	@DisplayName("Test of the getCellRange method with start and end address objects or strings as range")
	@ParameterizedTest(name = "Given start address {0} and end address {1} should lead to the cell addresses {2}")
	@CsvSource({
			"A1, A1, 'A1'",
			"A1, A4, 'A1,A2,A3,A4'",
			"A1, B3, 'A1,A2,A3,B1,B2,B3'",
			"B3, A2, 'A2,A3,B2,B3'", })
	void getCellRangeTest2(String startAddress, String endAddress, String expectedAddresses) {
		Address start = new Address(startAddress);
		Address end = new Address(endAddress);
		List<Address> addresses = Cell.getCellRange(startAddress, endAddress);
		TestUtils.assertCellRange(expectedAddresses, addresses);
		addresses = Cell.getCellRange(start, end);
		TestUtils.assertCellRange(expectedAddresses, addresses);
	}

	@DisplayName("Test of the getCellRange method with start/end row and column numbers as range")
	@ParameterizedTest(name = "Given start column {0}, start row {1}, end column {2} and end row {3} should lead to the cell addresses {4}")
	@CsvSource({
			"0, 0, 0, 0, 'A1'",
			"0, 0, 0, 3, 'A1,A2,A3,A4'",
			"0, 0, 1, 2, 'A1,A2,A3,B1,B2,B3'",
			"1, 2, 0, 1, 'A2,A3,B2,B3'", })
	void getCellRangeTest3(int startColumn, int startRow, int endColumn, int endRow, String expectedAddresses) {
		List<Address> addresses = Cell.getCellRange(startColumn, startRow, endColumn, endRow);
		TestUtils.assertCellRange(expectedAddresses, addresses);
	}

	@DisplayName("Test of the resolveCellAddress method")
	@ParameterizedTest(name = "Given column {0} and row {1} should lead to the address {2}")
	@CsvSource({
			"0, 0, A1",
			"5, 99, F100",
			"16383, 1048575, XFD1048576", })
	void resolveCellAddressTest(int column, int row, String expectedAddress) {
		String address = Cell.resolveCellAddress(column, row);
		assertEquals(expectedAddress, address);
	}

	@DisplayName("Test of the resolveCellAddress method")
	@ParameterizedTest(name = "Given column {0}, row {1} and type {2} should lead to the address {3}")
	@CsvSource({
			"0, 0, Default, A1",
			"0, 0, FixedColumn, $A1",
			"0, 0, FixedRow, A$1",
			"0, 0, FixedRowAndColumn, $A$1",
			"5, 99, Default, F100",
			"5, 99, FixedColumn, $F100",
			"5, 99, FixedRow, F$100",
			"5, 99, FixedRowAndColumn, $F$100",
			"16383, 1048575, Default, XFD1048576",
			"16383, 1048575, FixedColumn, $XFD1048576",
			"16383, 1048575, FixedRow, XFD$1048576",
			"16383, 1048575, FixedRowAndColumn, $XFD$1048576", })
	void resolveCellAddressTest2(int column, int row, Cell.AddressType type, String expectedAddress) {
		String address = Cell.resolveCellAddress(column, row, type);
		assertEquals(expectedAddress, address);
	}

	@DisplayName("Test of the  resolveCellCoordinate method with string as parameter")
	@ParameterizedTest(name = "Given address expression {0} should lead to the column {1}, row {2} and type {3}")
	@CsvSource({
			"A1, 0, 0, Default",
			"$A1, 0, 0, FixedColumn",
			"A$1, 0, 0, FixedRow",
			"$A$1, 0, 0, FixedRowAndColumn",
			"F100, 5, 99, Default",
			"$F100, 5, 99, FixedColumn",
			"F$100, 5, 99, FixedRow",
			"$F$100, 5, 99, FixedRowAndColumn",
			"XFD1048576, 16383, 1048575, Default",
			"$XFD1048576, 16383, 1048575, FixedColumn",
			"XFD$1048576, 16383, 1048575, FixedRow",
			"$XFD$1048576, 16383, 1048575, FixedRowAndColumn", })
	void resolveCellCoordinateTest(String addressString, int expectedColumn, int expectedRow, Cell.AddressType expectedType) {
		Address address = Cell.resolveCellCoordinate(addressString);
		assertEquals(expectedColumn, address.Column);
		assertEquals(expectedRow, address.Row);
		assertEquals(expectedType, address.Type);
	}

	@DisplayName("Test of the resolveCellRange method")
	@ParameterizedTest(name = "Given range expression {0} should lead to the start address {1} and end address {2}")
	@CsvSource({
			"a1:a1, A1, A1",
			"C3:C4, C3, C4",
			"$a1:Z$10, $A1, Z$10",
			"$R$9:a2, A2, $R$9", })
	void resolveCellRangeTest(String rangeString, String expectedStartAddress, String expectedEndAddress) {
		Range range = Cell.resolveCellRange(rangeString);
		Address start = new Address(expectedStartAddress);
		Address end = new Address(expectedEndAddress);
		assertEquals(start, range.StartAddress);
		assertEquals(end, range.EndAddress);
	}

	@DisplayName("Test of the failing resolveCellRange method")
	@Test()
	void resolveCellRangeTest2() {
		assertThrows(FormatException.class, () -> Cell.resolveCellRange(null));
		assertThrows(FormatException.class, () -> Cell.resolveCellRange(""));
		assertThrows(FormatException.class, () -> Cell.resolveCellRange("C3"));
	}

	@DisplayName("Test of the resolveColumn method")
	@ParameterizedTest(name = "Given character {0} should lead to the column number {1}")
	@CsvSource({
			"A, 0",
			"c, 2",
			"XFD, 16383", })
	void resolveColumnTest(String address, int expectedColumn) {
		int column = Cell.resolveColumn(address);
		assertEquals(expectedColumn, column);
	}

	@DisplayName("Test of the failing resolveColumn method")
	@Test()
	void resolveColumnTest2() {
		assertThrows(RangeException.class, () -> Cell.resolveColumn(null));
		assertThrows(RangeException.class, () -> Cell.resolveColumn(""));
		assertThrows(RangeException.class, () -> Cell.resolveColumn("XFE"));
	}

	@DisplayName("Test of the resolveColumnAddress method")
	@ParameterizedTest(name = "Given column number {0} should lead to the letter {1}")
	@CsvSource({
			"0, A",
			"2, C",
			"16383, XFD", })
	void resolveColumnAddressTest(int columnNumber, String expectedAddress) {
		String address = Cell.resolveColumnAddress(columnNumber);
		assertEquals(expectedAddress, address);
	}

	@DisplayName("Test of the failing ResolveColumnAddress method")
	@Test()
	void resolveColumnAddressTest2() {
		assertThrows(RangeException.class, () -> Cell.resolveColumnAddress(-1));
		assertThrows(RangeException.class, () -> Cell.resolveColumnAddress(16384));
	}

	@DisplayName("Test of the address scope check function")
	@ParameterizedTest(name = "Given address or range string: {0} should lead to: {1}")
	@CsvSource({
			"A1,SingleAddress",
			"$A$1, SingleAddress",
			"A1:B2, Range",
			"$A$1:$C5, Range",
			"A0, Invalid",
			"ZZZZZZZZZZZZZZZZZ0, Invalid",
			"A1:C0, Invalid",
			"A1:ZZZZZZZZZZZZ0, Invalid",
			"ZZZZZZZZZZ1:C1, Invalid",
			"ZZZZZZZZZZZZZZ:ZZZZZZZZZZZZZ, Invalid",
			":Z5, Invalid",
			"A2:, Invalid",
			":, Invalid", })
	void addressScopeTest(String addressString, Cell.AddressScope expectedScope) {
		Cell.AddressScope scope = Cell.getAddressScope(addressString);
		assertEquals(expectedScope, scope);
	}

}
