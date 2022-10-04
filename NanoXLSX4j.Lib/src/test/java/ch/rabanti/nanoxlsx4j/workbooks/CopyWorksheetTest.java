package ch.rabanti.nanoxlsx4j.workbooks;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.Set;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Column;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;

public class CopyWorksheetTest {

	@DisplayName("Test of the 'copyWorksheetIntoThis' function by name")
	@ParameterizedTest(name = "Given source worksheet {2}, copied as {3}, should lead to a copy, named {4}, with sanitation: {0}")
	@CsvSource({
			"false, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, worksheet1, worksheet2", })
	void copyWorksheetIntoThisTest(boolean sanitize, String givenWorksheetName1, String givenSourceWsName, String copyName, String expectedTargetWsName) {
		Workbook workbook1 = new Workbook(givenWorksheetName1);
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName(givenSourceWsName);
		workbook1.addWorksheet(worksheet2);
		workbook1.copyWorksheetIntoThis(givenSourceWsName, copyName, sanitize);
		assertCopy(workbook1, givenSourceWsName, workbook1, expectedTargetWsName);
	}

	@DisplayName("Test of the failing 'copyWorksheetIntoThis' (name) function when with disabled sanitation on duplicate worksheets")
	@Test()
	void copyWorksheetIntoThisFailTest() {
		Workbook workbook1 = new Workbook("worksheet1");
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName("worksheet2");
		workbook1.addWorksheet(worksheet2);
		assertThrows(WorksheetException.class,
				() -> workbook1.copyWorksheetIntoThis("worksheet2",
						"worksheet1",
						false));
	}

	@DisplayName("Test of the 'copyWorksheetIntoThis' function by index")
	@ParameterizedTest(name = "Given source worksheet {2}, copied as {3}, should lead to a copy, named {4}, with sanitation: {0}")
	@CsvSource({
			"false, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, worksheet1, worksheet2", })
	void copyWorksheetIntoThisTest2(boolean sanitize, String givenWorksheetName1, String givenSourceWsName, String copyName, String expectedTargetWsName) {
		Workbook workbook1 = new Workbook(givenWorksheetName1);
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName(givenSourceWsName);
		workbook1.addWorksheet(worksheet2);
		workbook1.copyWorksheetIntoThis(1, copyName, sanitize);
		assertCopy(workbook1, givenSourceWsName, workbook1, expectedTargetWsName);
	}

	@DisplayName("Test of the failing 'copyWorksheetIntoThis' (index) function when with disabled sanitation on duplicate worksheets")
	@Test()
	void copyWorksheetIntoThisFailTest2() {
		Workbook workbook1 = new Workbook("worksheet1");
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName("worksheet2");
		workbook1.addWorksheet(worksheet2);
		assertThrows(WorksheetException.class, () -> workbook1.copyWorksheetIntoThis(1, "worksheet1", false));
	}

	@DisplayName("Test of the 'copyWorksheetIntoThis' function by reference")
	@ParameterizedTest(name = "Given source worksheet {2}, copied as {3}, should lead to a copy, named {4}, with sanitation: {0}")
	@CsvSource({
			"false, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, worksheet1, worksheet2", })
	void copyWorksheetIntoThisTest3(boolean sanitize, String givenWorksheetName1, String givenSourceWsName, String copyName, String expectedTargetWsName) {
		Workbook workbook1 = new Workbook(givenWorksheetName1);
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName(givenSourceWsName);
		workbook1.addWorksheet(worksheet2);
		workbook1.copyWorksheetIntoThis(worksheet2, copyName, sanitize);
		assertCopy(workbook1, givenSourceWsName, workbook1, expectedTargetWsName);
	}

	@DisplayName("Test of the failing 'copyWorksheetIntoThis' (reference) function when with disabled sanitation on duplicate worksheets")
	@Test()
	void copyWorksheetIntoThisFailTest3() {
		Workbook workbook1 = new Workbook("worksheet1");
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName("worksheet2");
		workbook1.addWorksheet(worksheet2);
		assertThrows(WorksheetException.class, () -> workbook1.copyWorksheetIntoThis(worksheet2, "worksheet1", false));
	}

	@DisplayName("Test of the 'copyWorksheetTo' function by name")
	@ParameterizedTest(name = "Given source worksheet {2}, copied as {3}, should lead to a copy, named {4}, with sanitation: {0}")
	@CsvSource({
			"false, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, worksheet1, worksheet2", })

	void copyWorksheetToTest(boolean sanitize, String givenWorksheetName1, String givenSourceWsName, String copyName, String expectedTargetWsName) {
		Workbook workbook1 = new Workbook(givenWorksheetName1);
		Workbook workbook2 = new Workbook(givenWorksheetName1);
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName(givenSourceWsName);
		workbook1.addWorksheet(worksheet2);
		workbook1.copyWorksheetTo(givenSourceWsName, copyName, workbook2, sanitize);
		assertCopy(workbook1, givenSourceWsName, workbook2, expectedTargetWsName);
	}

	@DisplayName("Test of the failing 'copyWorksheetTo' (name) function when with disabled sanitation on duplicate worksheets")
	@Test()
	void copyWorksheetToFailTest() {
		Workbook workbook1 = new Workbook("worksheet1");
		Workbook workbook2 = new Workbook("worksheet1");
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName("worksheet2");
		workbook1.addWorksheet(worksheet2);
		assertThrows(WorksheetException.class,
				() -> workbook1.copyWorksheetTo("worksheet2",
						"worksheet1",
						workbook2,
						false));
	}

	@DisplayName("Test of the 'copyWorksheetTo' function by index")
	@ParameterizedTest(name = "Given source worksheet {2}, copied as {3}, should lead to a copy, named {4}, with sanitation: {0}")
	@CsvSource({
			"false, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, worksheet1, worksheet2", })
	void copyWorksheetToTest2(boolean sanitize, String givenWorksheetName1, String givenSourceWsName, String copyName, String expectedTargetWsName) {
		Workbook workbook1 = new Workbook(givenWorksheetName1);
		Workbook workbook2 = new Workbook(givenWorksheetName1);
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName(givenSourceWsName);
		workbook1.addWorksheet(worksheet2);
		workbook1.copyWorksheetTo(1, copyName, workbook2, sanitize);
		assertCopy(workbook1, givenSourceWsName, workbook2, expectedTargetWsName);
	}

	@DisplayName("Test of the failing 'copyWorksheetTo' (index) function when with disabled sanitation on duplicate worksheets")
	@Test()
	void copyWorksheetToFailTest2() {
		Workbook workbook1 = new Workbook("worksheet1");
		Workbook workbook2 = new Workbook("worksheet1");
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName("worksheet2");
		workbook1.addWorksheet(worksheet2);
		assertThrows(WorksheetException.class, () -> workbook1.copyWorksheetTo(1, "worksheet1", workbook2, false));
	}

	@DisplayName("Test of the 'copyWorksheetTo' function by reference")
	@ParameterizedTest(name = "Given source worksheet {2}, copied as {3}, should lead to a copy, named {4}, with sanitation: {0}")
	@CsvSource({
			"false, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, copy, copy",
			"true, worksheet1, worksheet2, worksheet1, worksheet2", })
	void copyWorksheetToTest3(boolean sanitize, String givenWorksheetName1, String givenSourceWsName, String copyName, String expectedTargetWsName) {
		Workbook workbook1 = new Workbook(givenWorksheetName1);
		Workbook workbook2 = new Workbook(givenWorksheetName1);
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName(givenSourceWsName);
		workbook1.addWorksheet(worksheet2);
		Workbook.copyWorksheetTo(worksheet2, copyName, workbook2, sanitize);
		assertCopy(workbook1, givenSourceWsName, workbook2, expectedTargetWsName);
	}

	@DisplayName("Test of the failing 'copyWorksheetTo' (reference) function when with disabled sanitation on duplicate worksheets")
	@Test()
	void copyWorksheetToFailTest4() {
		Workbook workbook1 = new Workbook("worksheet1");
		Workbook workbook2 = new Workbook("worksheet1");
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName("worksheet2");
		workbook1.addWorksheet(worksheet2);
		assertThrows(WorksheetException.class,
				() -> Workbook.copyWorksheetTo(worksheet2, "worksheet1", workbook2, false));
	}

	@DisplayName("Test of the failing 'copyWorksheetTo' function when no Workbook was defined")
	@Test()
	void copyWorksheetToFailTest5() {
		Workbook workbook1 = null;
		Worksheet worksheet2 = createWorksheet();
		assertThrows(WorksheetException.class, () -> Workbook.copyWorksheetTo(worksheet2, "copy", workbook1));
	}

	@DisplayName("Test of the failing 'copyWorksheetTo' function when no worksheet was defined")
	@Test()
	void copyWorksheetToFailTest6() {
		Workbook workbook1 = new Workbook("worksheet1");
		Worksheet worksheet2 = null;
		assertThrows(WorksheetException.class, () -> Workbook.copyWorksheetTo(worksheet2, "copy", workbook1));
	}

	@DisplayName("Test of the 'copy' function within the Worksheet class")
	@Test()
	void copyTest() {
		Worksheet worksheet = createWorksheet();
		Worksheet worksheet2 = worksheet.copy();
		assertWorksheetCopy(worksheet, worksheet2);
	}

	private void assertCopy(Workbook sourceWorkbook, String sourceName, Workbook targetWorkbook, String targetName) {
		Worksheet w1 = sourceWorkbook.getWorksheet(sourceName);
		Worksheet w2 = targetWorkbook.getWorksheet(targetName);
		assertWorksheetCopy(w1, w2);
	}

	@DisplayName("Test of the 'copyWorksheetTo' function for proper saving")
	@Test()
	void copyWorksheetSaveTest() throws Exception {
		Workbook workbook1 = new Workbook("worksheet1");
		Workbook workbook2 = new Workbook("worksheet1b");
		Worksheet worksheet2 = createWorksheet();
		worksheet2.setSheetName("worksheet2");
		workbook1.addWorksheet(worksheet2);
		Workbook.copyWorksheetTo(worksheet2, "copy", workbook2);

		Workbook newWorkbook = TestUtils.writeAndReadWorkbook(workbook2);
		assertEquals(workbook2.getWorksheets().size(), newWorkbook.getWorksheets().size());
	}

	private void assertWorksheetCopy(Worksheet w1, Worksheet w2) {
		Set<String> keys = w1.getCells().keySet();
		assertEquals(w2.getCells().size(), keys.size());
		for (String address : keys) {
			Cell c1 = w1.getCell(address);
			Cell c2 = w2.getCell(address);
			assertEquals(c2.getCellAddress(), c1.getCellAddress());
			assertEquals(c2.getValue(), c1.getValue());
			assertEquals(c2.getCellAddressType(), c1.getCellAddressType());
			assertStyle(c1.getCellStyle(), c2.getCellStyle());
			assertEquals(c2.getDataType(), c1.getDataType());
		}
		assertEquals(w2.getActivePane(), w1.getActivePane());
		assertStyle(w1.getActiveStyle(), w2.getActiveStyle());
		assertEquals(w2.getAutoFilterRange(), w1.getAutoFilterRange());
		// columns
		Set<Integer> keys2 = w1.getColumns().keySet();
		assertEquals(w2.getColumns().size(), keys2.size());
		for (int col : keys2) {
			Column c1 = w1.getColumns().get(col);
			Column c2 = w2.getColumns().get(col);
			assertEquals(c2.getColumnAddress(), c1.getColumnAddress());
			assertEquals(c2.hasAutoFilter(), c1.hasAutoFilter());
			assertEquals(c2.isHidden(), c1.isHidden());
			assertEquals(c2.getNumber(), c1.getNumber());
			assertEquals(c2.getWidth(), c1.getWidth());
		}
		assertEquals(w2.getCurrentCellDirection(), w1.getCurrentCellDirection());
		assertEquals(w2.getDefaultColumnWidth(), w1.getDefaultColumnWidth());
		assertEquals(w2.getDefaultRowHeight(), w1.getDefaultRowHeight());
		assertEquals(w2.getFreezeSplitPanes(), w1.getFreezeSplitPanes());
		assertEquals(w2.isHidden(), w1.isHidden());
		keys2 = w1.getHiddenRows().keySet();
		assertEquals(w2.getHiddenRows().size(), keys2.size());
		for (int row : keys2) {
			assertEquals(w2.getHiddenRows().get(row), w1.getHiddenRows().get(row));
		}
		keys = w1.getMergedCells().keySet();
		assertEquals(w2.getMergedCells().size(), keys.size());
		for (String address : keys) {
			Range r1 = w1.getMergedCells().get(address);
			Range r2 = w2.getMergedCells().get(address);
			assertEquals(r2.StartAddress, r1.StartAddress);
			assertEquals(r2.EndAddress, r1.EndAddress);
		}
		assertEquals(w2.getPaneSplitAddress(), w2.getPaneSplitAddress());
		assertEquals(w2.getPaneSplitLeftWidth(), w2.getPaneSplitLeftWidth());
		assertEquals(w2.getPaneSplitTopHeight(), w2.getPaneSplitTopHeight());
		assertEquals(w2.getPaneSplitTopLeftCell(), w2.getPaneSplitTopLeftCell());
		keys2 = w1.getRowHeights().keySet();
		assertEquals(w2.getRowHeights().size(), keys2.size());
		for (int row : keys2) {
			assertEquals(w2.getRowHeights().get(row), w1.getRowHeights().get(row));
		}
		assertEquals(w2.getSelectedCells(), w1.getSelectedCells());
		assertEquals(w2.getSheetProtectionPassword(), w1.getSheetProtectionPassword());
		assertEquals(w2.getSheetProtectionPasswordHash(), w1.getSheetProtectionPasswordHash());
		assertEquals(w2.getSheetProtectionValues().size(), w1.getSheetProtectionValues().size());
		for (int i = 0; i < w1.getSheetProtectionValues().size(); i++) {
			assertEquals(w2.getSheetProtectionValues().get(i), w1.getSheetProtectionValues().get(i));
		}
		assertEquals(w2.isUseSheetProtection(), w1.isUseSheetProtection());
	}

	private void assertStyle(Style style1, Style style2) {
		if (style1 == null) {
			assertNull(style2);
		}
		else {
			assertEquals(style2.hashCode(), style1.hashCode());
		}
	}

	private Worksheet createWorksheet() {
		Worksheet w = new Worksheet();
		Style s1 = BasicStyles.BoldItalic();
		Style s2 = BasicStyles.Bold().append(BasicStyles.DateFormat());
		w.addCell("A1",
				"A1",
				s1);
		w.addCell(true, "B2");
		w.addCell(100, "C3", s2);
		w.addCell(2.23f, "D4");
		w.addCell(false, "D5");
		w.addCellFormula("=A2",
				"E5");
		w.setColumnWidth(2, 31.2f);
		w.setRowHeight(2, 50.6f);
		w.addHiddenColumn(1);
		w.addHiddenColumn(3);
		w.addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.sort);
		w.addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.autoFilter);
		w.setSheetProtectionPassword("pwd");
		w.addHiddenRow(1);
		w.addHiddenRow(3);
		w.setCurrentCellDirection(Worksheet.CellDirection.Disabled);
		w.setDefaultColumnWidth(55.5f);
		w.setDefaultRowHeight(45.3f);
		w.setHidden(true);
		w.mergeCells(new Range("D4:D5"));
		w.setActiveStyle(s2);
		w.setAutoFilter("B1:C2");
		w.setCurrentCellAddress("D5");
		w.setSelectedCells(new Range("C3:C3"));
		w.setUseSheetProtection(true);
		w.setSplit(3, 2, true, new Address("F4"), Worksheet.WorksheetPane.bottomRight);
		return w;
	}
}
