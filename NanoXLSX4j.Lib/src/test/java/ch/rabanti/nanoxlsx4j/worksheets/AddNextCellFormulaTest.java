package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;

class AddNextCellFormulaTest {
	private Worksheet worksheet;

	@DisplayName("Test of the addNextCellFormula function with only the value")
	@Test()
	void addNextCellFormulaTest1() {
		worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
		assertEquals(0, worksheet.getCells().size());
		worksheet.addNextCellFormula("=B2");
		WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, null, "=B2", 3, 2);
		worksheet = WorksheetTest.initWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
		worksheet.addNextCellFormula("=B2");
		WorksheetTest.assertAddedCell(worksheet, 2, "E3", Cell.CellType.FORMULA, null, "=B2", 5, 2);
	}

	@DisplayName("Test of the addNextCellFormula function with value and Style")
	@Test()
	void addNextCellFormulaTest2() {
		worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
		assertEquals(0, worksheet.getCells().size());
		worksheet.addNextCellFormula("=B2", BasicStyles.BoldItalic());
		WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.BoldItalic(), "=B2", 3, 2);
		worksheet = WorksheetTest.initWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
		worksheet.addNextCellFormula("=B2", BasicStyles.Bold());
		WorksheetTest.assertAddedCell(worksheet, 2, "E3", Cell.CellType.FORMULA, BasicStyles.Bold(), "=B2", 5, 2);
	}

	@DisplayName("Test of the addNextCellFormula function with value and active worksheet style")
	@Test()
	void addNextCellFormulaTest3() {
		worksheet = WorksheetTest.initWorksheet(worksheet,
				"D2",
				Worksheet.CellDirection.RowToRow,
				BasicStyles.BorderFrameHeader());
		assertEquals(0, worksheet.getCells().size());
		worksheet.addNextCellFormula("=B2");
		WorksheetTest.assertAddedCell(worksheet,
				1,
				"D2",
				Cell.CellType.FORMULA,
				BasicStyles.BorderFrameHeader(),
				"=B2",
				3,
				2);
	}

	@DisplayName("Test of the addNextCell function for a nested cell object, if the cell is a formula")
	@Test()
	void addNextCellFormulaTest5() {
		Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
		worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
		worksheet.addNextCell(cell);
		WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, null, "=B2", 3, 2);
	}

	@DisplayName("Test of the addNextCell function for a nested cell object and style, if the cell is a formula")
	@Test()
	void addNextCellFormulaTest6() {
		Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
		worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
		worksheet.addNextCell(cell, BasicStyles.Bold());
		WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.Bold(), "=B2", 3, 2);
		cell = new Cell("=B2", Cell.CellType.FORMULA, "R2");
		cell.setStyle(BasicStyles.BorderFrame());
		Style mixedStyle = BasicStyles.BorderFrame();
		mixedStyle.append(BasicStyles.Bold());
		worksheet.addNextCell(cell, BasicStyles.Bold());
		WorksheetTest.assertAddedCell(worksheet, 2, "D3", Cell.CellType.FORMULA, mixedStyle, "=B2", 3, 3);
	}

	@DisplayName("Test of the addNextCell function for a nested cell object and active worksheet style, if the cell is a formula")
	@Test()
	void addNextCellFormulaTest7() {
		worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow, BasicStyles.Bold());
		Cell cell = new Cell("=B2", Cell.CellType.FORMULA, "R1"); // Address should be replaced
		worksheet.addNextCell(cell);
		WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, BasicStyles.Bold(), "=B2", 3, 2);
		cell = new Cell("=B2", Cell.CellType.FORMULA, "R2");
		cell.setStyle(BasicStyles.BorderFrame());
		Style mixedStyle = BasicStyles.BorderFrame();
		mixedStyle.append(BasicStyles.Bold());
		worksheet.addNextCell(cell);
		WorksheetTest.assertAddedCell(worksheet, 2, "D3", Cell.CellType.FORMULA, mixedStyle, "=B2", 3, 3);
	}

	@DisplayName("Test of the addNextCellFormula function with when changing the current cell direction")
	@Test()
	void addNextCellFormulaTest8() {
		Worksheet worksheet = new Worksheet();
		worksheet = WorksheetTest.initWorksheet(worksheet, "D2", Worksheet.CellDirection.RowToRow);
		worksheet.addNextCellFormula("=B2");
		WorksheetTest.assertAddedCell(worksheet, 1, "D2", Cell.CellType.FORMULA, null, "=B2", 3, 2);
		worksheet = WorksheetTest.initWorksheet(worksheet, "E3", Worksheet.CellDirection.ColumnToColumn);
		worksheet.addNextCellFormula("=B2");
		WorksheetTest.assertAddedCell(worksheet, 2, "E3", Cell.CellType.FORMULA, null, "=B2", 5, 2);
		worksheet = WorksheetTest.initWorksheet(worksheet, "F5", Worksheet.CellDirection.Disabled);
		worksheet.addNextCellFormula("=B2");
		WorksheetTest.assertAddedCell(worksheet, 3, "F5", Cell.CellType.FORMULA, null, "=B2", 5, 4);
	}
}
