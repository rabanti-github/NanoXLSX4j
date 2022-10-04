package ch.rabanti.nanoxlsx4j.cells.types;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;

public class BooleanCellTest {

	CellTypeUtils utils;

	public BooleanCellTest() {
		utils = new CellTypeUtils(Boolean.class);
	}

	@DisplayName("Bool value cell test: Test of the cell values, as well as proper modification")
	@Test()
	void booleanCellTest() {
		utils.assertCellCreation(true, true, Cell.CellType.BOOL, (current, other) -> {
			return current.equals(other);
		});
		utils.assertCellCreation(true, false, Cell.CellType.BOOL, (current, other) -> {
			return current.equals(other);
		});
	}

	@DisplayName("Bool value cell test with style")
	@Test()
	void booleanCellTest2() {
		Style style = BasicStyles.Bold();
		utils.assertStyledCellCreation(true, true, Cell.CellType.BOOL, (curent, other) -> {
			return curent.equals(other);
		}, style);
		utils.assertStyledCellCreation(true, false, Cell.CellType.BOOL, (curent, other) -> {
			return curent.equals(other);
		}, style);
	}

	@DisplayName("Test of the bool comparison method on cells")
	@ParameterizedTest(name = "Given values {0} and {1} should lead to a comparison result of {2}")
	@CsvSource({
			"true, true, 0",
			"false, false, 0",
			"true, false, 1",
			"false, true, -1", })
	void booleanCellComparisonTest(boolean value1, boolean value2, int expectedResult) {
		Cell cell1 = utils.createVariantCell(value1, utils.getCellAddress());
		Cell cell2 = utils.createVariantCell(value2, utils.getCellAddress());
		int comparison = ((Boolean) cell1.getValue()).compareTo((Boolean) cell2.getValue());
		assertEquals(comparison, expectedResult);
	}
}
