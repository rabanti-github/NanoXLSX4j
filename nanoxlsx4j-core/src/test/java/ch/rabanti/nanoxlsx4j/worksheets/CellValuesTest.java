package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.HashSet;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.CorePackageTestAccess;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;

class CellValuesTest {

    @Test
    @DisplayName("getCellValues contains the cell added via addCell")
    void cellValuesContainsAddedCell() {
        Worksheet ws = new Worksheet();
        ws.addCell("hello", 2, 3);
        List<Cell> values = List.copyOf(ws.getCellValues());
        assertEquals(1, values.size());
        assertEquals("hello", values.getFirst().getValue());
    }

    @Test
    @DisplayName("getCellValues size matches getCells().size()")
    void cellValuesCountMatchesCells() {
        Worksheet ws = new Worksheet();
        ws.addCell(1, 0, 0);
        ws.addCell(2, 1, 0);
        ws.addCell(3, 2, 0);
        assertEquals(ws.getCells().size(), ws.getCellValues().size());
    }

    @Test
    @DisplayName("getCellValues and getCells values yield the same Cell instances")
    void cellValuesSameInstancesAsCells() {
        Worksheet ws = new Worksheet();
        ws.addCell("a", 0, 0);
        ws.addCell("b", 1, 0);
        ws.addCell("c", 0, 1);
        HashSet<Cell> fromCellValues = new HashSet<>(ws.getCellValues());
        HashSet<Cell> fromCells = new HashSet<>(ws.getCells().values());
        assertEquals(fromCells, fromCellValues);
    }

    @Test
    @DisplayName("getCells resolves the correct cell after addCell")
    void cellsStringIndexerResolvesCorrectCell() {
        Worksheet ws = new Worksheet();
        ws.addCell(42, 0, 0);
        ws.addCell("world", 3, 5);
        assertEquals(42, ws.getCells().get("A1").getValue());
        assertEquals("world", ws.getCells().get("D6").getValue());
    }

    @Test
    @DisplayName("getCellValues is empty after removeCell")
    void cellValuesEmptyAfterRemove() {
        Worksheet ws = new Worksheet();
        ws.addCell(99, 0, 0);
        ws.removeCell(0, 0);
        assertTrue(ws.getCellValues().isEmpty());
    }

    @Test
    @DisplayName("getCells().containsKey returns true for added cell, false after removal")
    void cellsContainsKeyTracksMutations() {
        Worksheet ws = new Worksheet();
        ws.addCell(7, 1, 2);
        assertTrue(ws.getCells().containsKey("B3"));
        ws.removeCell(1, 2);
        assertFalse(ws.getCells().containsKey("B3"));
    }

    @Test
    @DisplayName("Cell formula type and value changes update worksheet and workbook features")
    void cellFormulaChangesUpdateFeatures() {
        Workbook workbook = new Workbook("Sheet1");
        Worksheet worksheet = workbook.getCurrentWorksheet();
        worksheet.addCell("Initial Value", "A1");
        Cell cell = worksheet.getCells().get("A1");

        assertEquals(0, CorePackageTestAccess.worksheetFormulaCount(worksheet));
        assertEquals(0, workbook.getFeatures().getFormulaCount());
        assertEquals(0, workbook.getFeatures().getExternalLinkCount());

        cell.setDataType(Cell.CellType.FORMULA);

        assertEquals(1, CorePackageTestAccess.worksheetFormulaCount(worksheet));
        assertEquals(1, workbook.getFeatures().getFormulaCount());
        assertEquals(0, workbook.getFeatures().getExternalLinkCount());

        cell.setValue("[1]ExternalSheet!A1");

        assertEquals(1, CorePackageTestAccess.worksheetExternalLinkCount(worksheet));
        assertEquals(1, workbook.getFeatures().getExternalLinkCount());

        cell.setValue("A1");

        assertEquals(0, CorePackageTestAccess.worksheetExternalLinkCount(worksheet));
        assertEquals(0, workbook.getFeatures().getExternalLinkCount());

        cell.setDataType(Cell.CellType.STRING);

        assertNull(cell.getFormula());
        assertEquals(0, CorePackageTestAccess.worksheetFormulaCount(worksheet));
        assertEquals(0, workbook.getFeatures().getFormulaCount());
    }
}
