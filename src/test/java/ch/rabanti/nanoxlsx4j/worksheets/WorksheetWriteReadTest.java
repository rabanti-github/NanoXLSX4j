package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Column;
import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class WorksheetWriteReadTest {

    @DisplayName("Test of the 'AutoFilterRange' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given range {0} should lead to the same range in the read worksheet with the index {1}")
    @CsvSource({
            ", 0",
            "A1:A1, 0",
            "A1:C1, 0",
            "B1:D1, 0",
            ", 1",
            "A1:A1, 1",
            "A1:C1, 2",
            "B1:D1, 3",
    })
    void autoFilterRangeWriteReadTest(String autoFilterRange, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        Range range = null;
        if (autoFilterRange != null) {
            range = new Range(autoFilterRange);
            for (int i = 0; i <= sheetIndex; i++) {
                if (sheetIndex == i) {
                    workbook.setCurrentWorksheet(i);
                    workbook.getCurrentWorksheet().setAutoFilter(range.StartAddress.Column, range.EndAddress.Column);
                }
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        if (autoFilterRange == null) {
            assertNull(givenWorksheet.getAutoFilterRange());
        } else {
            assertEquals(range, givenWorksheet.getAutoFilterRange());
        }
    }

    @DisplayName("Test of the 'Columns' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given column indices {0} with width definition: {1} and hidden definition: {2} should lead to the same columns in the read worksheet with the index {1}")
    @CsvSource({
            "'', 0, true, false",
            "'0', 0, true, false",
            "'0,1,2', 0, true, false",
            "'1,3,5', 0, true, false",
            "'', 1, true, false",
            "'0', 1, true, false",
            "'0,1,2', 2, true, false",
            "'1,3,5', 3, true, false",
            "'', 0, false, true",
            "'0', 0, false, true",
            "'0,1,2', 0, false, true",
            "'1,3,5', 0, false, true",
            "'', 1, false, true",
            "'0', 1, false, true",
            "'0,1,2', 2, false, true",
            "'1,3,5', 3, false, true",
    })
    void columnsWriteReadTest(String columnDefinitions, int sheetIndex, boolean setWidth, boolean setHidden) throws Exception {
        String[] tokens = columnDefinitions.split(",");
        List<Integer> columnIndices = new ArrayList<>();
        for (String token : tokens)
        {
            if (!token.equals(""))
            {
                columnIndices.add(Integer.parseInt(token));
            }
        }
        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                for(int index : columnIndices)
                {
                    if (setWidth)
                    {
                        workbook.getCurrentWorksheet().setColumnWidth(index, 99);
                    }
                    if (setHidden)
                    {
                        workbook.getCurrentWorksheet().addHiddenColumn(index);
                    }
                }
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(columnIndices.size(), givenWorksheet.getColumns().size());
        for(Map.Entry<Integer, Column> column : givenWorksheet.getColumns().entrySet())
        {
            assertTrue(columnIndices.stream().anyMatch(x -> x == column.getValue().getNumber()));
            if (setWidth)
            {
                assertTrue(Math.abs(column.getValue().getWidth() - Helper.getInternalColumnWidth(99)) < 0.001);
            }
            if (setHidden)
            {
                assertTrue(column.getValue().isHidden());
            }
        }
    }

    @DisplayName("Test of the 'DefaultColumnWidth' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given default column width {0} should lead to the same width in the read worksheet with the index {1}")
    @CsvSource({
            "1f, 0",
            "11f, 0",
            "55.55f, 0",
            "1f, 1",
            "11f, 2",
            "55.55f, 3",
    })
    void defaultColumnWidthWriteReadTest(float width, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setDefaultColumnWidth(width);
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertTrue(Math.abs(givenWorksheet.getDefaultColumnWidth() - width) < 0.001);
    }

    @DisplayName("Test of the 'DefaultRowHeight' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given default row height {0} should lead to the same height in the read worksheet with the index {1}")
    @CsvSource({
            "1f, 0",
            "11f, 0",
            "55.55f, 0",
            "1f, 1",
            "11f, 2",
            "55.55f, 3",
    })
    void defaultRowHeightWriteReadTest(float height, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setDefaultRowHeight(height);
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertTrue(Math.abs(givenWorksheet.getDefaultRowHeight() - height) < 0.001);
    }

    @DisplayName("Test of the 'HiddenRows' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given hidden row indices {0} with should lead to the same hidden rows in the read worksheet with the index {1}")
    @CsvSource({
            "'', 0",
            "'0', 0",
            "'0,1,2', 0",
            "'1,3,5', 0",
            "'', 1",
            "'0', 1",
            "'0,1,2', 2",
            "'1,3,5', 3",
    })
    void hiddenRowsWriteReadTest(String rowDefinitions, int sheetIndex) throws Exception {
        String[] tokens = rowDefinitions.split(",");
        List<Integer> rowIndices = new ArrayList<>();
        for (String token : tokens)
        {
            if (!token.equals(""))
            {
                rowIndices.add(Integer.parseInt(token));
            }
        }
        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                for(int index : rowIndices)
                {
                        workbook.getCurrentWorksheet().addHiddenRow(index);
                }
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(rowIndices.size(), givenWorksheet.getHiddenRows().size());
        for(Map.Entry<Integer, Boolean> hiddenRow : givenWorksheet.getHiddenRows().entrySet())
        {
            assertTrue(rowIndices.stream().anyMatch(x -> x == hiddenRow.getKey()));
            assertTrue(hiddenRow.getValue());
        }
    }

    @DisplayName("Test of the 'RowHeight' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given hidden row indices {0} with should lead to the same hidden rows in the read worksheet with the index {1}")
    @CsvSource({
            "'','',0",
            "'0','17',0",
            "'0,1,2','11,12,13.5',0",
            "'','',1",
            "'0','17.2',1",
            "'0','17',0,1,2','11.05,12.1,13.55',2",
            "'1,3,5','55.5,1.111,5.587',3",
    })
    void rowHeightsWriteReadTest(String rowDefinitions, String heightDefinitions, int sheetIndex) throws Exception {
        String[] tokens = rowDefinitions.split(",");
        String[] heightTokens = heightDefinitions.split(",");
        Map<Integer,Float> rows = new HashMap<>();
        for (int i = 0; i < tokens.length; i++)
        {
            if (!tokens[i].equals(""))
            {
                rows.put(Integer.parseInt(tokens[i]), Float.parseFloat(heightTokens[i]));
            }
        }
        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                for(Map.Entry<Integer,Float> row : rows.entrySet())
                {
                    workbook.getCurrentWorksheet().setRowHeight(row.getKey(), row.getValue());
                }
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(rows.size(), givenWorksheet.getRowHeights().size());
        for(Map.Entry<Integer, Float> rowHeight : givenWorksheet.getRowHeights().entrySet())
        {
            assertTrue(rows.keySet().stream().anyMatch(x -> x == rowHeight.getKey()));
            float expectedHeight = Helper.getInternalRowHeight(rows.get(rowHeight.getKey()));
            assertEquals(expectedHeight, rowHeight.getValue());
        }
    }

    @Test
    @DisplayName("Test of the 'RowHeight' property when writing and reading a worksheet, if a row already exists")
    void rowHeightsWriteReadTest2() throws Exception {
        Workbook workbook = new Workbook("worksheet1");
        workbook.getCurrentWorksheet().addCell(42, "C2");
        workbook.getCurrentWorksheet().setRowHeight(2, 22.55f);
        workbook.getCurrentWorksheet().addHiddenRow(2);
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, 0);
        assertEquals(Helper.getInternalRowHeight(22.55f), givenWorksheet.getRowHeights().get(2));
        assertTrue(givenWorksheet.getHiddenRows().get(2));
    }

    @DisplayName("Test of the 'MergedCells' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given range {0} should lead to the same range in the read worksheet with the index {1}")
    @CsvSource({
            ", 0",
            "'A1:A1', 0",
            "'A1:C1', 0",
            "'B1:D1', 0",
            "'B1:D1,E5:E7', 0",
            ", 1",
            "'A1:A1', 1",
            "'A1:C1', 2",
            "'B1:D1', 3",
            "'B1:D1,E5:E7', 3",
    })
    void mergedCellsWriteReadTest(String mergedCellsRanges, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        List<Range>ranges = new ArrayList<>();
        if (mergedCellsRanges != null) {
            String[] split = mergedCellsRanges.split(",");
            for(String range : split){
              ranges.add(new Range(range));
            }
            for (int i = 0; i <= sheetIndex; i++) {
                if (sheetIndex == i) {
                    workbook.setCurrentWorksheet(i);
                    for(Range range: ranges){
                        workbook.getCurrentWorksheet().mergeCells(range);
                    }
                }
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        if (mergedCellsRanges == null) {
            assertEquals(0, givenWorksheet.getMergedCells().size());
        } else {
            for(Range range : ranges){
                assertEquals(range, givenWorksheet.getMergedCells().get(range.toString()));
            }
        }
    }

    @DisplayName("Test of the 'SelectedCells' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given range {0} should lead to the same range in the read worksheet with the index {1}")
    @CsvSource({
            ", 0",
            "A1:A1, 0",
            "A1:C1, 0",
            "B1:D1, 0",
            ", 1",
            "A1:A1, 1",
            "A1:C1, 2",
            "B1:D1, 3",
    })
    void selectedCellsWriteReadTest(String selectedCellsRange, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        Range range = null;
        if (selectedCellsRange != null) {
            range = new Range(selectedCellsRange);
            for (int i = 0; i <= sheetIndex; i++) {
                if (sheetIndex == i) {
                    workbook.setCurrentWorksheet(i);
                    workbook.getCurrentWorksheet().setSelectedCells(range);
                }
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        if (selectedCellsRange == null) {
            assertNull(givenWorksheet.getSelectedCells());
        } else {
            assertEquals(range, givenWorksheet.getSelectedCells());
        }
    }

    @DisplayName("Test of the 'SheetID'  property when writing and reading a worksheet")
    @Test()
    void sheetIDWriteReadTest() throws Exception {
        Workbook workbook = new Workbook();
        String sheetName1 = "sheet_a";
        String sheetName2 = "sheet_b";
        String sheetName3 = "sheet_c";
        String sheetName4 = "sheet_d";
        int id1, id2, id3, id4;
        workbook.addWorksheet(sheetName1);
        id1 = workbook.getCurrentWorksheet().getSheetID();
        workbook.addWorksheet(sheetName2);
        id2 = workbook.getCurrentWorksheet().getSheetID();
        workbook.addWorksheet(sheetName3);
        id3 = workbook.getCurrentWorksheet().getSheetID();
        workbook.addWorksheet(sheetName4);
        id4 = workbook.getCurrentWorksheet().getSheetID();
        Workbook givenWorkbook = TestUtils.saveAndLoadWorkbook(workbook, null);
        assertEquals(id1, givenWorkbook.getWorksheets().stream().filter(w -> w.getSheetName().equals(sheetName1)).findFirst().map(Worksheet::getSheetID).get());
        assertEquals(id2, givenWorkbook.getWorksheets().stream().filter(w -> w.getSheetName().equals(sheetName2)).findFirst().map(Worksheet::getSheetID).get());
        assertEquals(id3, givenWorkbook.getWorksheets().stream().filter(w -> w.getSheetName().equals(sheetName3)).findFirst().map(Worksheet::getSheetID).get());
        assertEquals(id4, givenWorkbook.getWorksheets().stream().filter(w -> w.getSheetName().equals(sheetName4)).findFirst().map(Worksheet::getSheetID).get());
    }

    @DisplayName("Test of the 'SheetName'  property when writing and reading a worksheet")
    @Test()
    void sheetNameWriteReadTest() throws Exception {
        Workbook workbook = new Workbook();
        String sheetName1 = "sheet_a";
        String sheetName2 = "sheet_b";
        String sheetName3 = "sheet_c";
        String sheetName4 = "sheet_d";
        int id1, id2, id3, id4;
        workbook.addWorksheet(sheetName1);
        id1 = workbook.getCurrentWorksheet().getSheetID();
        workbook.addWorksheet(sheetName2);
        id2 = workbook.getCurrentWorksheet().getSheetID();
        workbook.addWorksheet(sheetName3);
        id3 = workbook.getCurrentWorksheet().getSheetID();
        workbook.addWorksheet(sheetName4);
        id4 = workbook.getCurrentWorksheet().getSheetID();
        Workbook givenWorkbook = TestUtils.saveAndLoadWorkbook(workbook, null);
        assertEquals(sheetName1, givenWorkbook.getWorksheets().stream().filter(w -> w.getSheetID() == id1).findFirst().map(Worksheet::getSheetName).get());
        assertEquals(sheetName2, givenWorkbook.getWorksheets().stream().filter(w -> w.getSheetID() == id2).findFirst().map(Worksheet::getSheetName).get());
        assertEquals(sheetName3, givenWorkbook.getWorksheets().stream().filter(w -> w.getSheetID() == id3).findFirst().map(Worksheet::getSheetName).get());
        assertEquals(sheetName4, givenWorkbook.getWorksheets().stream().filter(w -> w.getSheetID() == id4).findFirst().map(Worksheet::getSheetName).get());
    }

    private static Workbook prepareWorkbook(int numberOfWorksheets, Object a1Data){
        Workbook workbook = new Workbook();
        for(int i = 0; i < numberOfWorksheets; i++){
            workbook.addWorksheet("worksheet" + Integer.toString(i+1));
            workbook.getCurrentWorksheet().addCell(a1Data, "A1");
        }
        return workbook;
    }

    private static Worksheet writeAndReadWorksheet(Workbook workbook, int worksheetIndex) throws Exception {
        Workbook readWorkbook = TestUtils.saveAndLoadWorkbook(workbook, null);
        return readWorkbook.getWorksheets().get(worksheetIndex);
    }

}
