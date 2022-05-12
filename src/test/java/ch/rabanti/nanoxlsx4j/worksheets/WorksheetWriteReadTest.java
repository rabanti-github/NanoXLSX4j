package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Address;
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

    @DisplayName("Test of the 'SheetProtectionValues'  and 'UseSheetProtection' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given values {1} and enabled protection: {0} should lead to the values {2} in worksheet {3}")
    @CsvSource({
            "false, '', '', 0",
            "false, 'autoFilter:0,sort:0', '', 0",
            "true, '', 'objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'autoFilter:0', 'autoFilter:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'pivotTables:0', 'pivotTables:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'sort:0', 'sort:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'deleteRows:0', 'deleteRows:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'deleteColumns:0', 'deleteColumns:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'insertHyperlinks:0', 'insertHyperlinks:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'insertRows:0', 'insertRows:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'insertColumns:0', 'insertColumns:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'formatRows:0', 'formatRows:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'formatColumns:0', 'formatColumns:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'formatCells:0', 'formatCells:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'objects:0', 'scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'scenarios:0', 'objects:1,selectLockedCells:1,selectUnlockedCells:1', 0",
            "true, 'selectLockedCells:0', 'objects:1,scenarios:1', 0",
            "true, 'selectUnlockedCells:0', 'objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:0', 0",
            "false, '', '', 1",
            "false, 'autoFilter:0', '', 2",
            "true, '', 'objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 3",
            "true, 'autoFilter:0', 'autoFilter:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 1",
            "true, 'pivotTables:0,sort:0', 'pivotTables:0,sort:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 2",
            "true, 'sort:0,deleteColumns:0,formatCells:0', 'sort:0,deleteColumns:0,formatCells:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 3",
            "true, 'deleteRows:0,formatCells:0', 'deleteRows:0,formatCells:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 1",
            "true, 'deleteColumns:0,formatColumns:0,formatRows:0', 'deleteColumns:0,formatColumns:0,formatRows:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 2",
            "true, 'insertHyperlinks:0,formatCells:0', 'insertHyperlinks:0,formatCells:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 3",
            "true, 'insertRows:0,formatRows:0', 'insertRows:0,formatRows:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 1",
            "true, 'insertColumns:0,formatColumns:0', 'insertColumns:0,formatColumns:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 2",
            "true, 'formatRows:0,formatColumns:0', 'formatRows:0,formatColumns:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 3",
            "true, 'formatColumns:0,formatCells:0', 'formatColumns:0,formatCells:0,objects:1,scenarios:1,selectLockedCells:1,selectUnlockedCells:1', 1",

    })
    void sheetProtectionWriteReadTest(boolean useSheetProtection, String givenProtectionValues, String expectedProtectionValues, int sheetIndex) throws Exception {
        Map<Worksheet.SheetProtectionValue, Boolean> expectedProtection = prepareSheetProtectionValues(expectedProtectionValues);
        Map<Worksheet.SheetProtectionValue, Boolean> givenProtection = prepareSheetProtectionValues(givenProtectionValues);
        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                for (Map.Entry<Worksheet.SheetProtectionValue, Boolean> item : givenProtection.entrySet())
                {
                    workbook.getCurrentWorksheet().addAllowedActionOnSheetProtection(item.getKey());
                }
                // adding values will enable sheet protection in any case, can be deactivated afterwards
                workbook.getCurrentWorksheet().setUseSheetProtection(useSheetProtection);
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(expectedProtection.size(), givenWorksheet.getSheetProtectionValues().size());
        assertEquals(useSheetProtection, givenWorksheet.isUseSheetProtection());
        for (Map.Entry<Worksheet.SheetProtectionValue, Boolean> item : expectedProtection.entrySet())
        {
            if (item.getValue())
            {
                assertTrue(givenWorksheet.getSheetProtectionValues().stream().anyMatch(x -> item.getKey() == x));
            }
        }
    }

    @DisplayName("Test of the 'sheetProtectionPasswordHash' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given password \"{0}\" should lead to the same hash on worksheet {1} when writing and reading a worksheet")
    @CsvSource({
            "'x', 0",
            "'@test-1,23', 0",
            "'', 0",
            ", 0",
            "'x', 1",
            "'@test-1,23', 2",
            "'', 3",
            ", 4",
    })
    void sheetProtectionPasswordHashWriteReadTest(String givenPassword, int sheetIndex) throws Exception {
        String hash = null;
        Workbook workbook = prepareWorkbook(5, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.deleteRows);
                workbook.getCurrentWorksheet().setSheetProtectionPassword(givenPassword);
                hash = workbook.getCurrentWorksheet().getSheetProtectionPasswordHash();
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(hash, givenWorksheet.getSheetProtectionPasswordHash());
    }

    @DisplayName("Test of the 'Hidden' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given hidden state: {0} should lead to the same state on worksheet {1} when writing and reading a worksheet")
    @CsvSource({
            "false, 0",
            "true, 0",
            "false, 1",
            "true, 1",
            "false, 2",
            "true, 2",
    })
    void hiddenWriteReadTest(boolean hidden, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (i == 0 && i == sheetIndex)
            {
                // Prevents setting selected worksheet as hidden
                workbook.setSelectedWorksheet(1);
            }
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setHidden(hidden);
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(hidden, givenWorksheet.isHidden());
    }

    @DisplayName("Test of the 'PaneSplitTopHeight' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given height {0} should lead to same value on worksheet {1} when writing and reading a worksheet")
    @CsvSource({
            "27f, 0",
    })
    void paneSplitTopHeightWriteReadTest(float height, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setHorizontalSplit(height, new Address("A1"), Worksheet.WorksheetPane.topLeft);
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(height, givenWorksheet.getPaneSplitTopHeight());
    }

    private static Map<Worksheet.SheetProtectionValue, Boolean> prepareSheetProtectionValues(String tokenString)
    {
        Map<Worksheet.SheetProtectionValue, Boolean> map = new HashMap<>();
        String[] tokens = tokenString.split(",");
        for (String token : tokens)
        {
            if (token.equals(""))
            {
                continue;
            }
            String[] subTokens = token.split(":");
            Worksheet.SheetProtectionValue value = Worksheet.SheetProtectionValue.valueOf(subTokens[0]);
            if (subTokens[1].equals("1"))
            {
                map.put(value, true);
            }
            else
            {
                map.put(value, false);
            }
        }
        return map;
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
