package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Column;
import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.util.ArrayList;
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
            assertTrue(columnIndices.stream().anyMatch(x -> x + 1 == column.getValue().getNumber())); // Not zero-based
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
