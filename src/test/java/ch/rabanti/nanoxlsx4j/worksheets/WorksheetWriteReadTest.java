package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

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
        Workbook workbook = new Workbook("worksheet1");
        workbook.addWorksheet("worksheet2");
        workbook.addWorksheet("worksheet3");
        workbook.addWorksheet("worksheet4");
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


    private static Worksheet writeAndReadWorksheet(Workbook workbook, int worksheetIndex) throws Exception {
        Workbook readWorkbook = TestUtils.saveAndLoadWorkbook(workbook, null);
        return readWorkbook.getWorksheets().get(worksheetIndex);
    }

}
