package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class PaneWriteReadTest {
    @DisplayName("Test of the 'PaneSplitTopHeight' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given height {0} and active pane {1} should lead to same value on worksheet {2} when writing and reading a worksheet")
    @CsvSource({
            "27f, NULL, 0",
            "100f, NULL,  0",
            "0f, NULL, 0",
            "27f, topLeft, 0",
            "100f, bottomLeft,  0",
            "0f, topRight, 0"
    })
    void paneSplitTopHeightWriteReadTest(float height, String activePaneString, int sheetIndex) throws Exception {

        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setHorizontalSplit(height, new Address("A1"), parsePane(activePaneString));
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(height, givenWorksheet.getPaneSplitTopHeight());
    }

    @DisplayName("Test of the 'PaneSplitLeftWidth' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given width {0} and active pane {1} should lead to same value on worksheet {2} when writing and reading a worksheet")
    @CsvSource({
            "27f, NULL, 0",
            "100f, NULL,  0",
            "0f, NULL, 0",
            "27f, topLeft, 0",
            "100f, bottomLeft,  0",
            "0f, topRight, 0"
    })
    void paneSplitLeftWidthWriteReadTest(float width, String activePaneString, int sheetIndex) throws Exception {

        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setVerticalSplit(width, new Address("A2"), parsePane(activePaneString));
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        // There may be a deviation by rounding
        float delta = Math.abs(width -  givenWorksheet.getPaneSplitLeftWidth());
        assertTrue(delta < 0.1);
    }

    @DisplayName("Test of the 'ActivePane' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given height {0} and active pane {1} should lead to same value on worksheet {2} when writing and reading a worksheet")
    @CsvSource({
            "27f, NULL, 0",
            "100f, topLeft,  0",
            "0f, bottomLeft, 0",
            "27f, topRight, 0",
            "100f, bottomRight,  0"
    })
    void paneSplitActivePaneWriteReadTest(float height, String activePaneString, int sheetIndex) throws Exception {

        Workbook workbook = prepareWorkbook(4, "test");
        Worksheet.WorksheetPane activePane = parsePane(activePaneString);
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setHorizontalSplit(height, new Address("A1"), activePane);
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(activePane, givenWorksheet.getActivePane());
    }

    @DisplayName("Test of the 'TopLeftCell' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given height {0}, active pane {1} and top left cell {2} should lead to same value on worksheet {3} when writing and reading a worksheet")
    @CsvSource({
            "27f, NULL, A1, 0",
            "100f, topLeft, B2, 0",
            "0f, bottomLeft, Z15, 0",
            "27f, topRight, $A1, 0",
            "100f, bottomRight, $D$4, 0"
    })
    void paneSplitTopLeftCellWriteReadTest(float height, String activePaneString, String topLeftCellAddress, int sheetIndex) throws Exception {

        Address topLeftCell = new Address(topLeftCellAddress);
        Workbook workbook = prepareWorkbook(4, "test");
        Worksheet.WorksheetPane activePane = parsePane(activePaneString);
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setHorizontalSplit(height, topLeftCell, activePane);
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(topLeftCell, givenWorksheet.getPaneSplitTopLeftCell());
    }




    @DisplayName("Test of the the 'PaneSplitTopHeight' and 'PaneSplitLeftWidth' properties (combined X/Y-Split) when writing and reading a worksheet")
    @ParameterizedTest(name = "Given width {0}, height {1} and active pane {2} should lead to same value on worksheet {3} when writing and reading a worksheet")
    @CsvSource({
            "27, 0f, NULL, 0",
            "100f, 0f, NULL,  0",
            "0f, 0f, NULL, 0",
            "27f, 27f, topLeft, 0",
            "100f, 27f,bottomLeft,  0",
            "0f, 27f, topRight,  0",
            "27f, 100f, NULL, 0",
            "100f, 100f, NULL,  0",
            "0f, 100f, NULL, 0",
            "27f, NULL, topLeft, 0",
            "100f, NULL,bottomLeft,  0",
            "0f, NULL, topRight,  0",
            "NULL, 100f, NULL, 0",
            "NULL, 100f, NULL,  0",
            "NULL, 100f, NULL, 0",
            "NULL, NULL,bottomLeft,  0"
    })
    void paneSplitWidthHeightWriteReadTest(String widthString, String heightString, String activePaneString, int sheetIndex) throws Exception {
        Float width = null;
        Float height = null;
        if (!widthString.equalsIgnoreCase("NULL")){
            width = (Float) TestUtils.createInstance("FLOAT", widthString);
        }
        if (!heightString.equalsIgnoreCase("NULL")){
            height = (Float)TestUtils.createInstance("FLOAT", heightString);
        }
        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++)
        {
            if (sheetIndex == i)
            {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setSplit(width, height, new Address("A2"), parsePane(activePaneString));
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        if (width == null){
            assertNull(givenWorksheet.getPaneSplitLeftWidth());
        }
        else{
            // There may be a deviation by rounding
            float delta = Math.abs(width -  givenWorksheet.getPaneSplitLeftWidth());
            assertTrue(delta < 0.1);
        }
    }

    private static Worksheet.WorksheetPane parsePane(String activePaneString) {
        if (activePaneString == null || activePaneString.equalsIgnoreCase("null")){
            return null;
        }
        return Worksheet.WorksheetPane.valueOf(activePaneString);
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
