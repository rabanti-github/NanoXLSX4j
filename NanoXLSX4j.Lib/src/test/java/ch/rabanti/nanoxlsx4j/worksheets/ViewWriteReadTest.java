package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ViewWriteReadTest {
    @DisplayName("Test of the 'PaneSplitTopHeight' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given height {0} and active pane {1} should lead to same value on worksheet {2} when writing and reading a worksheet")
    @CsvSource(
            {
                    "27f, NULL, 0",
                    "100f, NULL,  0",
                    "0f, NULL, 0",
                    "27f, topLeft, 0",
                    "100f, bottomLeft,  0",
                    "0f, topRight, 0"}
    )
    void paneSplitTopHeightWriteReadTest(float height, String activePaneString, int sheetIndex) throws Exception {

        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++) {
            if (sheetIndex == i) {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setHorizontalSplit(height,
                                                                  new Address("A1"),
                                                                  parsePane(activePaneString));
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(height, givenWorksheet.getPaneSplitTopHeight());
    }

    @DisplayName("Test of the 'PaneSplitTopHeight' property defined by a split address, when writing and reading a worksheet")
    @ParameterizedTest(
            name = "Given row number {0}, freeze state {1}, top left cell address {2} and active pane {3} should lead to same value on worksheet {4} when writing and reading a worksheet"
    )
    @CsvSource(
            {
                    "0, false, A2, NULL, 0",
                    "1, false, A2, NULL, 0",
                    "15, false, A18, NULL, 0",
                    "0, false, A2, topLeft, 0",
                    "1, false, A2, bottomLeft, 0",
                    "15, false, A18, topRight, 0",
                    "0, true, A2, NULL, 0",
                    "1, true, A2, NULL, 0",
                    "15, true, A18, NULL, 0",
                    "0, true, A2, topLeft, 0",
                    "1, true, A2, bottomLeft, 0",
                    "15, true, A18, topRight, 0",}
    )
    void paneSplitTopHeightWriteReadTest2(int rowNumber, boolean freeze, String topLeftCellAddress, String activePaneString, int sheetIndex) throws Exception {

        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++) {
            if (sheetIndex == i) {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setHorizontalSplit(rowNumber,
                                                                  freeze,
                                                                  new Address(topLeftCellAddress),
                                                                  parsePane(activePaneString));
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertRowSplit(rowNumber, freeze, givenWorksheet);
    }

    @DisplayName("Test of the 'PaneSplitTopHeight' property defined by a split address with custom row heights, when writing and reading a worksheet")
    @Test()
    void paneSplitTopHeightWriteReadTest3() throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        workbook.setCurrentWorksheet(0);
        workbook.getCurrentWorksheet().setRowHeight(0, 18f);
        workbook.getCurrentWorksheet().setRowHeight(2, 22.5f);
        workbook.getCurrentWorksheet().setHorizontalSplit(4, false, new Address("D1"), Worksheet.WorksheetPane.topLeft);

        float expectedHeight = 0f;
        for (int i = 0; i < 4; i++) {
            if (workbook.getCurrentWorksheet().getRowHeights().containsKey(i)) {
                expectedHeight += Helper.getInternalRowHeight(workbook.getCurrentWorksheet().getRowHeights().get(i));
            } else {
                expectedHeight += Helper.getInternalRowHeight(Worksheet.DEFAULT_ROW_HEIGHT);
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, 0);
        // There may be a deviation by rounding
        float delta = Math.abs(expectedHeight - givenWorksheet.getPaneSplitTopHeight());
        assertTrue(delta < 0.15);
        assertNull(givenWorksheet.getFreezeSplitPanes());
    }

    @DisplayName("Test of the 'PaneSplitLeftWidth' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given width {0} and active pane {1} should lead to same value on worksheet {2} when writing and reading a worksheet")
    @CsvSource(
            {
                    "27f, NULL, 0",
                    "100f, NULL,  0",
                    "0f, NULL, 0",
                    "27f, topLeft, 0",
                    "100f, bottomLeft,  0",
                    "0f, topRight, 0"}
    )
    void paneSplitLeftWidthWriteReadTest(float width, String activePaneString, int sheetIndex) throws Exception {

        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++) {
            if (sheetIndex == i) {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setVerticalSplit(width, new Address("A2"), parsePane(activePaneString));
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        // There may be a deviation by rounding
        float delta = Math.abs(width - givenWorksheet.getPaneSplitLeftWidth());
        assertTrue(delta < 0.1);
    }

    @DisplayName("Test of the 'PaneSplitLeftWidth' property defined by a split address, when writing and reading a worksheet")
    @ParameterizedTest(
            name = "Given column number {0}, freeze state {1}, top left cell address {2} and active pane {3} should lead to same value on worksheet {4} when writing and reading a worksheet"
    )
    @CsvSource(
            {
                    "0, false, A2, NULL, 0",
                    "1, false, B2, NULL, 0",
                    "5, false, G2, NULL, 0",
                    "0, false, A2, topLeft, 0",
                    "1, false, B2, bottomLeft, 0",
                    "5, false, G2, topRight, 0",
                    "0, true, A2, NULL, 0",
                    "1, true, B2, NULL, 0",
                    "5, true, G2, NULL, 0",
                    "0, true, A2, topLeft, 0",
                    "1, true, B2, bottomLeft, 0",
                    "5, true, G2, topRight, 0",}
    )
    void paneSplitLeftWidthWriteReadTest2(int columnNumber, boolean freeze, String topLeftCellAddress, String activePaneString, int sheetIndex)
            throws Exception {

        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++) {
            if (sheetIndex == i) {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setVerticalSplit(columnNumber,
                                                                freeze,
                                                                new Address(topLeftCellAddress),
                                                                parsePane(activePaneString));
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertColumnSplit(columnNumber, freeze, givenWorksheet, false);
    }

    @DisplayName("Test of the 'PaneSplitLeftWidth' property defined by a split address with custom column widths, when writing and reading a worksheet")
    @Test()
    void paneSplitLeftWidthWriteReadTest3() throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        workbook.setCurrentWorksheet(0);
        workbook.getCurrentWorksheet().setColumnWidth(0, 18f);
        workbook.getCurrentWorksheet().setColumnWidth(2, 22.5f);
        workbook.getCurrentWorksheet().setVerticalSplit(4, false, new Address("D1"), Worksheet.WorksheetPane.topLeft);

        float expectedWidth = 0f;
        for (int i = 0; i < 4; i++) {
            if (workbook.getCurrentWorksheet().getColumns().containsKey(i)) {
                expectedWidth += Helper
                        .getInternalColumnWidth(workbook.getCurrentWorksheet().getColumns().get(i).getWidth());
            } else {
                expectedWidth += Helper.getInternalColumnWidth(Worksheet.DEFAULT_COLUMN_WIDTH);
            }

        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, 0);
        // There may be a deviation by rounding
        float delta = Math.abs(expectedWidth - givenWorksheet.getPaneSplitLeftWidth());
        assertTrue(delta < 0.15);
        assertNull(givenWorksheet.getFreezeSplitPanes());
    }

    @DisplayName("Test of the 'ActivePane' property when writing and reading a worksheet")
    @ParameterizedTest(name = "Given height {0} and active pane {1} should lead to same value on worksheet {2} when writing and reading a worksheet")
    @CsvSource(
            {
                    "27f, NULL, 0",
                    "100f, topLeft,  0",
                    "0f, bottomLeft, 0",
                    "27f, topRight, 0",
                    "100f, bottomRight,  0"}
    )
    void paneSplitActivePaneWriteReadTest(float height, String activePaneString, int sheetIndex) throws Exception {

        Workbook workbook = prepareWorkbook(4, "test");
        Worksheet.WorksheetPane activePane = parsePane(activePaneString);
        for (int i = 0; i <= sheetIndex; i++) {
            if (sheetIndex == i) {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setHorizontalSplit(height, new Address("A1"), activePane);
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(activePane, givenWorksheet.getActivePane());
    }

    @DisplayName("Test of the 'TopLeftCell' property when writing and reading a worksheet")
    @ParameterizedTest(
            name = "Given height {0}, active pane {1} and top left cell {2} should lead to same value on worksheet {3} when writing and reading a worksheet"
    )
    @CsvSource(
            {
                    "27f, NULL, A1, 0",
                    "100f, topLeft, B2, 0",
                    "0f, bottomLeft, Z15, 0",
                    "27f, topRight, $A1, 0",
                    "100f, bottomRight, $D$4, 0"}
    )
    void paneSplitTopLeftCellWriteReadTest(float height, String activePaneString, String topLeftCellAddress, int sheetIndex) throws Exception {

        Address topLeftCell = new Address(topLeftCellAddress);
        Workbook workbook = prepareWorkbook(4, "test");
        Worksheet.WorksheetPane activePane = parsePane(activePaneString);
        for (int i = 0; i <= sheetIndex; i++) {
            if (sheetIndex == i) {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setHorizontalSplit(height, topLeftCell, activePane);
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(topLeftCell, givenWorksheet.getPaneSplitTopLeftCell());
    }

    @DisplayName("Test of the the 'PaneSplitTopHeight' and 'PaneSplitLeftWidth' properties (combined X/Y-Split) when writing and reading a worksheet")
    @ParameterizedTest(name = "Given width {0}, height {1} and active pane {2} should lead to same value on worksheet {3} when writing and reading a worksheet")
    @CsvSource(
            {
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
                    "NULL, NULL,bottomLeft,  0"}
    )
    void paneSplitWidthHeightWriteReadTest(String widthString, String heightString, String activePaneString, int sheetIndex) throws Exception {
        Float width = null;
        Float height = null;
        if (!widthString.equalsIgnoreCase("NULL")) {
            width = (Float) TestUtils.createInstance("FLOAT", widthString);
        }
        if (!heightString.equalsIgnoreCase("NULL")) {
            height = (Float) TestUtils.createInstance("FLOAT", heightString);
        }
        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++) {
            if (sheetIndex == i) {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setSplit(width, height, new Address("A2"), parsePane(activePaneString));
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        if (width == null) {
            assertNull(givenWorksheet.getPaneSplitLeftWidth());
        } else {
            // There may be a deviation by rounding
            float delta = Math.abs(width - givenWorksheet.getPaneSplitLeftWidth());
            assertTrue(delta < 0.1);
        }
    }

    @DisplayName("Test of the 'PaneSplitLeftWidth' and the 'PaneSplitLeftWidth' properties (combined X/Y-Split) defined by a split address, when writing and reading a worksheet")
    @ParameterizedTest(
            name = "Given column number {0}, row number {1},freeze state {2}, top left cell address {3} and active pane {4} should lead to same value on worksheet {5} when writing and reading a worksheet"
    )
    @CsvSource(
            {
                    "0, 0, false, A2, NULL, 0",
                    "1, 0, false, B2, NULL, 0",
                    "5, 0, false, G2, NULL, 0",
                    "0, 0, false, A2, topLeft, 0",
                    "1, 0, false, B2, bottomLeft, 0",
                    "5, 0, false, G2, topRight, 0",
                    "0, 1, true, A2, NULL, 0",
                    "1, 1, true, B2, NULL, 0",
                    "5, 1, true, G2, NULL, 0",
                    "0, 1, true, A2, topLeft, 0",
                    "1, 1, true, B2, bottomLeft, 0",
                    "5, 1, true, G2, topRight, 0",
                    "0, 15, false, A20, NULL, 0",
                    "1, 15, false, B20, NULL, 0",
                    "5, 15, false, G20, NULL, 0",
                    "0, 15, false, A20, topLeft, 0",
                    "1, 15, false, B20, bottomLeft, 0",
                    "5, 15, false, G20, topRight, 0",}
    )
    void paneSplitWidthHeightWriteReadTest2(int columnNumber, int rowNumber, boolean freeze, String topLeftCellAddress, String activePaneString, int sheetIndex)
            throws Exception {

        Workbook workbook = prepareWorkbook(4, "test");
        for (int i = 0; i <= sheetIndex; i++) {
            if (sheetIndex == i) {
                workbook.setCurrentWorksheet(i);
                workbook.getCurrentWorksheet().setSplit(columnNumber,
                                                        rowNumber,
                                                        freeze,
                                                        new Address(topLeftCellAddress),
                                                        parsePane(activePaneString));
            }
        }
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertColumnSplit(columnNumber, freeze, givenWorksheet, true);
        assertRowSplit(rowNumber, freeze, givenWorksheet);
    }

    @DisplayName("Test of the 'showGridLines' field, when writing and reading a worksheet")
    @ParameterizedTest(name = "Given value {0} should lead to the same value when reading worksheet {1}")
    @CsvSource(
            {
                    "true, 0",
                    "false, 0",
                    "true, 2",
                    "false, 2",
            }
    )
    void showGridLinesWriteReadTest(boolean showGridLines, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        workbook.setCurrentWorksheet(sheetIndex);
        workbook.getCurrentWorksheet().setShowingGridLines(showGridLines);
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(showGridLines, givenWorksheet.isShowingGridLines());
    }

    @DisplayName("Test of the 'showRowColumnHeaders' field, when writing and reading a worksheet")
    @ParameterizedTest(name = "Given value {0} should lead to the same value when reading worksheet {1}")
    @CsvSource(
            {
                    "true, 0",
                    "false, 0",
                    "true, 2",
                    "false, 2",
            }
    )
    void showRowColumnHeadersWriteReadTest(boolean showRowColumnHeaders, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        workbook.setCurrentWorksheet(sheetIndex);
        workbook.getCurrentWorksheet().setShowingRowColumnHeaders(showRowColumnHeaders);
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(showRowColumnHeaders, givenWorksheet.isShowingRowColumnHeaders());
    }

    @DisplayName("Test of the 'showRuler' field, when writing and reading a worksheet")
    @ParameterizedTest(name = "Given value {0} and view type {2} should lead to {1} when reading worksheet {3} ")
    @CsvSource(
            {
                    "true, true, pageLayout, 0",
                    "false, true, pageBreakPreview, 0",
                    "true, true, normal, 2",
                    "false, false, pageLayout, 2",
                    "true, true, pageBreakPreview, 2",
                    "false, true, normal, 1",
            }
    )
    void showRulerWriteReadTest(boolean showRuler, boolean expectedShowRuler, Worksheet.SheetViewType viewType, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        workbook.setCurrentWorksheet(sheetIndex);
        workbook.getCurrentWorksheet().setViewType(viewType);
        workbook.getCurrentWorksheet().setShowingRuler(showRuler);
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(viewType, givenWorksheet.getViewType());
        assertEquals(expectedShowRuler, givenWorksheet.isShowingRuler());
    }

    @DisplayName("Test of the 'viewType' field, when writing and reading a worksheet")
    @ParameterizedTest(name = "Given view type {0} should lead to the same value when reading worksheet {1}")
    @CsvSource(
            {
                    "pageLayout, 0",
                    "pageBreakPreview, 0",
                    "normal,0",
                    "pageLayout, 2",
                    "pageBreakPreview, 2",
                    "normal, 2",
            }
    )
    void viewTypeWriteReadTest(Worksheet.SheetViewType viewType, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        workbook.setCurrentWorksheet(sheetIndex);
        workbook.getCurrentWorksheet().setViewType(viewType);
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(viewType, givenWorksheet.getViewType());
    }

    @DisplayName("Test of the 'zoomFactor' field, when writing and reading a worksheet")
    @ParameterizedTest(name = "Given view type {0} and zoom factor {1} should lead to the same values when reading worksheet {2}")
    @CsvSource(
            {
                    "normal, 0, 0",
                    "normal, 10, 2",
                    "normal, 100, 0",
                    "normal, 250, 2",
                    "normal, 400, 0",
                    "pageLayout, 0, 2",
                    "pageLayout, 10, 0",
                    "pageLayout, 100, 2",
                    "pageLayout, 250, 0",
                    "pageLayout, 400, 2",
                    "pageBreakPreview, 0, 0",
                    "pageBreakPreview, 10, 2",
                    "pageBreakPreview, 100, 0",
                    "pageBreakPreview, 250, 2",
                    "pageBreakPreview, 400, 0",
            }
    )
    void zoomFactorWriteReadTest(Worksheet.SheetViewType viewType, int zoomFactor, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        workbook.setCurrentWorksheet(sheetIndex);
        workbook.getCurrentWorksheet().setViewType(viewType);
        workbook.getCurrentWorksheet().setZoomFactor(zoomFactor);
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        assertEquals(viewType, givenWorksheet.getViewType());
        assertEquals(zoomFactor, givenWorksheet.getZoomFactor());
    }

    @DisplayName("Test of the 'setZoomFactor' function, when writing and reading a worksheet")
    @ParameterizedTest(name = "Given view type {0} and additional view type {1} and zoom factor {2} should lead to the same values when reading worksheet {3}")
    @CsvSource(
            {
                    "pageLayout, normal, 0, 0",
                    "pageBreakPreview, normal, 10, 2",
                    "pageLayout, normal, 100, 0",
                    "pageBreakPreview, normal, 250, 2",
                    "pageLayout, normal, 400, 0",
                    "normal, pageLayout, 0, 2",
                    "pageBreakPreview, pageLayout, 10, 0",
                    "normal, pageLayout, 100, 2",
                    "pageBreakPreview, pageLayout, 250, 0",
                    "normal, pageLayout, 400, 2",
                    "normal, pageBreakPreview, 0, 0",
                    "pageLayout, pageBreakPreview, 10, 2",
                    "normal, pageBreakPreview, 100, 0",
                    "pageLayout, pageBreakPreview, 250, 2",
                    "normal, pageBreakPreview, 400, 0",
            }
    )
    void setZoomFactorWriteReadTest(Worksheet.SheetViewType initialViewType, Worksheet.SheetViewType additionalViewType, int zoomFactor, int sheetIndex) throws Exception {
        Workbook workbook = prepareWorkbook(4, "test");
        workbook.setCurrentWorksheet(sheetIndex);
        workbook.getCurrentWorksheet().setViewType(initialViewType);
        workbook.getCurrentWorksheet().setZoomFactor(additionalViewType, zoomFactor);
        workbook.saveAs("c:\\purge-temp\\testZoom.xlsx");
        Worksheet givenWorksheet = writeAndReadWorksheet(workbook, sheetIndex);
        if (initialViewType != Worksheet.SheetViewType.normal && additionalViewType != Worksheet.SheetViewType.normal) {
            assertEquals(3, givenWorksheet.getZoomFactors().size());
            assertEquals(100, givenWorksheet.getZoomFactors().get(Worksheet.SheetViewType.normal));
        } else {
            assertEquals(2, givenWorksheet.getZoomFactors().size());
        }
        assertEquals(zoomFactor, givenWorksheet.getZoomFactors().get(additionalViewType));
        assertEquals(100, givenWorksheet.getZoomFactors().get(initialViewType));
    }

    private static void assertColumnSplit(int columnNumber, boolean freeze, Worksheet givenWorksheet, boolean xyApplied) {
        if (columnNumber == 0 && !xyApplied) {
            // No split at all (row 0)
            assertNull(givenWorksheet.getPaneSplitAddress());
            assertNull(givenWorksheet.getFreezeSplitPanes());
        } else {
            if (freeze) {
                assertEquals(columnNumber, givenWorksheet.getPaneSplitAddress().Column);
                assertEquals(freeze, givenWorksheet.getFreezeSplitPanes());
            } else {
                float width = Helper.getInternalColumnWidth(Worksheet.DEFAULT_COLUMN_WIDTH) * columnNumber;
                if (width == 0) {
                    // Not applied as x split
                    assertNull(givenWorksheet.getPaneSplitLeftWidth());
                } else {
                    // There may be a deviation by rounding
                    float delta = Math.abs(width - givenWorksheet.getPaneSplitLeftWidth());
                    assertTrue(delta < 0.1);
                }
                assertNull(givenWorksheet.getFreezeSplitPanes());
            }
        }
    }

    private static void assertRowSplit(int rowNumber, boolean freeze, Worksheet givenWorksheet) {
        if (rowNumber == 0) {
            // No split at all (row 0)
            assertNull(givenWorksheet.getPaneSplitAddress());
            assertNull(givenWorksheet.getFreezeSplitPanes());
        } else {
            if (freeze) {
                assertEquals(rowNumber, givenWorksheet.getPaneSplitAddress().Row);
                assertEquals(freeze, givenWorksheet.getFreezeSplitPanes());
            } else {
                float height = Worksheet.DEFAULT_ROW_HEIGHT * rowNumber;
                assertEquals(height, givenWorksheet.getPaneSplitTopHeight());
                assertNull(givenWorksheet.getFreezeSplitPanes());
            }

        }
    }

    private static Worksheet.WorksheetPane parsePane(String activePaneString) {
        if (activePaneString == null || activePaneString.equalsIgnoreCase("null")) {
            return null;
        }
        return Worksheet.WorksheetPane.valueOf(activePaneString);
    }

    private static Workbook prepareWorkbook(int numberOfWorksheets, Object a1Data) {
        Workbook workbook = new Workbook();
        for (int i = 0; i < numberOfWorksheets; i++) {
            workbook.addWorksheet("worksheet" + (i + 1));
            workbook.getCurrentWorksheet().addCell(a1Data, "A1");
        }
        return workbook;
    }

    private static Worksheet writeAndReadWorksheet(Workbook workbook, int worksheetIndex) throws Exception {
        Workbook readWorkbook = TestUtils.saveAndLoadWorkbook(workbook, null);
        return readWorkbook.getWorksheets().get(worksheetIndex);
    }
}
