package ch.rabanti.nanoxlsx4j.worksheets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;

public class ViewTest {
    @DisplayName("Test of the get function of the paneSplitTopHeight field")
    @Test()
    void paneSplitTopHeightTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getPaneSplitTopHeight().isEmpty());
        worksheet.setSplit(10f, 22.2f, new Address("A2"), Worksheet.WorksheetPane.BOTTOM_LEFT);
        assertTrue(worksheet.getPaneSplitTopHeight().isPresent());
        assertEquals(22.2f, worksheet.getPaneSplitTopHeight().orElseThrow());
        worksheet.resetSplit();
        assertTrue(worksheet.getPaneSplitTopHeight().isEmpty());
    }

    @DisplayName("Test of the get function of the paneSplitLeftWidth field")
    @Test()
    void paneSplitLeftWidthTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getPaneSplitLeftWidth().isEmpty());
        worksheet.setSplit(11.1f, 20f, new Address("A2"), Worksheet.WorksheetPane.BOTTOM_LEFT);
        assertTrue(worksheet.getPaneSplitLeftWidth().isPresent());
        assertEquals(11.1f, worksheet.getPaneSplitLeftWidth().orElseThrow());
        worksheet.resetSplit();
        assertTrue(worksheet.getPaneSplitLeftWidth().isEmpty());
    }

    @DisplayName("Test of the get function of the freezeSplitPanes field")
    @Test()
    void freezeSplitPanesTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getFreezeSplitPanes().isEmpty());
        worksheet.setSplit(2, 2, true, new Address("D4"), Worksheet.WorksheetPane.BOTTOM_RIGHT);
        assertTrue(worksheet.getFreezeSplitPanes().isPresent());
        assertEquals(true, worksheet.getFreezeSplitPanes().orElseThrow());
        worksheet.resetSplit();
        assertTrue(worksheet.getFreezeSplitPanes().isEmpty());
    }

    @DisplayName("Test of the get function of the paneSplitTopLeftCell field")
    @Test()
    void paneSplitTopLeftCellTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getPaneSplitTopLeftCell().isEmpty());
        worksheet.setSplit(10f, 22.2f, new Address("C4"), Worksheet.WorksheetPane.BOTTOM_LEFT);
        assertTrue(worksheet.getPaneSplitTopLeftCell().isPresent());
        assertEquals("C4", worksheet.getPaneSplitTopLeftCell().orElseThrow().getAddress());
        worksheet.resetSplit();
        assertTrue(worksheet.getPaneSplitTopLeftCell().isEmpty());
    }

    @DisplayName("Test of the get function of the paneSplitAddress field")
    @Test()
    void paneSplitAddressTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getPaneSplitAddress().isEmpty());
        worksheet.setSplit(2, 2, true, new Address("D4"), Worksheet.WorksheetPane.BOTTOM_RIGHT);
        assertTrue(worksheet.getPaneSplitAddress().isPresent());
        assertEquals("C3", worksheet.getPaneSplitAddress().orElseThrow().getAddress());
        worksheet.resetSplit();
        assertTrue(worksheet.getPaneSplitAddress().isEmpty());
    }

    @DisplayName("Test of the get function of the activePane field")
    @Test()
    void activePaneTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.getActivePane().isEmpty());
        worksheet.setSplit(2, 2, true, new Address("D4"), Worksheet.WorksheetPane.BOTTOM_RIGHT);
        assertTrue(worksheet.getActivePane().isPresent());
        assertEquals(Worksheet.WorksheetPane.BOTTOM_RIGHT, worksheet.getActivePane().orElseThrow());
        worksheet.resetSplit();
        assertTrue(worksheet.getActivePane().isEmpty());
    }

    @DisplayName("Test of the setHorizontalSplit function with height definition")
    @ParameterizedTest(name = "Given height {0}, top-left address {1} and active pane {2} should lead to a valid " +
            "horizontal split")
    @CsvSource(
            {
                    "22.2f, B2, BOTTOM_LEFT",
                    "0f, B2, BOTTOM_LEFT",
                    "500f, B2, BOTTOM_LEFT",
                    "22.2f, X1, BOTTOM_LEFT",
                    "0f, A1, BOTTOM_LEFT",
                    "500f, XFD1048576, BOTTOM_LEFT",
                    "22.2f, B2, TOP_RIGHT",
                    "0f, B2, BOTTOM_RIGHT",
                    "500f, B2, TOP_LEFT",}
    )
    void setHorizontalSplitTest(float height, String topLeftCellAddress, Worksheet.WorksheetPane activePane) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setHorizontalSplit(height, address, activePane);
        assertEquals(height, worksheet.getPaneSplitTopHeight().orElseThrow());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell().orElseThrow());
        assertEquals(activePane, worksheet.getActivePane().orElseThrow());
        assertTrue(worksheet.getFreezeSplitPanes().isEmpty());
        assertTrue(worksheet.getPaneSplitAddress().isEmpty());
    }

    @DisplayName("Test of the setHorizontalSplit function with row definition")
    @ParameterizedTest(name = "Given row {0}, freeze state {1}, top-left address {2} and active pane {3} should lead " +
            "to a valid horizontal split")
    @CsvSource(
            {
                    "3, false, D1, BOTTOM_LEFT",
                    "10, true, K11, BOTTOM_LEFT",
                    "3, false, E2, BOTTOM_RIGHT",
                    "10, true, L100, BOTTOM_RIGHT",
                    "3, false, F3, TOP_LEFT",
                    "10, true, M200, TOP_LEFT",
                    "3, false, F3, TOP_RIGHT",
                    "10, true, M11, TOP_RIGHT",}
    )
    void setHorizontalSplitTest2(
            int rowNumber, boolean freeze, String topLeftCellAddress, Worksheet.WorksheetPane activePane) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setHorizontalSplit(rowNumber, freeze, address, activePane);
        assertTrue(worksheet.getPaneSplitLeftWidth().isEmpty());
        assertTrue(worksheet.getPaneSplitTopHeight().isEmpty());
        Address expectedAddress = new Address(0, rowNumber);
        assertEquals(expectedAddress.getAddress(), worksheet.getPaneSplitAddress().orElseThrow().getAddress());
        assertEquals(freeze, worksheet.getFreezeSplitPanes().orElseThrow());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell().orElseThrow());
        assertEquals(activePane, worksheet.getActivePane().orElseThrow());
    }

    @DisplayName("Test of the failing setHorizontalSplit function")
    @ParameterizedTest(name = "Given row {0} and topLeftCell {2} with freezing state {1} should lead to an exception " +
            "= {3}")
    @CsvSource(
            {
                    "3, false, A1, true",
                    "3, true, A1, false",
                    "100, false, R100, true",
                    "100, true, R100, false",}
    )
    void setHorizontalSplitFailTest(int rowNumber, boolean freeze, String topLeftCellAddress, boolean expectedValid) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        if (expectedValid) {
            worksheet.setHorizontalSplit(rowNumber, freeze, address, Worksheet.WorksheetPane.BOTTOM_LEFT);
        } else {
            assertThrows(
                    WorksheetException.class,
                    () -> worksheet.setHorizontalSplit(rowNumber, freeze, address, Worksheet.WorksheetPane.BOTTOM_LEFT)
            );
        }
    }

    @DisplayName("Test of the setVerticalSplit function with width definition")
    @ParameterizedTest(name = "Given width {0}, top-left address {1} and active pane {2} should lead to a valid " +
            "horizontal split")
    @CsvSource(
            {
                    "22.2f, B2, BOTTOM_LEFT",
                    "0f, B2, BOTTOM_LEFT",
                    "500f, B2, BOTTOM_LEFT",
                    "22.2f, X1, BOTTOM_LEFT",
                    "0f, A1, BOTTOM_LEFT",
                    "500f, XFD1048576, BOTTOM_LEFT",
                    "22.2f, B2, TOP_RIGHT",
                    "0f, B2, BOTTOM_RIGHT",
                    "500f, B2, TOP_LEFT",}
    )
    void setVerticalSplitTest(float width, String topLeftCellAddress, Worksheet.WorksheetPane activePane) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setVerticalSplit(width, address, activePane);
        assertEquals(width, worksheet.getPaneSplitLeftWidth().orElseThrow());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell().orElseThrow());
        assertEquals(activePane, worksheet.getActivePane().orElseThrow());
        assertTrue(worksheet.getFreezeSplitPanes().isEmpty());
        assertTrue(worksheet.getPaneSplitAddress().isEmpty());
        assertTrue(worksheet.getPaneSplitTopHeight().isEmpty());
    }

    @DisplayName("Test of the setVerticalSplit function with column definition")
    @ParameterizedTest(name = "Given column {0}, freeze state {1}, top-left address {2} and active pane {3} should " +
            "lead to a valid vertical split")
    @CsvSource(
            {
                    "3, false, A1, BOTTOM_LEFT",
                    "10, true, K11, BOTTOM_LEFT",
                    "3, false, E2, BOTTOM_RIGHT",
                    "10, true, L100, BOTTOM_RIGHT",
                    "3, false, F3, TOP_LEFT",
                    "10, true, M200, TOP_LEFT",
                    "3, false, F3, TOP_RIGHT",
                    "10, true, M11, TOP_RIGHT",}
    )
    void setVerticalSplitTest2(
            int columnNumber, boolean freeze, String topLeftCellAddress, Worksheet.WorksheetPane activePane) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setVerticalSplit(columnNumber, freeze, address, activePane);
        assertTrue(worksheet.getPaneSplitLeftWidth().isEmpty());
        assertTrue(worksheet.getPaneSplitTopHeight().isEmpty());
        Address expectedAddress = new Address(columnNumber, 0);
        assertEquals(expectedAddress.getAddress(), worksheet.getPaneSplitAddress().orElseThrow().getAddress());
        assertEquals(freeze, worksheet.getFreezeSplitPanes().orElseThrow());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell().orElseThrow());
        assertEquals(activePane, worksheet.getActivePane().orElseThrow());
    }

    @DisplayName("Test of the failing setVerticalSplit function")
    @ParameterizedTest(name = "Given column {0} and topLeftCell {2} with freezing state {1} should lead to an " +
            "exception = {3}")
    @CsvSource(
            {
                    "3, false, A1, true",
                    "3, true, A1, false",
                    "100, false, R100, true",
                    "100, true, R100, false",}
    )
    void setVerticalSplitFailTest(int columnNumber, boolean freeze, String topLeftCellAddress, boolean expectedValid) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        if (expectedValid) {
            worksheet.setVerticalSplit(columnNumber, freeze, address, Worksheet.WorksheetPane.BOTTOM_LEFT);
        } else {
            assertThrows(
                    WorksheetException.class,
                    () -> worksheet.setVerticalSplit(
                            columnNumber,
                            freeze,
                            address,
                            Worksheet.WorksheetPane.BOTTOM_LEFT
                    )
            );
        }
    }

    @DisplayName("Test of the setSplit function with height and width definition")
    @ParameterizedTest(name = "Given height {0} width {1}, top-left address {2} and active pane {3} should lead to a " +
            "valid horizontal split")
    @CsvSource(
            {
                    "22.2f, 11.1f, B2, BOTTOM_LEFT",
                    "0f, 0f, B2, BOTTOM_LEFT",
                    "500f, 200f, B2, BOTTOM_LEFT",
                    "22.2f, 0f, X1, BOTTOM_LEFT",
                    ", 0f, A1, BOTTOM_LEFT",
                    "500f, , XFD1048576, BOTTOM_LEFT",
                    ", 22.2f, B2, TOP_RIGHT",
                    "0f, , B2, BOTTOM_RIGHT",
                    ", 500f, B2, TOP_LEFT",}
    )
    void setSplitTest(Float height, Float width, String topLeftCellAddress, Worksheet.WorksheetPane activePane) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setSplit(width, height, address, activePane);
        assertEquals(java.util.Optional.ofNullable(height), worksheet.getPaneSplitTopHeight());
        assertEquals(java.util.Optional.ofNullable(width), worksheet.getPaneSplitLeftWidth());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell().orElseThrow());
        assertEquals(activePane, worksheet.getActivePane().orElseThrow());
        assertTrue(worksheet.getFreezeSplitPanes().isEmpty());
        assertTrue(worksheet.getPaneSplitAddress().isEmpty());

    }

    @DisplayName("Test of the setSplit function with column and definition")
    @ParameterizedTest(name = "Given column {0}, row {1}, topLeftCell {3}  and active pane {4} with freezing state " +
            "{2} should lead to a valid split")
    @CsvSource(
            {
                    "3, 3, false, A1, BOTTOM_LEFT",
                    "10, 2, true, K11, BOTTOM_LEFT",
                    "3, 1, false, E2, BOTTOM_RIGHT",
                    "10, 99, true, L100, BOTTOM_RIGHT",
                    "3, , false, F3, TOP_LEFT",
                    ", 1, true, M200, TOP_LEFT",
                    "3, ,  false, F3, TOP_RIGHT",
                    ", 10, true, M11, TOP_RIGHT",}
    )
    void setSplitTest2(
            Integer columnNumber, Integer rowNumber, boolean freeze, String topLeftCellAddress,
            Worksheet.WorksheetPane activePane
    ) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setSplit(columnNumber, rowNumber, freeze, address, activePane);
        assertTrue(worksheet.getPaneSplitLeftWidth().isEmpty());
        assertTrue(worksheet.getPaneSplitTopHeight().isEmpty());
        int column = 0;
        if (columnNumber != null) {
            column = columnNumber;
        }
        int row = 0;
        if (rowNumber != null) {
            row = rowNumber;
        }
        Address expectedAddress = new Address(column, row);
        assertEquals(expectedAddress.getAddress(), worksheet.getPaneSplitAddress().orElseThrow().getAddress());
        assertEquals(freeze, worksheet.getFreezeSplitPanes().orElseThrow());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell().orElseThrow());
        assertEquals(activePane, worksheet.getActivePane().orElseThrow());
    }

    @DisplayName("Test of the failing setSplit function")
    @ParameterizedTest(name = "Given column {0}, row {1} and topLeftCell {3} with freezing state {2} should lead to " +
            "an exception = {4}")
    @CsvSource(
            {
                    "3, 3, false, A1, true",
                    "3, 0, true, A1, false",
                    "100, 1, false, R100, true",
                    "100, 1, true, R100, false",
                    "3, 3, false, B2, true",
                    "3, 0, true, B2, false",
                    "17, 1, false, R100, true",
                    "16, 100, true, R100, false",
                    "3, , true, E1, true",
                    ", 99, true, R100, true",
                    "3, , true, A1, false",
                    ", 101, true, R100, false",
                    ", , true, A1, true",

            }
    )
    void setSplitFailTest(
            Integer columnNumber, Integer rowNumber, boolean freeze, String topLeftCellAddress, boolean expectedValid) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        if (expectedValid) {
            worksheet.setSplit(columnNumber, rowNumber, freeze, address, Worksheet.WorksheetPane.BOTTOM_LEFT);
        } else {
            assertThrows(
                    WorksheetException.class,
                    () -> worksheet.setSplit(
                            columnNumber,
                            rowNumber,
                            freeze,
                            address,
                            Worksheet.WorksheetPane.BOTTOM_LEFT
                    )
            );
        }
    }

    @DisplayName("Test of the resetSplit function on a horizontal split with a height definition")
    @Test()
    void resetSplitTest() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setHorizontalSplit(22.2f, new Address("A1"), Worksheet.WorksheetPane.BOTTOM_LEFT);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the resetSplit function on a horizontal split with a row definition")
    @Test()
    void resetSplitTest2() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setHorizontalSplit(5, true, new Address("R6"), Worksheet.WorksheetPane.BOTTOM_LEFT);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the resetSplit function on a vertical split with a width definition")
    @Test()
    void resetSplitTest3() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setVerticalSplit(22.2f, new Address("A1"), Worksheet.WorksheetPane.BOTTOM_LEFT);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the resetSplit function on a vertical split with a column definition")
    @Test()
    void resetSplitTest4() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setVerticalSplit(5, true, new Address("R6"), Worksheet.WorksheetPane.BOTTOM_LEFT);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the resetSplit function on a split with a width and height definition")
    @Test()
    void resetSplitTest5() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setSplit(22.2f, 22.2f, new Address("A1"), Worksheet.WorksheetPane.BOTTOM_LEFT);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the resetSplit function on a split with a column and row definition")
    @Test()
    void resetSplitTest6() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setSplit(5, 5, true, new Address("R6"), Worksheet.WorksheetPane.BOTTOM_LEFT);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the get function of the showGridLine field")
    @Test()
    void showGridLinesTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.isShowGridLines());
        worksheet.setShowGridLines(false);
        assertFalse(worksheet.isShowGridLines());
    }

    @DisplayName("Test of the get function of the showRowColumnHeaders field")
    @Test()
    void showRowColumnHeadersTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.isShowRowColumnHeaders());
        worksheet.setShowRowColumnHeaders(false);
        assertFalse(worksheet.isShowRowColumnHeaders());
    }

    @DisplayName("Test of the get function of the showRuler field")
    @Test()
    void showRulerTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.isShowRuler());
        worksheet.setShowRuler(false);
        assertFalse(worksheet.isShowRuler());
    }

    @DisplayName("Test of the get function of the viewType field")
    @ParameterizedTest(name = "Given view type {0} should lead to the same value")
    @CsvSource(
            {
                    "NORMAL",
                    "PAGE_BREAK_PREVIEW",
                    "PAGE_LAYOUT",
            }
    )
    void viewTypeTest(Worksheet.SheetViewType viewType) {
        Worksheet worksheet = new Worksheet();
        assertEquals(Worksheet.SheetViewType.NORMAL, worksheet.getViewType());
        worksheet.setViewType(viewType);
        assertEquals(viewType, worksheet.getViewType());
    }

    @DisplayName("Test of the get function of the zoomFactor field on the current view type")
    @ParameterizedTest(name = "Given zoom factor {0} should lead to the same value")
    @CsvSource(
            {
                    "0",
                    "10",
                    "23",
                    "100",
                    "255",
                    "399",
                    "400",
            }
    )
    void zoomFactorTest(int zoomFactor) {
        Worksheet worksheet = new Worksheet();
        assertEquals(100, worksheet.getZoomFactor());
        worksheet.setZoomFactor(zoomFactor);
        assertEquals(zoomFactor, worksheet.getZoomFactor());
    }

    @DisplayName("Test of the get function of the zoomFactor and zoomFactors fields when the view type changes")
    @Test()
    void zoomFactorTest2() {
        int normalZoomFactor = 120;
        int pageBreakZoomFactor = 50;
        int pageLayoutZoomFactor = 400;

        Worksheet worksheet = new Worksheet();
        assertEquals(1, worksheet.getZoomFactors().size());
        assertEquals(100, worksheet.getZoomFactor());
        assertEquals(Worksheet.SheetViewType.NORMAL, worksheet.getViewType());
        worksheet.setZoomFactor(normalZoomFactor);
        worksheet.setViewType(Worksheet.SheetViewType.PAGE_BREAK_PREVIEW);
        worksheet.setZoomFactor(pageBreakZoomFactor);
        worksheet.setViewType(Worksheet.SheetViewType.PAGE_LAYOUT);
        worksheet.setZoomFactor(pageLayoutZoomFactor);

        assertEquals(3, worksheet.getZoomFactors().size());
        assertEquals(normalZoomFactor, worksheet.getZoomFactors().get(Worksheet.SheetViewType.NORMAL));
        assertEquals(pageBreakZoomFactor, worksheet.getZoomFactors().get(Worksheet.SheetViewType.PAGE_BREAK_PREVIEW));
        assertEquals(pageLayoutZoomFactor, worksheet.getZoomFactors().get(Worksheet.SheetViewType.PAGE_LAYOUT));
    }

    @DisplayName("Test of the failing zoomFactor set function")
    @ParameterizedTest(name = "Given zoom factor should lead to an exception")
    @CsvSource(
            {
                    "-1",
                    "-99",
                    "1",
                    "9",
                    "401",
                    "999",
            }
    )
    void zoomFactorFailTest(int zoomFactor) {
        Worksheet worksheet = new Worksheet();
        assertEquals(100, worksheet.getZoomFactor());
        assertThrows(WorksheetException.class, () -> worksheet.setZoomFactor(zoomFactor));
    }

    @DisplayName("Test of the setZoomFactor function")
    @ParameterizedTest(name = "Given zoom factor {0} and view type {1} should lead to the same zoom factor for the " +
            "given view types")
    @CsvSource(
            {
                    "0, NORMAL",
                    "10, PAGE_BREAK_PREVIEW",
                    "23, PAGE_LAYOUT",
                    "101, NORMAL",
                    "255, PAGE_BREAK_PREVIEW",
                    "399, PAGE_LAYOUT",
                    "400, NORMAL",
            }
    )
    void setZoomFactorTest(int zoomFactor, Worksheet.SheetViewType viewType) {
        Worksheet worksheet = new Worksheet();
        assertEquals(100, worksheet.getZoomFactor());
        worksheet.setZoomFactor(viewType, zoomFactor);
        assertEquals(zoomFactor, worksheet.getZoomFactors().get(viewType));
    }

    @DisplayName("Test of the failing zoomFactor set function")
    @ParameterizedTest(name = "Given zoom factor {0} and view type {1} should lead to an exception")
    @CsvSource(
            {
                    "-1, NORMAL",
                    "-99, PAGE_BREAK_PREVIEW",
                    "1, NORMAL",
                    "9, NORMAL",
                    "401, PAGE_LAYOUT",
                    "999, NORMAL",
            }
    )
    void setZoomFactorFailTest(int zoomFactor, Worksheet.SheetViewType viewType) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        assertEquals(100, worksheet.getZoomFactor());
        assertThrows(WorksheetException.class, () -> worksheet.setZoomFactor(viewType, zoomFactor));
    }

    private void assertInitializedPaneSplit(Worksheet worksheet) {
        assertTrue(worksheet.getPaneSplitTopHeight().isEmpty());
        assertTrue(worksheet.getPaneSplitLeftWidth().isEmpty());
        assertTrue(worksheet.getPaneSplitTopLeftCell().isEmpty());
        assertTrue(worksheet.getActivePane().isEmpty());
        assertTrue(worksheet.getFreezeSplitPanes().isEmpty());
        assertTrue(worksheet.getPaneSplitAddress().isEmpty());
    }
}




