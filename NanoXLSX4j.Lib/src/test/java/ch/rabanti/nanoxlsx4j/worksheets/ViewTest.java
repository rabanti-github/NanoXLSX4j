package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ViewTest {
    @DisplayName("Test of the get function of the paneSplitTopHeight field")
    @Test()
    void paneSplitTopHeightTest() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getPaneSplitTopHeight());
        worksheet.setSplit(10f, 22.2f, new Address("A2"), Worksheet.WorksheetPane.bottomLeft);
        assertNotNull(worksheet.getPaneSplitTopHeight());
        assertEquals(22.2f, worksheet.getPaneSplitTopHeight());
        worksheet.resetSplit();
        assertNull(worksheet.getPaneSplitTopHeight());
    }

    @DisplayName("Test of the get function of the paneSplitLeftWidth field")
    @Test()
    void paneSplitLeftWidthTest() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getPaneSplitLeftWidth());
        worksheet.setSplit(11.1f, 20f, new Address("A2"), Worksheet.WorksheetPane.bottomLeft);
        assertNotNull(worksheet.getPaneSplitLeftWidth());
        assertEquals(11.1f, worksheet.getPaneSplitLeftWidth());
        worksheet.resetSplit();
        assertNull(worksheet.getPaneSplitLeftWidth());
    }

    @DisplayName("Test of the get function of the freezeSplitPanes field")
    @Test()
    void freezeSplitPanesTest() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getFreezeSplitPanes());
        worksheet.setSplit(2, 2, true, new Address("D4"), Worksheet.WorksheetPane.bottomRight);
        assertNotNull(worksheet.getFreezeSplitPanes());
        assertEquals(true, worksheet.getFreezeSplitPanes());
        worksheet.resetSplit();
        assertNull(worksheet.getFreezeSplitPanes());
    }

    @DisplayName("Test of the get function of the paneSplitTopLeftCell field")
    @Test()
    void paneSplitTopLeftCellTest() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getPaneSplitTopLeftCell());
        worksheet.setSplit(10f, 22.2f, new Address("C4"), Worksheet.WorksheetPane.bottomLeft);
        assertNotNull(worksheet.getPaneSplitTopLeftCell());
        assertEquals("C4", worksheet.getPaneSplitTopLeftCell().getAddress());
        worksheet.resetSplit();
        assertNull(worksheet.getPaneSplitTopLeftCell());
    }

    @DisplayName("Test of the get function of the paneSplitAddress field")
    @Test()
    void paneSplitAddressTest() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getPaneSplitAddress());
        worksheet.setSplit(2, 2, true, new Address("D4"), Worksheet.WorksheetPane.bottomRight);
        assertNotNull(worksheet.getPaneSplitAddress());
        assertEquals("C3", worksheet.getPaneSplitAddress().getAddress());
        worksheet.resetSplit();
        assertNull(worksheet.getPaneSplitAddress());
    }

    @DisplayName("Test of the get function of the activePane field")
    @Test()
    void activePaneTest() {
        Worksheet worksheet = new Worksheet();
        assertNull(worksheet.getActivePane());
        worksheet.setSplit(2, 2, true, new Address("D4"), Worksheet.WorksheetPane.bottomRight);
        assertNotNull(worksheet.getActivePane());
        assertEquals(Worksheet.WorksheetPane.bottomRight, worksheet.getActivePane());
        worksheet.resetSplit();
        assertNull(worksheet.getActivePane());
    }

    @DisplayName("Test of the SetHorizontalSplit function with height definition")
    @ParameterizedTest(name = "Given height {0}, top-left address {1} and active pane {2} should lead to a valid horizontal split")
    @CsvSource(
            {
                    "22.2f, B2, bottomLeft",
                    "0f, B2, bottomLeft",
                    "500f, B2, bottomLeft",
                    "22.2f, X1, bottomLeft",
                    "0f, A1, bottomLeft",
                    "500f, XFD1048576, bottomLeft",
                    "22.2f, B2, topRight",
                    "0f, B2, bottomRight",
                    "500f, B2, topLeft",}
    )
    void setHorizontalSplitTest(float height, String topLeftCellAddress, Worksheet.WorksheetPane activePane) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setHorizontalSplit(height, address, activePane);
        assertEquals(height, worksheet.getPaneSplitTopHeight());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell());
        assertEquals(activePane, worksheet.getActivePane());
        assertNull(worksheet.getFreezeSplitPanes());
        assertNull(worksheet.getPaneSplitAddress());
    }

    @DisplayName("Test of the setHorizontalSplit function with row definition")
    @ParameterizedTest(name = "Given row {0}, freeze state {1}, top-left address {2} and active pane {3} should lead to a valid horizontal split")
    @CsvSource(
            {
                    "3, false, D1, bottomLeft",
                    "10, true, K11, bottomLeft",
                    "3, false, E2, bottomRight",
                    "10, true, L100, bottomRight",
                    "3, false, F3, topLeft",
                    "10, true, M200, topLeft",
                    "3, false, F3, topRight",
                    "10, true, M11, topRight",}
    )
    void setHorizontalSplitTest2(int rowNumber, boolean freeze, String topLeftCellAddress, Worksheet.WorksheetPane activePane) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setHorizontalSplit(rowNumber, freeze, address, activePane);
        assertNull(worksheet.getPaneSplitLeftWidth());
        assertNull(worksheet.getPaneSplitTopHeight());
        Address expectedAddress = new Address(0, rowNumber);
        assertEquals(expectedAddress.getAddress(), worksheet.getPaneSplitAddress().getAddress());
        assertEquals(freeze, worksheet.getFreezeSplitPanes());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell());
        assertEquals(activePane, worksheet.getActivePane());
    }

    @DisplayName("Test of the failing setHorizontalSplit function")
    @ParameterizedTest(name = "Given row {0} and topLeftCell {2} with freezing state {1} should lead to an exception = {3}")
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
            worksheet.setHorizontalSplit(rowNumber, freeze, address, Worksheet.WorksheetPane.bottomLeft);
        }
        else {
            assertThrows(
                    WorksheetException.class,
                    () -> worksheet.setHorizontalSplit(rowNumber, freeze, address, Worksheet.WorksheetPane.bottomLeft)
            );
        }
    }

    @DisplayName("Test of the setVerticalSplit function with width definition")
    @ParameterizedTest(name = "Given width {0}, top-left address {1} and active pane {2} should lead to a valid horizontal split")
    @CsvSource(
            {
                    "22.2f, B2, bottomLeft",
                    "0f, B2, bottomLeft",
                    "500f, B2, bottomLeft",
                    "22.2f, X1, bottomLeft",
                    "0f, A1, bottomLeft",
                    "500f, XFD1048576, bottomLeft",
                    "22.2f, B2, topRight",
                    "0f, B2, bottomRight",
                    "500f, B2, topLeft",}
    )
    void setVerticalSplitTest(float width, String topLeftCellAddress, Worksheet.WorksheetPane activePane) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setVerticalSplit(width, address, activePane);
        assertEquals(width, worksheet.getPaneSplitLeftWidth());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell());
        assertEquals(activePane, worksheet.getActivePane());
        assertNull(worksheet.getFreezeSplitPanes());
        assertNull(worksheet.getPaneSplitAddress());
        assertNull(worksheet.getPaneSplitTopHeight());
    }

    @DisplayName("Test of the setVerticalSplit function with column definition")
    @ParameterizedTest(name = "Given column {0}, freeze state {1}, top-left address {2} and active pane {3} should lead to a valid vertical split")
    @CsvSource(
            {
                    "3, false, A1, bottomLeft",
                    "10, true, K11, bottomLeft",
                    "3, false, E2, bottomRight",
                    "10, true, L100, bottomRight",
                    "3, false, F3, topLeft",
                    "10, true, M200, topLeft",
                    "3, false, F3, topRight",
                    "10, true, M11, topRight",}
    )
    void setVerticalSplitTest2(int columnNumber, boolean freeze, String topLeftCellAddress, Worksheet.WorksheetPane activePane) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setVerticalSplit(columnNumber, freeze, address, activePane);
        assertNull(worksheet.getPaneSplitLeftWidth());
        assertNull(worksheet.getPaneSplitTopHeight());
        Address expectedAddress = new Address(columnNumber, 0);
        assertEquals(expectedAddress.getAddress(), worksheet.getPaneSplitAddress().getAddress());
        assertEquals(freeze, worksheet.getFreezeSplitPanes());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell());
        assertEquals(activePane, worksheet.getActivePane());
    }

    @DisplayName("Test of the failing setVerticalSplit function")
    @ParameterizedTest(name = "Given column {0} and topLeftCell {2} with freezing state {1} should lead to an exception = {3}")
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
            worksheet.setVerticalSplit(columnNumber, freeze, address, Worksheet.WorksheetPane.bottomLeft);
        }
        else {
            assertThrows(
                    WorksheetException.class,
                    () -> worksheet.setVerticalSplit(
                            columnNumber,
                            freeze,
                            address,
                            Worksheet.WorksheetPane.bottomLeft
                    )
            );
        }
    }

    @DisplayName("Test of the setSplit function with height and width definition")
    @ParameterizedTest(name = "Given height {0} width {1}, top-left address {2} and active pane {3} should lead to a valid horizontal split")
    @CsvSource(
            {
                    "22.2f, 11.1f, B2, bottomLeft",
                    "0f, 0f, B2, bottomLeft",
                    "500f, 200f, B2, bottomLeft",
                    "22.2f, 0f, X1, bottomLeft",
                    ", 0f, A1, bottomLeft",
                    "500f, , XFD1048576, bottomLeft",
                    ", 22.2f, B2, topRight",
                    "0f, , B2, bottomRight",
                    ", 500f, B2, topLeft",}
    )
    void setSplitTest(Float height, Float width, String topLeftCellAddress, Worksheet.WorksheetPane activePane) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setSplit(width, height, address, activePane);
        assertEquals(height, worksheet.getPaneSplitTopHeight());
        assertEquals(width, worksheet.getPaneSplitLeftWidth());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell());
        assertEquals(activePane, worksheet.getActivePane());
        assertNull(worksheet.getFreezeSplitPanes());
        assertNull(worksheet.getPaneSplitAddress());

    }

    @DisplayName("Test of the setSplit function with column and definition")
    @ParameterizedTest(name = "Given column {0}, row {1}, topLeftCell {3}  and active pane {4} with freezing state {2} should lead to a valid split")
    @CsvSource(
            {
                    "3, 3, false, A1, bottomLeft",
                    "10, 2, true, K11, bottomLeft",
                    "3, 1, false, E2, bottomRight",
                    "10, 99, true, L100, bottomRight",
                    "3, , false, F3, topLeft",
                    ", 1, true, M200, topLeft",
                    "3, ,  false, F3, topRight",
                    ", 10, true, M11, topRight",}
    )
    void setSplitTest2(Integer columnNumber, Integer rowNumber, boolean freeze, String topLeftCellAddress, Worksheet.WorksheetPane activePane) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        worksheet.setSplit(columnNumber, rowNumber, freeze, address, activePane);
        assertNull(worksheet.getPaneSplitLeftWidth());
        assertNull(worksheet.getPaneSplitTopHeight());
        int column = 0;
        if (columnNumber != null) {
            column = columnNumber;
        }
        int row = 0;
        if (rowNumber != null) {
            row = rowNumber;
        }
        Address expectedAddress = new Address(column, row);
        assertEquals(expectedAddress.getAddress(), worksheet.getPaneSplitAddress().getAddress());
        assertEquals(freeze, worksheet.getFreezeSplitPanes());
        assertEquals(address, worksheet.getPaneSplitTopLeftCell());
        assertEquals(activePane, worksheet.getActivePane());
    }

    @DisplayName("Test of the failing setSplit function")
    @ParameterizedTest(name = "Given column {0}, row {1} and topLeftCell {3} with freezing state {2} should lead to an exception = {4}")
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
    void setSplitFailTest(Integer columnNumber, Integer rowNumber, boolean freeze, String topLeftCellAddress, boolean expectedValid) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        Address address = new Address(topLeftCellAddress);
        if (expectedValid) {
            worksheet.setSplit(columnNumber, rowNumber, freeze, address, Worksheet.WorksheetPane.bottomLeft);
        }
        else {
            assertThrows(
                    WorksheetException.class,
                    () -> worksheet.setSplit(
                            columnNumber,
                            rowNumber,
                            freeze,
                            address,
                            Worksheet.WorksheetPane.bottomLeft
                    )
            );
        }
    }

    @DisplayName("Test of the resetSplit function on a horizontal split with a height definition")
    @Test()
    void resetSplitTest() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setHorizontalSplit(22.2f, new Address("A1"), Worksheet.WorksheetPane.bottomLeft);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the resetSplit function on a horizontal split with a row definition")
    @Test()
    void resetSplitTest2() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setHorizontalSplit(5, true, new Address("R6"), Worksheet.WorksheetPane.bottomLeft);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the resetSplit function on a vertical split with a width definition")
    @Test()
    void resetSplitTest3() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setVerticalSplit(22.2f, new Address("A1"), Worksheet.WorksheetPane.bottomLeft);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the resetSplit function on a vertical split with a column definition")
    @Test()
    void resetSplitTest4() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setVerticalSplit(5, true, new Address("R6"), Worksheet.WorksheetPane.bottomLeft);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the ResetSplit function on a split with a width and height definition")
    @Test()
    void resetSplitTest5() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setSplit(22.2f, 22.2f, new Address("A1"), Worksheet.WorksheetPane.bottomLeft);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the ResetSplit function on a split with a column and row definition")
    @Test()
    void resetSplitTest6() {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        worksheet.setSplit(5, 5, true, new Address("R6"), Worksheet.WorksheetPane.bottomLeft);
        worksheet.resetSplit();
        assertInitializedPaneSplit(worksheet);
    }

    @DisplayName("Test of the get function of the ShowGridLine field")
    @Test()
    void showGridLinesTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.isShowingGridLines());
        worksheet.setShowingGridLines(false);
        assertFalse(worksheet.isShowingGridLines());
    }

    @DisplayName("Test of the get function of the ShowRowColumnHeaders field")
    @Test()
    void showRowColumnHeadersTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.isShowingRowColumnHeaders());
        worksheet.setShowingRowColumnHeaders(false);
        assertFalse(worksheet.isShowingRowColumnHeaders());
    }

    @DisplayName("Test of the get function of the ShowRuler field")
    @Test()
    void showRulerTest() {
        Worksheet worksheet = new Worksheet();
        assertTrue(worksheet.isShowingRuler());
        worksheet.setShowingRuler(false);
        assertFalse(worksheet.isShowingRuler());
    }

    @DisplayName("Test of the get function of the ViewType field")
    @ParameterizedTest(name = "Given view type {0} should lead to the same value")
    @CsvSource(
            {
                    "normal",
                    "pageBreakPreview",
                    "pageLayout",
            }
    )
    void viewTypeTest(Worksheet.SheetViewType viewType) {
        Worksheet worksheet = new Worksheet();
        assertEquals(Worksheet.SheetViewType.normal, worksheet.getViewType());
        worksheet.setViewType(viewType);
        assertEquals(viewType, worksheet.getViewType());
    }

    @DisplayName("Test of the get function of the ZoomFactor field on the current view type")
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

    @DisplayName("Test of the get function of the ZoomFactor and ZoomFactors fields when the view type changes")
    @Test()
    void zoomFactorTest2() {
        int normalZoomFactor = 120;
        int pageBreakZoomFactor = 50;
        int pageLayoutZoomFactor = 400;

        Worksheet worksheet = new Worksheet();
        assertEquals(1, worksheet.getZoomFactors().size());
        assertEquals(100, worksheet.getZoomFactor());
        assertEquals(Worksheet.SheetViewType.normal, worksheet.getViewType());
        worksheet.setZoomFactor(normalZoomFactor);
        worksheet.setViewType(Worksheet.SheetViewType.pageBreakPreview);
        worksheet.setZoomFactor(pageBreakZoomFactor);
        worksheet.setViewType(Worksheet.SheetViewType.pageLayout);
        worksheet.setZoomFactor(pageLayoutZoomFactor);

        assertEquals(3, worksheet.getZoomFactors().size());
        assertEquals(normalZoomFactor, worksheet.getZoomFactors().get(Worksheet.SheetViewType.normal));
        assertEquals(pageBreakZoomFactor, worksheet.getZoomFactors().get(Worksheet.SheetViewType.pageBreakPreview));
        assertEquals(pageLayoutZoomFactor, worksheet.getZoomFactors().get(Worksheet.SheetViewType.pageLayout));
    }

    @DisplayName("Test of the failing ZoomFactor set function")
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

    @DisplayName("Test of the SetZoomFactor function")
    @ParameterizedTest(name = "Given zoom factor {0} and view type {1} should lead to the same zoom factor for the given view types")
    @CsvSource(
            {
                    "0, normal",
                    "10, pageBreakPreview",
                    "23, pageLayout",
                    "101, normal",
                    "255, pageBreakPreview",
                    "399, pageLayout",
                    "400, normal",
            }
    )
    void setZoomFactorTest(int zoomFactor, Worksheet.SheetViewType viewType) {
        Worksheet worksheet = new Worksheet();
        assertEquals(100, worksheet.getZoomFactor());
        worksheet.setZoomFactor(viewType, zoomFactor);
        assertEquals(zoomFactor, worksheet.getZoomFactors().get(viewType));
    }

    @DisplayName("Test of the failing ZoomFactor set function")
    @ParameterizedTest(name = "Given zoom factor {0} and view type {1} should lead to an exception")
    @CsvSource(
            {
                    "-1, normal",
                    "-99, pageBreakPreview",
                    "1, normal",
                    "9, normal",
                    "401, pageLayout",
                    "999, normal",
            }
    )
    void setZoomFactorFailTest(int zoomFactor, Worksheet.SheetViewType viewType) {
        Worksheet worksheet = new Worksheet();
        assertInitializedPaneSplit(worksheet);
        assertEquals(100, worksheet.getZoomFactor());
        assertThrows(WorksheetException.class, () -> worksheet.setZoomFactor(viewType, zoomFactor));
    }

    private void assertInitializedPaneSplit(Worksheet worksheet) {
        assertNull(worksheet.getPaneSplitTopHeight());
        assertNull(worksheet.getPaneSplitLeftWidth());
        assertNull(worksheet.getPaneSplitTopLeftCell());
        assertNull(worksheet.getActivePane());
        assertNull(worksheet.getFreezeSplitPanes());
        assertNull(worksheet.getPaneSplitAddress());
    }
}
