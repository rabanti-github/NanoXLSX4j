package ch.rabanti.nanoxlsx4j.workbooks;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;

class SetWorksheetTest {

    @Test
    @DisplayName("Test of the setCurrentWorksheet function by index")
    void setCurrentWorksheetTest() {
        Workbook workbook = new Workbook();
        assertNull(workbook.getCurrentWorksheet());
        workbook.addWorksheet("test1");
        workbook.addWorksheet("test2");
        workbook.addWorksheet("test3");
        assertEquals("test3", workbook.getCurrentWorksheet().getSheetName());
        Worksheet worksheet = workbook.setCurrentWorksheet(1);
        assertEquals("test2", workbook.getCurrentWorksheet().getSheetName());
        assertEquals("test2", worksheet.getSheetName());
    }

    @Test
    @DisplayName("Test of the setCurrentWorksheet function by name")
    void setCurrentWorksheetTest2() {
        Workbook workbook = new Workbook();
        assertNull(workbook.getCurrentWorksheet());
        workbook.addWorksheet("test1");
        workbook.addWorksheet("test2");
        workbook.addWorksheet("test3");
        assertEquals("test3", workbook.getCurrentWorksheet().getSheetName());
        Worksheet worksheet = workbook.setCurrentWorksheet("test2");
        assertEquals("test2", workbook.getCurrentWorksheet().getSheetName());
        assertEquals("test2", worksheet.getSheetName());
    }

    @Test
    @DisplayName("Test of the setCurrentWorksheet function by reference")
    void setCurrentWorksheetTest3() {
        Workbook workbook = new Workbook();
        assertNull(workbook.getCurrentWorksheet());
        workbook.addWorksheet("test1");
        Worksheet worksheet = new Worksheet();
        worksheet.setSheetName("test2");
        workbook.addWorksheet(worksheet);
        workbook.addWorksheet("test3");
        assertEquals("test3", workbook.getCurrentWorksheet().getSheetName());
        workbook.setCurrentWorksheet(worksheet);
        assertEquals("test2", workbook.getCurrentWorksheet().getSheetName());
        assertEquals("test2", workbook.getWorksheets().get(1).getSheetName());
    }

    @Test
    @DisplayName("Test of the failing setCurrentWorksheet function on an invalid name")
    void setCurrentWorksheetFailTest() {
        Workbook workbook = new Workbook();
        assertNull(workbook.getCurrentWorksheet());
        workbook.addWorksheet("test1");
        String nullString = null;
        assertThrows(WorksheetException.class, () -> workbook.setCurrentWorksheet(nullString));
        assertThrows(WorksheetException.class, () -> workbook.setCurrentWorksheet(""));
        assertThrows(WorksheetException.class, () -> workbook.setCurrentWorksheet("test2"));
    }

    @Test
    @DisplayName("Test of the failing setCurrentWorksheet function on an invalid index")
    void setCurrentWorksheetFailTest2() {
        Workbook workbook = new Workbook();
        assertNull(workbook.getCurrentWorksheet());
        workbook.addWorksheet("test1");
        assertThrows(RangeException.class, () -> workbook.setCurrentWorksheet(-1));
        assertThrows(RangeException.class, () -> workbook.setCurrentWorksheet(1));
    }

    @Test
    @DisplayName("Test of the failing setCurrentWorksheet function on an invalid reference")
    void setCurrentWorksheetFailTest3() {
        Workbook workbook = new Workbook();
        assertNull(workbook.getCurrentWorksheet());
        workbook.addWorksheet("test1");
        Worksheet worksheet = new Worksheet();
        worksheet.setSheetName("test2");
        Worksheet nullWorksheet = null;
        assertThrows(WorksheetException.class, () -> workbook.setCurrentWorksheet(nullWorksheet));
        assertThrows(WorksheetException.class, () -> workbook.setCurrentWorksheet(worksheet));
    }

    @Test
    @DisplayName("Test of the setSelectedWorksheet function by name")
    void setSelectedWorksheetTest() {
        Workbook workbook = new Workbook();
        workbook.addWorksheet("test1");
        workbook.addWorksheet("test2");
        workbook.addWorksheet("test3");
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.setSelectedWorksheet("test2");
        assertEquals(1, workbook.getSelectedWorksheet());
    }

    @Test
    @DisplayName("Test of the setSelectedWorksheet function by index")
    void setSelectedWorksheetTest2() {
        Workbook workbook = new Workbook();
        workbook.addWorksheet("test1");
        workbook.addWorksheet("test2");
        workbook.addWorksheet("test3");
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.setSelectedWorksheet(1);
        assertEquals(1, workbook.getSelectedWorksheet());
    }

    @Test
    @DisplayName("Test of the setSelectedWorksheet function by reference")
    void setSelectedWorksheetTest3() {
        Workbook workbook = new Workbook();
        workbook.addWorksheet("test1");
        Worksheet worksheet = new Worksheet();
        worksheet.setSheetName("test2");
        workbook.addWorksheet(worksheet);
        workbook.addWorksheet("test3");
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.setSelectedWorksheet(worksheet);
        assertEquals(1, workbook.getSelectedWorksheet());
    }

    @Test
    @DisplayName("Test of the failing setSelectedWorksheet function on an invalid name")
    void setSelectedWorksheetFailTest() {
        Workbook workbook = new Workbook();
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.addWorksheet("test1");
        String nullString = null;
        assertThrows(WorksheetException.class, () -> workbook.setSelectedWorksheet(nullString));
        assertThrows(WorksheetException.class, () -> workbook.setSelectedWorksheet(""));
        assertThrows(WorksheetException.class, () -> workbook.setSelectedWorksheet("test2"));
    }

    @Test
    @DisplayName("Test of the failing setSelectedWorksheet function on an invalid index")
    void setSelectedWorksheetFailTest2() {
        Workbook workbook = new Workbook();
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.addWorksheet("test1");
        assertThrows(RangeException.class, () -> workbook.setSelectedWorksheet(-1));
        assertThrows(RangeException.class, () -> workbook.setSelectedWorksheet(1));
    }

    @Test
    @DisplayName("Test of the failing setSelectedWorksheet function on an invalid reference")
    void setSelectedWorksheetFailTest3() {
        Workbook workbook = new Workbook();
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.addWorksheet("test1");
        Worksheet worksheet = new Worksheet();
        worksheet.setSheetName("test2");
        Worksheet nullWorksheet = null;
        assertThrows(WorksheetException.class, () -> workbook.setSelectedWorksheet(nullWorksheet));
        assertThrows(WorksheetException.class, () -> workbook.setSelectedWorksheet(worksheet));
    }

    @Test
    @DisplayName("Test of the failing setSelectedWorksheet function when the worksheet is hidden")
    void selectedWorksheetFailTest() {
        Workbook workbook = new Workbook("test1");
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.addWorksheet("test2");
        workbook.setSelectedWorksheet(1);
        assertEquals(1, workbook.getSelectedWorksheet());
        workbook.getWorksheets().getFirst().setHidden(true);
        assertThrows(WorksheetException.class, () -> workbook.setSelectedWorksheet(0));
    }
}
