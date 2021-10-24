package ch.rabanti.nanoxlsx4j.workbooks;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

public class SetWorksheetTest {
    @DisplayName("Test of the setCurrentWorksheet function by index")
    @Test()
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

    @DisplayName("Test of the setCurrentWorksheet function by index")
    @Test()
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

    @DisplayName("Test of the SetCurrentWorksheet function by reference")
    @Test()
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

    @DisplayName("Test of the failing setCurrentWorksheet function on an invalid name")
    @Test()
    void setCurrentWorksheetFailTest() {
        Workbook workbook = new Workbook();
        assertNull(workbook.getCurrentWorksheet());
        workbook.addWorksheet("test1");
        String nullString = null;
        assertThrows(WorksheetException.class, () -> workbook.setCurrentWorksheet(nullString));
        assertThrows(WorksheetException.class, () -> workbook.setCurrentWorksheet(""));
        assertThrows(WorksheetException.class, () -> workbook.setCurrentWorksheet("test2"));
    }

    @DisplayName("Test of the failing setCurrentWorksheet function on an invalid index")
    @Test()
    void setCurrentWorksheetFailTest2() {
        Workbook workbook = new Workbook();
        assertNull(workbook.getCurrentWorksheet());
        workbook.addWorksheet("test1");
        assertThrows(RangeException.class, () -> workbook.setCurrentWorksheet(-1));
        assertThrows(RangeException.class, () -> workbook.setCurrentWorksheet(1));
    }

    @DisplayName("Test of the failing setCurrentWorksheet function on an invalid reference")
    @Test()
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

    @DisplayName("Test of the setSelectedWorksheet function by name")
    @Test()
    void setSelectedWorksheetTest() {
        Workbook workbook = new Workbook();
        workbook.addWorksheet("test1");
        workbook.addWorksheet("test2");
        workbook.addWorksheet("test3");
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.setSelectedWorksheet("test2");
        assertEquals(1, workbook.getSelectedWorksheet());
    }

    @DisplayName("Test of the setSelectedWorksheet function by index")
    @Test()
    void setSelectedWorksheetTest2() {
        Workbook workbook = new Workbook();
        workbook.addWorksheet("test1");
        workbook.addWorksheet("test2");
        workbook.addWorksheet("test3");
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.setSelectedWorksheet(1);
        assertEquals(1, workbook.getSelectedWorksheet());
    }

    @DisplayName("Test of the setSelectedWorksheet function by reference")
    @Test()
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

    @DisplayName("Test of the failing setSelectedWorksheet function on an invalid name")
    @Test()
    void setSelectedWorksheetFailTest() {
        Workbook workbook = new Workbook();
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.addWorksheet("test1");
        String nullString = null;
        assertThrows(WorksheetException.class, () -> workbook.setSelectedWorksheet(nullString));
        assertThrows(WorksheetException.class, () -> workbook.setSelectedWorksheet(""));
        assertThrows(WorksheetException.class, () -> workbook.setSelectedWorksheet("test2"));
    }

    @DisplayName("Test of the failing setSelectedWorksheet function on an invalid index")
    @Test()
    void setSelectedWorksheetFailTest2() {
        Workbook workbook = new Workbook();
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.addWorksheet("test1");
        assertThrows(RangeException.class, () -> workbook.setSelectedWorksheet(-1));
        assertThrows(RangeException.class, () -> workbook.setSelectedWorksheet(1));
    }

    @DisplayName("Test of the failing setSelectedWorksheet function on an invalid reference")
    @Test()
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
}
