package ch.rabanti.nanoxlsx4j.reader;

import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

public class ReadMetaDataTest {

    @DisplayName("Test of name property of worksheets when loading a workbook")
    @Test()
    void worksheetNameTest() throws Exception {
        Workbook workbook = new Workbook();
        workbook.addWorksheet("test1");
        workbook.addWorksheet("test2");
        workbook.addWorksheet("test3");

        Workbook givenWorkbook = TestUtils.saveAndLoadWorkbook(workbook, null);

        assertEquals(3, givenWorkbook.getWorksheets().size());
        assertEquals("test1", givenWorkbook.getWorksheets().get(0).getSheetName());
        assertEquals("test2", givenWorkbook.getWorksheets().get(1).getSheetName());
        assertEquals("test3", givenWorkbook.getWorksheets().get(2).getSheetName());
    }

    @DisplayName("Test of hidden property of worksheets when loading a workbook")
    @Test()
    void worksheetHiddenTest() throws Exception {
        Workbook workbook = new Workbook();
        workbook.addWorksheet("test1");
        workbook.addWorksheet("test2");
        workbook.addWorksheet("test3");
        workbook.setSelectedWorksheet(1);
        workbook.getWorksheets().get(0).setHidden(true);
        workbook.getWorksheets().get(2).setHidden(true);

        Workbook givenWorkbook = TestUtils.saveAndLoadWorkbook(workbook, null);

        assertEquals(3, givenWorkbook.getWorksheets().size());
        assertTrue(givenWorkbook.getWorksheets().get(0).isHidden());
        assertFalse(givenWorkbook.getWorksheets().get(1).isHidden());
        assertTrue(givenWorkbook.getWorksheets().get(2).isHidden());
    }




}
