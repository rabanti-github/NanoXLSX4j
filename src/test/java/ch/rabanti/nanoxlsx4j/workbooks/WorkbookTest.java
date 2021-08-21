package ch.rabanti.nanoxlsx4j.workbooks;

import ch.rabanti.nanoxlsx4j.*;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.io.File;
import java.io.FileOutputStream;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.*;

public class WorkbookTest {

    @DisplayName("Test of the Shortener getter")
    @Test()
    void shortenerTest() {
        Workbook workbook = new Workbook(false);
        assertNotNull(workbook.WS);
        workbook.addWorksheet("Sheet1");
        workbook.WS.value("Test");
        workbook.addWorksheet("Sheet2");
        workbook.WS.value("Test2");
        assertEquals("Test", workbook.getWorksheets().get(0).getCell(new Address("A1")).getValue().toString());
        assertEquals("Test2", workbook.getWorksheets().get(1).getCell(new Address("A1")).getValue().toString());
    }

    @DisplayName("Test of the currentWorksheet getter")
    @Test()
    void currentWorksheetTest() {
        Workbook workbook = new Workbook(false);
        assertNull(workbook.getCurrentWorksheet());
        workbook.addWorksheet("Test1");
        assertNotNull(workbook.getCurrentWorksheet());
        assertEquals("Test1", workbook.getCurrentWorksheet().getSheetName());
        workbook.addWorksheet("Test2");
        assertNotNull(workbook.getCurrentWorksheet());
        assertEquals("Test2", workbook.getCurrentWorksheet().getSheetName());
    }

    @DisplayName("Test of the Filename getter and setter")
    @Test()
    void filenameTest() {
        try {
            String filename = getRandomName();
            Workbook workbook = new Workbook(filename, "test");
            assertEquals(filename, workbook.getFilename());
            workbook.save();
            assertExistingFile(filename, true);
            filename = getRandomName();
            workbook.setFilename(filename);
            workbook.save();
            assertExistingFile(filename, true);
        } catch (Exception ex) {
            fail();
        }
    }

    @DisplayName("Test of the lockStructureIfProtected getter")
    @Test()
    void lockStructureIfProtectedTest() {
        Workbook workbook = new Workbook(false);
        assertFalse(workbook.isStructureLockedIfProtected());
        workbook.setWorkbookProtection(true, false, true, "");
        assertTrue(workbook.isStructureLockedIfProtected());
        workbook.setWorkbookProtection(false, false, false, "");
        assertFalse(workbook.isStructureLockedIfProtected());
    }

    @DisplayName("Test of the useWorkbookProtection getter and setter")
    @Test()
    void useWorkbookProtectionTest() {
        Workbook workbook = new Workbook(false);
        assertFalse(workbook.isWorkbookProtectionUsed());
        workbook.setWorkbookProtection(true);
        assertTrue(workbook.isWorkbookProtectionUsed());
        workbook.setWorkbookProtection(false);
        assertFalse(workbook.isWorkbookProtectionUsed());
    }

    @DisplayName("Test of the useWorkbookProtection getter using indirect measures")
    @Test()
    void useWorkbookProtectionTest2()
    {
        Workbook workbook = new Workbook(false);
        assertFalse(workbook.isWorkbookProtectionUsed());
        workbook.setWorkbookProtection(true, true, true, "");
        assertTrue(workbook.isWorkbookProtectionUsed());
        workbook.setWorkbookProtection(false, false, false, "");
        assertFalse(workbook.isWorkbookProtectionUsed());
    }

    @DisplayName("Test of the WorkbookMetadata getter and setter")
    @Test()
    void workbookMetadataTest() {
        Workbook workbook = new Workbook(false);
        assertNotNull(workbook.getWorkbookMetadata()); // Should be initialized
        workbook.getWorkbookMetadata().setTitle("Test");
        assertEquals("Test", workbook.getWorkbookMetadata().getTitle());
        Metadata newMetaData = new Metadata();
        workbook.setWorkbookMetadata(newMetaData);
        assertNotEquals("Test", workbook.getWorkbookMetadata().getTitle());
    }

    @DisplayName("Test of the selectedWorksheet getter")
    @Test()
    void selectedWorksheetTest() {
        Workbook workbook = new Workbook("test1");
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.addWorksheet("test2");
        workbook.setSelectedWorksheet(1);
        assertEquals(1, workbook.getSelectedWorksheet());
    }

    @DisplayName("Test of the UseWorkbookProtection getter")
    @Test()
    void lockWindowsIfProtectedTest() {
        Workbook workbook = new Workbook(false);
        assertFalse(workbook.isWindowsLockedIfProtected());
        workbook.setWorkbookProtection(false, true, true, "");
        assertTrue(workbook.isWindowsLockedIfProtected());
        workbook.setWorkbookProtection(false, false, false, "");
        assertFalse(workbook.isWindowsLockedIfProtected());
    }

    @DisplayName("Test of the workbookProtectionPassword getter")
    @Test()
    void workbookProtectionPasswordTest() {
        Workbook workbook = new Workbook(false);
        assertNull(workbook.getWorkbookProtectionPassword());
        workbook.setWorkbookProtection(false, true, true, "test");
        assertEquals("test", workbook.getWorkbookProtectionPassword());
        workbook.setWorkbookProtection(false, false, false, "");
        assertEquals("", workbook.getWorkbookProtectionPassword());
        workbook.setWorkbookProtection(false, false, false, null);
        assertNull(workbook.getWorkbookProtectionPassword());
    }

    @DisplayName("Test of the worksheets getter")
    @Test()
    void worksheetsTest() {
        Workbook workbook = new Workbook(false);
        assertEquals(0, workbook.getWorksheets().size());
        workbook.addWorksheet("test1");
        workbook.addWorksheet("test2");
        assertEquals(2, workbook.getWorksheets().size());
        workbook.removeWorksheet("test2");
        assertEquals(1, workbook.getWorksheets().size());
    }

    @DisplayName("Test of the hidden getter and setter")
    @Test()
    void hiddenTest() {
        Workbook workbook = new Workbook(false);
        assertFalse(workbook.isHidden());
        workbook.setHidden(true);
        assertTrue(workbook.isHidden());
        workbook.setHidden(false);
        assertFalse(workbook.isHidden());
    }

    @DisplayName("Test of the Workbook default constructor")
    @Test()
    void workbookConstructorTest() {
        Workbook workbook = new Workbook();
        assertEquals(0, workbook.getWorksheets().size());
        assertNotNull(workbook.getWorkbookMetadata());
        assertNull(workbook.getCurrentWorksheet());
        assertNull(workbook.getFilename());
        assertNotNull(workbook.WS);
        assertEquals(0, workbook.getWorksheets().size());
    }

    @DisplayName("Test of the Workbook constructor with an automatic option to create an initial worksheet")
    @ParameterizedTest(name = "Given parameter {0} should lead to a new workbook")
    @CsvSource({
            "true, Sheet1",
            "false, ",
    })
    void workbookConstructorTest2(boolean givenValue, String expectedName) {
        Workbook workbook = new Workbook(givenValue);
        if (givenValue) {
            assertNotNull(workbook.getCurrentWorksheet());
            assertEquals(expectedName, workbook.getWorksheets().get(0).getSheetName());
            assertEquals(1, workbook.getWorksheets().size());
        } else {
            assertEquals(0, workbook.getWorksheets().size());
            assertNull(workbook.getCurrentWorksheet());
        }
        assertNotNull(workbook.getWorkbookMetadata());
        assertNull(workbook.getFilename());
        assertNotNull(workbook.WS);
    }

    @DisplayName("Test of the Workbook constructor with the name of the initially crated worksheet")
    @ParameterizedTest(name = "Given value {0} should lead to a new workbook with an initial worksheet, called {1}")
    @CsvSource({
            "Sheet1, Sheet1",
            "?, _",
            "'', Sheet1",
            ", Sheet1",
    })
    void workbookConstructorTest3(String givenName, String expectedName) {
        Workbook workbook = new Workbook(givenName);
        assertNotNull(workbook.getCurrentWorksheet());
        assertEquals(expectedName, workbook.getWorksheets().get(0).getSheetName());
        assertEquals(1, workbook.getWorksheets().size());
        assertNotNull(workbook.getWorkbookMetadata());
        assertNull(workbook.getFilename());
        assertNotNull(workbook.WS);
    }

    @DisplayName("Test of the Workbook constructor with the file name of the workbook and name of the initially crated worksheet")
    @ParameterizedTest(name = "Given file name {0} and worksheet name {1} should lead to a new workbook")
    @CsvSource({
            "f1.xlsx, Sheet1, Sheet1",
            "'', ?, _",
            ",'', Sheet1",
            "?, , Sheet1",
    })
    void workbookConstructorTest4(String fileName, String givenSheetName, String expectedSheetName) {
        Workbook workbook = new Workbook(fileName, givenSheetName);
        assertNotNull(workbook.getCurrentWorksheet());
        assertEquals(expectedSheetName, workbook.getWorksheets().get(0).getSheetName());
        assertEquals(1, workbook.getWorksheets().size());
        assertNotNull(workbook.getWorkbookMetadata());
        assertNotNull(workbook.WS);
        assertEquals(fileName, workbook.getFilename());
    }

    @DisplayName("Test of the Workbook constructor with the file name of the workbook, the name of the initially created worksheet and a sanitizing option")
    @ParameterizedTest(name = "Given file name {1}, worksheet name {2} should lead to an exception: {4} or a workbook wit a worksheet {3}")
    @CsvSource({
            "false, f1.xlsx, Sheet1, Sheet1, false",
            "false, , ?, , true",
            "false, , , , true",
            "false, ?, , , true",
            "true, f1.xlsx, Sheet1, Sheet1, false",
            "true, , ?, _, false",
            "true, , , Sheet1, false",
            "true, ?, , Sheet1, false",
    })
    void workbookConstructorTest5(boolean sanitize, String fileName, String givenSheetName, String expectedSheetName, boolean expectException) {
        if (expectException) {
            assertThrows(FormatException.class, () -> new Workbook(fileName, givenSheetName, sanitize));
        } else {
            Workbook workbook = new Workbook(fileName, givenSheetName, sanitize);
            assertNotNull(workbook.getCurrentWorksheet());
            assertEquals(expectedSheetName, workbook.getWorksheets().get(0).getSheetName());
            assertEquals(1, workbook.getWorksheets().size());
            assertNotNull(workbook.getWorkbookMetadata());
            assertNotNull(workbook.WS);
            assertEquals(fileName, workbook.getFilename());
        }
    }

    @DisplayName("Test of the addWorksheet function with the worksheet name")
    @ParameterizedTest(name = "Given name {0} and optional name {1} should lead to new worksheet")
    @CsvSource({
            "test, ",
            "test, test2",
            "0, _",
    })
    void addWorksheetTest(String name1, String name2) {
        Workbook workbook = new Workbook();
        assertEquals(0, workbook.getWorksheets().size());
        workbook.addWorksheet(name1);
        assertEquals(1, workbook.getWorksheets().size());
        assertEquals(name1, workbook.getWorksheets().get(0).getSheetName());
        if (name2 != null) {
            workbook.addWorksheet(name2);
            assertEquals(2, workbook.getWorksheets().size());
            assertEquals(name2, workbook.getWorksheets().get(1).getSheetName());
        }
    }

    @DisplayName("Test of the failing addWorksheet function with an invalid worksheet name")
    @ParameterizedTest(name = "Given worksheet name {1} with initial worksheet {0} should lead to an exception")
    @CsvSource({
            "Sheet1, ",
            "Sheet1, ",
            "Sheet1, ?",
            "Sheet1, Sheet1",
            "Sheet1, --------------------------------",
    })
    void addWorksheetFailTest(String initialWorksheetName, String invalidName) {
        Workbook workbook = new Workbook();
        workbook.addWorksheet(initialWorksheetName);
        assertThrows(Exception.class, () -> workbook.addWorksheet(invalidName));
    }

    @DisplayName("Test of the addWorksheet function with the worksheet name and a sanitation option")
    @ParameterizedTest(name = "Given worksheet name {1} with initial worksheet {0} and sanitation: {2} should lead to an exception: {3}")
    @CsvSource({
            "Sheet1, , false, false, ",
            "test, test, false, false, ",
            "Sheet1, , false, false, ",
            "Sheet1, --------------------------------, false, false, ",
            "Sheet1, ?, false, false, ",
            "Sheet1, Sheet2, false, true, Sheet2",
            "Sheet1, , true, true, Sheet2",
            "test, test, true, true, test1",
            "Sheet1, , true, true, Sheet2",
            "Sheet1, --------------------------------, true, true, -------------------------------",
            "Sheet1, ?, true, true, _",
    })
    void addWorksheetTest2(String initialWorksheetName, String name2, boolean sanitize, boolean expectedValid, String expectedSheetName) {
        Workbook workbook = new Workbook();
        assertEquals(0, workbook.getWorksheets().size());
        workbook.addWorksheet(initialWorksheetName);
        assertEquals(1, workbook.getWorksheets().size());
        if (expectedValid) {
            workbook.addWorksheet(name2, sanitize);
            assertEquals(2, workbook.getWorksheets().size());
            assertEquals(expectedSheetName, workbook.getWorksheets().get(1).getSheetName());
            assertEquals(expectedSheetName, workbook.getCurrentWorksheet().getSheetName());
        } else {
            assertThrows(Exception.class, () -> workbook.addWorksheet(name2, sanitize));
        }
    }

    @DisplayName("Test of the addWorksheet function with a Worksheet object")
    @Test()
    void addWorksheetTest3() {
        Workbook workbook = new Workbook();
        assertEquals(0, workbook.getWorksheets().size());
        Worksheet worksheet = new Worksheet();
        worksheet.setSheetName("test");
        workbook.addWorksheet(worksheet);
        assertEquals(1, workbook.getWorksheets().size());
        assertEquals("test", workbook.getWorksheets().get(0).getSheetName());
        assertEquals("test", workbook.getCurrentWorksheet().getSheetName());
    }

    @DisplayName("Test of the failing AddWorksheet function with a null object")
    @Test()
    void addWorksheetFailTest3() {
        Workbook workbook = new Workbook();
        Worksheet worksheet = null;
        assertThrows(Exception.class, () -> workbook.addWorksheet(worksheet));
    }

    @DisplayName("Test of the failing addWorksheet function with a worksheet and an empty name")
    @Test()
    void addWorksheetFailTest3b() {
        Workbook workbook = new Workbook();
        Worksheet worksheet = new Worksheet();
        assertThrows(Exception.class, () -> workbook.addWorksheet(worksheet));
    }

    @DisplayName("Test of the failing addWorksheet function with a worksheet with an already defined name")
    @Test()
    void addWorksheetFailTest3c() {
        Workbook workbook = new Workbook();
        workbook.addWorksheet("Sheet1");
        Worksheet worksheet = new Worksheet();
        worksheet.setSheetName("Sheet1");
        assertThrows(Exception.class, () -> workbook.addWorksheet(worksheet));
    }

    @DisplayName("Test of the addWorksheet function with the worksheet object and a sanitation option")
    @ParameterizedTest(name = "Given worksheet {1} with initial worksheet {0} and sanitation: {2} should lead to an exception: {3}")

    @CsvSource({
            "Sheet1, Sheet1, false, false, ",
            "Sheet1, , false, false, ",
            "Sheet1, Sheet1, true, true, Sheet2",
            "Sheet1, , true, true, Sheet2",
    })
    void addWorksheetTest4(String initialWorksheetName, String name2, boolean sanitize, boolean expectedValid, String expectedSheetName) {
        Workbook workbook = new Workbook();
        assertEquals(0, workbook.getWorksheets().size());
        workbook.addWorksheet(initialWorksheetName);
        assertEquals(1, workbook.getWorksheets().size());
        Worksheet worksheet = new Worksheet();
        if (name2 != null) {
            worksheet.setSheetName(name2);
        }
        if (expectedValid) {
            workbook.addWorksheet(worksheet, sanitize);
            assertEquals(2, workbook.getWorksheets().size());
            assertEquals(expectedSheetName, workbook.getWorksheets().get(1).getSheetName());
            assertEquals(expectedSheetName, workbook.getCurrentWorksheet().getSheetName());
        } else {
            assertThrows(Exception.class, () -> workbook.addWorksheet(worksheet, sanitize));
        }
    }

    @DisplayName("Test of the addWorksheet function for a valid Sheet ID assignment with a name")
    @Test()
    void addWorksheetTest5() {
        Workbook workbook = new Workbook();
        workbook.addWorksheet("test");
        assertEquals(1, workbook.getWorksheets().get(0).getSheetID());
        workbook.addWorksheet("test2");
        assertEquals(2, workbook.getWorksheets().get(1).getSheetID());
        workbook.removeWorksheet("test");
        workbook.addWorksheet("test3");
        assertEquals(2, workbook.getWorksheets().get(1).getSheetID());
        workbook.removeWorksheet("test2");
        workbook.removeWorksheet("test3");
        workbook.addWorksheet("test4");
        assertEquals(1, workbook.getWorksheets().get(0).getSheetID());
    }

    @DisplayName("Test of the addWorksheet function for a valid Sheet ID assignment with a worksheet object")
    @Test()
    void addWorksheetTest6() {
        Workbook workbook = new Workbook();
        Worksheet worksheet1 = new Worksheet();
        worksheet1.setSheetName("test");
        workbook.addWorksheet(worksheet1, true);
        assertEquals(1, workbook.getWorksheets().get(0).getSheetID());
        Worksheet worksheet2 = new Worksheet();
        worksheet2.setSheetName("test2");
        workbook.addWorksheet(worksheet2, true);
        assertEquals(2, workbook.getWorksheets().get(1).getSheetID());
        workbook.removeWorksheet("test");
        Worksheet worksheet3 = new Worksheet();
        worksheet3.setSheetName("test3");
        workbook.addWorksheet(worksheet3, true);
        assertEquals(2, workbook.getWorksheets().get(1).getSheetID());
        workbook.removeWorksheet("test2");
        workbook.removeWorksheet("test3");
        Worksheet worksheet4 = new Worksheet();
        workbook.addWorksheet(worksheet4, true);
        assertEquals(1, workbook.getWorksheets().get(0).getSheetID());
    }

    @DisplayName("Test of the RemoveWorksheet function by name")
    @ParameterizedTest(name = "Given worksheet count {0} with current worksheet {1} and selected worksheet {2} on removal of {3} should lead to a current worksheet {4}  and selected worksheet {5}")
    @CsvSource({
            "2, 0, 1, 1, '0', 0",
            "2, 1, 0, 1, '0', 0",
            "2, 1, 1, 1, '0', 0",
            "2, 0, 0, 1, '0', 0",
            "1, 0, 0, 0, '', 0",
            "5, 2, 2, 2, '4', 3",
            "5, 0, 0, 4, '0', 0",
            "4, 3, 1, 3, '2', 1",
            "4, 3, 3, 3, '2', 2",
    })
    void removeWorksheetTest(int worksheetCount, int currentWorksheetIndex, int selectedWorksheetIndex, int worksheetToRemoveIndex, String expectedCurrentWorksheetIndex, int expectedSelectedWorksheetIndex) {
        Workbook workbook = new Workbook();
        String current = null;
        String toRemove = null;
        String expected = null;
        for (int i = 0; i < worksheetCount; i++) {
            String name = "Sheet" + Integer.toString(i + 1);
            workbook.addWorksheet(name);
            if (i == currentWorksheetIndex) {
                current = name;
            }
            if (i == worksheetToRemoveIndex) {
                toRemove = name;
            }
            if (Integer.toString(i).equals(expectedCurrentWorksheetIndex)) {
                expected = name;
            }
        }
        assertWorksheetRemoval(workbook, workbook::removeWorksheet, worksheetCount, current, selectedWorksheetIndex, toRemove, expected, expectedSelectedWorksheetIndex);
    }

    @DisplayName("Test of the RemoveWorksheet function by index")
    @ParameterizedTest(name = "Given worksheet count {0} with current worksheet {1} and selected worksheet {2} on removal of {3} should lead to a current worksheet {4}  and selected worksheet {5}")
    @CsvSource({
            "2, 0, 1, 1, '0', 0",
            "2, 1, 0, 1, '0', 0",
            "2, 1, 1, 1, '0', 0",
            "2, 0, 0, 1, '0', 0",
            "1, 0, 0, 0, '', 0",
            "5, 2, 2, 2, '4', 3",
            "5, 0, 0, 4, '0', 0",
            "4, 3, 1, 3, '2', 1",
            "4, 3, 3, 3, '2', 2",
    })
    void removeWorksheetTest2(int worksheetCount, int currentWorksheetIndex, int selectedWorksheetIndex, int worksheetToRemoveIndex, String expectedCurrentWorksheetIndex, int expectedSelectedWorksheetIndex) {
        Workbook workbook = new Workbook();
        String current = null;
        int toRemove = -1;
        String expected = null;
        for (int i = 0; i < worksheetCount; i++) {
            String name = "Sheet" + Integer.toString(i + 1);
            workbook.addWorksheet(name);
            if (i == currentWorksheetIndex) {
                current = name;
            }
            if (i == worksheetToRemoveIndex) {
                toRemove = i;
            }
            if (Integer.toString(i).equals(expectedCurrentWorksheetIndex)) {
                expected = name;
            }
        }
        assertWorksheetRemoval(workbook, workbook::removeWorksheet, worksheetCount, current, selectedWorksheetIndex, toRemove, expected, expectedSelectedWorksheetIndex);
    }

    @DisplayName("Test of the failing removeWorksheet function on an non-existing name")
    @ParameterizedTest(name = "Given existing worksheet {0} and non-existing {1} to remove should lead to an exception")
    @CsvSource({
            "test, ",
            "test, ''",
            "test, 'Test'",
            "test, 'Sheet1'",
    })
    void removeWorksheetFailTest(String existingWorksheet, String absentWorksheet) {
        Workbook workbook = new Workbook();
        workbook.addWorksheet(existingWorksheet);
        assertThrows(WorksheetException.class, () -> workbook.removeWorksheet(absentWorksheet));
    }

    @DisplayName("Test of the failing removeWorksheet function on an non-existing index")
    @ParameterizedTest(name = "Given existing worksheet {0} and non-existing {1} to remove should lead to an exception")
    @CsvSource({
            "test, -1",
            "test, 1",
            "test, 99",
    })
    void removeWorksheetFailTest2(String existingWorksheet, int absentIndex) {
        Workbook workbook = new Workbook();
        workbook.addWorksheet(existingWorksheet);
        assertThrows(WorksheetException.class, () -> workbook.removeWorksheet(absentIndex));
    }

    @DisplayName("Test of the SetWorkbookProtection function")
    @ParameterizedTest(name = "Given state {0}, protectWindow {1}, protectStructure {2} and password {3} should lead to a locked widows state {4}, a locked structure state {5} and a protection state {6}")
    @CsvSource({
            "false, false, false, , false, false, false",
            "true, false, false, '', false, false, false",
            "true, true, false, 'test', true, false, true",
            "true, false, true, , false, true, true",
            "true, true, true, ' ' , true, true, true",
            "false, true, false, '222', true, false, false",
            "false, false, true, '#*$', false, true, false",
            "false, true, true, '_-_', true, true, false",

    })
    void setWorkbookProtectionTest(boolean state, boolean protectWindows, boolean protectStructure, String password, boolean expectedLockWindowsState, boolean expectedLockStructureState, boolean expectedProtectionState)
    {
        Workbook workbook = new Workbook();
        workbook.setWorkbookProtection(state, protectWindows, protectStructure, password);
        assertEquals(expectedLockWindowsState, workbook.isWindowsLockedIfProtected());
        assertEquals(expectedLockStructureState, workbook.isStructureLockedIfProtected());
        assertEquals(expectedProtectionState, workbook.isWorkbookProtectionUsed());
        assertEquals(password, workbook.getWorkbookProtectionPassword());
    }

    private <T> void assertWorksheetRemoval(Workbook workbook, Consumer<T> removalFunction, int worksheetCount, String currentWorksheet, int selectedWorksheetIndex, T worksheetToRemove, String expectedCurrentWorksheet, int expectedSelectedWorksheetIndex) {
        workbook.setCurrentWorksheet(currentWorksheet);
        workbook.setSelectedWorksheet(selectedWorksheetIndex);
        assertEquals(worksheetCount, workbook.getWorksheets().size());
        assertEquals(currentWorksheet, workbook.getCurrentWorksheet().getSheetName());
        removalFunction.accept(worksheetToRemove);
        assertEquals(worksheetCount - 1, workbook.getWorksheets().size());
        if (expectedCurrentWorksheet == null) {
            assertNull(workbook.getCurrentWorksheet());
        } else {
            assertEquals(expectedCurrentWorksheet, workbook.getCurrentWorksheet().getSheetName());
        }
        assertEquals(expectedSelectedWorksheetIndex, workbook.getSelectedWorksheet());
    }

    public static void assertExistingFile(String expectedPath, boolean deleteAfterAssertion) {
        File file = new File(expectedPath);
        assertTrue(file.exists());
        if (deleteAfterAssertion) {
            try {
                file.delete();
            } catch (Exception ex) {
                System.out.println("Could not delete " + expectedPath);
            }
        }
    }

    public static String getRandomName() throws Exception {
        File tmp = new File(System.getProperty("java.io.tmpdir"));
        File file = File.createTempFile("tmp", ".xlsx", tmp);
        if (file.exists()){
            file.delete();
        }
        return file.getAbsolutePath();
    }


}
