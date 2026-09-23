package ch.rabanti.nanoxlsx4j.workbooks;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.function.Consumer;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Metadata;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.colors.Color;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;

class WorkbookTest {

    @Test
    @DisplayName("Test of the getShortener function")
    void shortenerTest() {
        Workbook workbook = new Workbook(false);
        assertNotNull(workbook.getShortener());
        workbook.addWorksheet("Sheet1");
        workbook.getShortener().value("Test");
        workbook.addWorksheet("Sheet2");
        workbook.getShortener().value("Test2");
        assertEquals("Test", workbook.getWorksheets().get(0).getCell(new Address("A1")).getValue().toString());
        assertEquals("Test2", workbook.getWorksheets().get(1).getCell(new Address("A1")).getValue().toString());
    }

    @Test
    @DisplayName("Test of the getCurrentWorksheet function")
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

    @Disabled("TODO: C# FilenameTest requires the Java Writer save integration")
    @Test
    @DisplayName("Test of the getFilename and setFilename functions")
    void filenameTest() {
        // TODO: C# reference: NanoXLSX.Writer-Reader.Test/Workbooks/WorkbookTest.cs
    }

    @Test
    @DisplayName("Test of the isLockStructureIfProtected function")
    void lockStructureIfProtectedTest() {
        Workbook workbook = new Workbook(false);
        assertFalse(workbook.isLockStructureIfProtected());
        workbook.setWorkbookProtection(true, false, true, "");
        assertTrue(workbook.isLockStructureIfProtected());
        workbook.setWorkbookProtection(false, false, false, "");
        assertFalse(workbook.isLockStructureIfProtected());
    }

    @Test
    @DisplayName("Test of the isLockWindowsIfProtected function")
    void lockWindowsIfProtectedTest() {
        Workbook workbook = new Workbook(false);
        assertFalse(workbook.isLockWindowsIfProtected());
        workbook.setWorkbookProtection(false, true, true, "");
        assertTrue(workbook.isLockWindowsIfProtected());
        workbook.setWorkbookProtection(false, false, false, "");
        assertFalse(workbook.isLockWindowsIfProtected());
    }

    @Test
    @DisplayName("Test of the getWorkbookMetadata and setWorkbookMetadata functions")
    void workbookMetadataTest() {
        Workbook workbook = new Workbook(false);
        assertNotNull(workbook.getWorkbookMetadata());
        workbook.getWorkbookMetadata().setTitle("Test");
        assertEquals("Test", workbook.getWorkbookMetadata().getTitle());
        Metadata newMetadata = new Metadata();
        workbook.setWorkbookMetadata(newMetadata);
        assertNotEquals("Test", workbook.getWorkbookMetadata().getTitle());
    }

    @Test
    @DisplayName("Test of the getSelectedWorksheet function")
    void selectedWorksheetTest() {
        Workbook workbook = new Workbook("test1");
        assertEquals(0, workbook.getSelectedWorksheet());
        workbook.addWorksheet("test2");
        workbook.setSelectedWorksheet(1);
        assertEquals(1, workbook.getSelectedWorksheet());
    }

    @Test
    @DisplayName("Test of the isUseWorkbookProtection and setUseWorkbookProtection functions")
    void useWorkbookProtectionTest() {
        Workbook workbook = new Workbook(false);
        assertFalse(workbook.isUseWorkbookProtection());
        workbook.setUseWorkbookProtection(true);
        assertTrue(workbook.isUseWorkbookProtection());
        workbook.setUseWorkbookProtection(false);
        assertFalse(workbook.isUseWorkbookProtection());
    }

    @Test
    @DisplayName("Test of the isUseWorkbookProtection function using indirect measures")
    void useWorkbookProtectionTest2() {
        Workbook workbook = new Workbook(false);
        assertFalse(workbook.isUseWorkbookProtection());
        workbook.setWorkbookProtection(true, true, true, "");
        assertTrue(workbook.isUseWorkbookProtection());
        workbook.setWorkbookProtection(false, false, false, "");
        assertFalse(workbook.isUseWorkbookProtection());
    }

    @Test
    @DisplayName("Test of the getWorkbookProtectionPassword function")
    void workbookProtectionPasswordTest() {
        Workbook workbook = new Workbook(false);
        assertNotNull(workbook.getWorkbookProtectionPassword());
        assertFalse(workbook.getWorkbookProtectionPassword().passwordIsSet());
        assertNull(workbook.getWorkbookProtectionPassword().getPassword());
        assertNull(workbook.getWorkbookProtectionPassword().getPasswordHash());

        workbook.setWorkbookProtection(false, true, true, "test");
        assertTrue(workbook.getWorkbookProtectionPassword().passwordIsSet());
        assertEquals("test", workbook.getWorkbookProtectionPassword().getPassword());
        assertNotNull(workbook.getWorkbookProtectionPassword().getPasswordHash());

        workbook.setWorkbookProtection(false, false, false, "");
        assertFalse(workbook.getWorkbookProtectionPassword().passwordIsSet());
        assertNull(workbook.getWorkbookProtectionPassword().getPassword());
        assertNull(workbook.getWorkbookProtectionPassword().getPasswordHash());

        workbook.setWorkbookProtection(false, false, false, null);
        assertFalse(workbook.getWorkbookProtectionPassword().passwordIsSet());
        assertNull(workbook.getWorkbookProtectionPassword().getPassword());
        assertNull(workbook.getWorkbookProtectionPassword().getPasswordHash());
    }

    @Test
    @DisplayName("Test of the getWorksheets function")
    void worksheetsTest() {
        Workbook workbook = new Workbook(false);
        assertTrue(workbook.getWorksheets().isEmpty());
        workbook.addWorksheet("test1");
        workbook.addWorksheet("test2");
        assertEquals(2, workbook.getWorksheets().size());
        workbook.removeWorksheet("test2");
        assertEquals(1, workbook.getWorksheets().size());
    }

    @Test
    @DisplayName("Test of the isHidden and setHidden functions")
    void hiddenTest() {
        Workbook workbook = new Workbook(false);
        assertFalse(workbook.isHidden());
        workbook.setHidden(true);
        assertTrue(workbook.isHidden());
        workbook.setHidden(false);
        assertFalse(workbook.isHidden());
    }

    @Test
    @DisplayName("Test of the Workbook default constructor")
    void workbookConstructorTest() {
        Workbook workbook = new Workbook();
        assertTrue(workbook.getWorksheets().isEmpty());
        assertNotNull(workbook.getWorkbookMetadata());
        assertNull(workbook.getCurrentWorksheet());
        assertNull(workbook.getFilename());
        assertNotNull(workbook.getShortener());
        assertTrue(workbook.getWorksheets().isEmpty());
    }

    @ParameterizedTest
    @DisplayName("Test of the Workbook constructor with an automatic option to create an initial worksheet")
    @CsvSource(
            value = {
                    "true|Sheet1",
                    "false|NULL"},
            delimiter = '|',
            nullValues = "NULL"
    )
    void workbookConstructorTest2(boolean givenValue, String expectedName) {
        Workbook workbook = new Workbook(givenValue);
        if (givenValue) {
            assertNotNull(workbook.getCurrentWorksheet());
            assertEquals(expectedName, workbook.getWorksheets().get(0).getSheetName());
            assertEquals(1, workbook.getWorksheets().size());
        } else {
            assertTrue(workbook.getWorksheets().isEmpty());
            assertNull(workbook.getCurrentWorksheet());
        }
        assertNotNull(workbook.getWorkbookMetadata());
        assertNull(workbook.getFilename());
        assertNotNull(workbook.getShortener());
    }

    @ParameterizedTest
    @DisplayName("Test of the Workbook constructor with the name of the initially crated worksheet")
    @CsvSource(
            value = {
                    "Sheet1|Sheet1",
                    "?|_",
                    "''|Sheet1",
                    "NULL|Sheet1"},
            delimiter = '|',
            nullValues = "NULL"
    )
    void workbookConstructorTest3(String givenName, String expectedName) {
        Workbook workbook = new Workbook(givenName);
        assertNotNull(workbook.getCurrentWorksheet());
        assertEquals(expectedName, workbook.getWorksheets().get(0).getSheetName());
        assertEquals(1, workbook.getWorksheets().size());
        assertNotNull(workbook.getWorkbookMetadata());
        assertNull(workbook.getFilename());
        assertNotNull(workbook.getShortener());
    }

    @ParameterizedTest
    @DisplayName("Test of the Workbook constructor with the file name and the name of the initially crated worksheet")
    @CsvSource(
            value = {
                    "f1.xlsx|Sheet1|Sheet1",
                    "''|?|_",
                    "NULL|''|Sheet1",
                    "?|NULL|Sheet1"},
            delimiter = '|',
            nullValues = "NULL"
    )
    void workbookConstructorTest4(String fileName, String givenSheetName, String expectedSheetName) {
        Workbook workbook = new Workbook(fileName, givenSheetName);
        assertNotNull(workbook.getCurrentWorksheet());
        assertEquals(expectedSheetName, workbook.getWorksheets().get(0).getSheetName());
        assertEquals(1, workbook.getWorksheets().size());
        assertNotNull(workbook.getWorkbookMetadata());
        assertNotNull(workbook.getShortener());
        assertEquals(fileName, workbook.getFilename());
    }

    @ParameterizedTest
    @DisplayName("Test of the Workbook constructor with the file name, worksheet name and a sanitizing option")
    @CsvSource(
            value = {
                    "false|f1.xlsx|Sheet1|Sheet1|false",
                    "false|''|?|NULL|true",
                    "false|NULL|''|NULL|true",
                    "false|?|NULL|NULL|true",
                    "true|f1.xlsx|Sheet1|Sheet1|false",
                    "true|''|?|_|false",
                    "true|NULL|''|Sheet1|false",
                    "true|?|NULL|Sheet1|false"
            },
            delimiter = '|',
            nullValues = "NULL"
    )
    void workbookConstructorTest5(
            boolean sanitize, String fileName, String givenSheetName,
            String expectedSheetName, boolean expectException
    ) {
        if (expectException) {
            assertThrows(FormatException.class, () -> new Workbook(fileName, givenSheetName, sanitize));
        } else {
            Workbook workbook = new Workbook(fileName, givenSheetName, sanitize);
            assertNotNull(workbook.getCurrentWorksheet());
            assertEquals(expectedSheetName, workbook.getWorksheets().get(0).getSheetName());
            assertEquals(1, workbook.getWorksheets().size());
            assertNotNull(workbook.getWorkbookMetadata());
            assertNotNull(workbook.getShortener());
            assertEquals(fileName, workbook.getFilename());
        }
    }

    @ParameterizedTest
    @DisplayName("Test of the addWorksheet function with the worksheet name")
    @CsvSource(
            value = {
                    "test|NULL",
                    "test|test2",
                    "0|_"},
            delimiter = '|',
            nullValues = "NULL"
    )
    void addWorksheetTest(String name1, String name2) {
        Workbook workbook = new Workbook();
        assertTrue(workbook.getWorksheets().isEmpty());
        workbook.addWorksheet(name1);
        assertEquals(1, workbook.getWorksheets().size());
        assertEquals(name1, workbook.getWorksheets().get(0).getSheetName());
        assertEquals(name1, workbook.getCurrentWorksheet().getSheetName());
        if (name2 != null) {
            workbook.addWorksheet(name2);
            assertEquals(2, workbook.getWorksheets().size());
            assertEquals(name2, workbook.getWorksheets().get(1).getSheetName());
            assertEquals(name2, workbook.getCurrentWorksheet().getSheetName());
        }
    }

    @ParameterizedTest
    @DisplayName("Test of the failing addWorksheet function with an invalid worksheet name")
    @CsvSource(
            value = {
                    "Sheet1|NULL",
                    "Sheet1|''",
                    "Sheet1|?",
                    "Sheet1|Sheet1",
                    "Sheet1|--------------------------------"},
            delimiter = '|',
            nullValues = "NULL"
    )
    void addWorksheetFailTest(String initialWorksheetName, String invalidName) {
        Workbook workbook = new Workbook();
        workbook.addWorksheet(initialWorksheetName);
        assertThrows(Exception.class, () -> workbook.addWorksheet(invalidName));
    }

    @ParameterizedTest
    @DisplayName("Test of the addWorksheet function with the worksheet name and a sanitation option")
    @CsvSource(
            value = {
                    "Sheet1|NULL|false|false|NULL",
                    "test|test|false|false|NULL",
                    "Sheet1|''|false|false|NULL",
                    "Sheet1|--------------------------------|false|false|NULL",
                    "Sheet1|?|false|false|NULL",
                    "Sheet1|Sheet2|false|true|Sheet2",
                    "Sheet1|NULL|true|true|Sheet2",
                    "test|test|true|true|test1",
                    "Sheet1|''|true|true|Sheet2",
                    "Sheet1|--------------------------------|true|true|-------------------------------",
                    "Sheet1|?|true|true|_"
            },
            delimiter = '|',
            nullValues = "NULL"
    )
    void addWorksheetTest2(
            String initialWorksheetName, String name2, boolean sanitize,
            boolean expectedValid, String expectedSheetName
    ) {
        Workbook workbook = new Workbook();
        assertTrue(workbook.getWorksheets().isEmpty());
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

    @Test
    @DisplayName("Test of the addWorksheet function with a Worksheet object")
    void addWorksheetTest3() {
        Workbook workbook = new Workbook();
        assertTrue(workbook.getWorksheets().isEmpty());
        Worksheet worksheet = new Worksheet();
        worksheet.setSheetName("test");
        workbook.addWorksheet(worksheet);
        assertEquals(1, workbook.getWorksheets().size());
        assertEquals("test", workbook.getWorksheets().get(0).getSheetName());
        assertEquals("test", workbook.getCurrentWorksheet().getSheetName());
    }

    @Test
    @DisplayName("Test of the failing addWorksheet function with a null object")
    void addWorksheetFailTest3() {
        Workbook workbook = new Workbook();
        Worksheet worksheet = null;
        assertThrows(Exception.class, () -> workbook.addWorksheet(worksheet));
    }

    @Test
    @DisplayName("Test of the failing addWorksheet function with a worksheet and an empty name")
    void addWorksheetFailTest3b() {
        Workbook workbook = new Workbook();
        Worksheet worksheet = new Worksheet();
        assertThrows(Exception.class, () -> workbook.addWorksheet(worksheet));
    }

    @Test
    @DisplayName("Test of the failing addWorksheet function with a worksheet with an already defined name")
    void addWorksheetFailTest3c() {
        Workbook workbook = new Workbook();
        workbook.addWorksheet("Sheet1");
        Worksheet worksheet = new Worksheet();
        worksheet.setSheetName("Sheet1");
        assertThrows(Exception.class, () -> workbook.addWorksheet(worksheet));
    }

    @ParameterizedTest
    @DisplayName("Test of the addWorksheet function with the worksheet object and a sanitation option")
    @CsvSource(
            value = {
                    "Sheet1|Sheet1|false|false|NULL",
                    "Sheet1|NULL|false|false|NULL",
                    "Sheet1|Sheet1|true|true|Sheet2",
                    "Sheet1|NULL|true|true|Sheet2"
            },
            delimiter = '|',
            nullValues = "NULL"
    )
    void addWorksheetTest4(
            String initialWorksheetName, String name2, boolean sanitize,
            boolean expectedValid, String expectedSheetName
    ) {
        Workbook workbook = new Workbook();
        assertTrue(workbook.getWorksheets().isEmpty());
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

    @Test
    @DisplayName("Test of the addWorksheet function for a valid sheet ID assignment with a name")
    void addWorksheetTest5() {
        Workbook workbook = new Workbook();
        workbook.addWorksheet("test");
        assertEquals(1, workbook.getWorksheets().get(0).getSheetId());
        workbook.addWorksheet("test2");
        assertEquals(2, workbook.getWorksheets().get(1).getSheetId());
        workbook.removeWorksheet("test");
        workbook.addWorksheet("test3");
        assertEquals(2, workbook.getWorksheets().get(1).getSheetId());
        workbook.removeWorksheet("test2");
        workbook.addWorksheet("test4");
        workbook.removeWorksheet("test3");
        assertEquals(1, workbook.getWorksheets().get(0).getSheetId());
    }

    @Test
    @DisplayName("Test of the addWorksheet function for a valid sheet ID assignment with a worksheet object")
    void addWorksheetTest6() {
        Workbook workbook = new Workbook();
        Worksheet worksheet1 = new Worksheet();
        worksheet1.setSheetName("test");
        workbook.addWorksheet(worksheet1, true);
        assertEquals(1, workbook.getWorksheets().get(0).getSheetId());
        Worksheet worksheet2 = new Worksheet();
        worksheet2.setSheetName("test2");
        workbook.addWorksheet(worksheet2, true);
        assertEquals(2, workbook.getWorksheets().get(1).getSheetId());
        workbook.removeWorksheet("test");
        Worksheet worksheet3 = new Worksheet();
        worksheet3.setSheetName("test3");
        workbook.addWorksheet(worksheet3, true);
        assertEquals(2, workbook.getWorksheets().get(1).getSheetId());
        workbook.removeWorksheet("test2");
        Worksheet worksheet4 = new Worksheet();
        workbook.addWorksheet(worksheet4, true);
        workbook.removeWorksheet("test3");
        assertEquals(1, workbook.getWorksheets().get(0).getSheetId());
    }

    @ParameterizedTest
    @DisplayName("Test of the removeWorksheet function by name")
    @CsvSource(
            {
                    "2,0,1,1,0,0",
                    "2,1,0,1,0,0",
                    "2,1,1,1,0,0",
                    "2,0,0,1,0,0",
                    "5,2,2,2,4,3",
                    "5,0,0,4,0,0",
                    "4,3,1,3,2,1",
                    "4,3,3,3,2,2"
            }
    )
    void removeWorksheetTest(
            int worksheetCount, int currentWorksheetIndex, int selectedWorksheetIndex,
            int worksheetToRemoveIndex, int expectedCurrentWorksheetIndex, int expectedSelectedWorksheetIndex
    ) {
        Workbook workbook = new Workbook();
        String current = null;
        String toRemove = null;
        String expected = null;
        for (int i = 0; i < worksheetCount; i++) {
            String name = "Sheet" + (i + 1);
            workbook.addWorksheet(name);
            if (i == currentWorksheetIndex) {
                current = name;
            }
            if (i == worksheetToRemoveIndex) {
                toRemove = name;
            }
            if (i == expectedCurrentWorksheetIndex) {
                expected = name;
            }
        }
        assertWorksheetRemoval(
                workbook, workbook::removeWorksheet, worksheetCount, current,
                selectedWorksheetIndex, toRemove, expected, expectedSelectedWorksheetIndex
        );
    }

    @ParameterizedTest
    @DisplayName("Test of the removeWorksheet function by index")
    @CsvSource(
            {
                    "2,0,1,1,0,0",
                    "2,1,0,1,0,0",
                    "2,1,1,1,0,0",
                    "2,0,0,1,0,0",
                    "5,2,2,2,4,3",
                    "5,0,0,4,0,0",
                    "4,3,1,3,2,1",
                    "4,3,3,3,2,2"
            }
    )
    void removeWorksheetTest2(
            int worksheetCount, int currentWorksheetIndex, int selectedWorksheetIndex,
            int worksheetToRemoveIndex, int expectedCurrentWorksheetIndex, int expectedSelectedWorksheetIndex
    ) {
        Workbook workbook = new Workbook();
        String current = null;
        int toRemove = -1;
        String expected = null;
        for (int i = 0; i < worksheetCount; i++) {
            String name = "Sheet" + (i + 1);
            workbook.addWorksheet(name);
            if (i == currentWorksheetIndex) {
                current = name;
            }
            if (i == worksheetToRemoveIndex) {
                toRemove = i;
            }
            if (i == expectedCurrentWorksheetIndex) {
                expected = name;
            }
        }
        assertWorksheetRemoval(
                workbook, workbook::removeWorksheet, worksheetCount, current,
                selectedWorksheetIndex, toRemove, expected, expectedSelectedWorksheetIndex
        );
    }

    @ParameterizedTest
    @DisplayName("Test of the failing removeWorksheet function on a non-existing name")
    @CsvSource(
            value = {
                    "test|NULL",
                    "test|''",
                    "test|Test",
                    "test|Sheet1"},
            delimiter = '|',
            nullValues = "NULL"
    )
    void removeWorksheetFailTest(String existingWorksheet, String absentWorksheet) {
        Workbook workbook = new Workbook();
        workbook.addWorksheet(existingWorksheet);
        assertThrows(WorksheetException.class, () -> workbook.removeWorksheet(absentWorksheet));
    }

    @ParameterizedTest
    @DisplayName("Test of the failing removeWorksheet function on a non-existing index")
    @CsvSource(
            {
                    "test,-1",
                    "test,1",
                    "test,99"}
    )
    void removeWorksheetFailTest2(String existingWorksheet, int absentIndex) {
        Workbook workbook = new Workbook();
        workbook.addWorksheet(existingWorksheet);
        assertThrows(WorksheetException.class, () -> workbook.removeWorksheet(absentIndex));
    }

    @Test
    @DisplayName("Test of the failing removeWorksheet function when the last worksheet is removed from a workbook")
    void removeWorksheetFailTest3() {
        Workbook workbook = new Workbook("worksheet1");
        assertThrows(WorksheetException.class, () -> workbook.removeWorksheet(0));
    }

    @ParameterizedTest
    @DisplayName("Test of the setWorkbookProtection function")
    @CsvSource(
            value = {
                    "false|false|false|NULL|false|false|false",
                    "true|false|false|''|false|false|false",
                    "true|true|false|test|true|false|true",
                    "true|false|true|NULL|false|true|true",
                    "true|true|true|' '|true|true|true",
                    "false|true|false|222|true|false|false",
                    "false|false|true|#*$|false|true|false",
                    "false|true|true|_-_|true|true|false"
            },
            delimiter = '|',
            nullValues = "NULL"
    )
    void setWorkbookProtectionTest(
            boolean state, boolean protectWindows, boolean protectStructure,
            String password, boolean expectedLockWindowsState, boolean expectedLockStructureState,
            boolean expectedProtectionState
    ) {
        Workbook workbook = new Workbook();
        workbook.setWorkbookProtection(state, protectWindows, protectStructure, password);
        assertEquals(expectedLockWindowsState, workbook.isLockWindowsIfProtected());
        assertEquals(expectedLockStructureState, workbook.isLockStructureIfProtected());
        assertEquals(expectedProtectionState, workbook.isUseWorkbookProtection());
        if (password == null || password.isEmpty()) {
            assertNull(workbook.getWorkbookProtectionPassword().getPassword());
        } else {
            assertEquals(password, workbook.getWorkbookProtectionPassword().getPassword());
        }
    }

    @Test
    @DisplayName("Test of the addMruColor function")
    void addMruColorTest() {
        Workbook workbook = new Workbook();
        assertTrue(workbook.getMruColors().isEmpty());
        workbook.addMruColor("AABBCC");
        assertEquals(1, workbook.getMruColors().size());
        assertEquals("FFAABBCC", workbook.getMruColors().getFirst().getArgbValue());
    }

    @Test
    @DisplayName("Test of the addMruColor function for ARGB values")
    void addMruColorTest2() {
        Workbook workbook = new Workbook();
        assertTrue(workbook.getMruColors().isEmpty());
        Color color = Color.createRgb("AABBCC");
        workbook.addMruColor(color);
        assertEquals(1, workbook.getMruColors().size());
        assertEquals("FFAABBCC", workbook.getMruColors().getFirst().getArgbValue());
    }

    @Test
    @DisplayName("Test of the addMruColor function for indexed values")
    void addMruColorTest3() {
        Workbook workbook = new Workbook();
        assertTrue(workbook.getMruColors().isEmpty());
        Color color = Color.createRgb("AABBCC");
        workbook.addMruColor(color);
        assertEquals(1, workbook.getMruColors().size());
        assertEquals("FFAABBCC", workbook.getMruColors().getFirst().getArgbValue());
    }

    @ParameterizedTest
    @NullSource
    @ValueSource(
            strings = {
                    "",
                    " ",
                    "GGGGGG",
                    "AABBCCDD22"}
    )
    @DisplayName("Test of the failing addMruColor function when adding an invalid color value")
    void addMruColorFailTest(String value) {
        Workbook workbook = new Workbook("worksheet1");
        assertThrows(StyleException.class, () -> workbook.addMruColor(value));
    }

    @Test
    @DisplayName("Test of the clearMruColors function")
    void clearMruColorsTest() {
        Workbook workbook = new Workbook();
        workbook.addMruColor("00AAFF");
        workbook.addMruColor("AABBCC");
        assertEquals(2, workbook.getMruColors().size());
        workbook.clearMruColors();
        assertTrue(workbook.getMruColors().isEmpty());
    }

    @Test
    @DisplayName("Test of the getMruColors function")
    void getMruColorsTest() {
        Workbook workbook = new Workbook();
        workbook.addMruColor("00AAFF");
        workbook.addMruColor("AABBCC");
        assertEquals(2, workbook.getMruColors().size());
        assertEquals("FF00AAFF", workbook.getMruColors().get(0).getArgbValue());
        assertEquals("FFAABBCC", workbook.getMruColors().get(1).getArgbValue());
    }

    @Test
    @DisplayName("Test of the getWorksheet function by name")
    void getWorksheetTest() {
        Workbook workbook = new Workbook("worksheet1");
        workbook.getCurrentWorksheet().addCell("WS1", "A1");
        workbook.addWorksheet("worksheet2");
        workbook.getCurrentWorksheet().addCell("WS2", "A1");
        workbook.addWorksheet("worksheet3");
        workbook.getCurrentWorksheet().addCell("WS3", "A1");
        Worksheet givenWorksheet = workbook.getWorksheet("worksheet2");
        assertEquals("WS2", givenWorksheet.getCells().get("A1").getValue());
    }

    @Test
    @DisplayName("Test of the getWorksheet function by index")
    void getWorksheetTest2() {
        Workbook workbook = new Workbook("worksheet1");
        workbook.getCurrentWorksheet().addCell("WS1", "A1");
        workbook.addWorksheet("worksheet2");
        workbook.getCurrentWorksheet().addCell("WS2", "A1");
        workbook.addWorksheet("worksheet3");
        workbook.getCurrentWorksheet().addCell("WS3", "A1");
        Worksheet givenWorksheet = workbook.getWorksheet(1);
        assertEquals("WS2", givenWorksheet.getCells().get("A1").getValue());
    }

    @Test
    @DisplayName("Test of the failing getWorksheet on a non existing worksheet name")
    void getWorksheetFailTest() {
        Workbook workbook = new Workbook("worksheet1");
        assertThrows(WorksheetException.class, () -> workbook.getWorksheet("worksheet3"));
    }

    @Test
    @DisplayName("Test of the failing getWorksheet on a non existing worksheet index")
    void getWorksheetFailTest2() {
        Workbook workbook = new Workbook("worksheet1");
        assertThrows(RangeException.class, () -> workbook.getWorksheet(1));
    }

    private <T> void assertWorksheetRemoval(
            Workbook workbook, Consumer<T> removalFunction, int worksheetCount,
            String currentWorksheet, int selectedWorksheetIndex, T worksheetToRemove,
            String expectedCurrentWorksheet, int expectedSelectedWorksheetIndex
    ) {
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
}
