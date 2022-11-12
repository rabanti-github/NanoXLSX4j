package ch.rabanti.nanoxlsx4j.workbooks;

import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.styles.Fill;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.util.Collections;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class WorkbookWriteReadTest {

    @DisplayName("Test of the (virtual) 'MruColors' property when writing and reading a workbook")
    @Test()
    void readMruColorsTest() throws Exception {
        Workbook workbook = new Workbook();
        String color1 = "AACC00";
        String color2 = "FFDD22";
        workbook.addMruColor(color1);
        workbook.addMruColor(color2);
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        List<String> mruColors = givenWorkbook.getMruColors();
        Collections.sort(mruColors);
        assertEquals(2, mruColors.size());
        assertEquals("FF" + color1, mruColors.get(0));
        assertEquals("FF" + color2, mruColors.get(1));
    }

    @DisplayName("Test of the (virtual) 'MruColors' property when writing and reading a workbook, neglecting the default color")
    @Test()
    void readMruColorsTest2() throws Exception {
        Workbook workbook = new Workbook();
        String color1 = "AACC00";
        String color2 = Fill.DEFAULT_COLOR; // Should not be added (black)
        workbook.addMruColor(color1);
        workbook.addMruColor(color2);
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        List<String> mruColors = givenWorkbook.getMruColors();
        Collections.sort(mruColors);
        assertEquals(1, mruColors.size());
        assertEquals("FF" + color1, mruColors.get(0));
    }

    @DisplayName("Test of the 'Hidden' property when writing and reading a workbook")
    @ParameterizedTest(name = "Given state {0} should lead to a loaded workbook with this state")
    @CsvSource(
            {
                    "true",
                    "false",}
    )
    void readWorkbookHiddenTest(boolean hidden) throws Exception {
        Workbook workbook = new Workbook();
        workbook.setHidden(hidden);
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        assertEquals(hidden, givenWorkbook.isHidden());
    }

    @DisplayName("Test of the 'SelectedWorksheet' property when writing and reading a workbook")
    @ParameterizedTest(name = "Given index {0} should lead to a loaded workbook with this index")
    @CsvSource(
            {
                    "0",
                    "1",
                    "2",}
    )
    void readWorkbookSelectedWorksheetTest(int index) throws Exception {
        Workbook workbook = new Workbook("sheet1");
        workbook.addWorksheet("sheet2");
        workbook.addWorksheet("sheet3");
        workbook.addWorksheet("sheet4");
        workbook.setSelectedWorksheet(index);
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        assertEquals(index, givenWorkbook.getSelectedWorksheet());
    }

    @DisplayName("Test of the 'LockWindowsIfProtected' property when writing and reading a workbook")
    @ParameterizedTest(name = "Given state {0} should lead to a loaded workbook with this state")
    @CsvSource(
            {
                    "true",
                    "false",}
    )
    void readWorkbookLockWindowsTest(boolean locked) throws Exception {
        Workbook workbook = new Workbook("sheet1");
        workbook.setWorkbookProtection(true, locked, false, null);
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        assertEquals(locked, givenWorkbook.isWindowsLockedIfProtected());
    }

    @DisplayName("Test of the 'LockStructureIfProtected' property when writing and reading a workbook")
    @ParameterizedTest(name = "Given state {0} should lead to a loaded workbook with this state")
    @CsvSource(
            {
                    "true",
                    "false",}
    )
    void readWorkbookLockStructureTest(boolean locked) throws Exception {
        Workbook workbook = new Workbook("sheet1");
        workbook.setWorkbookProtection(true, false, locked, null);
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        assertEquals(locked, givenWorkbook.isStructureLockedIfProtected());
    }

    @DisplayName("Test of the 'WorkbookProtectionPasswordHash' property when writing and reading a workbook")
    @ParameterizedTest(name = "Given password {0} should lead to a loaded workbook with a hash of this value")
    @CsvSource(
            {
                    "NULL, null",
                    "STRING, ''",
                    "STRING,A",
                    "STRING,123",
                    "STRING,test",}
    )
    void readWorkbookPasswordHashTest(String sourceType, String sourceValue) throws Exception {
        String plainText = (String) TestUtils.createInstance(sourceType, sourceValue);
        Workbook workbook = new Workbook("sheet1");
        workbook.setWorkbookProtection(true, false, true, plainText);
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        String hash = Helper.generatePasswordHash(plainText);
        if (hash.equals("")) {
            hash = null;
        }
        assertEquals(hash, givenWorkbook.getWorkbookProtectionPasswordHash());
    }

}
