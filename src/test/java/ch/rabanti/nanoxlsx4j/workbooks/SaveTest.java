package ch.rabanti.nanoxlsx4j.workbooks;

import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.io.File;
import java.io.FileOutputStream;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.fail;

public class SaveTest {

    @DisplayName("Test of the save function (file System)")
    @Test()
    void saveTest() {
        try {
            String fileName = WorkbookTest.getRandomName();
            Workbook workbook = new Workbook(fileName, "test");
            File fi = new File(fileName);
            assertFalse(fi.exists());
            workbook.save();
            WorkbookTest.assertExistingFile(fileName, true);
        } catch (Exception ex) {
            fail();
        }
    }

    @DisplayName("Test of the failing save function (file System)")
    @ParameterizedTest(name = "Given file name {0} should lead to an exception")
    @CsvSource({
            "NULL, ''",
            "STRING, '?'",
            "STRING, ''",
    })
    void saveFailTest(String sourceType, String sourceValue) {
        String fileName = (String) TestUtils.createInstance(sourceType, sourceValue);
        Workbook workbook = new Workbook(fileName, "test");
        assertThrows(Exception.class, () -> workbook.save());
    }

    @DisplayName("Test of the saveAs function (file System)")
    @Test()
    void saveAsTest() {
        try {
            String fileName = WorkbookTest.getRandomName();
            Workbook workbook = new Workbook("test");
            File fi = new File(fileName);
            assertFalse(fi.exists());
            workbook.saveAs(fileName);
            WorkbookTest.assertExistingFile(fileName, true);
        } catch (Exception ex) {
            fail();
        }
    }

    @DisplayName("Test of the failing saveAs function (file System)")
    @ParameterizedTest(name = "Given file name {0} should lead to an exception")
    @CsvSource({
            "NULL, ''",
            "STRING, '?'",
            "STRING, ''",
    })
    void saveAsFailTest(String sourceType, String sourceValue) {
        String fileName = (String) TestUtils.createInstance(sourceType, sourceValue);
        Workbook workbook = new Workbook("test");
        assertThrows(Exception.class, () -> workbook.saveAs(fileName));
    }

    @DisplayName("Test of the saveAsStream function with an output stream")
    @Test()
    void saveAsStreamTest() {
        try {
            String fileName = WorkbookTest.getRandomName();
            Workbook workbook = new Workbook("test");
            FileOutputStream fs = new FileOutputStream(fileName);
            workbook.saveAsStream(fs);
            WorkbookTest.assertExistingFile(fileName, true);
        } catch (Exception ex) {
            throw new AssertionError("Stream exception", ex);
        }
    }

    @DisplayName("Test of the failing saveAsStream function with a already closed stream")
    @Test()
    void saveAsStreamFailTest() {
        try {
            String fileName = WorkbookTest.getRandomName();
            Workbook workbook = new Workbook("test");
            FileOutputStream fs = new FileOutputStream(fileName);
            fs.write(0);
            fs.close();
            assertThrows(Exception.class, () -> workbook.saveAsStream(fs));
        } catch (Exception ex) {
            throw new AssertionError("Stream exception", ex);
        }
    }

    @DisplayName("Test of the failing saveAsStream function with a null stream")
    @Test()
    void saveAsStreamFailTest2() {
        try {
            Workbook workbook = new Workbook("test");
            assertThrows(Exception.class, () -> workbook.saveAsStream(null));
        } catch (Exception ex) {
            throw new AssertionError("Stream exception", ex);
        }
    }
}
