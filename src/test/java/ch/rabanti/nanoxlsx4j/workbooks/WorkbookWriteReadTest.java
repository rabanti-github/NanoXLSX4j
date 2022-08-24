package ch.rabanti.nanoxlsx4j.workbooks;

import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

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
        Workbook givenWorkbook = writeAndReadWorkbook(workbook);
        List<String> mruColors = ((List<String>)givenWorkbook.getMruColors());
        Collections.sort(mruColors);
        assertEquals(2, mruColors.size());
        assertEquals("FF" + color1, mruColors.get(0));
        assertEquals("FF" + color2, mruColors.get(1));
    }

    private static Workbook writeAndReadWorkbook(Workbook workbook) throws Exception {
        return TestUtils.saveAndLoadWorkbook(workbook, null);
    }
}
