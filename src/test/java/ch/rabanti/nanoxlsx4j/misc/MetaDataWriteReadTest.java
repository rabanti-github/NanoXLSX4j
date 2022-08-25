package ch.rabanti.nanoxlsx4j.misc;

import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class MetaDataWriteReadTest {

    @DisplayName("Test of the 'Application' property when writing and reading a workbook")
    @Test()
    void readApplicationTest() throws Exception {
        Workbook workbook = new Workbook();
        workbook.getWorkbookMetadata().setApplication("testApp");
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        assertEquals("testApp", givenWorkbook.getWorkbookMetadata().getApplication());
    }

}
