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

    @DisplayName("Test of the 'ApplicationVersion' property when writing and reading a workbook")
    @Test()
    void readApplicationVersionTest() throws Exception {
        Workbook workbook = new Workbook();
        workbook.getWorkbookMetadata().setApplicationVersion("3.456");
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        assertEquals("3.456", givenWorkbook.getWorkbookMetadata().getApplicationVersion());
    }


    @DisplayName("Test of the 'Category' property when writing and reading a workbook")
    @Test()
    void readCategoryTest() throws Exception {
        Workbook workbook = new Workbook();
        workbook.getWorkbookMetadata().setCategory("cat1");
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        assertEquals("cat1", givenWorkbook.getWorkbookMetadata().getCategory());
    }

    @DisplayName("Test of the 'Company' property when writing and reading a workbook")
    @Test()
    void readCompanyTest() throws Exception {
        Workbook workbook = new Workbook();
        workbook.getWorkbookMetadata().setCompany("company1");
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        assertEquals("company1", givenWorkbook.getWorkbookMetadata().getCompany());
    }

    @DisplayName("Test of the 'ContentStatus' property when writing and reading a workbook")
    @Test()
    void readContentStatusTest() throws Exception {
        Workbook workbook = new Workbook();
        workbook.getWorkbookMetadata().setContentStatus("status1");
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        assertEquals("status1", givenWorkbook.getWorkbookMetadata().getContentStatus());
    }

    @DisplayName("Test of the 'Creator' property when writing and reading a workbook")
    @Test()
    void readCreatorTest() throws Exception {
        Workbook workbook = new Workbook();
        workbook.getWorkbookMetadata().setCreator("creator1");
        Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
        assertEquals("creator1", givenWorkbook.getWorkbookMetadata().getCreator());
    }


}
