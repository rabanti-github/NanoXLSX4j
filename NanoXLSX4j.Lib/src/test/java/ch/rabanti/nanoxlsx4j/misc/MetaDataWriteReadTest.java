package ch.rabanti.nanoxlsx4j.misc;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;

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

	@DisplayName("Test of the 'Description' property when writing and reading a workbook")
	@Test()
	void readDescriptionTest() throws Exception {
		Workbook workbook = new Workbook();
		workbook.getWorkbookMetadata().setDescription("description1");
		Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
		assertEquals("description1", givenWorkbook.getWorkbookMetadata().getDescription());
	}

	@DisplayName("Test of the 'HyperlinkBase' property when writing and reading a workbook")
	@Test()
	void readHyperlinkBaseTest() throws Exception {
		Workbook workbook = new Workbook();
		workbook.getWorkbookMetadata().setHyperlinkBase("hyperlinkBase1");
		Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
		assertEquals("hyperlinkBase1", givenWorkbook.getWorkbookMetadata().getHyperlinkBase());
	}

	@DisplayName("Test of the 'Keywords' property when writing and reading a workbook")
	@Test()
	void readKeywordsTest() throws Exception {
		Workbook workbook = new Workbook();
		workbook.getWorkbookMetadata().setKeywords("keyword1;keyword2");
		Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
		assertEquals("keyword1;keyword2", givenWorkbook.getWorkbookMetadata().getKeywords());
	}

	@DisplayName("Test of the 'Manager' property when writing and reading a workbook")
	@Test()
	void readManagerTest() throws Exception {
		Workbook workbook = new Workbook();
		workbook.getWorkbookMetadata().setManager("manager1");
		Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
		assertEquals("manager1", givenWorkbook.getWorkbookMetadata().getManager());
	}

	@DisplayName("Test of the 'Subject' property when writing and reading a workbook")
	@Test()
	void readSubjectTest() throws Exception {
		Workbook workbook = new Workbook();
		workbook.getWorkbookMetadata().setSubject("subject1");
		Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
		assertEquals("subject1", givenWorkbook.getWorkbookMetadata().getSubject());
	}

	@DisplayName("Test of the 'Title' property when writing and reading a workbook")
	@Test()
	void readTitleTest() throws Exception {
		Workbook workbook = new Workbook();
		workbook.getWorkbookMetadata().setTitle("title1");
		Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
		assertEquals("title1", givenWorkbook.getWorkbookMetadata().getTitle());
	}

	@DisplayName("Test of writing and reading a workbook with a null WorkbookMetadata object")
	@Test()
	void readNullTest() throws Exception {
		Workbook workbook = new Workbook();
		workbook.setWorkbookMetadata(null);
		Workbook givenWorkbook = TestUtils.writeAndReadWorkbook(workbook);
		assertNull(givenWorkbook.getWorkbookMetadata().getApplication());
		assertNull(givenWorkbook.getWorkbookMetadata().getCreator());
		assertNull(givenWorkbook.getWorkbookMetadata().getTitle());
	}

}
