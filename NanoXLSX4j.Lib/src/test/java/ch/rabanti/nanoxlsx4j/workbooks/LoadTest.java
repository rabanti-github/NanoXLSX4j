package ch.rabanti.nanoxlsx4j.workbooks;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.fail;

import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Workbook;

public class LoadTest {

	@DisplayName("Test of the Load function with a file name")
	@Test()
	void loadTest1() {
		Map<String, Object> data = createSampleData();
		Workbook workbook;
		try {
			String name = createWorksheet("test1", data);
			workbook = Workbook.load(name);
			assertEquals("test1", workbook.getWorksheets().get(0).getSheetName());
			for (Map.Entry<String, Object> item : data.entrySet()) {
				assertEquals(item.getValue(),
						workbook.getWorksheets().get(0).getCell(new Address(item.getKey())).getValue());
			}
			WorkbookTest.assertExistingFile(name, true);
		}
		catch (Exception ex) {
			fail();
		}
	}

	@DisplayName("Test of the Load function with a stream")
	@Test()
	void loadTest2() {
		Map<String, Object> data = createSampleData();
		Workbook workbook;
		try {
			String name = createWorksheet("test1", data);
			FileInputStream fs = new FileInputStream(name);
			workbook = Workbook.load(fs);
			assertEquals("test1", workbook.getWorksheets().get(0).getSheetName());
			for (Map.Entry<String, Object> item : data.entrySet()) {
				assertEquals(item.getValue(),
						workbook.getWorksheets().get(0).getCell(new Address(item.getKey())).getValue());
			}
			WorkbookTest.assertExistingFile(name, true);
		}
		catch (Exception ex) {
			fail();
		}
	}

	private static String createWorksheet(String worksheetName, Map<String, Object> data) throws Exception {
		String name = WorkbookTest.getRandomName();
		Workbook workbook = new Workbook(worksheetName);
		for (Map.Entry<String, Object> cell : data.entrySet()) {
			workbook.getCurrentWorksheet().addCell(cell.getValue(), cell.getKey());
		}
		workbook.saveAs(name);
		return name;
	}

	private static Map<String, Object> createSampleData() {
		Map<String, Object> data = new HashMap<String, Object>();
		data.put("A1",
				"test");
		data.put("A2", 22);
		data.put("A3", 11.1f);
		return data;
	}

}
