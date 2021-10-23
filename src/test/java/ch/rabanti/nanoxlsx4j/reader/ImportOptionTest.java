package ch.rabanti.nanoxlsx4j.reader;

import ch.rabanti.nanoxlsx4j.*;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.LocalTime;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.function.BiConsumer;

import static org.junit.jupiter.api.Assertions.*;

public class ImportOptionTest {

    @DisplayName("Test of the reader functionality with the global import option to cast everything to string")
    @Test()
    void castAllToStringTest() throws Exception {
        Map<String, Object> cells = new HashMap<String, Object>();
        cells.put("A1", "test");
        cells.put("A2", true);
        cells.put("A3", false);
        cells.put("A4", 42);
        cells.put("A5", 0.55f);
        cells.put("A6", -0.111d);
        cells.put("A7",getDate(2020, 10, 10, 9, 8, 7)); // month -1
        cells.put("A8", LocalTime.of(18,15,12));
        cells.put("A9", null);
        Map<String, String> expectedCells = new HashMap<String, String>();
        expectedCells.put("A1", "test");
        expectedCells.put("A2", "true");
        expectedCells.put("A3", "false");
        expectedCells.put("A4", "42");
        expectedCells.put("A5", "0.55");
        expectedCells.put("A6", "-0.111");
        expectedCells.put("A7", "2020-11-10 09:08:07");
        expectedCells.put("A8", "18:15:12");
        expectedCells.put("A9", null);
        ImportOptions options = new ImportOptions();
        options.setGlobalEnforcingType(ImportOptions.GlobalType.EverythingToString);
        assertValues(cells, options, ImportOptionTest::assertEqualsFunction, expectedCells);
    }


    @DisplayName("Test of the reader functionality with the global import option to cast all number to double")
    @Test()
    void castToDoubleTest() throws Exception {
        Map<String, Object> cells = new HashMap<String, Object>();
        cells.put("A1", "test");
        cells.put("A2", true);
        cells.put("A3", false);
        cells.put("A4", 42);
        cells.put("A5", 0.55f);
        cells.put("A6", -0.111d);
        cells.put("A7", getDate(2020, 11, 10, 9, 8, 7));
        cells.put("A8", LocalTime.of(18,15,12));
        cells.put("A9", null);
        cells.put("A10", "27");
        Map<String, Object> expectedCells = new HashMap<>();
        expectedCells.put("A1", "test");
        expectedCells.put("A2", 1d);
        expectedCells.put("A3", 0d);
        expectedCells.put("A4", 42d);
        expectedCells.put("A5", 0.55d);
        expectedCells.put("A6", -0.111d);
        expectedCells.put("A7", Double.valueOf(Helper.getOADateString(getDate(2020,11,10,9,8,7))));
        expectedCells.put("A8", Double.valueOf(Helper.getOATimeString(LocalTime.of(18,15,12))));
        expectedCells.put("A9", null);
        expectedCells.put("A10", 27d);
        ImportOptions options = new ImportOptions();
        options.setGlobalEnforcingType(ImportOptions.GlobalType.AllNumbersToDouble);
        assertValues(cells, options,ImportOptionTest::assertApproximateFunction, expectedCells);
    }


    @DisplayName("Test of the reader functionality with the global import option to cast all number to int")
    @Test()
    void castToIntTest() throws Exception {
        Map<String, Object> cells = new HashMap<>();
        cells.put("A1", "test");
        cells.put("A2", true);
        cells.put("A3", false);
        cells.put("A4", 42);
        cells.put("A5", 0.55f);
        cells.put("A6", -3.111d);
        cells.put("A7", getDate(2020, 11, 10, 9, 8, 7));
        cells.put("A8", LocalTime.of(18,15,12));
        cells.put("A9", -4.9f);
        cells.put("A10", 0.49d);
        cells.put("A11", null);
        cells.put("A12", "28");
        Map<String, Object> expectedCells = new HashMap<String, Object>();
        expectedCells.put("A1", "test");
        expectedCells.put("A2", 1);
        expectedCells.put("A3", 0);
        expectedCells.put("A4", 42);
        expectedCells.put("A5", 1);
        expectedCells.put("A6", -3);
        expectedCells.put("A7", (int)Math.round(Double.parseDouble(Helper.getOADateString(getDate(2020, 11, 10, 9, 8, 7)))));
        expectedCells.put("A8", (int)Math.round(Double.parseDouble(Helper.getOATimeString(LocalTime.of(18,15,12)))));
        expectedCells.put("A9", -5);
        expectedCells.put("A10", 0);
        expectedCells.put("A11", null);
        expectedCells.put("A12", 28);
        ImportOptions options = new ImportOptions();
        options.setGlobalEnforcingType(ImportOptions.GlobalType.AllNumbersToInt);
        assertValues(cells, options, ImportOptionTest::assertApproximateFunction, expectedCells);
    }

    @DisplayName("Test of the enforcingStartRowNumber functionality on global enforcing rules")
    @Test()
    void enforcingStartRowNumberTest() throws Exception {
        Map<String, Object> cells = new HashMap<String, Object>();
        cells.put("A1", 22);
        cells.put("A2", true);
        cells.put("A3", 22);
        cells.put("A4", true);
        cells.put("A5", 22.5d);
        Map<String, Object> expectedCells = new HashMap<String, Object>();
        expectedCells.put("A1", 22);
        expectedCells.put("A2", true);
        expectedCells.put("A3", "22");
        expectedCells.put("A4", "true");
        expectedCells.put("A5", "22.5");
        ImportOptions options = new ImportOptions();
        options.setEnforcingStartRowNumber(2);
        options.setGlobalEnforcingType(ImportOptions.GlobalType.EverythingToString);
        assertValues(cells, options, ImportOptionTest::assertApproximateFunction, expectedCells);
    }

    @DisplayName("Test of the import options for the import column type: Double")
    @Test()
    void enforcingColumnAsNumberTest() throws Exception {
        LocalTime time = LocalTime.of(11, 12, 13);
        Date date = getDate(2021, 8, 14, 18, 22, 13);
        Map<String, Object> cells = new HashMap<>();
        cells.put("A1", 22);
        cells.put("A2", "21");
        cells.put("A3", true);
        cells.put("B1", 23);
        cells.put("B2", "20");
        cells.put("B3", true);
        cells.put("B4", time);
        cells.put("B5", date);
        cells.put("C1", "2");
        cells.put("C2", LocalTime.of(12, 14, 16));
        Map<String, Object> expectedCells = new HashMap<>();
        expectedCells.put("A1", 22);
        expectedCells.put("A2", "21");
        expectedCells.put("A3", true);
        expectedCells.put("B1", 23d);
        expectedCells.put("B2", 20d);
        expectedCells.put("B3", 1d);
        expectedCells.put("B4", Double.parseDouble(Helper.getOATimeString(time)));
        expectedCells.put("B5", Double.parseDouble(Helper.getOADateString(date)));
        expectedCells.put("C1", "2");
        expectedCells.put("C2", LocalTime.of(12, 14, 16));
        ImportOptions options = new ImportOptions();
        options.addEnforcedColumn(1, ImportOptions.ColumnType.Double);
        assertValues(cells, options, ImportOptionTest::assertApproximateFunction, expectedCells);
    }



    private static Date getDate(int year, int month, int day, int hour, int minute, int second){
        Calendar calendar = Calendar.getInstance();
        calendar.set(year, month, day, hour, minute, second);
        calendar.set(Calendar.MILLISECOND, 0);
       return calendar.getTime();
    }

    private static <T, D> void assertValues(Map<String, T> givenCells, ImportOptions importOptions, BiConsumer<Object, Object> assertionAction) throws Exception {
        assertValues(givenCells, importOptions, assertionAction, null);
    }

    private static <T, D> void assertValues(Map<String, T> givenCells, ImportOptions importOptions, BiConsumer<Object, Object> assertionAction, Map<String, D> expectedCells) throws Exception {
        Workbook workbook = new Workbook("worksheet1");
        for (Map.Entry<String, T> cell : givenCells.entrySet()) {
            workbook.getCurrentWorksheet().addCell(cell.getValue(), cell.getKey());
        }
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        workbook.saveAsStream(stream);
        ByteArrayInputStream inputStream = new ByteArrayInputStream(stream.toByteArray());
        stream.close();
        Workbook givenWorkbook = Workbook.load(inputStream, importOptions);
        inputStream.close();


        assertNotNull(givenWorkbook);
        Worksheet givenWorksheet = givenWorkbook.setCurrentWorksheet(0);
        assertEquals("worksheet1", givenWorksheet.getSheetName());
        for (String address : givenCells.keySet()) {
            Cell givenCell = givenWorksheet.getCell(new Address(address));
            D expectedValue = expectedCells.get(address);
            if (expectedValue == null) {
                assertEquals(Cell.CellType.EMPTY, givenCell.getDataType());
            } else {
                assertionAction.accept(expectedValue, (D) givenCell.getValue());
            }
        }
    }

    private static <T> void assertEqualsFunction(T expected, T given) {
        assertEquals(expected, given);
    }

    private static void assertApproximateFunction(Object expected, Object given) {
        double threshold = 0.000012; // The precision may vary (roughly one second)
        if (given instanceof Double) {
            assertTrue(Math.abs((Double) given - (Double) expected) < threshold);
        } else if (given instanceof Float) {
            assertTrue(Math.abs((Float) given - (Float) expected) < threshold);
        } else if (given instanceof Date) {

            double e = Double.parseDouble(Helper.getOADateString((Date)expected));
            double g = Double.parseDouble(Helper.getOADateString((Date)given));
            assertApproximateFunction(e, g);
        } else if (given instanceof LocalTime){
            double g = Double.parseDouble(Helper.getOATimeString((LocalTime)given));
            double e = Double.parseDouble(Helper.getOATimeString((LocalTime)expected));
            assertApproximateFunction(e, g);
        }
        else {
            assertEqualsFunction(expected, given);
        }

    }
}
