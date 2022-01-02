package ch.rabanti.nanoxlsx4j.misc;

import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.*;

public class HelperTest {

    @DisplayName("Test of the getOADateString function")
    @ParameterizedTest(name = "Given date {0} should lead to the OADate number {1}")
    @CsvSource({
            "01.01.1900 00:00:00, 1",
            "02.01.1900 12:35:20, 2.5245370370370401",
            "27.02.1900 00:00:00, 58",
            "28.02.1900 00:00:00, 59",
            "28.02.1900 12:30:32, 59.521203703703705",
            "01.03.1900 00:00:00, 61",
            "01.03.1900 08:08:11, 61.339016203703707",
            "20.05.1960 22:11:05, 22056.924363425926",
            "01.01.2021 00:00:00, 44197",
            "12.12.5870 11:30:12, 1450360.47930556",
    })
    void getOADateStringTest(String dateString, String expectedOaDate) throws ParseException {
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss", Locale.US);
        Date date = formatter.parse(dateString);
        String oaDate = Helper.getOADateString(date);
        float expected = Float.parseFloat(expectedOaDate);
        float given = Float.parseFloat(oaDate);
        float threshold = 0.000000001f; // Ignore everything below a millisecond
        assertTrue(Math.abs(expected - given) < threshold);
    }

    @DisplayName("Test of the getOADate function")
    @ParameterizedTest(name = "Given date {0} should lead to the OADate number {1}")
    @CsvSource({
            "01.01.1900 00:00:00, 1",
            "02.01.1900 12:35:20, 2.5245370370370401",
            "27.02.1900 00:00:00, 58",
            "28.02.1900 00:00:00, 59",
            "28.02.1900 12:30:32, 59.521203703703705",
            "01.03.1900 00:00:00, 61",
            "01.03.1900 08:08:11, 61.339016203703707",
            "20.05.1960 22:11:05, 22056.924363425926",
            "01.01.2021 00:00:00, 44197",
            "12.12.5870 11:30:12, 1450360.47930556",
    })
    void getOADateTest(String dateString, double expectedOaDate) throws ParseException {
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss", Locale.US);
        Date date = formatter.parse(dateString);
        double oaDate = Helper.getOADate(date);
        float threshold = 0.00000001f; // Ignore everything below a millisecond (double precision may vary)
        assertTrue(Math.abs(expectedOaDate - oaDate) < threshold);
    }

    @DisplayName("Test of the successful getOADate function on invalid dates when checks are disabled")
    @ParameterizedTest(name = "Given date {0} should not lead to an exception")
    @CsvSource({
            "01.01.0001 00:00:00",
            "18.05.0712 11:15:02",
            "31.12.1899 23:59:59",
            "01.01.10000 00:00:00",
    })
    void getOADateTest2(String dateString) throws ParseException {
        // Note: Dates beyond the year 10000 may not be tested wit other frameworks
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss", Locale.US);
        Date date = formatter.parse(dateString);
        double given = Helper.getOADate(date, true);
        assertNotEquals(0d, given);
    }

    @DisplayName("Test of the failing getOADateString function on invalid dates")
    @ParameterizedTest(name = "Given date {0} should lead to an exception")
    @CsvSource({
            "01.01.0001 00:00:00",
            "18.05.0712 11:15:02",
            "31.12.1899 23:59:59",
            "01.01.10000 00:00:00",
    })
    void getOADateStringFailTest(String dateString) throws ParseException {
        // Note: Dates beyond the year 10000 may not be tested wit other frameworks
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss", Locale.US);
        Date date = formatter.parse(dateString);
        assertThrows(FormatException.class, () -> Helper.getOADateString(date));
    }

    @DisplayName("Test of the failing getOADate function on invalid dates")
    @ParameterizedTest(name = "Given date {0} should lead to an exception")
    @CsvSource({
            "01.01.0001 00:00:00",
            "18.05.0712 11:15:02",
            "31.12.1899 23:59:59",
            "01.01.10000 00:00:00",
    })
    void getOADateFailTest(String dateString) throws ParseException {
        // Note: Dates beyond the year 10000 may not be tested wit other frameworks
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss", Locale.US);
        Date date = formatter.parse(dateString);
        assertThrows(FormatException.class, () -> Helper.getOADate(date));
    }

    @DisplayName("Test of the failing getOADateString function on a null Date")
    @Test()
    void getOADateStringFailTest2() {
        assertThrows(FormatException.class, () -> Helper.getOADateString(null));
    }

    @DisplayName("Test of the failing getOADate function on a null Date")
    @Test()
    void getOADateFailTest() {
        assertThrows(FormatException.class, () -> Helper.getOADate(null));
    }

    @DisplayName("Test of the getOATimeString function")
    @ParameterizedTest(name = "Given value {0} should lead to the OaDate {1}")
    @CsvSource({
            "00:00:00, 0.0",
            "12:00:00, 0.5",
            "23:59:59,  0.999988425925926",
            "13:11:10, 0.549421296296296",
            "18:00:00, 0.75",
    })
    void getOATimeStringTest(String timeString, String expectedOaTime) throws ParseException {
        Duration time = Helper.parseTime(timeString, "HH:mm:ss", Locale.US);
        String oaDate = Helper.getOATimeString(time);
        float expected = Float.parseFloat(expectedOaTime);
        float given = Float.parseFloat(oaDate);
        float threshold = 0.000000001f; // Ignore everything below a millisecond
        assertTrue(Math.abs(expected - given) < threshold);
    }

    @DisplayName("Test of the getOATime function")
    @ParameterizedTest(name = "Given value {0} should lead to the OaDate {1}")
    @CsvSource({
            "00:00:00, 0.0",
            "12:00:00, 0.5",
            "23:59:59,  0.999988425925926",
            "13:11:10, 0.549421296296296",
            "18:00:00, 0.75",
    })
    void getOATimeTest(String timeString, double expectedOaTime) throws ParseException {
        Duration time = Helper.parseTime(timeString, "HH:mm:ss", Locale.US);
        double oaTime = Helper.getOATime(time);
        float threshold = 0.000000001f; // Ignore everything below a millisecond
        assertTrue(Math.abs(expectedOaTime - oaTime) < threshold);
    }

    @DisplayName("Test of the failing getOATimeString function on a null LocalTime")
    @Test()
    void getOATimeStringFailTest2() {
        assertThrows(FormatException.class, () -> Helper.getOATimeString(null));
    }

    @DisplayName("Test of the failing getOATime function on a null LocalTime")
    @Test()
    void getOATimeFailTest() {
        assertThrows(FormatException.class, () -> Helper.getOATime(null));
    }


    @DisplayName("Test of the getInternalColumnWidth function")
    @ParameterizedTest(name = "Given value {0} should lead to the internal width {1}")
    @CsvSource({
            "0.5, 0.85546875",
            "1, 1.7109375",
            "10, 10.7109375",
            "15, 15.7109375",
            "60, 60.7109375",
            "254, 254.7109375",
            "255, 255.7109375",
            "0, 0f",
    })
    void getInternalColumnWidthTest(float width, float expectedInternalWidth) {
        float internalWidth = Helper.getInternalColumnWidth(width);
        assertEquals(expectedInternalWidth, internalWidth);
    }

    @DisplayName("Test of the failing GetInternalColumnWidth function on invalid column widths")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "-0.1",
            "-10",
            "255.01",
            "10000",
    })
    void getInternalColumnWidthFailTest(float width) {
        assertThrows(FormatException.class, () -> Helper.getInternalColumnWidth(width));
    }

    @DisplayName("Test of the getInternalRowHeight function")
    @ParameterizedTest(name = "Given value {0} should lead to the internal row height {1}")
    @CsvSource({
            "0.1, 0f",
            "0.5, 0.75",
            "1, 0.75",
            "10, 9.75",
            "15, 15",
            "409, 408.75",
            "409.5, 409.5",
            "0, 0f",
    })
    void getInternalRowHeightTest(float height, float expectedInternalHeight) {
        float internalHeight = Helper.getInternalRowHeight(height);
        assertEquals(expectedInternalHeight, internalHeight);
    }

    @DisplayName("Test of the failing getInternalRowHeight function on invalid row heights")
    @ParameterizedTest(name = "Given value {0} should lead to n exception")
    @CsvSource({
            "-0.1",
            "-10",
            "409.6",
            "10000",
    })
    void getInternalRowHeightFailTest(float height) {
        assertThrows(FormatException.class, () -> Helper.getInternalRowHeight(height));
    }

    @DisplayName("Test of the getInternalPaneSplitWidth function")
    @ParameterizedTest(name = "Given value {0} should lead to width {1}")
    @CsvSource({
            "0.1f, 390f",
            "1f, 390f",
            "18.5, 2415f",
            "32f, 3825f",
            "255, 27240f",
            "256, 27345f",
            "1000, 105465f",
            "0, 390f",
            "-1, 390f",
            "-10, 390f",
    })
    void getInternalPaneSplitWidthTest(float width, float expectedSplitWidth) {
        float splitWidth = Helper.getInternalPaneSplitWidth(width);
        assertEquals(expectedSplitWidth, splitWidth);
    }

    @DisplayName("Test of the getInternalPaneSplitHeight function")
    @ParameterizedTest(name = "Given value {0} should lead to the height {1}")
    @CsvSource({
            "0.1f, 302f",
            "0.5f, 310f",
            "1f, 320f",
            "15f, 600f",
            "409.5, 8490f",
            "500, 10300f",
            "0, 300f",
            "-1, 300f",
            "-10, 300f",
    })
    void getInternalPaneSplitHeightTest(float height, float expectedSplitHeight) {
        float splitHeight = Helper.getInternalPaneSplitHeight(height);
        assertEquals(expectedSplitHeight, splitHeight);
    }

    @DisplayName("Test of the getDateFromOA function")
    @ParameterizedTest(name = "Given value {0} should lead to the date {1}")
    @CsvSource({
            "1, '01.01.1900 00:00:00'",
            "2.5245370370370401, '02.01.1900 12:35:20'",
            "58, '27.02.1900 00:00:00'",
            "59, '28.02.1900 00:00:00'",
            "59.521203703703705, '28.02.1900 12:30:32'",
            "61, '01.03.1900 00:00:00'",
            "61.339016203703707, '01.03.1900 08:08:11'",
            "22056.924363425926, '20.05.1960 22:11:05'",
            "44197, '01.01.2021 00:00:00'",
            "1450360.47930556, '12.12.5870 11:30:12'",
    })
    void getDateFromOATest(double givenValue, String expectedDateString) throws ParseException {
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss", Locale.US);
        Date expectedDate = formatter.parse(expectedDateString);
        Date date = Helper.getDateFromOA(givenValue);
        assertEquals(expectedDate, date);
    }

    @DisplayName("Test of the parseTime function")
    @ParameterizedTest(name = "Given value {0}  and pattern {1} should lead to the duration {2}")
    @CsvSource({
            "'12:13:14', 'HH:mm:ss', 'en-US', 'PT12H13M14S'",
            "'00-00-00', 'HH-mm-ss', 'CA-fr', 'PT0S'",
            "'59/59/23', 'mm/ss/HH', 'en-US', 'PT23H59M59S'",
            "'10:11', 'HH:mm', 'de-DE', 'PT10H11M'",
            "'09:10:12 PM', 'hh:mm:ss a', 'en-US', 'PT21H10M12S'",
            "'09:10:12 AM', 'hh:mm:ss a', 'en-US', 'PT9H10M12S'",
            "'09:10:12p.m.', 'hh:mm:ssa', 'ca-FR', 'PT21H10M12S'",
            "'01.01.2022', 'dd.MM.yyyy', 'en-US', 'PT0S'", // Not properly formatted to get the number of days
            "'01.01.2022 22-11-05', 'dd.MM.yyyy HH-mm-ss', 'ca-FR', 'PT22H11M5S'", // Not properly formatted to get the number of days
            "'13 09:10:12 PM', 'n hh:mm:ss a', 'en-US', 'PT333H10M12S'", // Adds 13x24h
            "'00-00-00 10', 'HH-mm-ss n', 'CA-fr', 'PT240H'", // Adds 10x24h
    })
    void parseTimeTest(String givenValue, String givenPattern, String localeString, String expectedDuration) {

        String[] localeParts = localeString.split("-");
        Locale locale = new Locale(localeParts[1], localeParts[0]);
        Duration time = Helper.parseTime(givenValue, givenPattern, locale);
        assertEquals(expectedDuration, time.toString());
    }

    @DisplayName("Test of the failing parseTime function on a missing formatter")
    @Test()
    void parseTimeFailTest() {
        assertThrows(FormatException.class, () -> Helper.parseTime("12.12.2021", null));
    }

    @DisplayName("Test of the failing parseTime function on missing or invalid values")
    @ParameterizedTest(name = "Given time value {1} of type {0} should lead to an exception")
    @CsvSource({
            "NULL, '', 'HH:mm:ss'",
            "STRING, '?', 'HH:mm:ss'",
            "STRING, '', 'HH:mm:ss'",
            "STRING, '12-10-09', 'HH:mm:ss'",
            "STRING, '12-10-09 1', 'HH-mm-ss'",
            "STRING, '12-10-09', 'n HH-mm-ss'",
            "STRING, '-1 12-10-09', 'n HH-mm-ss'",
            "STRING, 'x 12-10-09', 'n HH-mm-ss'",
            "STRING, '10 zz-10-09', 'n HH-mm-ss'",
    })
    void parseTimeFailTest2(String sourceType, String sourceValue, String givenPattern) {
        String value = (String) TestUtils.createInstance(sourceType, sourceValue);
        assertThrows(FormatException.class, () -> Helper.parseTime(value, givenPattern, Locale.US));
    }

    @DisplayName("Test of the createDuration function without days")
    @ParameterizedTest(name = "Given hours {1}, minutes {2} and seconds {3} should lead to the duration {4}")
    @CsvSource({
            "0, 0, 0, 'PT0S'",
            "12, 10, 9, 'PT12H10M9S'",
            "23, 59, 59, 'PT23H59M59S'",
    })
    void createDurationTest(int hours, int minutes, int seconds, String expectedDuration){
        Duration time = Helper.createDuration(hours,minutes,seconds);
        assertEquals(expectedDuration, time.toString());
    }

    @DisplayName("Test of the createDuration  with days")
    @ParameterizedTest(name = "Given days {0}, hours {1}, minutes {2} and seconds {3} should lead to the duration {4}")
    @CsvSource({
            "0, 0, 0, 0, 'PT0S'",
            "0, 12, 10, 9, 'PT12H10M9S'",
            "0, 23, 59, 59, 'PT23H59M59S'",
            "1, 0, 0, 0, 'PT24H'",
            "2, 3, 4, 5, 'PT51H4M5S'",
    })
    void createDurationTest2(int days, int hours, int minutes, int seconds, String expectedDuration){
        Duration time = Helper.createDuration(days, hours,minutes,seconds);
        assertEquals(expectedDuration, time.toString());
    }

    @DisplayName("Test of the failing createDuration function on invalid values")
    @ParameterizedTest(name = "Given hours {0}, minutes {1} and seconds {2} should lead to an exception")
    @CsvSource({
            "-1, 0, 0",
            "10, -1, 0",
            "1, 0, -5",
            "25, 1, 0",
            "10, 62, 30",
            "0, 12, 120",
    })
    void createDurationFailTest(int hours, int minutes, int seconds) {
        assertThrows(FormatException.class, () -> Helper.createDuration(hours,minutes,seconds));
    }

    @DisplayName("Test of the failing createDuration function on invalid values with days")
    @ParameterizedTest(name = "Given days {0}, hours {1}, minute {2} and second {3} should lead to an exception")
    @CsvSource({
            "-1, 0, 0, 0",
            "-10, 10, -1, 0",
            "5, 1, 0, -5",
            "10, 25, 1, 0",
            "0, 10, 62, 30",
            "-1, 0, 12, 120",
            "2958466, 0, 0, 0",
            "2958465, 10, -1, 0",
            "2958465, 1, 0, -5",
            "2958465, 25, 1, 0",
            "2958465, 10, 62, 30",
            "2958465, 0, 12, 120",
            "2958466, 1 , 1, 1",
    })
    void createDurationFailTest2(int days, int hours, int minutes, int seconds) {
        assertThrows(FormatException.class, () -> Helper.createDuration(days,hours,minutes,seconds));
    }

}
