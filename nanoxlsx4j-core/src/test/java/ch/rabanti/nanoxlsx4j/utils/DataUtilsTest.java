/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.Duration;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.ValueSource;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

public class DataUtilsTest {

    private static final DateTimeFormatter DATE_TIME_FORMAT = DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm:ss");
    private static final DateTimeFormatter TIME_FORMAT = DateTimeFormatter.ofPattern("HH:mm:ss");

    @ParameterizedTest
    @DisplayName("Test of the getOADateTimestring function")
    @CsvSource(
            {
                    "01.01.1900 00:00:00, 1",
                    "02.01.1900 12:35:20, 2.52453703703704",
                    "27.02.1900 00:00:00, 58",
                    "28.02.1900 00:00:00, 59",
                    "28.02.1900 12:30:32, 59.5212037037037",
                    "01.03.1900 00:00:00, 61",
                    "01.03.1900 08:08:11, 61.3390162037037",
                    "20.05.1960 22:11:05, 22056.9243634259",
                    "01.01.2021 00:00:00, 44197",
                    "12.12.5870 11:30:12, 1450360.47930556"
            }
    )
    public void getOADateTimeStringTest(String dateString, String expectedOaDate) {
        Date date = parseDate(dateString);
        String oaDate = DataUtils.getOADateTimeString(date);
        float expected = Float.parseFloat(expectedOaDate);
        float given = Float.parseFloat(oaDate);
        float threshold = 0.00000001f; // Ignore everything below a millisecond (double precision may vary)
        assertTrue(Math.abs(expected - given) < threshold);
    }

    @ParameterizedTest
    @DisplayName("Test of the getOADateTime function")
    @CsvSource(
            {
                    "01.01.1900 00:00:00, 1",
                    "02.01.1900 12:35:20, 2.5245370370370401",
                    "27.02.1900 00:00:00, 58",
                    "28.02.1900 00:00:00, 59",
                    "28.02.1900 12:30:32, 59.521203703703705",
                    "01.03.1900 00:00:00, 61",
                    "01.03.1900 08:08:11, 61.339016203703707",
                    "20.05.1960 22:11:05, 22056.924363425926",
                    "01.01.2021 00:00:00, 44197",
                    "12.12.5870 11:30:12, 1450360.47930556"
            }
    )
    public void getOADateTimeString(String dateString, double expectedOaDate) {
        Date date = parseDate(dateString);
        double oaDate = DataUtils.getOADateTime(date);
        float threshold = 0.00000001f; // Ignore everything below a millisecond
        assertTrue(Math.abs(expectedOaDate - oaDate) < threshold);
    }

    @ParameterizedTest
    @DisplayName("Test of the successful getOADateTime function on invalid dates when checks are disabled")
    @ValueSource(
            strings = {
                    "02.01.0001 00:00:00",
                    "18.05.0712 11:15:02",
                    "31.12.1899 23:59:59"
            }
    )
    public void getOADateTimeTest2(String dateString) {
        Date date = parseDate(dateString);
        double given = DataUtils.getOADateTime(date, true);
        assertNotEquals(0d, given);
    }

    @ParameterizedTest
    @DisplayName("Test of the failing getOADateTimeString function on invalid dates")
    @ValueSource(
            strings = {
                    "01.01.0001 00:00:00",
                    "18.05.0712 11:15:02",
                    "31.12.1899 23:59:59"
            }
    )
    public void getOADateTimeStringFailTest(String dateString) {
        Date date = parseDate(dateString);
        assertThrows(FormatException.class, () -> DataUtils.getOADateTimeString(date));
    }

    @ParameterizedTest
    @DisplayName("Test of the failing getOADateTime function on invalid dates")
    @ValueSource(
            strings = {
                    "01.01.0001 00:00:00",
                    "18.05.0712 11:15:02",
                    "31.12.1899 23:59:59"
            }
    )
    public void getOADateTimeFailTest(String dateString) {
        Date date = parseDate(dateString);
        assertThrows(FormatException.class, () -> DataUtils.getOADateTime(date));
    }

    @ParameterizedTest
    @DisplayName("Test of the getOATimeString function")
    @CsvSource(
            {
                    "00:00:00, 0.0",
                    "12:00:00, 0.5",
                    "23:59:59, 0.999988425925926",
                    "13:11:10, 0.549421296296296",
                    "18:00:00, 0.75"
            }
    )
    public void getOATimeStringTest(String timeString, String expectedOaTime) {
        Duration time = Duration.between(LocalTime.MIDNIGHT, LocalTime.parse(timeString, TIME_FORMAT));
        String oaDate = DataUtils.getOATimeString(time);
        float expected = Float.parseFloat(expectedOaTime);
        float given = Float.parseFloat(oaDate);
        float threshold = 0.000000001f; // Ignore everything below a millisecond
        assertTrue(Math.abs(expected - given) < threshold);
    }

    @ParameterizedTest
    @DisplayName("Test of the getOATime function")
    @CsvSource(
            {
                    "00:00:00, 0.0",
                    "12:00:00, 0.5",
                    "23:59:59, 0.999988425925926",
                    "13:11:10, 0.549421296296296",
                    "18:00:00, 0.75"
            }
    )
    public void getOATimeTest(String timeString, double expectedOaTime) {
        Duration time = Duration.between(LocalTime.MIDNIGHT, LocalTime.parse(timeString, TIME_FORMAT));
        double oaTime = DataUtils.getOATime(time);
        float threshold = 0.000000001f; // Ignore everything below a millisecond
        assertTrue(Math.abs(expectedOaTime - oaTime) < threshold);
    }

    @ParameterizedTest
    @DisplayName("Test of the getDateFromOA function")
    @CsvSource(
            {
                    "1, 01.01.1900 00:00:00",
                    "2.5245370370370401, 02.01.1900 12:35:20",
                    "58, 27.02.1900 00:00:00",
                    "59, 28.02.1900 00:00:00",
                    "59.521203703703705, 28.02.1900 12:30:32",
                    "61, 01.03.1900 00:00:00",
                    "61.339016203703707, 01.03.1900 08:08:11",
                    "22056.924363425926, 20.05.1960 22:11:05",
                    "44197, 01.01.2021 00:00:00",
                    "1450360.47930556, 12.12.5870 11:30:12"
            }
    )
    public void getDateFromOATest(double givenValue, String expectedDateString) {
        Date expectedDate = parseDate(expectedDateString);
        Date date = DataUtils.getDateFromOA(givenValue);
        assertTrue(Math.abs(expectedDate.getTime() - date.getTime()) < 1d);
    }

    @ParameterizedTest
    @DisplayName("Test of the getInternalColumnWidth function")
    @CsvSource(
            {
                    "0.5, 0.85546875",
                    "1, 1.7109375",
                    "10, 10.7109375",
                    "15, 15.7109375",
                    "60, 60.7109375",
                    "254, 254.7109375",
                    "255, 255.7109375",
                    "0, 0"
            }
    )
    public void getInternalColumnWidthTest(float width, float expectedInternalWidth) {
        float internalWidth = DataUtils.getInternalColumnWidth(width);
        assertEquals(expectedInternalWidth, internalWidth);
    }

    @ParameterizedTest
    @DisplayName("Test of the failing getInternalColumnWidth function on invalid column widths")
    @ValueSource(
            floats = {
                    -0.1f,
                    -10f,
                    255.01f,
                    10000f}
    )
    public void getInternalColumnWidthFailTest(float width) {
        assertThrows(FormatException.class, () -> DataUtils.getInternalColumnWidth(width));
    }

    @ParameterizedTest
    @DisplayName("Test of the getInternalRowHeight function")
    @CsvSource(
            {
                    "0.1, 0",
                    "0.5, 0.75",
                    "1, 0.75",
                    "10, 9.75",
                    "15, 15",
                    "409, 408.75",
                    "409.5, 409.5",
                    "0, 0"
            }
    )
    public void getInternalRowHeightTest(float height, float expectedInternalHeight) {
        float internalHeight = DataUtils.getInternalRowHeight(height);
        assertEquals(expectedInternalHeight, internalHeight);
    }

    @ParameterizedTest
    @DisplayName("Test of the failing getInternalRowHeight function on invalid row heights")
    @ValueSource(
            floats = {
                    -0.1f,
                    -10f,
                    409.6f,
                    10000f}
    )
    public void getInternalRowHeightFailTest(float height) {
        assertThrows(FormatException.class, () -> DataUtils.getInternalRowHeight(height));
    }

    @ParameterizedTest
    @DisplayName("Test of the getInternalPaneSplitWidth function")
    @CsvSource(
            {
                    "0.1, 390",
                    "1, 390",
                    "18.5, 2415",
                    "32, 3825",
                    "255, 27240",
                    "256, 27345",
                    "1000, 105465",
                    "0, 390",
                    "-1, 390",
                    "-10, 390"
            }
    )
    public void getInternalPaneSplitWidthTest(float width, float expectedSplitWidth) {
        float splitWidth = DataUtils.getInternalPaneSplitWidth(width);
        assertEquals(expectedSplitWidth, splitWidth);
    }

    @ParameterizedTest
    @DisplayName("Test of the getInternalPaneSplitHeight function")
    @CsvSource(
            {
                    "0.1, 302",
                    "0.5, 310",
                    "1, 320",
                    "15, 600",
                    "409.5, 8490",
                    "500, 10300",
                    "0, 300",
                    "-1, 300",
                    "-10, 300"
            }
    )
    public void getInternalPaneSplitHeightTest(float height, float expectedSplitHeight) {
        float splitHeight = DataUtils.getInternalPaneSplitHeight(height);
        assertEquals(expectedSplitHeight, splitHeight);
    }

    @ParameterizedTest
    @DisplayName("Test of the getPaneSplitHeight function")
    @CsvSource(
            {
                    "301, 0.05",
                    "320, 1",
                    "600, 15",
                    "310, 0.5",
                    "8490, 409.5",
                    "10300, 500",
                    "300, 0",
                    "299.9, 0",
                    "-10, 0"
            }
    )
    public void getPaneSplitHeightTest(float height, float expectedSplitHeight) {
        float splitHeight = DataUtils.getPaneSplitHeight(height);
        assertEquals(expectedSplitHeight, splitHeight);
    }

    @ParameterizedTest
    @DisplayName("Test of the getPaneSplitWidth function")
    @CsvSource(
            {
                    "390, 0",
                    "2415, 18.5",
                    "1680, 11.5",
                    "3825, 31.9286",
                    "27240, 254.9286",
                    "27345, 255.9286",
                    "105465, 999.9286"
            }
    )
    public void getPaneSplitWidthTest(float width, float expectedSplitWidth) {
        float splitWidth = DataUtils.getPaneSplitWidth(width);
        float delta = Math.abs(splitWidth - expectedSplitWidth);
        assertTrue(delta < 0.001);
    }

    @Test
    @DisplayName("Test of the mergeRange overload")
    public void mergeRangeOverloadTest() {
        List<Range> givenRanges = List.of(new Range("B2:B4"));
        Range rangeToAdd = new Range("B5:B6");

        List<Range> resultRanges = DataUtils.mergeRange(givenRanges, rangeToAdd);

        assertEquals(1, resultRanges.size());
        assertTrue(resultRanges.stream().anyMatch(r -> r.toString().equals("B2:B6")));
    }

    @ParameterizedTest
    @DisplayName("Test of the mergeRange function")
    @CsvSource(
            delimiter = '|',
            value = {
                    "A2:A2 | A2:A2 | MERGE_COLUMNS | A2:A2",
                    "A2:A2 | A5:A6 | MERGE_COLUMNS | A2:A2,A5:A6",
                    "B2:B3 | B5:B6 | MERGE_COLUMNS | B2:B3,B5:B6",
                    "B2:B4 | B5:B6 | MERGE_COLUMNS | B2:B6",
                    "B2:C2 | D2:E2 | MERGE_COLUMNS | B2:E2",
                    "B5:B6 | B2:B4 | MERGE_COLUMNS | B2:B6",
                    "D2:E2 | B2:C2 | MERGE_COLUMNS | B2:E2",
                    "B2:C5 | C4:D6 | MERGE_COLUMNS | B2:B5,C2:C6,D4:D6",
                    "B2:C5,E2:F2 | C4:D6 | MERGE_COLUMNS | B2:B5,C2:C6,D4:D6,E2:F2",
                    "B2:C5,E3:F4 | C4:E6 | MERGE_COLUMNS | B2:B5,C2:C6,D4:D6,E3:E6,F3:F4",
                    "A2:A2 | A2:A2 | MERGE_ROWS | A2:A2",
                    "A2:A2 | A5:A6 | MERGE_ROWS | A2:A2,A5:A6",
                    "B2:B3 | B5:B6 | MERGE_ROWS | B2:B3,B5:B6",
                    "B2:B4 | B5:B6 | MERGE_ROWS | B2:B6",
                    "B2:C2 | D2:E2 | MERGE_ROWS | B2:E2",
                    "B5:B6 | B2:B4 | MERGE_ROWS | B2:B6",
                    "D2:E2 | B2:C2 | MERGE_ROWS | B2:E2",
                    "B2:C5 | C4:D6 | MERGE_ROWS | B2:C3,B4:D5,C6:D6",
                    "B2:C5,E2:F2 | C4:D6 | MERGE_ROWS | B2:C3,B4:D5,C6:D6,E2:F2",
                    "B2:C5,E3:F4 | C4:E6 | MERGE_ROWS | B2:C3,B4:F4,B5:E5,C6:E6,E3:F3",
                    "A2:A2 | A2:A2 | NO_MERGE | A2:A2",
                    "A2:A2 | A5:A6 | NO_MERGE | A2:A2,A5:A6",
                    "B2:B3 | B5:B6 | NO_MERGE | B2:B3,B5:B6",
                    "B2:B4 | B5:B6 | NO_MERGE | B2:B4,B5:B6",
                    "B2:C2 | D2:E2 | NO_MERGE | B2:C2,D2:E2",
                    "B5:B6 | B2:B4 | NO_MERGE | B2:B4,B5:B6",
                    "D2:E2 | B2:C2 | NO_MERGE | B2:C2,D2:E2",
                    "B2:C5 | C4:D6 | NO_MERGE | B2:B3,C2:C3,B4:B5,C4:C5,D4:D5,C6:C6,D6:D6"
            }
    )
    public void mergeRangeTest(
            String givenRangesString, String rangeToAddString,
            DataUtils.RangeMergeStrategy mergeStrategy, String expectedRangesString
    ) {
        List<Range> givenRanges = Arrays.stream(givenRangesString.split(","))
                .map(Range::new)
                .toList();

        Range rangeToAdd = new Range(rangeToAddString);

        List<Range> expectedRanges = Arrays.stream(expectedRangesString.split(","))
                .map(Range::new)
                .toList();

        List<Range> resultRanges = DataUtils.mergeRange(givenRanges, rangeToAdd, mergeStrategy);

        assertEquals(resultRanges.size(), expectedRanges.size());
        for (Range range : expectedRanges) {
            assertTrue(resultRanges.stream().anyMatch(r -> r.toString().equals(range.toString())));
        }
    }

    @Test
    @DisplayName("Test of the subtractRange overload")
    public void subtractRangeOverloadTest() {
        List<Range> givenRanges = List.of(new Range("B2:B5"));
        Range rangeToRemove = new Range("B4:B6");

        List<Range> resultRanges = DataUtils.subtractRange(givenRanges, rangeToRemove);

        assertEquals(1, resultRanges.size());
        assertTrue(resultRanges.stream().anyMatch(r -> r.toString().equals("B2:B3")));
    }

    @ParameterizedTest
    @DisplayName("Test of the subtractRange function")
    @CsvSource(
            delimiter = '|',
            value = {
                    "B5:C6 | A2:B3 | MERGE_COLUMNS | B5:C6",
                    "A2:A2 | A2:A2 | MERGE_COLUMNS | ''",
                    "B3:D5 | A2:E6 | MERGE_COLUMNS | ''",
                    "A2:A2 | A5:A6 | MERGE_COLUMNS | A2:A2",
                    "B2:B5 | B4:B6 | MERGE_COLUMNS | B2:B3",
                    "B4:B7 | B2:B5 | MERGE_COLUMNS | B6:B7",
                    "B2:B7 | A3:C4 | MERGE_COLUMNS | B2:B2,B5:B7",
                    "B3:D5 | A4:E4 | MERGE_COLUMNS | B3:D3,B5:D5",
                    "B3:D5 | C2:C6 | MERGE_COLUMNS | B3:B5,D3:D5",
                    "B3:D5 | A1:B3 | MERGE_COLUMNS | B4:B5,C3:D5",
                    "B3:D5 | A5:B6 | MERGE_COLUMNS | B3:B4,C3:D5",
                    "B5:C6,E2:F4 | E4:F4 | MERGE_COLUMNS | B5:C6,E2:F3",
                    "B3:C8,D3:E5 | C5:D6 | MERGE_COLUMNS | B3:B8,C3:D4,E3:E5,C7:C8",
                    "B3:C8,D3:F5,E7:F7 | C5:E7 | MERGE_COLUMNS | B3:B8,C3:E4,C8:C8,F3:F5,F7:F7",
                    "B5:C6 | A2:B3 | MERGE_ROWS | B5:C6",
                    "A2:A2 | A2:A2 | MERGE_ROWS | ''",
                    "B3:D5 | A2:E6 | MERGE_ROWS | ''",
                    "A2:A2 | A5:A6 | MERGE_ROWS | A2:A2",
                    "B2:B5 | B4:B6 | MERGE_ROWS | B2:B3",
                    "B4:B7 | B2:B5 | MERGE_ROWS | B6:B7",
                    "B2:B7 | A3:C4 | MERGE_ROWS | B2:B2,B5:B7",
                    "B3:D5 | A4:E4 | MERGE_ROWS | B3:D3,B5:D5",
                    "B3:D5 | C2:C6 | MERGE_ROWS | B3:B5,D3:D5",
                    "B3:D5 | A1:B3 | MERGE_ROWS | C3:D3,B4:D5",
                    "B3:D5 | A5:B6 | MERGE_ROWS | B3:D4,C5:D5",
                    "B5:C6,E2:F4 | E4:F4 | MERGE_ROWS | B5:C6,E2:F3",
                    "B3:C8,D3:E5 | C5:D6 | MERGE_ROWS | B3:E4,E5:E5,B5:B6,B7:C8",
                    "B3:C8,D3:F5,E7:F7 | C5:E7 | MERGE_ROWS | B3:F4,F5:F5,B5:B7,B8:C8,F7:F7",
                    "B5:C6 | A2:B3 | NO_MERGE | B5:C6",
                    "A2:A2 | A2:A2 | NO_MERGE | ''",
                    "B3:D5 | A2:E6 | NO_MERGE | ''",
                    "A2:A2 | A5:A6 | NO_MERGE | A2:A2",
                    "B2:B5 | B4:B6 | NO_MERGE | B2:B3",
                    "B4:B7 | B2:B5 | NO_MERGE | B6:B7",
                    "B2:B7 | A3:C4 | NO_MERGE | B2:B2,B5:B7",
                    "B3:D5 | A4:E4 | NO_MERGE | B3:D3,B5:D5",
                    "B3:D5 | C2:C6 | NO_MERGE | B3:B5,D3:D5",
                    "B3:D5 | A1:B3 | NO_MERGE | C3:D3,B4:B5,C4:D5",
                    "B3:D5 | A5:B6 | NO_MERGE | B3:B4,C3:D4,C5:D5",
                    "B5:C6,E2:F4 | E4:F4 | NO_MERGE | B5:C6,E2:F3",
                    "B3:C8,D3:E5 | C5:D6 | NO_MERGE | B3:B4,C3:C4,D3:D4,E3:E4,B5:B5,B6:B6,E5:E5,B7:B8,C7:C8",
                    "B3:C8,D3:F5,E7:F7 | C5:E7 | NO_MERGE | B3:B4,C3:C4,D3:E4,F3:F4,B5:B5,F5:F5,B6:B6,B7:B7,F7:F7," +
                            "B8:B8,C8:C8"
            }
    )
    public void subtractRangeTest(
            String givenRangesString, String rangeToRemoveString,
            DataUtils.RangeMergeStrategy mergeStrategy, String expectedRangesString
    ) {
        List<Range> givenRanges = Arrays.stream(givenRangesString.split(","))
                .map(Range::new)
                .toList();

        Range rangeToRemove = new Range(rangeToRemoveString);

        List<Range> expectedRanges = expectedRangesString.isEmpty()
                ? List.of()
                : Arrays.stream(expectedRangesString.split(","))
                .map(Range::new)
                .toList();

        List<Range> resultRanges = DataUtils.subtractRange(givenRanges, rangeToRemove, mergeStrategy);

        assertEquals(resultRanges.size(), expectedRanges.size());
        for (Range range : expectedRanges) {
            assertTrue(resultRanges.stream().anyMatch(r -> r.toString().equals(range.toString())));
        }
    }

    private static Date parseDate(String dateString) {
        LocalDateTime date = LocalDateTime.parse(dateString, DATE_TIME_FORMAT);
        return Date.from(date.toInstant(ZoneOffset.UTC));
    }
}
