package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import ch.rabanti.nanoxlsx4j.styles.Style;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.NumberFormat;
import java.time.Duration;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class TestUtils {

    private static final Calendar CALENADR;

    static {
        CALENADR = new GregorianCalendar();
    }

    public static Object createInstance(String sourceType, String stringValue) {
        switch (sourceType.toUpperCase()) {
            case "BIGDECIMAL":
                return  new BigDecimal(stringValue);
            case "INTEGER":
                try {
                    Number number = NumberFormat.getInstance().parse(stringValue);
                    return number.intValue();
                }
                catch(Exception ex){
                    throw new IllegalArgumentException("Cannot cast to int: " + sourceType);
            }
            case "LONG":
                return Long.parseLong(stringValue);
            case "BYTE":
                return Byte.parseByte(stringValue);
            case "DOUBLE":
                return Double.parseDouble(stringValue);
            case "FLOAT":
                return Float.parseFloat(stringValue);
            case "BOOLEAN":
                return Boolean.parseBoolean(stringValue);
            case "STRING":
                return stringValue;
            case "NULL":
                return null;
            default:
                throw new IllegalArgumentException("Not implemented source type: " + sourceType);
        }
    }

    public static void assertCellRange(String expectedAddresses, List<Address> addresses) {
        String[] addressStrings = splitValues(expectedAddresses);
        List<Address> expected = new ArrayList<Address>();
        for (String address : addressStrings) {
            expected.add(new Address(address));
        }
        assertEquals(expected.size(), addresses.size());
        for (int i = 0; i < expected.size(); i++) {
            assertEquals(expected.get(i), addresses.get(i));
        }
    }

    public static String[] splitValues(String valueString) {
        List<String> list = splitValuesAsList(valueString);
        String[] output = new String[list.size()];
        list.toArray(output);
        return output;
    }

    public static List<String> splitValuesAsList(String valueString) {
        if (valueString == null || valueString.isEmpty()) {
            return new ArrayList<>();
        }
        String[] valueStrings = valueString.split(",|\\s");
        List<String> output = new ArrayList<>();
        for (String value : valueStrings) {
            if (value != null && !value.isEmpty()) {
                output.add(value);
            }
        }
        return output;
    }

    public static <T> void assertCellsEqual(T value1, T value2, T inequalValue, Address cellAddress) {
        Cell cell1 = new Cell(value1, Cell.CellType.DEFAULT, cellAddress);
        Cell cell2 = new Cell(value2, Cell.CellType.DEFAULT, cellAddress);
        Cell cell3 = new Cell(inequalValue, Cell.CellType.DEFAULT, cellAddress);
        assertTrue(cell1.equals(cell2));
        assertFalse(cell1.equals(cell3));
    }

    public static <K, V> void assertMapEntry(K expectedKey, V expectedValue, Map<K, V> map) {
        assertNotNull(map);
        assertNotEquals(0, map.size());
        assertTrue(map.containsKey(expectedKey));
        assertTrue(map.get(expectedKey).equals(expectedValue));
    }

    public static <K, MV, V> void assertMapEntry(K expectedKey, V expectedValue, Map<K, MV> map, Function<MV, V> method) {
        assertNotNull(map);
        assertNotEquals(0, map.size());
        assertTrue(map.containsKey(expectedKey));
        MV mapValue = map.get(expectedKey);
        V actualValue = method.apply(mapValue);
        if (actualValue == null) {
            assertNull(expectedValue);
        } else {
            assertTrue(actualValue.equals(expectedValue));
        }
    }

    public static <T> void assertListEntry(T expectedEntry, List<T> list) {
        assertNotNull(list);
        assertNotEquals(0, list.size());
        assertTrue(list.contains(expectedEntry));
    }

    public static Workbook saveAndLoadWorkbook(Workbook workbook, ImportOptions options) throws IOException, java.io.IOException {
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        workbook.saveAsStream(stream);
        ByteArrayInputStream inputStream = new ByteArrayInputStream(stream.toByteArray());
        stream.close();
        Workbook givenWorkbook = Workbook.load(inputStream, options);
        inputStream.close();
        return givenWorkbook;
    }

    public interface TriConsumer<T1, T2, T3> {
        void accept(T1 t1, T2 t2, T3 t3);

        default TriConsumer<T1, T2, T3> andThen(TriConsumer<? super T1, ? super T2, ? super T3> after) {
            Objects.requireNonNull(after);
            return (l, r, s) -> {
                accept(l, r, s);
                after.accept(l, r, s);
            };
        }
    }

    public static InputStream getResource(String resourceName) {
        return ClassLoader.getSystemResourceAsStream(resourceName);
    }

    public interface QuadConsumer<T1, T2, T3, T4> {
        void accept(T1 t1, T2 t2, T3 t3, T4 t4);

        default QuadConsumer<T1, T2, T3, T4> andThen(QuadConsumer<? super T1, ? super T2, ? super T3, ? super T4> after) {
            Objects.requireNonNull(after);
            return (l, r, s, t) -> {
                accept(l, r, s, t);
                after.accept(l, r, s, t);
            };
        }
    }

    public static Cell saveAndReadStyledCell(Object value, Style style, String targetCellAddress) {
        return saveAndReadStyledCell(value, value, style, targetCellAddress);
    }

    public static Cell saveAndReadStyledCell(Object givenValue, Object expectedValue, Style style, String targetCellAddress) {
        try {
            Workbook workbook = new Workbook(false);
            workbook.addWorksheet("sheet1");
            workbook.getCurrentWorksheet().addCell(givenValue, targetCellAddress, style);

            Workbook givenWorkbook = TestUtils.saveAndLoadWorkbook(workbook, null);

            Cell cell = givenWorkbook.getCurrentWorksheet().getCell(new Address(targetCellAddress));
            assertEquals(expectedValue, cell.getValue());
            return cell;
        } catch (Exception ex) {
            return null;
        }
    }

    public static Date buildDate(int year, int month, int day) {
        return buildDate(year, month, day, 0, 0, 0, 0);
    }

    public static Date buildDate(int year, int month, int day, int hour, int minute, int second) {
        return buildDate(year, month, day, hour, minute, second, 0);
    }

    public static Date buildDate(int year, int month, int day, int hour, int minute, int second, int millisSecond) {
        CALENADR.set(year, month, day, hour, minute, second);
        CALENADR.set(Calendar.MILLISECOND, millisSecond);
        return CALENADR.getTime();
    }

    public static Duration buildTimeWithDays(int days, int hour, int minute, int second) {
        Duration duration = buildTime(hour, minute, second, 0);
        duration = duration.plusDays(days);
        return duration;
    }

    public static Duration buildTime(int hour, int minute, int second) {
        return buildTime(hour, minute, second, 0);
    }

    public static Duration buildTime(int hour, int minute, int second, int milliSecond) {
        int millis = milliSecond + (1000 * second) + (minute * 60000) + (hour * 3600000);
        return Duration.ofMillis(millis);
    }

    public static String formatTime(Object given, String expectedPattern) {
        if (given == null || !(given instanceof Duration) || expectedPattern == null || expectedPattern.isEmpty()) {
            return null;
        }
        long totalSeconds = ((Duration) given).getSeconds();
        long days = totalSeconds / 86400;
        long hours = (totalSeconds - (days * 86400)) / 3600;
        long minutes = (totalSeconds - (days * 86400) - (hours * 3600)) / 60;
        long seconds = (totalSeconds - (days * 86400) - (hours * 3600) - (minutes * 60));
        LocalTime tempTime = LocalTime.of((int) hours, (int) minutes, (int) seconds, (int) days);
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(expectedPattern);
        return formatter.format(tempTime);
    }

}
