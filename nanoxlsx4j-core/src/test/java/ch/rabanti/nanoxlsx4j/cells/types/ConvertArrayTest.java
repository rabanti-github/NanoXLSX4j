package ch.rabanti.nanoxlsx4j.cells.types;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.math.BigDecimal;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import ch.rabanti.nanoxlsx4j.Cell;

class ConvertArrayTest {

    @Test
    @DisplayName("Test of the convertArray method on bools")
    void convertBoolArrayTest() {
        Boolean[] array = {
                true,
                true,
                false,
                true,
                false};
        assertArray(array, Boolean.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on bytes")
    void convertByteArrayTest() {
        // C# byte values above 127 use Java Short to preserve the unsigned range.
        Short[] array = {
                12,
                55,
                127,
                0,
                1,
                255,
                0,
                255};
        assertArray(array, Short.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on unsigned byte values") // equivalent to C# sb
    void convertSByteArrayTest() {
        Byte[] array = {
                12,
                55,
                127,
                -128,
                -1,
                0,
                -128,
                127};
        assertArray(array, Byte.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on BigDecimal")
    void convertDecimalArrayTest() {
        BigDecimal[] array = {
                new BigDecimal("0"),
                new BigDecimal("11.7"),
                new BigDecimal("0.00001"),
                new BigDecimal("-22.5"),
                new BigDecimal("100"),
                new BigDecimal("-99"),
                new BigDecimal("-79228162514264337593543950335"), // min
                new BigDecimal("79228162514264337593543950335") // max
        };
        assertArray(array, BigDecimal.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on double")
    void convertDoubleArrayTest() {
        Double[] array = {
                0d,
                11.7d,
                0.00001d,
                -22.5d,
                100d,
                -99d,
                -Double.MAX_VALUE,
                Double.MAX_VALUE};
        assertArray(array, Double.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on float")
    void convertFloatArrayTest() {
        Float[] array = {
                0f,
                11.7f,
                0.00001f,
                -22.5f,
                100f,
                -99f,
                -Float.MAX_VALUE,
                Float.MAX_VALUE};
        assertArray(array, Float.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on int")
    void convertIntArrayTest() {
        Integer[] array = {
                12,
                55,
                -1,
                0,
                Integer.MAX_VALUE,
                Integer.MIN_VALUE};
        assertArray(array, Integer.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on unsigned int values") // Equivalent to C# uint
    void convertUintArrayTest() {
        Long[] array = {
                12L,
                55L,
                777L,
                0L,
                4294967295L,
                0L};
        assertArray(array, Long.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on long")
    void convertLongArrayTest() {
        Long[] array = {
                12L,
                55L,
                -1L,
                0L,
                Long.MAX_VALUE,
                Long.MIN_VALUE};
        assertArray(array, Long.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on unsigned long values") // Equivalent to C# ulong
    void convertULongArrayTest() {
        BigDecimal[] array = {
                new BigDecimal("12"),
                new BigDecimal("55"),
                new BigDecimal("777"),
                new BigDecimal("0"),
                new BigDecimal("18446744073709551615"),
                new BigDecimal("0")
        };
        assertArray(array, BigDecimal.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on short")
    void convertShortArrayTest() {
        Short[] array = {
                12,
                55,
                -1,
                0,
                Short.MAX_VALUE,
                Short.MIN_VALUE};
        assertArray(array, Short.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on unsigned short values") // Equivalent to C# ushort
    void convertUShortArrayTest() {
        Integer[] array = {
                12,
                55,
                777,
                0,
                65535,
                0};
        assertArray(array, Integer.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on Date")
    void convertDateArrayTest() {
        Date[] array = {
                date(1901, 1, 12, 12, 12, 12),
                date(2200, 1, 12, 12, 12, 12),
                date(2020, 11, 12, 0, 0, 0),
                date(1950, 5, 1, 0, 0, 0)
        };
        assertArray(array, Date.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on Duration")
    void convertDurationArrayTest() {
        Duration[] array = {
                Duration.ZERO,
                Duration.ofHours(12).plusMinutes(10).plusSeconds(50),
                Duration.ofHours(23).plusMinutes(59).plusSeconds(59),
                Duration.ofHours(11).plusMinutes(11).plusSeconds(11)
        };
        assertArray(array, Duration.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on nested Cell objects")
    void convertCellArrayTest() {
        Cell[] array = {
                new Cell("", Cell.CellType.STRING),
                new Cell("test", Cell.CellType.STRING),
                new Cell("x", Cell.CellType.STRING),
                new Cell(" ", Cell.CellType.STRING)
        };
        assertArray(
                array, String.class, new Object[]{
                        "",
                        "test",
                        "x",
                        " "}
        );
    }

    @Test
    @DisplayName("Test of the convertArray method on string")
    void convertStringArrayTest() {
        String[] array = {
                "",
                "test",
                "X",
                "Ø",
                null,
                " "};
        assertArray(array, String.class);
    }

    @Test
    @DisplayName("Test of the convertArray method on other object types")
    void convertObjectArrayTest() {
        DummyArrayClass[] array = {
                new DummyArrayClass(""),
                new DummyArrayClass(null),
                new DummyArrayClass(" "),
                new DummyArrayClass("test")
        };
        Object[] actualValues = new Object[array.length];
        for (int i = 0; i < actualValues.length; i++) {
            actualValues[i] = array[i].toString();
        }
        assertArray(array, String.class, actualValues);
    }

    @Test
    @DisplayName("Test of the convertArray method on null and empty arrays")
    void convertObjectArrayEmptyTest() {
        List<String> nullArray = null;
        List<Cell> cells = Cell.convertArray(nullArray);
        assertTrue(cells.isEmpty());

        List<String> emptyArray = List.of();
        List<Cell> cells2 = Cell.convertArray(emptyArray);
        assertTrue(cells2.isEmpty());
    }

    private static <T> void assertArray(T[] array, Class<?> expectedValueType) {
        assertArray(array, expectedValueType, null);
    }

    private static <T> void assertArray(T[] array, Class<?> expectedValueType, Object[] actualValues) {
        List<T> list = Arrays.asList(array);
        List<Cell> cells = Cell.convertArray(list);
        assertNotNull(cells);
        assertEquals(array.length, cells.size());
        for (int i = 0; i < array.length; i++) {
            Cell cell = cells.get(i);
            if (cell.getValue() != null) {
                assertEquals(expectedValueType, cell.getValue().getClass());
            }
            assertEquals(actualValues == null ? array[i] : actualValues[i], cell.getValue());
        }
    }

    private static Date date(int year, int month, int day, int hour, int minute, int second) {
        return Date.from(LocalDateTime.of(year, month, day, hour, minute, second).toInstant(ZoneOffset.UTC));
    }

    private static class DummyArrayClass {
        private final String value;

        DummyArrayClass(String value) {
            this.value = value;
        }

        @Override
        public String toString() {
            return value;
        }
    }
}
