package ch.rabanti.nanoxlsx4j.cells.types;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.math.BigDecimal;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.ValueSource;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;

class NumericCellTest {
    private final CellTypeUtils utils = new CellTypeUtils();

    // C# byte is unsigned; short preserves its 0..255 test values in Java.
    @ParameterizedTest
    @DisplayName("Byte value cell test: Test of the cell values, as well as proper modification")
    @ValueSource(
            shorts = {
                    0,
                    16,
                    0,
                    255}
    )
    void byteCellTest(short value) {
        utils.assertCellCreation((short) 8, value, Cell.CellType.NUMBER, Short::equals);
    }

    @Test
    @DisplayName("Byte value cell test with style")
    void byteCellTest2() {
        Style style = BasicStyles.getItalic();
        utils.assertStyledCellCreation((short) 0, (short) 8, Cell.CellType.NUMBER, Short::equals, style);
    }

    @ParameterizedTest
    @DisplayName("Signed Byte value cell test: Test of the cell values, as well as proper modification")
    @ValueSource(
            bytes = {
                    0,
                    -22,
                    -128,
                    127}
    )
    void signedByteCellTest(byte value) {
        utils.assertCellCreation((byte) 8, value, Cell.CellType.NUMBER, Byte::equals);
    }

    @ParameterizedTest
    @DisplayName("Short value cell test: Test of the cell values, as well as proper modification")
    @ValueSource(
            shorts = {
                    0,
                    -127,
                    Short.MIN_VALUE,
                    Short.MAX_VALUE}
    )
    void shortCellTest(short value) {
        utils.assertCellCreation((short) -6, value, Cell.CellType.NUMBER, Short::equals);
    }

    // C# ushort maps to int for its 0..65535 test range.
    @ParameterizedTest
    @DisplayName("Unsigned Short value cell test: Test of the cell values, as well as proper modification")
    @ValueSource(
            ints = {
                    0,
                    127,
                    0,
                    65535}
    )
    void unsignedShortCellTest(int value) {
        utils.assertCellCreation(3398, value, Cell.CellType.NUMBER, Integer::equals);
    }

    @ParameterizedTest
    @DisplayName("Int value cell test: Test of the cell values, as well as proper modification")
    @ValueSource(
            ints = {
                    0,
                    -42,
                    Integer.MIN_VALUE,
                    Integer.MAX_VALUE}
    )
    void intCellTest(int value) {
        utils.assertCellCreation(99, value, Cell.CellType.NUMBER, Integer::equals);
    }

    // C# uint maps to long for its 0..4294967295 test range.
    @ParameterizedTest
    @DisplayName("Unsigned Int value cell test: Test of the cell values, as well as proper modification")
    @ValueSource(
            longs = {
                    0L,
                    42L,
                    0L,
                    4294967295L}
    )
    void unsignedIntCellTest(long value) {
        utils.assertCellCreation(98L, value, Cell.CellType.NUMBER, Long::equals);
    }

    @ParameterizedTest
    @DisplayName("Long value cell test: Test of the cell values, as well as proper modification")
    @ValueSource(
            longs = {
                    0L,
                    -999999999L,
                    Long.MIN_VALUE,
                    Long.MAX_VALUE}
    )
    void longCellTest(long value) {
        utils.assertCellCreation(-6L, value, Cell.CellType.NUMBER, Long::equals);
    }

    // BigDecimal preserves C# ulong's full value range without inventing an unsigned Java type.
    @ParameterizedTest
    @DisplayName("Unsigned Long value cell test: Test of the cell values, as well as proper modification")
    @ValueSource(
            strings = {
                    "0",
                    "55555",
                    "0",
                    "18446744073709551615"}
    )
    void unsignedLongCellTest(String value) {
        utils.assertCellCreation(
                new BigDecimal("99"), new BigDecimal(value),
                Cell.CellType.NUMBER, BigDecimal::equals
        );
    }

    @Test
    @DisplayName("BigDecimal value cell test: Test of the cell values, as well as proper modification")
    void bigDecimalCellTest() {
        BigDecimal initial = new BigDecimal("-2.338");
        utils.assertCellCreation(initial, new BigDecimal("0"), Cell.CellType.NUMBER, NumericCellTest::compareDecimal);
        utils.assertCellCreation(
                initial, new BigDecimal("-0.0057"), Cell.CellType.NUMBER,
                NumericCellTest::compareDecimal
        );
        utils.assertCellCreation(
                initial, new BigDecimal("-79228162514264337593543950335"), Cell.CellType.NUMBER,
                NumericCellTest::compareDecimal
        );
        utils.assertCellCreation(
                initial, new BigDecimal("79228162514264337593543950335"), Cell.CellType.NUMBER,
                NumericCellTest::compareDecimal
        );
    }

    @ParameterizedTest
    @DisplayName("Float value cell test: Test of the cell values, as well as proper modification")
    @ValueSource(
            floats = {
                    0f,
                    779.254f,
                    -Float.MAX_VALUE,
                    Float.MAX_VALUE}
    )
    void floatCellTest(float value) {
        utils.assertCellCreation(-2.338f, value, Cell.CellType.NUMBER, NumericCellTest::compareFloat);
    }

    @ParameterizedTest
    @DisplayName("Double value cell test: Test of the cell values, as well as proper modification")
    @ValueSource(
            doubles = {
                    0d,
                    1.22d,
                    -Double.MAX_VALUE,
                    Double.MAX_VALUE}
    )
    void doubleCellTest(double value) {
        utils.assertCellCreation(42.778d, value, Cell.CellType.NUMBER, NumericCellTest::compareDouble);
    }

    @ParameterizedTest
    @DisplayName("Test of the byte comparison method on cells")
    @CsvSource(
            {
                    "42,42,0",
                    "100,24,1",
                    "0,127,-1"}
    )
    void byteCellComparisonTest(short value1, short value2, int expectedResult) {
        assertNumericType(value1, value2, expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the signed byte comparison method on cells") // Equivalent to C# sbyte
    @CsvSource(
            {
                    "42,42,0",
                    "100,-20,1",
                    "-127,127,-1"}
    )
    void signedByteCellComparisonTest(byte value1, byte value2, int expectedResult) {
        assertNumericType(value1, value2, expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the int comparison method on cells")
    @CsvSource(
            {
                    "42,42,0",
                    "9999,-999999,1",
                    "0,18720,-1"}
    )
    void intCellComparisonTest(int value1, int value2, int expectedResult) {
        assertNumericType(value1, value2, expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the unsigned int comparison method on cells") // Equivalent to C# uint
    @CsvSource(
            {
                    "42,42,0",
                    "10000,2004,1",
                    "5000,12700,-1"}
    )
    void unsignedIntCellComparisonTest(long value1, long value2, int expectedResult) {
        assertNumericType(value1, value2, expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the long comparison method on cells")
    @CsvSource(
            {
                    "42,42,0",
                    "9999,-999999,1",
                    "0,18720,-1"}
    )
    void longCellComparisonTest(long value1, long value2, int expectedResult) {
        assertNumericType(value1, value2, expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the unsigned long comparison method on cells") // Equivalent to C# ulong
    @CsvSource(
            {
                    "42,42,0",
                    "10000,2004,1",
                    "5000,12700,-1"}
    )
    void unsignedLongCellComparisonTest(long value1, long value2, int expectedResult) {
        assertNumericType(new BigDecimal(value1), new BigDecimal(value2), expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the short comparison method on cells")
    @CsvSource(
            {
                    "42,42,0",
                    "9999,-9999,1",
                    "0,18720,-1"}
    )
    void shortCellComparisonTest(short value1, short value2, int expectedResult) {
        assertNumericType(value1, value2, expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the unsigned short comparison method on cells")
    @CsvSource(
            {
                    "42,42,0",
                    "1000,204,1",
                    "500,1200,-1"}
    )
    void unsignedShortCellComparisonTest(int value1, int value2, int expectedResult) { // Equivalent to C# ushort
        assertNumericType(value1, value2, expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the BigDecimal comparison method on cells")
    @CsvSource(
            {
                    "22.223,22.223,0",
                    "9999.1,-0.001,1",
                    "0.23,18720.0,-1"}
    )
    void bigDecimalCellComparisonTest(BigDecimal value1, BigDecimal value2, int expectedResult) {
        assertNumericType(value1, value2, expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the double comparison method on cells")
    @CsvSource(
            {
                    "22.223,22.223,0",
                    "9999.1,-0.001,1",
                    "0.23,18720.0,-1"}
    )
    void doubleCellComparisonTest(double value1, double value2, int expectedResult) {
        assertNumericType(value1, value2, expectedResult);
    }

    @ParameterizedTest
    @DisplayName("Test of the float comparison method on cells")
    @CsvSource(
            {
                    "22.223,22.223,0",
                    "9999.1,-0.001,1",
                    "0.23,18720.0,-1"}
    )
    void floatCellComparisonTest(float value1, float value2, int expectedResult) {
        assertNumericType(value1, value2, expectedResult);
    }

    @SuppressWarnings("unchecked")
    private <T extends Comparable<T>> void assertNumericType(T value1, T value2, int expectedResult) {
        Cell cell1 = utils.createVariantCell(value1, utils.getCellAddress());
        Cell cell2 = utils.createVariantCell(value2, utils.getCellAddress());
        int comparison = ((T) cell1.getValue()).compareTo((T) cell2.getValue());
        assertEquals(Integer.signum(expectedResult), Integer.signum(comparison));
    }

    private static boolean compareDouble(double current, double other) {
        return Math.abs(current - other) < 0.0000001d;
    }

    private static boolean compareDecimal(BigDecimal current, BigDecimal other) {
        return current.subtract(other).abs().compareTo(new BigDecimal("0.0000001")) < 0;
    }

    private static boolean compareFloat(float current, float other) {
        return Math.abs(current - other) < 0.0000001f;
    }
}
