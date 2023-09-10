package ch.rabanti.nanoxlsx4j.cells.types;

import ch.rabanti.nanoxlsx4j.Cell;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.math.BigDecimal;
import java.util.Objects;
import java.util.function.BiFunction;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.assertTrue;

public class NumericCellTest {

    CellTypeUtils byteUtils;
    CellTypeUtils bigDecimalUtils;
    CellTypeUtils doubleUtils;
    CellTypeUtils floatUtils;
    CellTypeUtils integerUtils;
    CellTypeUtils longUtils;
    CellTypeUtils shortUtils;

    public NumericCellTest() {
        byteUtils = new CellTypeUtils(Byte.class);
        bigDecimalUtils = new CellTypeUtils(BigDecimal.class);
        doubleUtils = new CellTypeUtils(Double.class);
        floatUtils = new CellTypeUtils(Float.class);
        integerUtils = new CellTypeUtils(Integer.class);
        longUtils = new CellTypeUtils(Long.class);
        shortUtils = new CellTypeUtils(Short.class);
    }

    @DisplayName("Byte value cell test: Test of the cell values, as well as proper modification")
    @ParameterizedTest(name = "Given value {0} should lead to a valid cell")
    @CsvSource(
            {
                    "0",
                    "16",
                    "-32",
                    "Byte.MIN_VALUE",
                    "Byte.MAX_VALUE",}
    )
    void byteCellTest(String byteString) {
        genericAssertion(
                byteUtils,
                byteString,
                (byte) 8,
                Byte.MIN_VALUE,
                Byte.MAX_VALUE,
                Byte::parseByte,
                Byte::equals
        );
    }

    @DisplayName("BigDecimal value cell test: Test of the cell values, as well as proper modification")
    @ParameterizedTest(name = "Given value {0} should lead to a valid cell")
    @CsvSource(
            {
                    "0",
                    "16",
                    "-32",
                    "Integer.MIN_VALUE",
                    "Integer.MAX_VALUE",}
    )
    void bigDecimalCellTest(String byteString) {
        // Note: BigDecimal has no defined MIN and MAX
        genericAssertion(
                bigDecimalUtils,
                byteString,
                BigDecimal.valueOf(8),
                BigDecimal.valueOf(Integer.MIN_VALUE),
                BigDecimal.valueOf(Integer.MAX_VALUE),
                BigDecimal::new,
                BigDecimal::equals
        );
    }

    @DisplayName("Double value cell test: Test of the cell values, as well as proper modification")
    @ParameterizedTest(name = "Given value {0} should lead to a valid cell")
    @CsvSource(
            {
                    "0",
                    "1.22",
                    "Double.MIN_VALUE",
                    "Double.MAX_VALUE",}
    )
    void doubleCellTest(String doubleString) {
        genericAssertion(
                doubleUtils,
                doubleString,
                8d,
                Double.MIN_VALUE,
                Double.MAX_VALUE,
                Double::parseDouble,
                NumericCellTest::compareDouble
        );
    }

    @DisplayName("Float value cell test: Test of the cell values, as well as proper modification")
    @ParameterizedTest(name = "Given value {0} should lead to a valid cell")
    @CsvSource(
            {
                    "0",
                    "779.254",
                    "Float.MIN_VALUE",
                    "Float.MAX_VALUE",}
    )
    void floatCellTest(String floatString) {
        genericAssertion(
                floatUtils,
                floatString,
                8f,
                Float.MIN_VALUE,
                Float.MAX_VALUE,
                Float::parseFloat,
                NumericCellTest::compareFloat
        );
    }

    @DisplayName("Integer value cell test: Test of the cell values, as well as proper modification")
    @ParameterizedTest(name = "Given value {0} should lead to a valid cell")
    @CsvSource(
            {
                    "0",
                    "-42",
                    "42",
                    "Integer.MIN_VALUE",
                    "Integer.MAX_VALUE",}
    )
    void integerCellTest(String integerString) {
        genericAssertion(
                integerUtils,
                integerString,
                8,
                Integer.MIN_VALUE,
                Integer.MAX_VALUE,
                Integer::parseInt,
                Objects::equals
        );
    }

    @DisplayName("Long value cell test: Test of the cell values, as well as proper modification")
    @ParameterizedTest(name = "Given value {0} should lead to a valid cell")
    @CsvSource(
            {
                    "0",
                    "-999999999",
                    "55555",
                    "Long.MIN_VALUE",
                    "Long.MAX_VALUE",}
    )
    void longCellTest(String longString) {
        genericAssertion(longUtils, longString, 8L, Long.MIN_VALUE, Long.MAX_VALUE, Long::parseLong, Objects::equals);
    }

    @DisplayName("Short value cell test: Test of the cell values, as well as proper modification")
    @ParameterizedTest(name = "Given value {0} should lead to a valid cell")
    @CsvSource(
            {
                    "0",
                    "127",
                    "-127",
                    "Short.MIN_VALUE",
                    "Short.MAX_VALUE",}
    )
    void shortCellTest(String shortString) {
        genericAssertion(
                shortUtils,
                shortString,
                (short) 8,
                Short.MIN_VALUE,
                Short.MAX_VALUE,
                Short::parseShort,
                Objects::equals
        );
    }

    @DisplayName("Test of the byte comparison method on cells")
    @ParameterizedTest(name = "Given values of {0} and {1} should lead to a comparison result of {2}")
    @CsvSource(
            {
                    "42, 42, 0",
                    "42.00001, 42.00001, 0",
                    "100, 24, 1",
                    "0, 127, -1",
                    "-127, 127, -1",
                    "-0.00001, 0, -1",}
    )
    void bigDecimalCellComparisonTest(double value1, double value2, int expectedResult) {
        BigDecimal v1 = BigDecimal.valueOf(value1);
        BigDecimal v2 = BigDecimal.valueOf(value2);
        assertNumericType(bigDecimalUtils, value1, value2, expectedResult);
    }

    @DisplayName("Test of the byte comparison method on cells")
    @ParameterizedTest(name = "Given values of {0} and {1} should lead to a comparison result of {2}")
    @CsvSource(
            {
                    "42, 42, 0",
                    "100, 24, 1",
                    "0, 127, -1",
                    "-127, 127, -1",}
    )
    void byteCellComparisonTest(byte value1, byte value2, int expectedResult) {
        assertNumericType(byteUtils, value1, value2, expectedResult);
    }

    @DisplayName("Test of the double comparison method on cells")
    @ParameterizedTest(name = "Given values of {0} and {1} should lead to a comparison result of {2}")
    @CsvSource(
            {
                    "22.223, 22.223, 0",
                    "9999.1, -0.001, 1",
                    "0.23, 18720.0, -1",}
    )
    void doubleCellComparisonTest(double value1, double value2, int expectedResult) {
        assertNumericType(doubleUtils, value1, value2, expectedResult);
    }

    @DisplayName("Test of the float comparison method on cells")
    @ParameterizedTest(name = "Given values of {0} and {1} should lead to a comparison result of {2}")
    @CsvSource(
            {
                    "22.223, 22.223, 0",
                    "9999.1, -0.001, 1",
                    "0.23, 18720.0, -1",}
    )
    void floatCellComparisonTest(float value1, float value2, int expectedResult) {
        assertNumericType(floatUtils, value1, value2, expectedResult);
    }

    @DisplayName("Test of the int comparison method on cells")
    @ParameterizedTest(name = "Given values of {0} and {1} should lead to a comparison result of {2}")
    @CsvSource(
            {
                    "42, 42, 0",
                    "9999, -999999, 1",
                    "0, 18720, -1",}
    )
    void intCellComparisonTest(int value1, int value2, int expectedResult) {
        assertNumericType(integerUtils, value1, value2, expectedResult);
    }

    @DisplayName("Test of the long comparison method on cells")
    @ParameterizedTest(name = "Given values of {0} and {1} should lead to a comparison result of {2}")
    @CsvSource(
            {
                    "42, 42, 0",
                    "9999, -999999, 1",
                    "0, 18720, -1",}
    )
    void longCellComparisonTest(long value1, long value2, int expectedResult) {
        assertNumericType(longUtils, value1, value2, expectedResult);
    }

    @DisplayName("Test of the short comparison method on cells")
    @ParameterizedTest(name = "Given values of {0} and {1} should lead to a comparison result of {2}")
    @CsvSource(
            {
                    "42, 42, 0",
                    "9999, -9999, 1",
                    "0, 18720, -1",}
    )
    void shortCellComparisonTest(short value1, short value2, int expectedResult) {
        assertNumericType(shortUtils, value1, value2, expectedResult);
    }

    private <T> void genericAssertion(CellTypeUtils utilsInstance, String valueString, T initialValue, T min, T max, Function<String, T> parser, BiFunction<T, T, Boolean> comparer) {
        T value;
        if (valueString.toUpperCase().contains("MIN_VALUE")) {
            value = min;
        }
        else if (valueString.toUpperCase().contains("MAX_VALUE")) {
            value = max;
        }
        else {
            value = parser.apply(valueString);
        }
        utilsInstance.assertCellCreation(initialValue, value, Cell.CellType.NUMBER, comparer);
    }

    private <T extends Comparable<T>> void assertNumericType(CellTypeUtils utilsInstance, T value1, T value2, int expectedResult) {
        Cell cell1 = utilsInstance.createVariantCell(value1, utilsInstance.getCellAddress());
        Cell cell2 = utilsInstance.createVariantCell(value2, utilsInstance.getCellAddress());
        int comparison = ((T) cell1.getValue()).compareTo((T) cell2.getValue());
        assertTrue(variantCompareTo(comparison, expectedResult));
    }

    private static boolean variantCompareTo(int comparison, int expectedResult) {
        if (comparison == expectedResult) {
            return true;
        }
        if (comparison < 0 && expectedResult < 0) {
            return true;
        }
        return comparison > 0 && expectedResult > 0;
    }

    private static boolean compareDouble(double current, double other) {
        final double threshold = 0.0000001;
        return Math.abs(current - other) < threshold;
    }

    private static boolean compareFloat(float current, float other) {
        final float threshold = 0.0000001f;
        return Math.abs(current - other) < threshold;
    }

}
