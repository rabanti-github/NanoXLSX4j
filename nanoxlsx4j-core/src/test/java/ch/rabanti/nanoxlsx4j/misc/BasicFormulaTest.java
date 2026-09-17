package ch.rabanti.nanoxlsx4j.misc;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.math.BigDecimal;
import java.time.LocalDateTime;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.BasicFormulas;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

public class BasicFormulaTest {

    @ParameterizedTest
    @DisplayName("Test of the average function on a Range object")
    @CsvSource(
            {
                    "A1:A1, 'AVERAGE(A1)'",
                    "A1:C2, 'AVERAGE(A1:C2)'",
                    "'$A1:C2', 'AVERAGE($A1:C2)'",
                    "'$A$1:C2', 'AVERAGE($A$1:C2)'",
                    "'$A$1:$C2', 'AVERAGE($A$1:$C2)'",
                    "'$A$1:$C$2', 'AVERAGE($A$1:$C$2)'"
            }
    )
    void averageTest(String rangeExpression, String expectedFormula) {
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.average(range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the average function on a Range object and a target worksheet")
    @CsvSource(
            {
                    "worksheet1, A1:A1, 'AVERAGE(worksheet1!A1)'",
                    "worksheet1, A1:C2, 'AVERAGE(worksheet1!A1:C2)'",
                    "worksheet1, '$A1:C2', 'AVERAGE(worksheet1!$A1:C2)'",
                    "worksheet1, '$A$1:C2', 'AVERAGE(worksheet1!$A$1:C2)'",
                    "worksheet1, '$A$1:$C2', 'AVERAGE(worksheet1!$A$1:$C2)'",
                    "worksheet1, '$A$1:$C$2', 'AVERAGE(worksheet1!$A$1:$C$2)'"
            }
    )
    void averageTest2(String worksheetName, String rangeExpression, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.average(worksheet, range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the ceil function on a value and a number of decimals")
    @CsvSource(
            {
                    "A1, 1, 'ROUNDUP(A1,1)'",
                    "C4, 0, 'ROUNDUP(C4,0)'",
                    "'$A1', 10, 'ROUNDUP($A1,10)'",
                    "'$A$1', 5, 'ROUNDUP($A$1,5)'",
                    "A1, -2, 'ROUNDUP(A1,-2)'"
                    // This seems to be valid
            }
    )
    void ceilTest(String addressExpression, int numberOfDecimals, String expectedFormula) {
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.ceil(address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the ceil function on a value and a number of decimals")
    @CsvSource(
            {
                    "worksheet3, A1, 1, 'ROUNDUP(worksheet3!A1,1)'",
                    "worksheet3, C4, 0, 'ROUNDUP(worksheet3!C4,0)'",
                    "worksheet3, '$A1', 10, 'ROUNDUP(worksheet3!$A1,10)'",
                    "worksheet3, '$A$1', 5, 'ROUNDUP(worksheet3!$A$1,5)'",
                    "worksheet3, A1, -2, 'ROUNDUP(worksheet3!A1,-2)'"
                    // This seems to be valid
            }
    )
    void ceilTest2(
            String worksheetName,
            String addressExpression,
            int numberOfDecimals,
            String expectedFormula
    ) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.ceil(worksheet, address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the floor function on a value and a number of decimals")
    @CsvSource(
            {
                    "A1, 1, 'ROUNDDOWN(A1,1)'",
                    "C4, 0, 'ROUNDDOWN(C4,0)'",
                    "'$A1', 10, 'ROUNDDOWN($A1,10)'",
                    "'$A$1', 5, 'ROUNDDOWN($A$1,5)'",
                    "A1, -2, 'ROUNDDOWN(A1,-2)'"
                    // This seems to be valid
            }
    )
    void floorTest(String addressExpression, int numberOfDecimals, String expectedFormula) {
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.floor(address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the floor function on a value and a number of decimals")
    @CsvSource(
            {
                    "worksheet3, A1, 1, 'ROUNDDOWN(worksheet3!A1,1)'",
                    "worksheet3, C4, 0, 'ROUNDDOWN(worksheet3!C4,0)'",
                    "worksheet3, '$A1', 10, 'ROUNDDOWN(worksheet3!$A1,10)'",
                    "worksheet3, '$A$1', 5, 'ROUNDDOWN(worksheet3!$A$1,5)'",
                    "worksheet3, A1, -2, 'ROUNDDOWN(worksheet3!A1,-2)'"
                    // This seems to be valid
            }
    )
    void floorTest2(
            String worksheetName,
            String addressExpression,
            int numberOfDecimals,
            String expectedFormula
    ) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.floor(worksheet, address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the max function on a Range object")
    @CsvSource(
            {
                    "A1:A1, 'MAX(A1)'",
                    "A1:C2, 'MAX(A1:C2)'",
                    "'$A1:C2', 'MAX($A1:C2)'",
                    "'$A$1:C2', 'MAX($A$1:C2)'",
                    "'$A$1:$C2', 'MAX($A$1:$C2)'",
                    "'$A$1:$C$2', 'MAX($A$1:$C$2)'"
            }
    )
    void maxTest(String rangeExpression, String expectedFormula) {
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.max(range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the max function on a Range object and a target worksheet")
    @CsvSource(
            {
                    "worksheet1, A1:A1, 'MAX(worksheet1!A1)'",
                    "worksheet1, A1:C2, 'MAX(worksheet1!A1:C2)'",
                    "worksheet1, '$A1:C2', 'MAX(worksheet1!$A1:C2)'",
                    "worksheet1, '$A$1:C2', 'MAX(worksheet1!$A$1:C2)'",
                    "worksheet1, '$A$1:$C2', 'MAX(worksheet1!$A$1:$C2)'",
                    "worksheet1, '$A$1:$C$2', 'MAX(worksheet1!$A$1:$C$2)'"
            }
    )
    void maxTest2(String worksheetName, String rangeExpression, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.max(worksheet, range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the min function on a Range object")
    @CsvSource(
            {
                    "A1:A1, 'MIN(A1)'",
                    "A1:C2, 'MIN(A1:C2)'",
                    "'$A1:C2', 'MIN($A1:C2)'",
                    "'$A$1:C2', 'MIN($A$1:C2)'",
                    "'$A$1:$C2', 'MIN($A$1:$C2)'",
                    "'$A$1:$C$2', 'MIN($A$1:$C$2)'"
            }
    )
    void minTest(String rangeExpression, String expectedFormula) {
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.min(range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the min function on a Range object and a target worksheet")
    @CsvSource(
            {
                    "worksheet1, A1:A1, 'MIN(worksheet1!A1)'",
                    "worksheet1, A1:C2, 'MIN(worksheet1!A1:C2)'",
                    "worksheet1, '$A1:C2', 'MIN(worksheet1!$A1:C2)'",
                    "worksheet1, '$A$1:C2', 'MIN(worksheet1!$A$1:C2)'",
                    "worksheet1, '$A$1:$C2', 'MIN(worksheet1!$A$1:$C2)'",
                    "worksheet1, '$A$1:$C$2', 'MIN(worksheet1!$A$1:$C$2)'"
            }
    )
    void minTest2(String worksheetName, String rangeExpression, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.min(worksheet, range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the median function on a Range object")
    @CsvSource(
            {
                    "A1:A1, 'MEDIAN(A1)'",
                    "A1:C2, 'MEDIAN(A1:C2)'",
                    "'$A1:C2', 'MEDIAN($A1:C2)'",
                    "'$A$1:C2', 'MEDIAN($A$1:C2)'",
                    "'$A$1:$C2', 'MEDIAN($A$1:$C2)'",
                    "'$A$1:$C$2', 'MEDIAN($A$1:$C$2)'"
            }
    )
    void medianTest(String rangeExpression, String expectedFormula) {
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.median(range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the median function on a Range object and a target worksheet")
    @CsvSource(
            {
                    "worksheet1, A1:A1, 'MEDIAN(worksheet1!A1)'",
                    "worksheet1, A1:C2, 'MEDIAN(worksheet1!A1:C2)'",
                    "worksheet1, '$A1:C2', 'MEDIAN(worksheet1!$A1:C2)'",
                    "worksheet1, '$A$1:C2', 'MEDIAN(worksheet1!$A$1:C2)'",
                    "worksheet1, '$A$1:$C2', 'MEDIAN(worksheet1!$A$1:$C2)'",
                    "worksheet1, '$A$1:$C$2', 'MEDIAN(worksheet1!$A$1:$C$2)'"
            }
    )
    void medianTest2(String worksheetName, String rangeExpression, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.median(worksheet, range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the round function on a value and a number of decimals")
    @CsvSource(
            {
                    "A1, 1, 'ROUND(A1,1)'",
                    "C4, 0, 'ROUND(C4,0)'",
                    "'$A1', 10, 'ROUND($A1,10)'",
                    "'$A$1', 5, 'ROUND($A$1,5)'",
                    "A1, -2, 'ROUND(A1,-2)'"
            }
    )
    void roundTest(String addressExpression, int numberOfDecimals, String expectedFormula) {
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.round(address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the round function on a value and a number of decimals")
    @CsvSource(
            {
                    "worksheet3, A1, 1, 'ROUND(worksheet3!A1,1)'",
                    "worksheet3, C4, 0, 'ROUND(worksheet3!C4,0)'",
                    "worksheet3, '$A1', 10, 'ROUND(worksheet3!$A1,10)'",
                    "worksheet3, '$A$1', 5, 'ROUND(worksheet3!$A$1,5)'",
                    "worksheet3, A1, -2, 'ROUND(worksheet3!A1,-2)'"
            }
    )
    void roundTest2(
            String worksheetName,
            String addressExpression,
            int numberOfDecimals,
            String expectedFormula
    ) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.round(worksheet, address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the sum function on a Range object")
    @CsvSource(
            {
                    "A1:A1, 'SUM(A1)'",
                    "A1:C2, 'SUM(A1:C2)'",
                    "'$A1:C2', 'SUM($A1:C2)'",
                    "'$A$1:C2', 'SUM($A$1:C2)'",
                    "'$A$1:$C2', 'SUM($A$1:$C2)'",
                    "'$A$1:$C$2', 'SUM($A$1:$C$2)'"
            }
    )
    void sumTest(String rangeExpression, String expectedFormula) {
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.sum(range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName("Test of the sum function on a Range object and a target worksheet")
    @CsvSource(
            {
                    "worksheet1, A1:A1, 'SUM(worksheet1!A1)'",
                    "worksheet1, A1:C2, 'SUM(worksheet1!A1:C2)'",
                    "worksheet1, '$A1:C2', 'SUM(worksheet1!$A1:C2)'",
                    "worksheet1, '$A$1:C2', 'SUM(worksheet1!$A$1:C2)'",
                    "worksheet1, '$A$1:$C2', 'SUM(worksheet1!$A$1:$C2)'",
                    "worksheet1, '$A$1:$C$2', 'SUM(worksheet1!$A$1:$C$2)'"
            }
    )
    void sumTest2(String worksheetName, String rangeExpression, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.sum(worksheet, range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName(
            "Test of the vLookup function on a Range object with an arbitrary number, the column index and the " +
                    "option of an exact match"
    )
    @CsvSource(
            {
                    "INTEGER, 11, A1:A1, 1, false, 'VLOOKUP(11,A1:A1,1,FALSE)'",
                    "FLOAT, 0.5, A1:C4, 3, false, 'VLOOKUP(0.5,A1:C4,3,FALSE)'",
                    "LONG, -800, A10:XFD999999, 200, true, 'VLOOKUP(-800,A10:XFD999999,200,TRUE)'",
                    "INTEGER, 0, X100:A1, 5, true, 'VLOOKUP(0,A1:X100,5,TRUE)'"
            }
    )
    void vLookupTest(
            String numberType,
            String numberValue,
            String rangeExpression,
            int columnIndex,
            boolean exactMatch,
            String expectedFormula
    ) {
        Object number = createNumber(numberType, numberValue);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.vLookup(number, range, columnIndex, exactMatch);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName(
            "Test of the vLookup function on a Range object with an arbitrary number, the column index, the option " +
                    "of an exact match and a target worksheet"
    )
    @CsvSource(
            {
                    "worksheet1, INTEGER, 11, '$A$1:A1', 1, false, 'VLOOKUP(11,worksheet1!$A$1:A1,1,FALSE)'",
                    "worksheet1, DOUBLE, 0.5, 'A1:$C4', 3, false, 'VLOOKUP(0.5,worksheet1!A1:$C4,3,FALSE)'",
                    "worksheet1, DOUBLE, 2.22, '$A10:XFD999999', 200, true, " +
                            "'VLOOKUP(2.22,worksheet1!$A10:XFD999999,200,TRUE)'",
                    "worksheet1, INTEGER, 0, X100:A1, 5, true, 'VLOOKUP(0,worksheet1!A1:X100,5,TRUE)'"
            }
    )
    void vLookupTest2(
            String worksheetName,
            String numberType,
            String numberValue,
            String rangeExpression,
            int columnIndex,
            boolean exactMatch,
            String expectedFormula
    ) {
        Object number = createNumber(numberType, numberValue);
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.vLookup(number, worksheet, range, columnIndex, exactMatch);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName(
            "Test of the vLookup function on a Range object with reference address, the column index and the option " +
                    "of an exact match"
    )
    @CsvSource(
            {
                    "C5, 'A1:$A$1', 1, false, 'VLOOKUP(C5,A1:$A$1,1,FALSE)'",
                    "A1, 'A1:C$4', 3, false, 'VLOOKUP(A1,A1:C$4,3,FALSE)'",
                    "'$F4', A10:XFD999999, 200, true, 'VLOOKUP($F4,A10:XFD999999,200,TRUE)'",
                    "'$XFD$99999', X100:A1, 5, true, 'VLOOKUP($XFD$99999,A1:X100,5,TRUE)'"
            }
    )
    void vLookupTest3(
            String addressExpression,
            String rangeExpression,
            int columnIndex,
            boolean exactMatch,
            String expectedFormula
    ) {
        Address address = new Address(addressExpression);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.vLookup(address, range, columnIndex, exactMatch);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @ParameterizedTest
    @DisplayName(
            "Test of the vLookup function on a Range object with reference address, the column index, the option of " +
                    "an exact match and tow target worksheets"
    )
    @CsvSource(
            {
                    "worksheet1, C5, worksheet1, 'A1:$A$1', 1, false, " +
                            "'VLOOKUP(worksheet1!C5,worksheet1!A1:$A$1,1,FALSE)'",
                    "worksheet2, A1, worksheet1, 'A1:C$4', 3, false, " +
                            "'VLOOKUP(worksheet2!A1,worksheet1!A1:C$4,3,FALSE)'",
                    "worksheet1, '$F4', worksheet2, A10:XFD999999, 200, true, " +
                            "'VLOOKUP(worksheet1!$F4,worksheet2!A10:XFD999999,200,TRUE)'",
                    "worksheet2, '$XFD$99999', worksheet2, X100:A1, 5, true, " +
                            "'VLOOKUP(worksheet2!$XFD$99999,worksheet2!A1:X100,5,TRUE)'"
            }
    )
    void vLookupTest4(
            String valueWorksheetName,
            String addressExpression,
            String rangesWorksheetName,
            String rangeExpression,
            int columnIndex,
            boolean exactMatch,
            String expectedFormula
    ) {
        Worksheet valueWorksheet = new Worksheet(valueWorksheetName);
        Worksheet rangeWorksheet = new Worksheet(rangesWorksheetName);
        Address address = new Address(addressExpression);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.vLookup(
                valueWorksheet,
                address,
                rangeWorksheet,
                range,
                columnIndex,
                exactMatch
        );
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @Test
    @DisplayName("Test of the vLookup function for byte as value")
    void vLookupByteTest() {
        assertVLookup((byte) 0, "0");
        assertVLookup((byte) 15, "15");
        assertVLookup((byte) 15, "15");
    }

    // The C# sbyte case has no distinct Java runtime type and is covered by the signed byte tests.

    @Test
    @DisplayName("Test of the vLookup function for decimal as value")
    void vLookupDecimalTest() {
        BigDecimal d1 = new BigDecimal("0");
        BigDecimal d2 = new BigDecimal("-0.005");
        BigDecimal d3 = new BigDecimal("22.78");
        assertVLookup(d1, "0");
        assertVLookup(d2, "-0.005");
        assertVLookup(d3, "22.78");
    }

    @Test
    @DisplayName("Test of the vLookup function for double as value")
    void vLookupDoubleTest() {
        assertVLookup(0.0d, "0");
        assertVLookup(222.5d, "222.5");
        assertVLookup(-0.101d, "-0.101");
    }

    @Test
    @DisplayName("Test of the vLookup function for float as value")
    void vLookupFloatTest() {
        assertVLookup(0.0f, "0");
        assertVLookup(22.5f, "22.5");
        assertVLookup(-0.01f, "-0.01");
    }

    @Test
    @DisplayName("Test of the vLookup function for int as value")
    void vLookupIntTest() {
        assertVLookup(0, "0");
        assertVLookup(-77, "-77");
        assertVLookup(77, "77");
    }

    // The C# uint and ulong cases have no distinct Java runtime type and are covered by the signed integral tests.

    @Test
    @DisplayName("Test of the vLookup function for long as value")
    void vLookupLongTest() {
        assertVLookup(0L, "0");
        assertVLookup(-999999L, "-999999");
        assertVLookup(999999L, "999999");
    }

    @Test
    @DisplayName("Test of the vLookup function for short as value")
    void vLookupShortTest() {
        assertVLookup((short) 0, "0");
        assertVLookup((short) -128, "-128");
        assertVLookup((short) 255, "255");
    }

    // The C# ushort case has no distinct Java runtime type and is covered by the int test.

    @Test
    @DisplayName("Test of the failing vLookup function on an invalid value type")
    void vLookupFailTest() {
        Range range = new Range("A1:D100");
        int column = 2;
        boolean exactMatch = true;
        assertThrows(FormatException.class, () -> BasicFormulas.vLookup("test", range, column, exactMatch));
        assertThrows(FormatException.class, () -> BasicFormulas.vLookup(false, range, column, exactMatch));
        // C# Address is a value type, so null selects the numeric object overload there.
        assertThrows(FormatException.class, () -> BasicFormulas.vLookup((Object) null, range, column, exactMatch));
        assertThrows(
                FormatException.class,
                () -> BasicFormulas.vLookup(LocalDateTime.of(1, 1, 1, 0, 0), range, column, exactMatch)
        );
    }

    @Test
    @DisplayName("Test of the failing vLookup function on an invalid index column")
    void vLookupFailTest2() {
        Range range = new Range("A1:D100");
        Range range2 = new Range("C1:D100");
        assertThrows(FormatException.class, () -> BasicFormulas.vLookup(22, range, 0, true));
        assertThrows(FormatException.class, () -> BasicFormulas.vLookup(22, range, -2, false));
        assertThrows(FormatException.class, () -> BasicFormulas.vLookup(22, range, 100, true));
        assertThrows(FormatException.class, () -> BasicFormulas.vLookup(22, range2, 3, true));
        assertThrows(FormatException.class, () -> BasicFormulas.vLookup(22, range2, 4, false));
    }

    private void assertVLookup(Object number, String expectedLookupValue) {
        Range range = new Range("A1:D100");
        int column = 2;
        boolean exactMatch = true;
        Cell formula = BasicFormulas.vLookup(number, range, column, exactMatch);
        String expectedFormula = "VLOOKUP(" + expectedLookupValue + "," + range + "," + column + "," +
                Boolean.toString(exactMatch).toUpperCase() + ")";
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    private Object createNumber(String numberType, String numberValue) {
        return switch (numberType) {
            case "BYTE" -> Byte.valueOf(numberValue);
            case "FLOAT" -> Float.valueOf(numberValue);
            case "DOUBLE" -> Double.valueOf(numberValue);
            case "INTEGER" -> Integer.valueOf(numberValue);
            case "LONG" -> Long.valueOf(numberValue);
            default -> throw new IllegalArgumentException("Unsupported test number type: " + numberType);
        };
    }
}
