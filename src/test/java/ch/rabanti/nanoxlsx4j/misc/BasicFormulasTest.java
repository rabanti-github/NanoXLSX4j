package ch.rabanti.nanoxlsx4j.misc;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.BasicFormulas;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.math.BigDecimal;
import java.util.Date;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

public class BasicFormulasTest {
    @DisplayName("Test of the average function on a Range object")
    @ParameterizedTest(name = "Given range {0} should lead to the formula {1}")
    @CsvSource({
            "'A1:A1', 'AVERAGE(A1)'",
            "'A1:C2', 'AVERAGE(A1:C2)'",
            "'$A1:C2', 'AVERAGE($A1:C2)'",
            "'$A$1:C2', 'AVERAGE($A$1:C2)'",
            "'$A$1:$C2', 'AVERAGE($A$1:$C2)'",
            "'$A$1:$C$2', 'AVERAGE($A$1:$C$2)'",
    })
    void averageTest(String rangeExpression, String expectedFormula) {
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.Average(range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the average function on a Range object and a target worksheet")
    @ParameterizedTest(name = "Given worksheet {0} and range {1} should lead to the formula {2}")
    @CsvSource({
            "'worksheet1', 'A1:A1', 'AVERAGE(worksheet1!A1)'",
            "'worksheet1', 'A1:C2', 'AVERAGE(worksheet1!A1:C2)'",
            "'worksheet1', '$A1:C2', 'AVERAGE(worksheet1!$A1:C2)'",
            "'worksheet1', '$A$1:C2', 'AVERAGE(worksheet1!$A$1:C2)'",
            "'worksheet1', '$A$1:$C2', 'AVERAGE(worksheet1!$A$1:$C2)'",
            "'worksheet1', '$A$1:$C$2', 'AVERAGE(worksheet1!$A$1:$C$2)'",
    })
    void averageTest2(String worksheetName, String rangeExpression, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.Average(worksheet, range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the ceil function on a value and a number of decimals")
    @ParameterizedTest(name = "Given Address {0}  and number of decimals {1} should lead to to the formula {2}")
    @CsvSource({
            "'A1', 1, 'ROUNDUP(A1,1)'",
            "'C4', 0, 'ROUNDUP(C4,0)'",
            "'$A1', 10, 'ROUNDUP($A1,10)'",
            "'$A$1', 5, 'ROUNDUP($A$1,5)'",
            "'A1', -2, 'ROUNDUP(A1,-2)'",
            // This seems to be valid
    })
    void ceilTest(String addressExpression, int numberOfDecimals, String expectedFormula) {
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.Ceil(address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the ceil function on a value and a number of decimals")
    @ParameterizedTest(name = "Given worksheet {0}, address {1} and number of decimals {2} should lead to the formula {3}")
    @CsvSource({
            "'worksheet3', 'A1', 1, 'ROUNDUP(worksheet3!A1,1)'",
            "'worksheet3', 'C4', 0, 'ROUNDUP(worksheet3!C4,0)'",
            "'worksheet3', '$A1', 10, 'ROUNDUP(worksheet3!$A1,10)'",
            "'worksheet3', '$A$1', 5, 'ROUNDUP(worksheet3!$A$1,5)'",
            "'worksheet3', 'A1', -2, 'ROUNDUP(worksheet3!A1,-2)'",
            // This seems to be valid
    })
    void ceilTest2(String worksheetName, String addressExpression, int numberOfDecimals, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.Ceil(worksheet, address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the floor function on a value and a number of decimals")
    @ParameterizedTest(name = "Given address {0} and number of decimals {1} should lead to to the formula {2}")
    @CsvSource({
            "'A1', 1, 'ROUNDDOWN(A1,1)'",
            "'C4', 0, 'ROUNDDOWN(C4,0)'",
            "'$A1', 10, 'ROUNDDOWN($A1,10)'",
            "'$A$1', 5, 'ROUNDDOWN($A$1,5)'",
            "'A1', -2, 'ROUNDDOWN(A1,-2)'",
            // This seems to be valid
    })
    void floorTest(String addressExpression, int numberOfDecimals, String expectedFormula) {
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.Floor(address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the floor function on a value and a number of decimals")
    @ParameterizedTest(name = "Given worksheet {0}, address {1} and number of decimals {2} should lead to the formula {3}")
    @CsvSource({
            "'worksheet3', 'A1', 1, 'ROUNDDOWN(worksheet3!A1,1)'",
            "'worksheet3', 'C4', 0, 'ROUNDDOWN(worksheet3!C4,0)'",
            "'worksheet3', '$A1', 10, 'ROUNDDOWN(worksheet3!$A1,10)'",
            "'worksheet3', '$A$1', 5, 'ROUNDDOWN(worksheet3!$A$1,5)'",
            "'worksheet3', 'A1', -2, 'ROUNDDOWN(worksheet3!A1,-2)'",
            // This seems to be valid
    })
    void floorTest2(String worksheetName, String addressExpression, int numberOfDecimals, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.Floor(worksheet, address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the max function on a Range object")
    @ParameterizedTest(name = "Given range {0} should lead to the formula {1}")
    @CsvSource({
            "'A1:A1', 'MAX(A1)'",
            "'A1:C2', 'MAX(A1:C2)'",
            "'$A1:C2', 'MAX($A1:C2)'",
            "'$A$1:C2', 'MAX($A$1:C2)'",
            "'$A$1:$C2', 'MAX($A$1:$C2)'",
            "'$A$1:$C$2', 'MAX($A$1:$C$2)'",
    })
    void maxTest(String rangeExpression, String expectedFormula) {
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.Max(range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the max function on a Range object and a target worksheet")
    @ParameterizedTest(name = "Given worksheet {0} and range {1} should lead to the formula {2}")
    @CsvSource({
            "'worksheet1', 'A1:A1', 'MAX(worksheet1!A1)'",
            "'worksheet1', 'A1:C2', 'MAX(worksheet1!A1:C2)'",
            "'worksheet1', '$A1:C2', 'MAX(worksheet1!$A1:C2)'",
            "'worksheet1', '$A$1:C2', 'MAX(worksheet1!$A$1:C2)'",
            "'worksheet1', '$A$1:$C2', 'MAX(worksheet1!$A$1:$C2)'",
            "'worksheet1', '$A$1:$C$2', 'MAX(worksheet1!$A$1:$C$2)'",
    })
    void maxTest2(String worksheetName, String rangeExpression, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.Max(worksheet, range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the min function on a Range object")
    @ParameterizedTest(name = "Given range {0} should lead to the formula {1}")
    @CsvSource({
            "'A1:A1', 'MIN(A1)'",
            "'A1:C2', 'MIN(A1:C2)'",
            "'$A1:C2', 'MIN($A1:C2)'",
            "'$A$1:C2', 'MIN($A$1:C2)'",
            "'$A$1:$C2', 'MIN($A$1:$C2)'",
            "'$A$1:$C$2', 'MIN($A$1:$C$2)'",
    })
    void minTest(String rangeExpression, String expectedFormula) {
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.Min(range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the min function on a Range object and a target worksheet")
    @ParameterizedTest(name = "Given worksheet {0} and range {1} should lead to the formula {2}")
    @CsvSource({
            "'worksheet1', 'A1:A1', 'MIN(worksheet1!A1)'",
            "'worksheet1', 'A1:C2', 'MIN(worksheet1!A1:C2)'",
            "'worksheet1', '$A1:C2', 'MIN(worksheet1!$A1:C2)'",
            "'worksheet1', '$A$1:C2', 'MIN(worksheet1!$A$1:C2)'",
            "'worksheet1', '$A$1:$C2', 'MIN(worksheet1!$A$1:$C2)'",
            "'worksheet1', '$A$1:$C$2', 'MIN(worksheet1!$A$1:$C$2)'",
    })
    void minTest2(String worksheetName, String rangeExpression, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.Min(worksheet, range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the median function on a Range object")
    @ParameterizedTest(name = "Given range {0} should lead to the formula {1}")
    @CsvSource({
            "'A1:A1', 'MEDIAN(A1)'",
            "'A1:C2', 'MEDIAN(A1:C2)'",
            "'$A1:C2', 'MEDIAN($A1:C2)'",
            "'$A$1:C2', 'MEDIAN($A$1:C2)'",
            "'$A$1:$C2', 'MEDIAN($A$1:$C2)'",
            "'$A$1:$C$2', 'MEDIAN($A$1:$C$2)'",
    })
    void medianTest(String rangeExpression, String expectedFormula) {
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.Median(range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the median function on a Range object and a target worksheet")
    @ParameterizedTest(name = "Given worksheet {0} and range {1} should lead to the formula {2}")
    @CsvSource({
            "'worksheet1', 'A1:A1', 'MEDIAN(worksheet1!A1)'",
            "'worksheet1', 'A1:C2', 'MEDIAN(worksheet1!A1:C2)'",
            "'worksheet1', '$A1:C2', 'MEDIAN(worksheet1!$A1:C2)'",
            "'worksheet1', '$A$1:C2', 'MEDIAN(worksheet1!$A$1:C2)'",
            "'worksheet1', '$A$1:$C2', 'MEDIAN(worksheet1!$A$1:$C2)'",
            "'worksheet1', '$A$1:$C$2', 'MEDIAN(worksheet1!$A$1:$C$2)'",
    })
    void medianTest2(String worksheetName, String rangeExpression, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.Median(worksheet, range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the round function on a value and a number of decimals")
    @ParameterizedTest(name = "Given address {0} and number of decimals {1} should lead to the formula {2}")
    @CsvSource({
            "'A1', 1, 'ROUND(A1,1)'",
            "'C4', 0, 'ROUND(C4,0)'",
            "'$A1', 10, 'ROUND($A1,10)'",
            "'$A$1', 5, 'ROUND($A$1,5)'",
            "'A1', -2, 'ROUND(A1,-2)'", // This seems to be valid
    })
    void roundTest(String addressExpression, int numberOfDecimals, String expectedFormula) {
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.Round(address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the round function on a value and a number of decimals")
    @ParameterizedTest(name = "Given worksheet {0}, address {1} and number of decimals {2} should lead to the formula {3}")
    @CsvSource({
            "'worksheet3', 'A1', 1, 'ROUND(worksheet3!A1,1)'",
            "'worksheet3', 'C4', 0, 'ROUND(worksheet3!C4,0)'",
            "'worksheet3', '$A1', 10, 'ROUND(worksheet3!$A1,10)'",
            "'worksheet3', '$A$1', 5, 'ROUND(worksheet3!$A$1,5)'",
            "'worksheet3', 'A1', -2, 'ROUND(worksheet3!A1,-2)'", // This seems to be valid
    })
    void roundTest2(String worksheetName, String addressExpression, int numberOfDecimals, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Address address = new Address(addressExpression);
        Cell formula = BasicFormulas.Round(worksheet, address, numberOfDecimals);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the sum function on a Range object")
    @ParameterizedTest(name = "Given range {0} should lead to the formula {1}")
    @CsvSource({
            "'A1:A1', 'SUM(A1)'",
            "'A1:C2', 'SUM(A1:C2)'",
            "'$A1:C2', 'SUM($A1:C2)'",
            "'$A$1:C2', 'SUM($A$1:C2)'",
            "'$A$1:$C2', 'SUM($A$1:$C2)'",
            "'$A$1:$C$2', 'SUM($A$1:$C$2)'",
    })
    void sumTest(String rangeExpression, String expectedFormula) {
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.Sum(range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the sum function on a Range object and a target worksheet")
    @ParameterizedTest(name = "Given worksheet {0} and range {1} should lead to the formula {2}")
    @CsvSource({
            "'worksheet1', 'A1:A1', 'SUM(worksheet1!A1)'",
            "'worksheet1', 'A1:C2', 'SUM(worksheet1!A1:C2)'",
            "'worksheet1', '$A1:C2', 'SUM(worksheet1!$A1:C2)'",
            "'worksheet1', '$A$1:C2', 'SUM(worksheet1!$A$1:C2)'",
            "'worksheet1', '$A$1:$C2', 'SUM(worksheet1!$A$1:$C2)'",
            "'worksheet1', '$A$1:$C$2', 'SUM(worksheet1!$A$1:$C$2)'",
    })
    void sumTest2(String worksheetName, String rangeExpression, String expectedFormula) {
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.Sum(worksheet, range);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the vLookup function on a Range object with an arbitrary number, the column index and the option of an exact match")
    @ParameterizedTest(name = "Given value {0}, range {1}, column index {2} and exact match flag = {3} should lead to the formula {4}")
    @CsvSource({
            "'DOUBLE', '11', 'A1:A1', 1, false, 'VLOOKUP(11,A1:A1,1,FALSE)'",
            "'FLOAT', '0.5', 'A1:C4', 3, false, 'VLOOKUP(0.5,A1:C4,3,FALSE)'",
            "'LONG', '-800', 'A10:XFD999999', 200, true, 'VLOOKUP(-800,A10:XFD999999,200,TRUE)'",
            "'BYTE', '0', 'X100:A1', 5, true, 'VLOOKUP(0,A1:X100,5,TRUE)'",
    })
    void vLookupTest(String sourceType, String sourceValue, String rangeExpression, int columnIndex, boolean exactMatch, String expectedFormula) {
        Object number = TestUtils.createInstance(sourceType, sourceValue);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.VLookup(number, range, columnIndex, exactMatch);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the vLookup function on a Range object with an arbitrary number, the column index, the option of an exact match and a target worksheet")
    @ParameterizedTest(name = "Given worksheet {0}, value {2}, range {3}, column index {4} and exact match flag = {5} should lead to the formula {6}")
    @CsvSource({
            "'worksheet1', 'INTEGER', '11', '$A$1:A1', 1, false, 'VLOOKUP(11,worksheet1!$A$1:A1,1,FALSE)'",
            "'worksheet1', 'FLOAT', '0.5', 'A1:$C4', 3, false, 'VLOOKUP(0.5,worksheet1!A1:$C4,3,FALSE)'",
            "'worksheet1', 'DOUBLE', '2.22', '$A10:XFD999999', 200, true, 'VLOOKUP(2.22,worksheet1!$A10:XFD999999,200,TRUE)'",
            "'worksheet1', 'BYTE', '0', 'X100:A1', 5, true, 'VLOOKUP(0,worksheet1!A1:X100,5,TRUE)'",
    })
    void vLookupTest2(String worksheetName, String sourceType, String sourceValue, String rangeExpression, int columnIndex, boolean exactMatch, String expectedFormula) {
        Object number = TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet(worksheetName);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.VLookup(number, worksheet, range, columnIndex, exactMatch);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the vLookup function on a Range object with reference address, the column index and the option of an exact match")
    @ParameterizedTest(name = "Given address {0}, range {1}, column index {2} and exact match flag = {3} should lead to the formula {4}")
    @CsvSource({
            "'C5', 'A1:$A$1', 1, false, 'VLOOKUP(C5,A1:$A$1,1,FALSE)'",
            "'A1', 'A1:C$4', 3, false, 'VLOOKUP(A1,A1:C$4,3,FALSE)'",
            "'$F4', 'A10:XFD999999', 200, true, 'VLOOKUP($F4,A10:XFD999999,200,TRUE)'",
            "'$XFD$99999', 'X100:A1', 5, true,'VLOOKUP($XFD$99999,A1:X100,5,TRUE)'",
    })
    void vLookupTest3(String addressExpression, String rangeExpression, int columnIndex, boolean exactMatch, String expectedFormula) {
        Address address = new Address(addressExpression);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.VLookup(address, range, columnIndex, exactMatch);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the vLookup function on a Range object with reference address, the column index, the option of an exact match and tow target worksheets")
    @ParameterizedTest(name = "Given worksheet {0}, address {1}, range {2}, column index {3} and exact match flag = {4} should lead to the formula {5}")
    @CsvSource({
            "'worksheet1','C5', 'worksheet1', 'A1:$A$1', 1, false, 'VLOOKUP(worksheet1!C5,worksheet1!A1:$A$1,1,FALSE)'",
            "'worksheet2','A1', 'worksheet1', 'A1:C$4', 3, false, 'VLOOKUP(worksheet2!A1,worksheet1!A1:C$4,3,FALSE)'",
            "'worksheet1','$F4', 'worksheet2', 'A10:XFD999999', 200, true, 'VLOOKUP(worksheet1!$F4,worksheet2!A10:XFD999999,200,TRUE)'",
            "'worksheet2','$XFD$99999', 'worksheet2',  'X100:A1', 5, true, 'VLOOKUP(worksheet2!$XFD$99999,worksheet2!A1:X100,5,TRUE)'",
    })
    void vLookupTest4(String valueWorksheetName, String addressExpression, String rangesWorksheetName, String rangeExpression, int columnIndex, boolean exactMatch, String expectedFormula) {
        Worksheet valueWorksheet = new Worksheet(valueWorksheetName);
        Worksheet rangeWorksheet = new Worksheet(rangesWorksheetName);
        Address address = new Address(addressExpression);
        Range range = new Range(rangeExpression);
        Cell formula = BasicFormulas.VLookup(valueWorksheet, address, rangeWorksheet, range, columnIndex, exactMatch);
        assertEquals(expectedFormula, formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

    @DisplayName("Test of the vLookup function for byte / Byte as value")
    @Test()
    void vLookupByteTest() {
        assertVlookup((byte) 0b0, "0");
        assertVlookup((byte) -13, "-13");
        assertVlookup((byte) 15, "15");
        assertVlookup(Byte.valueOf((byte) 0), "0");
        assertVlookup(Byte.valueOf((byte) 15), "15");
    }

    @DisplayName("Test of the vLookup function for BigDecimal as value")
    @Test()
    void vLookupDecimalTest() {
        BigDecimal d1 = new BigDecimal("0");
        BigDecimal d2 = new BigDecimal("-0.0052");
        BigDecimal d3 = new BigDecimal("22.78");
        BigDecimal d4 = new BigDecimal("0.12345678901234567890");
        assertVlookup(d1, "0");
        assertVlookup(d2, "-0.0052");
        assertVlookup(d3, "22.78");
        assertVlookup(d4, "0.123456789012345678");
    }

    @DisplayName("Test of the vLookup function for double / Double as value")
    @Test()
    void vLookupDoubleTest() {
        assertVlookup(0.00d, "0");
        assertVlookup(222.5d, "222.5");
        assertVlookup(-0.101d, "-0.101");
        assertVlookup(Double.valueOf(222.5d), "222.5");
        assertVlookup(Double.valueOf(-0.101d), "-0.101");
    }

    @DisplayName("Test of the vLookup function for float / Float as value")
    @Test()
    void vLookupFloatTest() {
        assertVlookup(0.0f, "0");
        assertVlookup(22.5f, "22.5");
        assertVlookup(-0.01f, "-0.01");
        assertVlookup(Float.valueOf(22.5f), "22.5");
        assertVlookup(Float.valueOf(-0.01f), "-0.01");
    }

    @DisplayName("Test of the vLookup function for int / Integer as value")
    @Test()
    void vLookupIntTest() {
        assertVlookup((int) 0, "0");
        assertVlookup((int) -77, "-77");
        assertVlookup((int) 77, "77");
        assertVlookup(Integer.valueOf(-77), "-77");
        assertVlookup(Integer.valueOf(77), "77");
    }

    @DisplayName("Test of the vLookup function for long / Long as value")
    @Test()
    void vLookupLongTest() {
        assertVlookup(0L, "0");
        assertVlookup(-999999L, "-999999");
        assertVlookup(999999L, "999999");
        assertVlookup(Long.valueOf(-999999L), "-999999");
        assertVlookup(Long.valueOf(999999L), "999999");
    }

    @DisplayName("Test of the vLookup function for short / Short as value")
    @Test()
    void vLookupShortTest() {
        assertVlookup((short) 0, "0");
        assertVlookup((short) -128, "-128");
        assertVlookup((short) 255, "255");
        assertVlookup(Short.valueOf((short) -128), "-128");
        assertVlookup(Short.valueOf((short) 255), "255");
    }

    @DisplayName("Test of the failing vLookup function on an invalid value type")
    @Test()
    void vLookupFailTest() {
        Range range = new Range("A1:D100");
        int column = 2;
        boolean exactMatch = true;
        Long lValue = null; // Differentiates from null rage object
        assertThrows(FormatException.class, () -> BasicFormulas.VLookup("test", range, column, exactMatch));
        assertThrows(FormatException.class, () -> BasicFormulas.VLookup(false, range, column, exactMatch));
        assertThrows(FormatException.class, () -> BasicFormulas.VLookup(null, range, column, exactMatch));
        assertThrows(FormatException.class, () -> BasicFormulas.VLookup(lValue, range, column, exactMatch));
        assertThrows(FormatException.class, () -> BasicFormulas.VLookup(new Date(), range, column, exactMatch));
    }


    @DisplayName("Test of the failing vLookup function on an invalid index column")
    @Test()
    void vLookupFailTest2() {
        Range range = new Range("A1:D100");
        Range range2 = new Range("C1:D100");
        assertThrows(FormatException.class, () -> BasicFormulas.VLookup(22, range, 0, true));
        assertThrows(FormatException.class, () -> BasicFormulas.VLookup(22, range, -2, false));
        assertThrows(FormatException.class, () -> BasicFormulas.VLookup(22, range, 100, true));
        assertThrows(FormatException.class, () -> BasicFormulas.VLookup(22, range2, 3, true));
        assertThrows(FormatException.class, () -> BasicFormulas.VLookup(22, range2, 4, false));
    }

    private void assertVlookup(Object number, String expectedLookupValue) {
        Range range = new Range("A1:D100");
        int column = 2;
        boolean exactMatch = true;
        Cell formula = BasicFormulas.VLookup(number, range, column, exactMatch);
        StringBuilder sb = new StringBuilder();
        sb.append("VLOOKUP(")
                .append(expectedLookupValue)
                .append(",")
                .append(range)
                .append(",")
                .append(column)
                .append(",")
                .append(Boolean.toString(exactMatch).toUpperCase())
                .append(")");
        assertEquals(sb.toString(), formula.getValue().toString());
        assertEquals(Cell.CellType.FORMULA, formula.getDataType());
    }

}
