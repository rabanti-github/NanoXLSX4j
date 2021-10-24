package ch.rabanti.nanoxlsx4j.worksheets;

import ch.rabanti.nanoxlsx4j.*;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static org.junit.jupiter.api.Assertions.*;

public class SetStyleTest {

    public enum RangeRepresentation {
        StringExpression,
        RangeObject
    }

    @DisplayName("Test of the setStyle function on an empty worksheet with a Range object or its string representation")
    @ParameterizedTest(name = "Given range {0} as {1} should lead to the encoded number of empty cells with a style")
    @CsvSource({
            "A1:A1, RangeObject",
            "A1:A5, RangeObject",
            "A1:C1, RangeObject",
            "A1:C3, RangeObject",
            "R17:N22, RangeObject",
            "A1:A1, StringExpression",
            "A1:A5, StringExpression",
            "A1:C1, StringExpression",
            "A1:C3, StringExpression",
            "R17:N22, StringExpression",
    })
    void setStyleTest1(String rangeString, RangeRepresentation rangeRepresentation) {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCells().size());
        Range range = new Range(rangeString);
        if (rangeRepresentation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(range, BasicStyles.BoldItalic());
        } else {
            worksheet.setStyle(rangeString, BasicStyles.BoldItalic());
        }
        List<String> emptyCells = range.resolveEnclosedAddresses().stream().map(i -> i.getAddress()).collect(Collectors.toList());
        assertCellRange(rangeString, BasicStyles.BoldItalic(), worksheet, emptyCells, emptyCells.size());
    }

    @DisplayName("Test of the setStyle function on an empty worksheet with a Range object or its string representation with null as style")
    @ParameterizedTest(name = "Given range {0} as {1} should lead to no new cells")
    @CsvSource({
            "A1:A1, RangeObject",
            "A1:A5, RangeObject",
            "A1:C1, RangeObject",
            "A1:C3, RangeObject",
            "R17:N22, RangeObject",
            "A1:A1, StringExpression",
            "A1:A5, StringExpression",
            "A1:C1, StringExpression",
            "A1:C3, StringExpression",
            "R17:N22, StringExpression",
    })
    void setStyleTest1b(String rangeString, RangeRepresentation rangeRepresentation) {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCells().size());
        Range range = new Range(rangeString);
        if (rangeRepresentation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(range, null);
        } else {
            worksheet.setStyle(rangeString, null);
        }
        assertEquals(0, worksheet.getCells().size()); // Should not create empty cells
    }

    @DisplayName("Test of the setStyle function on a worksheet with existing cells and a Range object or its string representation")
    @ParameterizedTest(name = "Given range {0} with existing value {3} on {1} should lead to a worksheet with these cells are set to null as style and no additional, empty cells")
    @CsvSource({
            "A1:A1, A1, INTEGER, '22', '', RangeObject",
            "A1:A5, A2, BOOLEAN, 'true', 'A1,A3,A4,A5', RangeObject",
            "A1:C1, B1, STRING, 'test', 'A1,C1', RangeObject",
            "A1:C3, B2, FLOAT, '-0.25', 'A1,A2,A3,B1,B3,C1,C2,C3', RangeObject",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', 'R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21', RangeObject",
            "A1:A1, A1, INTEGER, '22', '', StringExpression",
            "A1:A5, A2, BOOLEAN, 'true', 'A1,A3,A4,A5', StringExpression",
            "A1:C1, B1, STRING, 'test', 'A1,C1', StringExpression",
            "A1:C3, B2, FLOAT, '-0.25', 'A1,A2,A3,B1,B3,C1,C2,C3', StringExpression",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', 'R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21', StringExpression",
    })
    void setStyleTest2(String rangeString, String definedCells, String sourceType, String sourceValue, String expectedEmptyCells, RangeRepresentation representation) {
        Object cellValue = TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        addCells(worksheet, cellValue, definedCells);
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Range range = new Range(rangeString);
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(range, BasicStyles.Bold());
        } else {
            worksheet.setStyle(rangeString, BasicStyles.Bold());
        }
        List<String> emptyCells = TestUtils.splitValuesAsList(expectedEmptyCells);
        assertCellRange(rangeString, BasicStyles.Bold(), worksheet, emptyCells, emptyCells.size() + cellCount);
    }

    @DisplayName("Test of the setStyle function on a worksheet with existing cells and a Range object or its string representation with null as style")
    @ParameterizedTest(name = "Given range {0} with existing value {3} on {1} should lead to a worksheet with these cells are set to null as style and no additional, empty cells")
    @CsvSource({
            "A1:A1, A1, INTEGER, '22', RangeObject",
            "A1:A5, A2, BOOLEAN, 'true', RangeObject",
            "A1:C1, B1, STRING, 'test', RangeObject",
            "A1:C3, B2, FLOAT, '-0.25',  RangeObject",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', RangeObject",
            "A1:A1, A1, INTEGER, '22', StringExpression",
            "A1:A5, A2, BOOLEAN, 'true', StringExpression",
            "A1:C1, B1, STRING, 'test',  StringExpression",
            "A1:C3, B2, FLOAT, '-0.25',  StringExpression",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', StringExpression",
    })
    void setStyleTest2b(String rangeString, String definedCells, String sourceType, String sourceValue, RangeRepresentation representation) {
        Object cellValue = TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        addCells(worksheet, cellValue, definedCells);
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Range range = new Range(rangeString);
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(range, null);
        } else {
            worksheet.setStyle(rangeString, null);
        }
        assertRemovedStyles(worksheet, cellCount);
    }

    @DisplayName("Test of the setStyle function on a worksheet with existing cells that have a style defined, and a Range object or its string representation")
    @ParameterizedTest(name = "Given range {0} with existing value {3} on {1} should lead to a worksheet with these cells are set to a style and additional, empty cells")
    @CsvSource({
            "A1:A1, A1, INTEGER, '22', '', RangeObject",
            "A1:A5, A2, BOOLEAN, 'true', 'A1,A3,A4,A5', RangeObject",
            "A1:C1, B1, STRING, 'test', 'A1,C1', RangeObject",
            "A1:C3, B2, FLOAT, '-0.25', 'A1,A2,A3,B1,B3,C1,C2,C3', RangeObject",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', 'R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21', RangeObject",
            "A1:A1, A1, INTEGER, '22', '', StringExpression",
            "A1:A5, A2, BOOLEAN, 'true', 'A1,A3,A4,A5', StringExpression",
            "A1:C1, B1, STRING, 'test', 'A1,C1', StringExpression",
            "A1:C3, B2, FLOAT, '-0.25', 'A1,A2,A3,B1,B3,C1,C2,C3', StringExpression",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', 'R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21', StringExpression",
            "A1:A1, A1, INTEGER, '22', '', RangeObject",
            "A1:A5, A2, BOOLEAN, 'true', 'A1,A3,A4,A5', RangeObject",
            "A1:C1, B1, STRING, 'test', 'A1,C1', RangeObject",
            "A1:C3, B2, FLOAT, '-0.25', 'A1,A2,A3,B1,B3,C1,C2,C3', RangeObject",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', 'R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21', RangeObject",
            "A1:A1, A1, INTEGER, '22', '', StringExpression",
            "A1:A5, A2, BOOLEAN, 'true', 'A1,A3,A4,A5', StringExpression",
            "A1:C1, B1, STRING, 'test', 'A1,C1', StringExpression",
            "A1:C3, B2, FLOAT, '-0.25', 'A1,A2,A3,B1,B3,C1,C2,C3', StringExpression",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', 'R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21', StringExpression",
    })
    void setStyleTest3(String rangeString, String definedCells, String sourceType, String sourceValue, String expectedEmptyCells, RangeRepresentation representation) {
        Object cellValue = TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        addCells(worksheet, cellValue, definedCells, BasicStyles.Italic());
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Range range = new Range(rangeString);
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(range, BasicStyles.BoldItalic());
        } else {
            worksheet.setStyle(rangeString, BasicStyles.BoldItalic());
        }

        List<String> emptyCells = TestUtils.splitValuesAsList(expectedEmptyCells);
        assertCellRange(rangeString, BasicStyles.BoldItalic(), worksheet, emptyCells, emptyCells.size() + cellCount);
    }

    @DisplayName("Test of the setStyle function on a worksheet with existing cells that have a style defined, and a Range object or its string representation with null as style")
    @ParameterizedTest(name = "Given range {0} with existing value {3} on {1} should lead to a worksheet with these cells are set to null as style and no additional, empty cells")
    @CsvSource({
            "A1:A1, A1, INTEGER, '22', RangeObject",
            "A1:A5, A2, BOOLEAN, 'true', RangeObject",
            "A1:C1, B1, STRING, 'test', RangeObject",
            "A1:C3, B2, FLOAT, '-0.25', RangeObject",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', RangeObject",
            "A1:A1, A1, INTEGER, '22', StringExpression",
            "A1:A5, A2, BOOLEAN, 'true', StringExpression",
            "A1:C1, B1, STRING, 'test', StringExpression",
            "A1:C3, B2, FLOAT, '-0.25', StringExpression",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', StringExpression",
            "A1:A1, A1, INTEGER, '22', RangeObject",
            "A1:A5, A2, BOOLEAN, 'true', RangeObject",
            "A1:C1, B1, STRING, 'test', RangeObject",
            "A1:C3, B2, FLOAT, '-0.25', RangeObject",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', RangeObject",
            "A1:A1, A1, INTEGER, '22', StringExpression",
            "A1:A5, A2, BOOLEAN, 'true', StringExpression",
            "A1:C1, B1, STRING, 'test', StringExpression",
            "A1:C3, B2, FLOAT, '-0.25', StringExpression",
            "R17:T21, 'R18,R19,R20,S19', LONG, '99999', StringExpression",
    })
    void setStyleTest3b(String rangeString, String definedCells, String sourceType, String sourceValue, RangeRepresentation representation) {
        Object cellValue = TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        addCells(worksheet, cellValue, definedCells, BasicStyles.Italic());
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Range range = new Range(rangeString);
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(range, null);
        } else {
            worksheet.setStyle(rangeString, null);
        }
        assertRemovedStyles(worksheet, cellCount);
    }

    @DisplayName("Test of the setStyle function on a worksheet with existing date and time cells and a Range object")
    @ParameterizedTest(name = "Representation {0} should lead date and time objects with override styles")
    @CsvSource({
            "RangeObject",
            "StringExpression",
    })
    void setStyleTest4(RangeRepresentation representation) {
        Worksheet worksheet = new Worksheet();
        Calendar calendar = Calendar.getInstance();
        calendar.set(2020, 11, 10, 9, 8, 7);
        addCells(worksheet, calendar.getTime(), "B2");
        addCells(worksheet, LocalTime.of(10, 11, 12), "B3");
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Range range = new Range("A1:C3");
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(range, BasicStyles.BorderFrame());
        } else {
            worksheet.setStyle("A1:C3", BasicStyles.BorderFrame());
        }

        List<String> emptyCells = TestUtils.splitValuesAsList("A1,A2,A3,B1,C1,C2,C3");
        assertCellRange("A1:C3", BasicStyles.BorderFrame(), worksheet, emptyCells, emptyCells.size() + cellCount);
    }

    @DisplayName("Test of the setStyle function on a worksheet with existing date and time cells and a Range object with null as style")
    @ParameterizedTest(name = "Representation {0} should lead date and time objects with null as style (invalid)")
    @CsvSource({
            "RangeObject",
            "StringExpression",
    })
    void setStyleTest4b(RangeRepresentation representation) {
        Worksheet worksheet = new Worksheet();
        Calendar calendar = Calendar.getInstance();
        calendar.set(2020, 11, 10, 9, 8, 7);
        addCells(worksheet, calendar.getTime(), "B2");
        addCells(worksheet, LocalTime.of(10, 11, 12), "B3");
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Range range = new Range("A1:C3");
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(range, null);
        } else {
            worksheet.setStyle("A1:C3", null);
        }
        assertRemovedStyles(worksheet, cellCount);
    }

    @DisplayName("Test of the SetStyle function on an empty worksheet with a start and end address")
    @ParameterizedTest(name = "Given start address {0} and end address {1} should lead to the encoded number of empty cells with a style")
    @CsvSource({
            "A1, A1",
            "A1, A5",
            "A1, C1",
            "A1, C3",
            "R17, N22",
    })
    void setStyleTest5(String startAddressString, String endAddressString) {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCells().size());
        Address startAddress = new Address(startAddressString);
        Address endAddress = new Address(endAddressString);
        worksheet.setStyle(startAddress, endAddress, BasicStyles.BoldItalic());
        Range range = new Range(startAddress, endAddress);
        List<String> emptyCells = range.resolveEnclosedAddresses().stream().map(i -> i.getAddress()).collect(Collectors.toList());
        assertCellRange(range.toString(), BasicStyles.BoldItalic(), worksheet, emptyCells, emptyCells.size());
    }

    @DisplayName("Test of the SetStyle function on an empty worksheet with a start and end address with null as style")
    @ParameterizedTest(name = "Given start address {0} and end address {1} should lead to no empty cells")
    @CsvSource({
            "A1, A1",
            "A1, A5",
            "A1, C1",
            "A1, C3",
            "R17, N22",
    })
    void setStyleTest5b(String startAddressString, String endAddressString) {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCells().size());
        Address startAddress = new Address(startAddressString);
        Address endAddress = new Address(endAddressString);
        worksheet.setStyle(startAddress, endAddress, null);
        assertEquals(0, worksheet.getCells().size()); // Should not create empty cells
    }

    @DisplayName("Test of the SetStyle function on a worksheet with existing cells, and a start and end address")
    @ParameterizedTest(name = "Given start address {0} and end address {1} with existing value {4} on {2} should lead to a worksheet with these cells are set to a style and additional, empty cells")
    @CsvSource({
            "A1, A1, A1, INTEGER, '22', ''",
            "A1, A5, A2, BOOLEAN, 'true', 'A1,A3,A4,A5'",
            "A1, C1, B1, STRING, 'test', 'A1,C1'",
            "A1, C3, B2, FLOAT, '-0.25', 'A1,A2,A3,B1,B3,C1,C2,C3'",
            "R17, T21, 'R18,R19,R20,S19', LONG, '99999', 'R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21'",
    })
    void setStyleTest6(String startAddressString, String endAddressString, String definedCells, String sourceType, String sourceValue, String expectedEmptyCells) {
        Object cellValue = TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        addCells(worksheet, cellValue, definedCells);
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Address startAddress = new Address(startAddressString);
        Address endAddress = new Address(endAddressString);
        worksheet.setStyle(startAddress, endAddress, BasicStyles.Bold());
        List<String> emptyCells = TestUtils.splitValuesAsList(expectedEmptyCells);
        Range range = new Range(startAddress, endAddress);
        assertCellRange(range.toString(), BasicStyles.Bold(), worksheet, emptyCells, emptyCells.size() + cellCount);
    }

    @DisplayName("Test of the SetStyle function on a worksheet with existing cells, and a start and end address with null as style")
    @ParameterizedTest(name = "Given start address {0} and end address {1} with existing value {4} on {2} should lead to a worksheet with these cells are set to null as style and no additional, empty cells")
    @CsvSource({
            "A1, A1, A1, INTEGER, '22'",
            "A1, A5, A2, BOOLEAN, 'true'",
            "A1, C1, B1, STRING, 'test'",
            "A1, C3, B2, FLOAT, '-0.25'",
            "R17, T21, 'R18,R19,R20,S19', LONG, '99999'",
    })
    void setStyleTest6b(String startAddressString, String endAddressString, String definedCells, String sourceType, String sourceValue) {
        Object cellValue = TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        addCells(worksheet, cellValue, definedCells);
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Address startAddress = new Address(startAddressString);
        Address endAddress = new Address(endAddressString);
        worksheet.setStyle(startAddress, endAddress, null);
        assertRemovedStyles(worksheet, cellCount);
    }

    @DisplayName("Test of the SetStyle function on a worksheet with existing cells that have a style defined, and a start and end address")
    @ParameterizedTest(name = "Given start address {0} and end address {1} with existing value {4} on {2} should lead to a worksheet with these cells are set to a style and additional, empty cells")
    @CsvSource({
            "A1, A1, A1, INTEGER, '22', ''",
            "A1, A5, A2, BOOLEAN, 'true', 'A1,A3,A4,A5'",
            "A1, C1, B1, STRING, 'test', 'A1,C1'",
            "A1, C3, B2, FLOAT, '-0.25', 'A1,A2,A3,B1,B3,C1,C2,C3'",
            "R17, T21, 'R18,R19,R20,S19', LONG, '99999', 'R17,R21,S17,S18,S20,S21,T17,T18,T19,T20,T21'",
    })
    void setStyleTest7(String startAddressString, String endAddressString, String definedCells, String sourceType, String sourceValue, String expectedEmptyCells) {
        Object cellValue = TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        addCells(worksheet, cellValue, definedCells, BasicStyles.Italic());
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Address startAddress = new Address(startAddressString);
        Address endAddress = new Address(endAddressString);
        worksheet.setStyle(startAddress, endAddress, BasicStyles.Bold());
        List<String> emptyCells = TestUtils.splitValuesAsList(expectedEmptyCells);
        Range range = new Range(startAddress, endAddress);
        assertCellRange(range.toString(), BasicStyles.Bold(), worksheet, emptyCells, emptyCells.size() + cellCount);
    }

    @DisplayName("Test of the SetStyle function on a worksheet with existing cells that have a style defined, and a start and end address with null as style")
    @ParameterizedTest(name = "Given start address {0} and end address {1} with existing value {4} on {2} should lead to a worksheet with these cells are set to null as style and no additional, empty cells")
    @CsvSource({
            "A1, A1, A1, INTEGER, '22'",
            "A1, A5, A2, BOOLEAN, 'true'",
            "A1, C1, B1, STRING, 'test'",
            "A1, C3, B2, FLOAT, '-0.25'",
            "R17, T21, 'R18,R19,R20,S19', LONG, '99999'",
    })
    void setStyleTest7b(String startAddressString, String endAddressString, String definedCells, String sourceType, String sourceValue) {
        Object cellValue = TestUtils.createInstance(sourceType, sourceValue);
        Worksheet worksheet = new Worksheet();
        addCells(worksheet, cellValue, definedCells, BasicStyles.Italic());
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Address startAddress = new Address(startAddressString);
        Address endAddress = new Address(endAddressString);
        worksheet.setStyle(startAddress, endAddress, null);
        assertRemovedStyles(worksheet, cellCount);
    }


    @DisplayName("Test of the setStyle function on a worksheet with existing date and time cells, and a start and end address")
    @Test()
    void setStyleTest8() {
        Worksheet worksheet = new Worksheet();
        Calendar calendar = Calendar.getInstance();
        calendar.set(2020, 11, 10, 9, 8, 7);
        addCells(worksheet, calendar.getTime(), "B2");
        addCells(worksheet, LocalTime.of(10, 11, 12), "B3");
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Address startAddress = new Address("A1");
        Address endAddress = new Address("C3");
        worksheet.setStyle(startAddress, endAddress, BasicStyles.BorderFrame());
        List<String> emptyCells = TestUtils.splitValuesAsList("A1,A2,A3,B1,C1,C2,C3");
        assertCellRange("A1:C3", BasicStyles.BorderFrame(), worksheet, emptyCells, emptyCells.size() + cellCount);
    }

    @DisplayName("Test of the setStyle function on a worksheet with existing date and time cells, and a start and end address with null as style")
    @Test()
    void setStyleTest8b() {
        Worksheet worksheet = new Worksheet();
        Calendar calendar = Calendar.getInstance();
        calendar.set(2020, 11, 10, 9, 8, 7);
        addCells(worksheet, calendar.getTime(), "B2");
        addCells(worksheet, LocalTime.of(10, 11, 12), "B3");
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        Address startAddress = new Address("A1");
        Address endAddress = new Address("C3");
        worksheet.setStyle(startAddress, endAddress, null);
        assertRemovedStyles(worksheet, cellCount);
    }

    @DisplayName("Test of the SetStyle function on an empty worksheet with a singular address or its string representation")
    @ParameterizedTest(name = "Representation {0} should lead to one override cell")
    @CsvSource({
            "RangeObject",
            "StringExpression",
    })
    void setStyleTest9(RangeRepresentation representation) {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCells().size());
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(new Address("C2"), BasicStyles.BoldItalic());
        } else {
            worksheet.setStyle("C2", BasicStyles.BoldItalic());
        }
        Range range = new Range("C2:C2");
        List<String> emptyCells = range.resolveEnclosedAddresses().stream().map(i -> i.getAddress()).collect(Collectors.toList());
        assertCellRange(range.toString(), BasicStyles.BoldItalic(), worksheet, emptyCells, emptyCells.size());
    }

    @DisplayName("Test of the SetStyle function on an empty worksheet with a singular address or its string representation with null as style")
    @ParameterizedTest(name = "Representation {0} should lead to no new cells")
    @CsvSource({
            "RangeObject",
            "StringExpression",
    })
    void setStyleTest9b(RangeRepresentation representation) {
        Worksheet worksheet = new Worksheet();
        assertEquals(0, worksheet.getCells().size());
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(new Address("C2"), null);
        } else {
            worksheet.setStyle("C2", null);
        }
        assertEquals(0, worksheet.getCells().size());
    }

    @DisplayName("Test of the SetStyle function on a worksheet with existing cells, and a singular address or its string representation")
    @ParameterizedTest(name = "Representation {0} should lead to one override cell")
    @CsvSource({
            "RangeObject",
            "StringExpression",
    })
    void setStyleTest10(RangeRepresentation representation) {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(22, "B2");
        worksheet.addCell(false, "B3");
        worksheet.addCell("test", "B4");
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(new Address("B2"), BasicStyles.Bold());
        } else {
            worksheet.setStyle("B2", BasicStyles.Bold());
        }
        assertCellRange("B2:B2", BasicStyles.Bold(), worksheet, new ArrayList<>(), 3);
    }

    @DisplayName("Test of the SetStyle function on a worksheet with existing cells, and a singular address or its string representation with null as style")
    @ParameterizedTest(name = "Representation {0} should lead to one override cell where the style is set to null")
    @CsvSource({
            "RangeObject",
            "StringExpression",
    })
    void setStyleTest10b(RangeRepresentation representation) {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(22, "B2");
        worksheet.addCell(false, "B3");
        worksheet.addCell("test", "B4");
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(new Address("B2"), null);
        } else {
            worksheet.setStyle("B2", null);
        }
        assertRemovedStyles(worksheet, cellCount);
    }

    @DisplayName("Test of the SetStyle function on a worksheet with existing cells that have a style defined, and a singular address or its string representation")
    @ParameterizedTest(name = "Representation {0} should lead to one override cell")
    @CsvSource({
            "RangeObject",
            "StringExpression",
    })
    void setStyleTest11(RangeRepresentation representation) {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(22, "B2", BasicStyles.Bold());
        worksheet.addCell(false, "B3", BasicStyles.Bold());
        worksheet.addCell("test", "B4", BasicStyles.Bold());
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(new Address("B2"), BasicStyles.Bold());
        } else {
            worksheet.setStyle("B2", BasicStyles.Bold());
        }
        assertCellRange("B2:B2", BasicStyles.Bold(), worksheet, new ArrayList<>(), 3);
    }

    @DisplayName("Test of the SetStyle function on a worksheet with existing cells that have a style defined, and a singular address or its string representation with null as style")
    @ParameterizedTest(name = "Representation {0} should lead to one override cell where the style is set to null")
    @CsvSource({
            "RangeObject",
            "StringExpression",
    })
    void setStyleTest11b(RangeRepresentation representation) {
        Worksheet worksheet = new Worksheet();
        worksheet.addCell(22, "B2", BasicStyles.Bold());
        worksheet.addCell(false, "B3", BasicStyles.Bold());
        worksheet.addCell("test", "B4", BasicStyles.Bold());
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(new Address("B2"), null);
        } else {
            worksheet.setStyle("B2", null);
        }
        assertNull(worksheet.getCells().get("B2").getCellStyle());
        assertTrue(worksheet.getCells().get("B3").getCellStyle().equals(BasicStyles.Bold()));
        assertTrue(worksheet.getCells().get("B4").getCellStyle().equals(BasicStyles.Bold()));
    }

    @DisplayName("Test of the setStyle function on a worksheet with existing date and time cells, and a singular address or its string representation")
    @ParameterizedTest(name = "Representation {0} should lead to one override cell")
    @CsvSource({
            "RangeObject",
            "StringExpression",
    })
    void setStyleTest12(RangeRepresentation representation) {
        Worksheet worksheet = new Worksheet();
        Calendar calendar = Calendar.getInstance();
        calendar.set(2020, 11, 10, 9, 8, 7);
        addCells(worksheet, calendar.getTime(), "B2");
        addCells(worksheet, LocalTime.of(10, 11, 12), "B3");
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(new Address("B2"), BasicStyles.BorderFrame());
            worksheet.setStyle(new Address("B3"), BasicStyles.BorderFrame());
        } else {
            worksheet.setStyle("B2", BasicStyles.BorderFrame());
            worksheet.setStyle("B3", BasicStyles.BorderFrame());
        }
        assertCellRange("B2:B3", BasicStyles.BorderFrame(), worksheet, new ArrayList<>(), cellCount);
    }

    @DisplayName("Test of the setStyle function on a worksheet with existing date and time cells, and a singular address or its string representation with null as style")
    @ParameterizedTest(name = "Representation {0} should lead to one override cell where the style is set to null")
    @CsvSource({
            "RangeObject",
            "StringExpression",
    })
    void setStyleTest12b(RangeRepresentation representation) {
        Worksheet worksheet = new Worksheet();
        Calendar calendar = Calendar.getInstance();
        calendar.set(2020, 11, 10, 9, 8, 7);
        addCells(worksheet, calendar.getTime(), "B2");
        addCells(worksheet, LocalTime.of(10, 11, 12), "B3");
        int cellCount = worksheet.getCells().size();
        assertNotEquals(0, cellCount);
        if (representation == RangeRepresentation.RangeObject) {
            worksheet.setStyle(new Address("B2"), null);
            worksheet.setStyle(new Address("B3"), null);
        } else {
            worksheet.setStyle("B2", null);
            worksheet.setStyle("B3", null);
        }
        assertRemovedStyles(worksheet, cellCount);
    }

    @Test()
    @DisplayName("Test of the failing setStyle function, when no range was defined")
    void setStyleFailTest() {
        Worksheet worksheet = new Worksheet();
        Range range = null;
        assertThrows(RangeException.class, () -> worksheet.setStyle(range, BasicStyles.Bold()));
    }

    @Test()
    @DisplayName("Test of the failing setStyle function, when no range as string was defined")
    void setStyleFailTest2() {
        Worksheet worksheet = new Worksheet();
        String range = null;
        assertThrows(FormatException.class, () -> worksheet.setStyle(range, BasicStyles.Bold()));
    }

    private void assertCellRange(String range, Style expectedStyle, Worksheet worksheet, List<String> createdCells, int expectedSize) {
        assertEquals(expectedSize, worksheet.getCells().size());
        Range setRange = new Range(range);
        for (Address address : setRange.resolveEnclosedAddresses()) {
            assertTrue(worksheet.getCells().containsKey(address.getAddress()));
            if (expectedStyle == null) {
                assertNull(worksheet.getCells().get(address.getAddress()).getCellStyle());
            } else {
                assertTrue(expectedStyle.equals(worksheet.getCells().get(address.getAddress()).getCellStyle()));
            }
            if (createdCells != null && createdCells.contains(address.getAddress())) {
                assertEquals(Cell.CellType.EMPTY, worksheet.getCells().get(address.getAddress()).getDataType());
            }
        }
    }

    private void assertRemovedStyles(Worksheet worksheet, int expectedSize) {
        assertEquals(expectedSize, worksheet.getCells().size());
        for (Map.Entry<String, Cell> cell : worksheet.getCells().entrySet()) {
            assertNull(cell.getValue().getCellStyle());
        }
    }

    private static void addCells(Worksheet worksheet, Object sample, String addressString) {
        addCells(worksheet, sample, addressString, null);
    }

    private static void addCells(Worksheet worksheet, Object sample, String addressString, Style style) {
        List<String> addresses = TestUtils.splitValuesAsList(addressString);
        for (String address : addresses) {
            Cell cell = new Cell(sample, Cell.CellType.DEFAULT);
            worksheet.addCell(cell, address, style);
        }
    }
}
