package ch.rabanti.nanoxlsx4j.cells;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.TestUtils;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

public class RangeTest {

    @DisplayName("Test of the Range constructor with start and end address")
    @ParameterizedTest(name = "Given start address {0} and end address {1} should lead to the range {2}")
    @CsvSource(
            {
                    "A1, A1, A1:A1",
                    "A1, C4, A1:C4",
                    "C3, A1, A1:C3",
                    "$A1, $A$2, $A1:$A$2",
                    "A$1, C$4, A$1:C$4",
                    "$C$3, $A1, $A1:$C$3",}
    )
    void constructorTest(String startAddress, String endAddress, String expectedRange) {
        Address start = new Address(startAddress);
        Address end = new Address(endAddress);
        Range range = new Range(start, end);
        assertEquals(expectedRange, range.toString());
    }

    @DisplayName("Test of the Range constructor with range expression string")
    @ParameterizedTest(name = "Given range expression {0} should lead to should lead to the range {1}")
    @CsvSource(
            {
                    "A1:A1, A1:A1",
                    "c2:C3, C2:C3",
                    "$A1:$F10, $A1:$F10",
                    "$r$1:$b$2, $B$2:$R$1",}
    )
    void constructorTest2(String rangeExpression, String expectedRange) {
        Range range = new Range(rangeExpression);
        assertEquals(expectedRange, range.toString());
    }

    @DisplayName("Test of the ResolveEnclosedAddressesTest method")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource(
            {
                    "A1:A1, 'A1'",
                    "A1:A4, 'A1,A2,A3,A4'",
                    "A1:B3, 'A1,A2,A3,B1,B2,B3'",
                    "B3:A2, 'A2,A3,B2,B3'",}
    )
    void resolveEnclosedAddressesTest(String rangeExpression, String expectedAddresses) {
        Range range = new Range(rangeExpression);
        List<Address> addresses = range.resolveEnclosedAddresses();
        TestUtils.assertCellRange(expectedAddresses, addresses);
    }

    @DisplayName("Test of the equals() method")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource(
            {
                    "A1:A1, A1:A1, true",
                    "A1:A4, A$1:A$4, false",
                    "A1:B3, A1:B4, false",
                    "B3:A2, A2:B3, true",
                    "B$3:A2, A2:B$3, true",}
    )
    void equalsTest(String rangeExpression1, String rangeExpression2, boolean expectedEquality) {
        Range range1 = new Range(rangeExpression1);
        Range range2 = new Range(rangeExpression2);
        boolean result = range1.equals(range2);
        assertEquals(expectedEquality, result);
        assertEquals(range1, range1); // Self-test
    }

    @DisplayName("Test of the equals() method returning false on invalid values")
    @Test()
    void equalsTest2() {
        Range range1 = new Range("A1:A7");
        boolean result = range1.equals(null);
        assertFalse(result);
        result = range1.equals("Wrong type");
        assertFalse(result);
    }

    @DisplayName("Test of the hashCode() method")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource(
            {
                    "A1:A1,  A1:A1, true",
                    "A1:A4, A$1:A$4, false",
                    "A1:B3, A1:B4, false",
                    "B3:A2, A2:B3, true",
                    "B$3:A2, A2:B$3, true",}
    )
    void hashCodeTest(String rangeExpression1, String rangeExpression2, boolean expectedEquality) {
        Range range1 = new Range(rangeExpression1);
        Range range2 = new Range(rangeExpression2);
        if (expectedEquality) {
            assertEquals(range1.hashCode(), range2.hashCode());
        }
        else {
            assertNotEquals(range1.hashCode(), range2.hashCode());
        }
    }

}
