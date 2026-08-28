/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.enums.FormulaError;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.math.BigDecimal;
import java.time.Duration;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class FormulaDataTest {

    @Test
    @DisplayName("Test of the FormulaData default constructor")
    public void formulaDataConstructorTest() {
        FormulaData data = new FormulaData();

        assertEquals(FormulaData.FormulaType.NORMAL, data.getType());
        assertEquals(Cell.CellType.DEFAULT, data.getCachedValueType());
        assertNull(data.getCachedValue());
    }

    @ParameterizedTest
    @MethodSource("formulaDataConstructorTest2Data")
    @DisplayName("Test of cached value type inference in the FormulaData constructor")
    public void formulaDataConstructorTest2(Object givenCachedValue, Cell.CellType expectedType) {
        FormulaData data = new FormulaData("A1", givenCachedValue);

        assertEquals("A1", data.getExpression());
        assertEquals(givenCachedValue, data.getCachedValue());
        assertEquals(expectedType, data.getCachedValueType());
    }

    @Test
    @DisplayName("Test of special cached value type inference in the FormulaData constructor")
    public void formulaDataConstructorTest3() {
        assertEquals(Cell.CellType.NUMBER, new FormulaData("A1", new BigDecimal("1.25")).getCachedValueType());
        Date date = new GregorianCalendar(2026, 6, 30).getTime();
        assertEquals(Cell.CellType.DATE, new FormulaData("A1", date).getCachedValueType());
        assertEquals(Cell.CellType.TIME, new FormulaData("A1", Duration.ofHours(2)).getCachedValueType());
        assertEquals(Cell.CellType.ERROR,
            new FormulaData("A1", FormulaError.DIVISION_BY_ZERO).getCachedValueType());
        assertEquals(Cell.CellType.STRING, new FormulaData("A1", new Object()).getCachedValueType());
    }

    @Test
    @DisplayName("Test that external workbook reference detection follows expression changes")
    public void externalReferenceExpressionChangeTest() {
        FormulaData data = new FormulaData("A1");
        assertFalse(data.hasExternalReferences());

        data.setExpression("[1]Sheet1!A1");
        assertTrue(data.hasExternalReferences());

        data.setExpression("Table1[Column]");
        assertFalse(data.hasExternalReferences());
    }

    @Test
    @DisplayName("Test that copying FormulaData preserves external workbook reference detection")
    public void externalReferenceCopyTest() {
        FormulaData data = new FormulaData("[1]Sheet1!A1");

        FormulaData copy = data.copy();

        assertTrue(copy.hasExternalReferences());
        assertEquals(data, copy);
    }

    @Test
    @DisplayName("Test of the FormulaData Copy function with cached value metadata")
    public void copyTest() {
        FormulaData data = new FormulaData("A1", "0");
        data.setCachedValueType(Cell.CellType.NUMBER);
        data.setFormulaRange("A1:A2");
        data.setMasterCellAddress("A1");
        data.setType(FormulaData.FormulaType.ARRAY);

        FormulaData copy = data.copy();

        assertNotSame(data, copy);
        assertEquals(data, copy);
        assertEquals(Cell.CellType.NUMBER, copy.getCachedValueType());
    }

    @Test
    @DisplayName("Test of FormulaData equality, comparison and hashing with cached value metadata")
    public void cachedValueTypeComparisonTest() {
        FormulaData number = new FormulaData("A1", "0");
        number.setCachedValueType(Cell.CellType.NUMBER);
        FormulaData numberCopy = number.copy();
        FormulaData text = new FormulaData("A1", "0");
        text.setCachedValueType(Cell.CellType.STRING);

        assertTrue(number.equals(numberCopy));
        assertEquals(0, number.compareTo(numberCopy));
        assertEquals(number.hashCode(), numberCopy.hashCode());
        assertFalse(number.equals(text));
        assertNotEquals(0, number.compareTo(text));
        assertNotEquals(number.hashCode(), text.hashCode());
    }

    @Test
    @DisplayName("Test of the FormulaData CompareTo method")
    public void compareToTest() {
        FormulaData data = createFormulaData();

        assertEquals(1, data.compareTo(null));
        assertEquals(0, data.compareTo(data.copy()));
        assertNotEquals(0, data.compareTo(createFormulaData("B1", FormulaData.FormulaType.ARRAY,
            "A1:A2", createDefinedNameReference(), null, Cell.CellType.NUMBER, "A1")));
        assertNotEquals(0, data.compareTo(createFormulaData("A1", FormulaData.FormulaType.SHARED,
            "A1:A2", createDefinedNameReference(), null, Cell.CellType.NUMBER, "A1")));
        assertNotEquals(0, data.compareTo(createFormulaData("A1", FormulaData.FormulaType.ARRAY,
            "A1:A3", createDefinedNameReference(), null, Cell.CellType.NUMBER, "A1")));
        // DefinedName construction is not ported yet; null versus a reference exercises the same FormulaData branch.
        assertNotEquals(0, data.compareTo(createFormulaData("A1", FormulaData.FormulaType.ARRAY,
            "A1:A2", null, null, Cell.CellType.NUMBER, "A1")));
        assertNotEquals(0, data.compareTo(createFormulaData("A1", FormulaData.FormulaType.ARRAY,
            "A1:A2", createDefinedNameReference(), null, Cell.CellType.STRING, "A1")));
        assertNotEquals(0, data.compareTo(createFormulaData("A1", FormulaData.FormulaType.ARRAY,
            "A1:A2", createDefinedNameReference(), 2, Cell.CellType.NUMBER, "A1")));
        assertNotEquals(0, data.compareTo(createFormulaData("A1", FormulaData.FormulaType.ARRAY,
            "A1:A2", createDefinedNameReference(), null, Cell.CellType.NUMBER, "A2")));
    }

    @Test
    @DisplayName("Test of the strongly typed FormulaData Equals method")
    public void equalsFormulaDataTest() {
        FormulaData data = createFormulaData();

        assertFalse(data.equals((FormulaData) null));
        assertTrue(data.equals(data));
        assertTrue(data.equals(data.copy()));
        assertFalse(data.equals(createFormulaData("B1", FormulaData.FormulaType.ARRAY,
            "A1:A2", data.getDefinedNameReference(), null, Cell.CellType.NUMBER, "A1")));
        assertFalse(data.equals(createFormulaData("A1", FormulaData.FormulaType.SHARED,
            "A1:A2", data.getDefinedNameReference(), null, Cell.CellType.NUMBER, "A1")));
        assertFalse(data.equals(createFormulaData("A1", FormulaData.FormulaType.ARRAY,
            "A1:A3", data.getDefinedNameReference(), null, Cell.CellType.NUMBER, "A1")));
        assertFalse(data.equals(createFormulaData("A1", FormulaData.FormulaType.ARRAY,
            "A1:A2", null, null, Cell.CellType.NUMBER, "A1")));
        assertFalse(data.equals(createFormulaData("A1", FormulaData.FormulaType.ARRAY,
            "A1:A2", data.getDefinedNameReference(), 2, Cell.CellType.NUMBER, "A1")));
        assertFalse(data.equals(createFormulaData("A1", FormulaData.FormulaType.ARRAY,
            "A1:A2", data.getDefinedNameReference(), null, Cell.CellType.STRING, "A1")));
        assertFalse(data.equals(createFormulaData("A1", FormulaData.FormulaType.ARRAY,
            "A1:A2", data.getDefinedNameReference(), null, Cell.CellType.NUMBER, "A2")));
    }

    @Test
    @DisplayName("Test of the object FormulaData Equals method")
    public void equalsObjectTest() {
        FormulaData data = createFormulaData();

        assertTrue(data.equals((Object) data.copy()));
        assertFalse(data.equals((Object) null));
        assertFalse(data.equals("Wrong type"));
    }

    @Test
    @DisplayName("Test of the FormulaData GetHashCode method")
    public void getHashCodeTest() {
        FormulaData data = createFormulaData();
        FormulaData copy = data.copy();
        FormulaData empty = new FormulaData();

        assertEquals(data.hashCode(), copy.hashCode());
        assertNotEquals(data.hashCode(), empty.hashCode());
        assertEquals(empty.hashCode(), new FormulaData().hashCode());
    }

    private static Stream<Arguments> formulaDataConstructorTest2Data() {
        // Java has no ushort, uint, or ulong equivalents. Its signed byte covers the C# byte and sbyte cases.
        return Stream.of(
            Arguments.of((Object) null, Cell.CellType.DEFAULT),
            Arguments.of("0", Cell.CellType.STRING),
            Arguments.of('x', Cell.CellType.STRING),
            Arguments.of(true, Cell.CellType.BOOL),
            Arguments.of((byte) 1, Cell.CellType.NUMBER),
            Arguments.of((byte) -1, Cell.CellType.NUMBER),
            Arguments.of((short) -2, Cell.CellType.NUMBER),
            Arguments.of(-3, Cell.CellType.NUMBER),
            Arguments.of((long) -4, Cell.CellType.NUMBER),
            Arguments.of(1.25f, Cell.CellType.NUMBER),
            Arguments.of(2.5d, Cell.CellType.NUMBER)
        );
    }

    private static FormulaData createFormulaData() {
        return createFormulaData("A1", FormulaData.FormulaType.ARRAY, "A1:A2", createDefinedNameReference(), null,
            Cell.CellType.NUMBER, "A1");
    }

    private static FormulaData createFormulaData(
            String expression,
            FormulaData.FormulaType type,
            String formulaRange,
            DefinedName definedName,
            Object cachedValue,
            Cell.CellType cachedValueType,
            String masterCellAddress) {
        FormulaData data = new FormulaData(expression, cachedValue != null ? cachedValue : 1);
        data.setType(type);
        data.setFormulaRange(formulaRange);
        data.setDefinedNameReference(definedName);
        data.setCachedValueType(cachedValueType);
        data.setMasterCellAddress(masterCellAddress);
        return data;
    }

    private static DefinedName createDefinedNameReference() {
        Workbook workbook = new Workbook("Sheet1");
       return new DefinedName(workbook, DefinedName.NameType.FORMULA, "definedName1", "A1+A2", workbook.getWorksheets().get(0));
    }
}
