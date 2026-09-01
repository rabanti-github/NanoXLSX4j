/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Optional;
import java.util.stream.Stream;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.ValueSource;

class CellXfTest {

    private final CellXf exampleStyle;

    CellXfTest() {
        exampleStyle = new CellXf();
        exampleStyle.setHidden(true);
        exampleStyle.setLocked(true);
        exampleStyle.setForceApplyAlignment(true);
        exampleStyle.setHorizontalAlign(CellXf.HorizontalAlignValue.LEFT);
        exampleStyle.setVerticalAlign(CellXf.VerticalAlignValue.CENTER);
        exampleStyle.setTextDirection(CellXf.TextDirectionValue.HORIZONTAL);
        exampleStyle.setAlignment(CellXf.TextBreakValue.SHRINK_TO_FIT);
        exampleStyle.setTextRotation(75);
        exampleStyle.setIndent(3);
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Hidden property")
    @ValueSource(booleans = {true, false})
    void hiddenTest(boolean value) {
        CellXf cellXf = new CellXf();
        assertFalse(cellXf.isHidden());
        cellXf.setHidden(value);
        assertEquals(value, cellXf.isHidden());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Locked property")
    @ValueSource(booleans = {true, false})
    void lockedTest(boolean value) {
        CellXf cellXf = new CellXf();
        assertTrue(cellXf.isLocked()); // Locked is set to true by default (has no effect until protection is enabled)
        cellXf.setLocked(value);
        assertEquals(value, cellXf.isLocked());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the ForceApplyAlignment property")
    @ValueSource(booleans = {true, false})
    void forceApplyAlignmentTest(boolean value) {
        CellXf cellXf = new CellXf();
        assertFalse(cellXf.isForceApplyAlignment());
        cellXf.setForceApplyAlignment(value);
        assertEquals(value, cellXf.isForceApplyAlignment());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the HorizontalAlign property")
    @CsvSource({
            "CENTER",
            "CENTER_CONTINUOUS",
            "DISTRIBUTED",
            "FILL",
            "GENERAL",
            "JUSTIFY",
            "LEFT",
            "NONE",
            "RIGHT"
    })
    void horizontalAlignTest(CellXf.HorizontalAlignValue value) {
        CellXf cellXf = new CellXf();
        assertEquals(CellXf.DEFAULT_HORIZONTAL_ALIGNMENT, cellXf.getHorizontalAlign()); // none is default
        cellXf.setHorizontalAlign(value);
        assertEquals(value, cellXf.getHorizontalAlign());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the VerticalAlign property")
    @CsvSource({
            "BOTTOM",
            "CENTER",
            "DISTRIBUTED",
            "JUSTIFY",
            "NONE",
            "TOP"
    })
    void verticalAlignTest(CellXf.VerticalAlignValue value) {
        CellXf cellXf = new CellXf();
        assertEquals(CellXf.DEFAULT_VERTICAL_ALIGNMENT, cellXf.getVerticalAlign()); // none is default
        cellXf.setVerticalAlign(value);
        assertEquals(value, cellXf.getVerticalAlign());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the HorizontalAlign property")
    @CsvSource({
            "HORIZONTAL",
            "VERTICAL"
    })
    void textDirectionTest(CellXf.TextDirectionValue value) {
        CellXf cellXf = new CellXf();
        assertEquals(CellXf.DEFAULT_TEXT_DIRECTION, cellXf.getTextDirection()); // horizontal is default
        cellXf.setTextDirection(value);
        assertEquals(value, cellXf.getTextDirection());
        if (value == CellXf.TextDirectionValue.VERTICAL) {
            assertEquals(255, cellXf.getTextRotation());
        }
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the TextRotation property")
    @ValueSource(ints = {0, 33, 90, -33, -90})
    void textRotationTest(int value) {
        CellXf cellXf = new CellXf();
        assertEquals(0, cellXf.getTextRotation()); // 0 is default
        cellXf.setTextRotation(value);
        assertEquals(value, cellXf.getTextRotation());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing get and set function of the TextRotation property on out-of-range values")
    @ValueSource(ints = {91, -91, -360, 360, 720})
    void textRotationFailTest(int value) {
        CellXf cellXf = new CellXf();
        assertEquals(0, cellXf.getTextRotation()); // 0 is default
        assertThrows(FormatException.class, () -> cellXf.setTextRotation(value));
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Align property")
    @CsvSource({
            "NONE",
            "SHRINK_TO_FIT",
            "WRAP_TEXT"
    })
    void alignTest(CellXf.TextBreakValue value) {
        CellXf cellXf = new CellXf();
        assertEquals(CellXf.DEFAULT_ALIGNMENT, cellXf.getAlignment()); // none is default
        cellXf.setAlignment(value);
        assertEquals(value, cellXf.getAlignment());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the Indent property")
    @ValueSource(ints = {0, 1, 99})
    void indentTest(int value) {
        CellXf cellXf = new CellXf();
        assertEquals(0, cellXf.getIndent()); // 0 is default
        cellXf.setIndent(value);
        assertEquals(value, cellXf.getIndent());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing set function of the Indent property when an invalid value was passed")
    @ValueSource(ints = {-1, -999})
    void indentFailTest(int value) {
        StyleException exception = assertThrows(StyleException.class, () -> exampleStyle.setIndent(value));
        assertEquals(StyleException.class, exception.getClass());
    }

    @Test
    @DisplayName("Test of the Equals method")
    void equalsTest() {
        CellXf style2 = (CellXf) exampleStyle.copy();
        assertTrue(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Locked)")
    void equalsTest2() {
        CellXf style2 = (CellXf) exampleStyle.copy();
        style2.setLocked(false);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Hidden)")
    void equalsTest2b() {
        CellXf style2 = (CellXf) exampleStyle.copy();
        style2.setHidden(false);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of HorizontalAlign)")
    void equalsTest2c() {
        CellXf style2 = (CellXf) exampleStyle.copy();
        style2.setHorizontalAlign(CellXf.HorizontalAlignValue.RIGHT);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of VerticalAlign)")
    void equalsTest2d() {
        CellXf style2 = (CellXf) exampleStyle.copy();
        style2.setVerticalAlign(CellXf.VerticalAlignValue.TOP);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of ForceApplyAlignment)")
    void equalsTest2e() {
        CellXf style2 = (CellXf) exampleStyle.copy();
        style2.setForceApplyAlignment(false);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of TextDirection)")
    void equalsTest2f() {
        CellXf style2 = (CellXf) exampleStyle.copy();
        style2.setTextDirection(CellXf.TextDirectionValue.VERTICAL);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of TextRotation)")
    void equalsTest2g() {
        CellXf style2 = (CellXf) exampleStyle.copy();
        style2.setTextRotation(27);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Alignment)")
    void equalsTest2h() {
        CellXf style2 = (CellXf) exampleStyle.copy();
        style2.setAlignment(CellXf.TextBreakValue.NONE);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of Indent)")
    void equalsTest2i() {
        CellXf style2 = (CellXf) exampleStyle.copy();
        style2.setIndent(77);
        assertFalse(exampleStyle.equals(style2));
    }

    @ParameterizedTest
    @DisplayName("Test of the Equals method (inequality on null or different objects)")
    @MethodSource("differentObjects") // see helper method differentObjects()
    void equalsTest3(Object obj) {
        assertFalse(exampleStyle.equals(obj));
    }

    @ParameterizedTest
    @DisplayName("Test of the Equals method when the origin object is null or not of the same type")
    @MethodSource("differentOriginObjects") // see helper method differentOriginObjects()
    void equalsTest5(Object origin) {
        CellXf copy =  (CellXf) exampleStyle.copy();
        assertFalse(copy.equals(origin));
    }

    @Test
    @DisplayName("Test of the GetHashCode method (equality of two identical objects)")
    void getHashCodeTest() {
        CellXf copy = (CellXf) exampleStyle.copy();
        copy.setInternalId(Optional.of(99)); // Should not influence
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @Test
    @DisplayName("Test of the GetHashCode method (inequality of two different objects)")
    void getHashCodeTest2() {
        CellXf copy = (CellXf) exampleStyle.copy();
        copy.setHidden(false);
        assertNotEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @Test
    @DisplayName("Test of the CompareTo method")
    void compareToTest() {
        CellXf cellXf = new CellXf();
        CellXf other = new CellXf();
        cellXf.setInternalId(Optional.empty());
        other.setInternalId(Optional.empty());
        assertEquals(-1, cellXf.compareTo(other));
        cellXf.setInternalId(Optional.of(5));
        assertEquals(1, cellXf.compareTo(other));
        assertEquals(1, cellXf.compareTo(null));
        other.setInternalId(Optional.of(5));
        assertEquals(0, cellXf.compareTo(other));
        other.setInternalId(Optional.of(4));
        assertEquals(1, cellXf.compareTo(other));
        other.setInternalId(Optional.of(6));
        assertEquals(-1, cellXf.compareTo(other));
    }

    // For code coverage
    @Test
    @DisplayName("Test of the ToString function")
    void toStringTest() {
        CellXf cellXf = new CellXf();
        String s1 = cellXf.toString();
        cellXf.setTextRotation(12);
        assertNotEquals(s1, cellXf.toString()); // An explicit value comparison is probably not sensible
    }

    private static Stream<Arguments> differentObjects() {
        return Stream.of(Arguments.of((Object) null), Arguments.of("text"), Arguments.of(true));
    }

    private static Stream<Arguments> differentOriginObjects() {
        return Stream.of(Arguments.of((Object) null), Arguments.of(true), Arguments.of("origin"));
    }
}
