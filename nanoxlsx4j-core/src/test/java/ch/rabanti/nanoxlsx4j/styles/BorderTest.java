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

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;

class BorderTest {

    private final Border exampleStyle;

    BorderTest() {
        exampleStyle = new Border();
        exampleStyle.setBottomColor("11001100");
        exampleStyle.setBottomStyle(Border.StyleValue.DASH_DOT);
        exampleStyle.setDiagonalColor("8877AA00");
        exampleStyle.setDiagonalDown(true);
        exampleStyle.setDiagonalStyle(Border.StyleValue.THICK);
        exampleStyle.setDiagonalUp(true);
        exampleStyle.setLeftColor("9911DD00");
        exampleStyle.setLeftStyle(Border.StyleValue.MEDIUM_DASH_DOT_DOT);
        exampleStyle.setRightColor("FF00AA00");
        exampleStyle.setRightStyle(Border.StyleValue.DASH_DOT_DOT);
        exampleStyle.setTopColor("22222200");
        exampleStyle.setTopStyle(Border.StyleValue.DASHED);
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the BottomColor property")
    @NullSource // Adds null as test value
    @ValueSource(strings = {"", "FFAA3300"})
    void bottomColorTest(String value) {
        Border border = new Border();
        assertTrue(border.getBottomColor().isEmpty());
        border.setBottomColor(value);
        assertEquals(value, border.getBottomColor());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing set function of the BottomColor property with invalid values")
    @ValueSource(strings = {"77BB00", "0002200000", "XXXXXXXX"})
    void bottomColorFailTest(String value) {
        Border border = new Border();
        StyleException exception = assertThrows(StyleException.class, () -> border.setBottomColor(value));
        assertEquals(StyleException.class, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the BottomStyle property")
    @EnumSource(Border.StyleValue.class) // Adds all enum values
    void bottomStyleTest(Border.StyleValue value) {
        Border border = new Border();
        assertEquals(Border.DEFAULT_BORDER_STYLE, border.getBottomStyle()); // none is default
        border.setBottomStyle(value);
        assertEquals(value, border.getBottomStyle());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the DiagonalColor property")
    @NullSource // Adds null as test value
    @ValueSource(strings = {"", "FFAA3300"})
    void diagonalColorTest(String value) {
        Border border = new Border();
        assertTrue(border.getDiagonalColor().isEmpty());
        border.setDiagonalColor(value);
        assertEquals(value, border.getDiagonalColor());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing set function of the DiagonalColor property with invalid values")
    @ValueSource(strings = {"77BB00", "0002200000", "XXXXXXXX"})
    void diagonalColorFailTest(String value) {
        Border border = new Border();
        StyleException exception = assertThrows(StyleException.class, () -> border.setDiagonalColor(value));
        assertEquals(StyleException.class, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the DiagonalStyle property")
    @EnumSource(Border.StyleValue.class) // Adds all enum values
    void diagonalStyleTest(Border.StyleValue value) {
        Border border = new Border();
        assertEquals(Border.DEFAULT_BORDER_STYLE, border.getDiagonalStyle()); // none is default
        border.setDiagonalStyle(value);
        assertEquals(value, border.getDiagonalStyle());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the LeftColor property")
    @NullSource // Adds null as test value
    @ValueSource(strings = {"", "FFAA3300"})
    void leftColorTest(String value) {
        Border border = new Border();
        assertTrue(border.getLeftColor().isEmpty());
        border.setLeftColor(value);
        assertEquals(value, border.getLeftColor());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing set function of the LeftColor property with invalid values")
    @ValueSource(strings = {"77BB00", "0002200000", "XXXXXXXX"})
    void leftColorFailTest(String value) {
        Border border = new Border();
        StyleException exception = assertThrows(StyleException.class, () -> border.setLeftColor(value));
        assertEquals(StyleException.class, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the LeftColor property")
    @EnumSource(Border.StyleValue.class) // Adds all enum values
    void leftStyleTest(Border.StyleValue value) {
        Border border = new Border();
        assertEquals(Border.DEFAULT_BORDER_STYLE, border.getLeftStyle()); // none is default
        border.setLeftStyle(value);
        assertEquals(value, border.getLeftStyle());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the RightColor property")
    @NullSource // Adds null as test value
    @ValueSource(strings = {"", "FFAA3300"})
    void rightColorTest(String value) {
        Border border = new Border();
        assertTrue(border.getRightColor().isEmpty());
        border.setRightColor(value);
        assertEquals(value, border.getRightColor());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing set function of the RightColor property with invalid values")
    @ValueSource(strings = {"77BB00", "0002200000", "XXXXXXXX"})
    void rightColorFailTest(String value) {
        Border border = new Border();
        StyleException exception = assertThrows(StyleException.class, () -> border.setRightColor(value));
        assertEquals(StyleException.class, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the RightStyle property")
    @EnumSource(Border.StyleValue.class)
    void rightStyleTest(Border.StyleValue value) { // Adds all enum values
        Border border = new Border();
        assertEquals(Border.DEFAULT_BORDER_STYLE, border.getRightStyle()); // none is default
        border.setRightStyle(value);
        assertEquals(value, border.getRightStyle());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the TopColor property")
    @NullSource // Adds null as test value
    @ValueSource(strings = {"", "FFAA3300"})
    void topColorTest(String value) {
        Border border = new Border();
        assertTrue(border.getTopColor().isEmpty());
        border.setTopColor(value);
        assertEquals(value, border.getTopColor());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing set function of the TopColor property with invalid values")
    @ValueSource(strings = {"77BB00", "0002200000", "XXXXXXXX"})
    void topColorFailTest(String value) {
        Border border = new Border();
        StyleException exception = assertThrows(StyleException.class, () -> border.setTopColor(value));
        assertEquals(StyleException.class, exception.getClass());
    }

    @ParameterizedTest
    @DisplayName("Test of the get and set function of the TopStyle property")
    @EnumSource(Border.StyleValue.class) // Adds all enum values
    void topStyleTest(Border.StyleValue value) {
        Border border = new Border();
        assertEquals(Border.DEFAULT_BORDER_STYLE, border.getTopStyle()); // none is default
        border.setTopStyle(value);
        assertEquals(value, border.getTopStyle());
    }

    @Test
    @DisplayName("Test of the CopyBorder function")
    void copyBorderTest() {
        Border copy = exampleStyle.copyBorder();
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @Test
    @DisplayName("Test of the Equals method")
    void equalsTest() {
        Border style2 = (Border) exampleStyle.copy();
        assertTrue(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of BottomColor)")
    void equalsTest2() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setBottomColor("");
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of BottomStyle)")
    void equalsTest2b() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setBottomStyle(Border.StyleValue.DOUBLE);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of TopColor)")
    void equalsTest2c() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setTopColor("");
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of TopStyle)")
    void equalsTest2d() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setTopStyle(Border.StyleValue.DOUBLE);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of LeftColor)")
    void equalsTest2e() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setLeftColor("");
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of LeftStyle)")
    void equalsTest2f() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setLeftStyle(Border.StyleValue.DOUBLE);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of RightColor)")
    void equalsTest2g() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setRightColor("");
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of RightStyle)")
    void equalsTest2h() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setRightStyle(Border.StyleValue.DOUBLE);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of DiagonalColor)")
    void equalsTest2i() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setDiagonalColor("");
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of DiagonalStyle)")
    void equalsTest2j() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setDiagonalStyle(Border.StyleValue.DOUBLE);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of DiagonalDown)")
    void equalsTest2k() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setDiagonalDown(false);
        assertFalse(exampleStyle.equals(style2));
    }

    @Test
    @DisplayName("Test of the Equals method (inequality of DiagonalUp)")
    void equalsTest2l() {
        Border style2 = (Border) exampleStyle.copy();
        style2.setDiagonalUp(false);
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
        Border copy = (Border) exampleStyle.copy();
        assertFalse(copy.equals(origin));
    }

    @Test
    @DisplayName("Test of the GetHashCode method (equality of two identical objects)")
    void getHashCodeTest() {
        Border copy = (Border) exampleStyle.copy();
        copy.setInternalId(Optional.of(99)); // Should not influence
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @Test
    @DisplayName("Test of the GetHashCode method (inequality of two different objects)")
    void getHashCodeTest2() {
        Border copy = (Border) exampleStyle.copy();
        copy.setBottomColor("AACCDD00");
        assertNotEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    @Test
    @DisplayName("Test of the CompareTo method")
    void compareToTest() {
        Border border = new Border();
        Border other = new Border();
        border.setInternalId(Optional.empty());
        other.setInternalId(Optional.empty());
        assertEquals(-1, border.compareTo(other));
        border.setInternalId(Optional.of(5));
        assertEquals(1, border.compareTo(other));
        assertEquals(1, border.compareTo(null));
        other.setInternalId(Optional.of(5));
        assertEquals(0, border.compareTo(other));
        other.setInternalId(Optional.of(4));
        assertEquals(1, border.compareTo(other));
        other.setInternalId(Optional.of(6));
        assertEquals(-1, border.compareTo(other));
    }

    // For code coverage
    @Test
    @DisplayName("Test of the ToString function")
    void toStringTest() {
        Border border = new Border();
        String s1 = border.toString();
        border.setBottomColor("FFAABBCC");
        assertNotEquals(s1, border.toString()); // An explicit value comparison is probably not sensible
    }

    private static Stream<Arguments> differentObjects() {
        return Stream.of(Arguments.of((Object) null), Arguments.of("text"), Arguments.of(true));
    }

    private static Stream<Arguments> differentOriginObjects() {
        return Stream.of(Arguments.of((Object) null), Arguments.of(true), Arguments.of("origin"));
    }
}
