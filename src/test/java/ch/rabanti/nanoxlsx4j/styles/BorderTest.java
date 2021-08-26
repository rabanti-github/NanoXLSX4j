package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.TestUtils;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.*;

public class BorderTest {

    private Border exampleStyle;

    public BorderTest()
    {
        exampleStyle = new Border();
        exampleStyle.setBottomColor("11001100");
        exampleStyle.setBottomStyle(Border.StyleValue.dashDot);
        exampleStyle.setDiagonalColor("8877AA00");
        exampleStyle.setDiagonalDown(true);
        exampleStyle.setDiagonalStyle(Border.StyleValue.thick);
        exampleStyle.setDiagonalUp(true);
        exampleStyle.setLeftColor("9911DD00");
        exampleStyle.setLeftStyle(Border.StyleValue.mediumDashDotDot);
        exampleStyle.setRightColor("FF00AA00");
        exampleStyle.setRightStyle(Border.StyleValue.dashDotDot);
        exampleStyle.setTopColor("22222200");
        exampleStyle.setTopStyle(Border.StyleValue.dashed);
    }

    @DisplayName("Test of the get and set function of the BottomColor property")
    @ParameterizedTest(name = "Given value {1} should not lead to an exception")
    @CsvSource({
            "STRING, ''",
            "NULL, ''",
            "STRING, 'FFAA3300'",
    })
    void bottomColorTest(String sourceType, String sourceValue)
    {
        String value = (String)TestUtils.createInstance(sourceType, sourceValue);
        Border border = new Border();
        assertEquals(0, border.getBottomColor().length());
        border.setBottomColor(value);
        assertEquals(value, border.getBottomColor());
    }

    @DisplayName("Test of the failing set function of the BottomColor property with invalid values")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "77BB00",
            "0002200000",
            "XXXXXXXX",
    })
    void bottomColorFailTest(String value)
    {
        Border border = new Border();
        assertThrows(Exception.class, () -> border.setBottomColor(value));
    }

    @DisplayName("Test of the get and set function of the BottomColor property")
    @ParameterizedTest(name = "Given value {0} should lead to the defined filed value")
    @CsvSource({
            "dashDot",
            "dashDotDot",
            "dashed",
            "dotted",
            "hair",
            "medium",
            "mediumDashDot",
            "mediumDashDotDot",
            "mediumDashed",
            "none",
            "slantDashDot",
            "s_double",
            "thick",
            "thin",
    })
    void bottomStyleTest(Border.StyleValue value)
    {
        Border border = new Border();
        assertEquals(Border.DEFAULT_BORDER_STYLE, border.getBottomStyle()); // none is default
        border.setBottomStyle(value);
        assertEquals(value, border.getBottomStyle());
    }

    @DisplayName("Test of the get and set function of the DiagonalColor property")
    @ParameterizedTest(name = "Given value {1} should not lead to an exception")
    @CsvSource({
            "STRING, ''",
            "NULL, ''",
            "STRING, 'FFAA3300'",
    })
    void diagonalColorTest(String sourceType, String sourceValue)
    {
        String value = (String)TestUtils.createInstance(sourceType, sourceValue);
        Border border = new Border();
        assertEquals(0, border.getDiagonalColor().length());
        border.setDiagonalColor(value);
        assertEquals(value, border.getDiagonalColor());
    }

    @DisplayName("Test of the failing set function of the DiagonalColor property with invalid values")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "77BB00",
            "0002200000",
            "XXXXXXXX",
    })
    void diagonalColorFailTest(String value)
    {
        Border border = new Border();
        assertThrows(Exception.class, () -> border.setDiagonalColor(value));
    }

    @DisplayName("Test of the get and set function of the BottomColor property")
    @ParameterizedTest(name = "Given value {0} should lead to the defined filed value")
    @CsvSource({
            "dashDot",
            "dashDotDot",
            "dashed",
            "dotted",
            "hair",
            "medium",
            "mediumDashDot",
            "mediumDashDotDot",
            "mediumDashed",
            "none",
            "slantDashDot",
            "s_double",
            "thick",
            "thin",
    })
    void diagonalStyleTest(Border.StyleValue value)
    {
        Border border = new Border();
        assertEquals(Border.DEFAULT_BORDER_STYLE, border.getDiagonalStyle()); // none is default
        border.setDiagonalStyle(value);
        assertEquals(value, border.getDiagonalStyle());
    }

    @DisplayName("Test of the get and set function of the LeftColor property")
    @ParameterizedTest(name = "Given value {1} should not lead to an exception")
    @CsvSource({
            "STRING, ''",
            "NULL, ''",
            "STRING, 'FFAA3300'",
    })
    void leftColorTest(String sourceType, String sourceValue)
    {
        String value = (String)TestUtils.createInstance(sourceType, sourceValue);
        Border border = new Border();
        assertEquals(0, border.getLeftColor().length());
        border.setLeftColor(value);
        assertEquals(value, border.getLeftColor());
    }

    @DisplayName("Test of the failing set function of the LeftColor property with invalid values")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "77BB00",
            "0002200000",
            "XXXXXXXX",
    })
    void leftColorFailTest(String value)
    {
        Border border = new Border();
        assertThrows(Exception.class, () -> border.setLeftColor(value));
    }

    @DisplayName("Test of the get and set function of the LeftColor property")
    @ParameterizedTest(name = "Given value {0} should lead to the defined filed value")
    @CsvSource({
            "dashDot",
            "dashDotDot",
            "dashed",
            "dotted",
            "hair",
            "medium",
            "mediumDashDot",
            "mediumDashDotDot",
            "mediumDashed",
            "none",
            "slantDashDot",
            "s_double",
            "thick",
            "thin",
    })
    void leftStyleTest(Border.StyleValue value)
    {
        Border border = new Border();
        assertEquals(Border.DEFAULT_BORDER_STYLE, border.getLeftStyle()); // none is default
        border.setLeftStyle(value);
        assertEquals(value, border.getLeftStyle());
    }

    @DisplayName("Test of the get and set function of the RightColor property")
    @ParameterizedTest(name = "Given value {1} should not lead to an exception")
    @CsvSource({
            "STRING, ''",
            "NULL, ''",
            "STRING, 'FFAA3300'",
    })
    void rightColorTest(String sourceType, String sourceValue)
    {
        String value = (String)TestUtils.createInstance(sourceType, sourceValue);
        Border border = new Border();
        assertEquals(0, border.getRightColor().length());
        border.setRightColor(value);
        assertEquals(value, border.getRightColor());
    }

    @DisplayName("Test of the failing set function of the RightColor property with invalid values")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "77BB00",
            "0002200000",
            "XXXXXXXX",
    })
    void rightColorFailTest(String value)
    {
        Border border = new Border();
        assertThrows(Exception.class, () -> border.setRightColor(value));
    }

    @DisplayName("Test of the get and set function of the RightStyle property")
    @ParameterizedTest(name = "Given value {0} should lead to the defined filed value")
    @CsvSource({
            "dashDot",
            "dashDotDot",
            "dashed",
            "dotted",
            "hair",
            "medium",
            "mediumDashDot",
            "mediumDashDotDot",
            "mediumDashed",
            "none",
            "slantDashDot",
            "s_double",
            "thick",
            "thin",
    })
    void rightStyleTest(Border.StyleValue value)
    {
        Border border = new Border();
        assertEquals(Border.DEFAULT_BORDER_STYLE, border.getRightStyle()); // none is default
        border.setRightStyle(value);
        assertEquals(value, border.getRightStyle());
    }

    @DisplayName("Test of the get and set function of the TopColor property")
    @ParameterizedTest(name = "Given value {1} should not lead to an exception")
    @CsvSource({
            "STRING, ''",
            "NULL, ''",
            "STRING, 'FFAA3300'",
    })
    void topColorTest(String sourceType, String sourceValue)
    {
        String value = (String)TestUtils.createInstance(sourceType, sourceValue);
        Border border = new Border();
        assertEquals(0, border.getTopColor().length());
        border.setTopColor(value);
        assertEquals(value, border.getTopColor());
    }

    @DisplayName("Test of the failing set function of the TopColor property with invalid values")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "77BB00",
            "0002200000",
            "XXXXXXXX",
    })
    void topColorFailTest(String value)
    {
        Border border = new Border();
        assertThrows(Exception.class, () -> border.setTopColor(value));
    }

    @DisplayName("Test of the get and set function of the TopStyle property")
    @ParameterizedTest(name = "Given value {0} should lead to the defined filed value")
    @CsvSource({
            "dashDot",
            "dashDotDot",
            "dashed",
            "dotted",
            "hair",
            "medium",
            "mediumDashDot",
            "mediumDashDotDot",
            "mediumDashed",
            "none",
            "slantDashDot",
            "s_double",
            "thick",
            "thin",
    })
    void topStyleTest(Border.StyleValue value)
    {
        Border border = new Border();
        assertEquals(Border.DEFAULT_BORDER_STYLE, border.getTopStyle()); // none is default
        border.setTopStyle(value);
        assertEquals(value, border.getTopStyle());
    }

    @DisplayName("Test of the CopyBorder function")
    @Test()
    void copyBorderTest()
    {
        Border copy = exampleStyle.copy();
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }


    @DisplayName("Test of the Equals method")
    @Test()
    void equalsTest()
    {
        Border style2 = (Border)exampleStyle.copy();
        assertTrue(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of BottomColor)")
    @Test()
    void equalsTest2()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setBottomColor("");
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of BottomStyle)")
    @Test()
    void equalsTest2b()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setBottomStyle(Border.StyleValue.s_double);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of TopColor)")
    @Test()
    void equalsTest2c()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setTopColor("");
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of TopStyle)")
    @Test()
    void equalsTest2d()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setTopStyle(Border.StyleValue.s_double);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of LeftColor)")
    @Test()
    void equalsTest2e()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setLeftColor("");
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of LeftStyle)")
    @Test()
    void equalsTest2f()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setLeftStyle(Border.StyleValue.s_double);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of RightColor)")
    @Test()
    void equalsTest2g()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setRightColor("");
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of RightStyle)")
    @Test()
    void equalsTest2h()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setRightStyle(Border.StyleValue.s_double);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of DiagonalColor)")
    @Test()
    void equalsTest2i()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setDiagonalColor("");
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of DiagonalStyle)")
    @Test()
    void equalsTest2j()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setDiagonalStyle(Border.StyleValue.s_double);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of DiagonalDown)")
    @Test()
    void equalsTest2k()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setDiagonalDown(false);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality of DiagonalUp)")
    @Test()
    void equalsTest2l()
    {
        Border style2 = (Border)exampleStyle.copy();
        style2.setDiagonalUp(false);
        assertFalse(exampleStyle.equals(style2));
    }


    @DisplayName("Test of the Equals method (inequality on null or different objects)")
    @ParameterizedTest(name = "Given value {1} should lead to an unequal result")
    @CsvSource({
            "NULL, ''",
            "STRING, 'text'",
            "BOOLEAN, 'true'",
    })
    void equalsTest3(String sourceType, String sourceValue)
    {
        Object obj = TestUtils.createInstance(sourceType, sourceValue);
        assertFalse(exampleStyle.equals(obj));
    }

    @DisplayName("Test of the Equals method when the origin object is null or not of the same type")
    @ParameterizedTest(name = "Given value {1} should lead to an unequal result")
    @CsvSource({
            "NULL, ''",
            "BOOLEAN, 'true'",
            "STRING, 'origin'",
    })
    void equalsTest5(String sourceType, String sourceValue)
    {
        Object origin = TestUtils.createInstance(sourceType, sourceValue);
        assertFalse(exampleStyle.equals(origin));
    }

    @DisplayName("Test of the GetHashCode method (equality of two identical objects)")
    @Test()
    void getHashCodeTest()
    {
        Border copy = (Border)exampleStyle.copy();
        copy.setInternalID(99);  // Should not influence
        assertEquals(exampleStyle.hashCode(), copy.hashCode());
    }


    @DisplayName("Test of the GetHashCode method (inequality of two different objects)")
    @Test()
    void getHashCodeTest2()
    {
        Border copy = (Border)exampleStyle.copy();
        copy.setBottomColor("AACCDD00");
        assertNotEquals(exampleStyle.hashCode(), copy.hashCode());
    }

    // For code coverage

    @DisplayName("Test of the ToString function")
    @Test()
    void toStringTest()
    {
        Border border = new Border();
        String s1 = border.toString();
        border.setBottomColor("FFAABBCC");
        assertNotEquals(s1, border.toString()); // An explicit value comparison is probably not sensible
    }

    @DisplayName("Test of the getValue function of the StyleValue enum")
    @Test()
    void styleValueTest(){
        assertEquals(0, Border.StyleValue.none.getValue());
        assertEquals(1, Border.StyleValue.hair.getValue());
        assertEquals(2, Border.StyleValue.dotted.getValue());
        assertEquals(3, Border.StyleValue.dashDotDot.getValue());
        assertEquals(4, Border.StyleValue.dashDot.getValue());
        assertEquals(5, Border.StyleValue.dashed.getValue());
        assertEquals(6, Border.StyleValue.thin.getValue());
        assertEquals(7, Border.StyleValue.mediumDashDotDot.getValue());
        assertEquals(8, Border.StyleValue.slantDashDot.getValue());
        assertEquals(9, Border.StyleValue.mediumDashDot.getValue());
        assertEquals(10, Border.StyleValue.mediumDashed.getValue());
        assertEquals(11, Border.StyleValue.medium.getValue());
        assertEquals(12, Border.StyleValue.thick.getValue());
        assertEquals(13, Border.StyleValue.s_double.getValue());
    }

}
