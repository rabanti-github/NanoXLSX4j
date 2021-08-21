package ch.rabanti.nanoxlsx4j.cells;
import ch.rabanti.nanoxlsx4j.Column;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.*;

public class ColumnTest {

    @DisplayName("Test of the columnAddress field, as well as the constructor and proper modification by getter and setter")
    @ParameterizedTest(name = "Given initial value {0} should lead to {1} and changed value {2} should lead to {3}")
    @CsvSource({
            "A,A,B,B",
            "a, A, b, B",
            "AAB, AAB, A, A",
            "a, A, XFD, XFD",
    })
    void columnAddressTest(String initialValue, String expectedValue, String changedValue, String expectedChangedValue)
    {
        Column column = new Column(initialValue);
        assertEquals(expectedValue, column.getColumnAddress());
        column.setColumnAddress(changedValue);
        assertEquals(expectedChangedValue, column.getColumnAddress());
    }

    @DisplayName("Test of the failing columnAddress field setter")
    @ParameterizedTest(name = "Given value should lead to an exception")
    @CsvSource({
            "STRING, ''",
            "NULL, ",
            "STRING, '4'",
            "STRING, -",
            "STRING, .",
            "STRING, $",
            "STRING, XFE",
    })
    void columnAddressTest2(String sourceType, String sourceValue)
    {
        String value = (String)TestUtils.createInstance(sourceType, sourceValue);
        Column column = new Column("A");
        assertThrows(RangeException.class, () -> column.setColumnAddress(value));
    }

    @DisplayName("Test of the hasAutoFilter field, as well as proper modification by getter and setter")
    @Test()
    void hasAutoFilterTest()
    {
        Column column = new Column("A");
        assertFalse(column.hasAutoFilter());
        column.setAutoFilter(true);
        assertTrue(column.hasAutoFilter());
        column.setAutoFilter(false);
        assertFalse(column.hasAutoFilter());
    }

    @DisplayName("Test of the isHidden field, as well as proper modification by getter and setter")
    @Test()
    void isHiddenTest()
    {
        Column column = new Column("A");
        assertFalse(column.isHidden());
        column.setHidden(true);
        assertTrue(column.isHidden());
        column.setHidden(false);
        assertFalse(column.isHidden());
    }




    @DisplayName("Test of the number field, as well as the constructor and proper modification by getter and setter")
    @ParameterizedTest(name = "Given initial value {0} should lead to {1} and changed value {2} should lead to {3}")
    @CsvSource({
            "0, 0, 1, 1",
            "999, 999, 5, 5",
            "0, 0, 16383, 16383",
    })
    void numberTest(int initialValue, int expectedValue, int changedValue, int expectedChangedValue) {
        Column column = new Column(initialValue);
        assertEquals(expectedValue, column.getNumber());
        column.setNumber(changedValue);
        assertEquals(expectedChangedValue, column.getNumber());
    }

    @DisplayName("Test of the failing number field setter")
    @ParameterizedTest(name = "Given value should lead to an exception")
    @CsvSource({
            "INTEGER, -1",
            "INTEGER, 16384",
    })
    void numberTest2(String sourceType, String sourceValue)
    {
        int value = (int)TestUtils.createInstance(sourceType, sourceValue);
        Column column = new Column(2);
        assertThrows(RangeException.class, () -> column.setNumber(value));
    }

    @DisplayName("Test of the width field, as well as proper modification by getter and setter")
    @ParameterizedTest(name = "Given value {0} should lead to {1}")
    @CsvSource({
            "15f, 15f",
            "11.1f, 11.1f",
            "0f, 0f",
            "255f, 255f",
    })
    void widthTest(float initialValue, float expectedValue)
    {
        Column column = new Column(0);
        assertEquals(Worksheet.DEFAULT_COLUMN_WIDTH, column.getWidth());
        column.setWidth(initialValue);
        assertEquals(expectedValue, column.getWidth());
    }

    @DisplayName("Test of the failing Width property")
    @ParameterizedTest(name = "Given value {0} should lead to an exception")
    @CsvSource({
            "-1f",
            "255.1f",
    })
    void widthTest2(float value)
    {
        Column column = new Column(0);
        assertThrows(RangeException.class, () -> column.setWidth(value));
    }


}
