package ch.rabanti.nanoxlsx4j.colors;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.ValueSource;

class IndexedColorTest {

    @ParameterizedTest
    @DisplayName("Test of the getter and setter of the ColorValue property on valid values")
    @EnumSource(IndexedColor.Value.class)
    void colorValueTest(IndexedColor.Value value) {
        IndexedColor color = new IndexedColor();
        assertEquals(IndexedColor.DEFAULT_INDEXED_COLOR, color.getColorValue());
        color.setColorValue(value);
        assertEquals(value, color.getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the getter of the StringValue property")
    @CsvSource({
            "BLACK_0, 0",
            "WHITE_1, 1",
            "RED_2, 2",
            "BRIGHT_GREEN_3, 3",
            "BLUE_4, 4",
            "YELLOW_5, 5",
            "MAGENTA_6, 6",
            "CYAN_7, 7",
            "BLACK, 8",
            "WHITE, 9",
            "RED, 10",
            "BRIGHT_GREEN, 11",
            "BLUE, 12",
            "YELLOW, 13",
            "MAGENTA, 14",
            "CYAN, 15",
            "DARK_RED, 16",
            "DARK_GREEN, 17",
            "DARK_BLUE, 18",
            "OLIVE, 19",
            "PURPLE, 20",
            "TEAL, 21",
            "LIGHT_GRAY, 22",
            "GRAY, 23",
            "LIGHT_CORNFLOWER_BLUE, 24",
            "DARK_ROSE, 25",
            "LIGHT_YELLOW, 26",
            "LIGHT_CYAN, 27",
            "DARK_PURPLE, 28",
            "SALMON, 29",
            "MEDIUM_BLUE, 30",
            "LIGHT_LAVENDER, 31",
            "NAVY, 32",
            "STRONG_MAGENTA, 33",
            "STRONG_YELLOW, 34",
            "STRONG_CYAN, 35",
            "DARK_VIOLET, 36",
            "DARK_MAROON, 37",
            "DARK_TEAL, 38",
            "PURE_BLUE, 39",
            "SKY_BLUE, 40",
            "PALE_CYAN, 41",
            "LIGHT_MINT, 42",
            "PASTEL_YELLOW, 43",
            "LIGHT_SKY_BLUE, 44",
            "ROSE, 45",
            "LAVENDER, 46",
            "PEACH, 47",
            "ROYAL_BLUE, 48",
            "TURQUOISE, 49",
            "LIGHT_OLIVE, 50",
            "GOLD, 51",
            "ORANGE, 52",
            "DARK_ORANGE, 53",
            "BLUE_GRAY, 54",
            "MEDIUM_GRAY, 55",
            "DARK_SLATE_BLUE, 56",
            "SEA_GREEN, 57",
            "VERY_DARK_GREEN, 58",
            "DARK_OLIVE, 59",
            "BROWN, 60",
            "DARK_ROSE_DUPLICATE, 61",
            "INDIGO, 62",
            "VERY_DARK_GRAY, 63",
            "SYSTEM_FOREGROUND, 64",
            "SYSTEM_BACKGROUND, 65"
    })
    void stringValueTest(IndexedColor.Value givenValue, String expectedValue) {
        IndexedColor color = new IndexedColor();
        assertEquals(IndexedColor.DEFAULT_INDEXED_COLOR, color.getColorValue());
        color.setColorValue(givenValue);
        assertEquals(expectedValue, color.getStringValue());
    }

    @Test
    @DisplayName("Test of the default Constructor")
    void constructorTest() {
        IndexedColor color = new IndexedColor();
        assertEquals(IndexedColor.DEFAULT_INDEXED_COLOR, color.getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the Constructor with an enum value")
    @EnumSource(value = IndexedColor.Value.class, names = {
            "BLACK_0", "WHITE_1", "RED_2", "BRIGHT_GREEN_3", "SYSTEM_BACKGROUND", "SYSTEM_FOREGROUND"
    })
    void constructorTest2(IndexedColor.Value value) {
        IndexedColor color = new IndexedColor(value);
        assertEquals(value, color.getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the Constructor with an index")
    @CsvSource({
            "0, BLACK_0",
            "1, WHITE_1",
            "2, RED_2",
            "3, BRIGHT_GREEN_3",
            "4, BLUE_4",
            "5, YELLOW_5",
            "6, MAGENTA_6",
            "7, CYAN_7",
            "8, BLACK",
            "9, WHITE",
            "10, RED",
            "11, BRIGHT_GREEN",
            "12, BLUE",
            "13, YELLOW",
            "14, MAGENTA",
            "15, CYAN",
            "16, DARK_RED",
            "17, DARK_GREEN",
            "18, DARK_BLUE",
            "19, OLIVE",
            "20, PURPLE",
            "21, TEAL",
            "22, LIGHT_GRAY",
            "23, GRAY",
            "24, LIGHT_CORNFLOWER_BLUE",
            "25, DARK_ROSE",
            "26, LIGHT_YELLOW",
            "27, LIGHT_CYAN",
            "28, DARK_PURPLE",
            "29, SALMON",
            "30, MEDIUM_BLUE",
            "31, LIGHT_LAVENDER",
            "32, NAVY",
            "33, STRONG_MAGENTA",
            "34, STRONG_YELLOW",
            "35, STRONG_CYAN",
            "36, DARK_VIOLET",
            "37, DARK_MAROON",
            "38, DARK_TEAL",
            "39, PURE_BLUE",
            "40, SKY_BLUE",
            "41, PALE_CYAN",
            "42, LIGHT_MINT",
            "43, PASTEL_YELLOW",
            "44, LIGHT_SKY_BLUE",
            "45, ROSE",
            "46, LAVENDER",
            "47, PEACH",
            "48, ROYAL_BLUE",
            "49, TURQUOISE",
            "50, LIGHT_OLIVE",
            "51, GOLD",
            "52, ORANGE",
            "53, DARK_ORANGE",
            "54, BLUE_GRAY",
            "55, MEDIUM_GRAY",
            "56, DARK_SLATE_BLUE",
            "57, SEA_GREEN",
            "58, VERY_DARK_GREEN",
            "59, DARK_OLIVE",
            "60, BROWN",
            "61, DARK_ROSE_DUPLICATE",
            "62, INDIGO",
            "63, VERY_DARK_GRAY",
            "64, SYSTEM_FOREGROUND",
            "65, SYSTEM_BACKGROUND"
    })
    void constructorTest3(int index, IndexedColor.Value expectedValue) {
        IndexedColor color = new IndexedColor(index);
        assertEquals(expectedValue, color.getColorValue());
    }

    @ParameterizedTest
    @DisplayName("Test of the failing Constructor on invalid values")
    @ValueSource(ints = {-1, 66, 255, -100})
    void constructorFailTest(int value) {
        assertThrows(StyleException.class, () -> new IndexedColor(value));
    }

    @ParameterizedTest
    @DisplayName("Test of the GetSrgbColor method")
    @CsvSource({
            "BLACK_0, FF000000",
            "WHITE_1, FFFFFFFF",
            "RED_2, FFFF0000",
            "BRIGHT_GREEN_3, FF00FF00",
            "BLUE_4, FF0000FF",
            "YELLOW_5, FFFFFF00",
            "MAGENTA_6, FFFF00FF",
            "CYAN_7, FF00FFFF",
            "BLACK, FF000000",
            "WHITE, FFFFFFFF",
            "RED, FFFF0000",
            "BRIGHT_GREEN, FF00FF00",
            "BLUE, FF0000FF",
            "PURE_BLUE, FF0000FF",
            "YELLOW, FFFFFF00",
            "STRONG_YELLOW, FFFFFF00",
            "MAGENTA, FFFF00FF",
            "STRONG_MAGENTA, FFFF00FF",
            "CYAN, FF00FFFF",
            "STRONG_CYAN, FF00FFFF",
            "DARK_RED, FF800000",
            "DARK_MAROON, FF800000",
            "DARK_GREEN, FF008000",
            "DARK_BLUE, FF000080",
            "NAVY, FF000080",
            "OLIVE, FF808000",
            "PURPLE, FF800080",
            "DARK_VIOLET, FF800080",
            "TEAL, FF008080",
            "DARK_TEAL, FF008080",
            "LIGHT_GRAY, FFC0C0C0",
            "GRAY, FF808080",
            "LIGHT_CORNFLOWER_BLUE, FF9999FF",
            "DARK_ROSE, FF993366",
            "DARK_ROSE_DUPLICATE, FF993366",
            "LIGHT_YELLOW, FFFFFFCC",
            "LIGHT_CYAN, FFCCFFFF",
            "PALE_CYAN, FFCCFFFF",
            "DARK_PURPLE, FF660066",
            "SALMON, FFFF8080",
            "MEDIUM_BLUE, FF0066CC",
            "LIGHT_LAVENDER, FFCCCCFF",
            "SKY_BLUE, FF00CCFF",
            "LIGHT_MINT, FFCCFFCC",
            "PASTEL_YELLOW, FFFFFF99",
            "LIGHT_SKY_BLUE, FF99CCFF",
            "ROSE, FFFF99CC",
            "LAVENDER, FFCC99FF",
            "PEACH, FFFFCC99",
            "ROYAL_BLUE, FF3366FF",
            "TURQUOISE, FF33CCCC",
            "LIGHT_OLIVE, FF99CC00",
            "GOLD, FFFFCC00",
            "ORANGE, FFFF9900",
            "DARK_ORANGE, FFFF6600",
            "BLUE_GRAY, FF666699",
            "MEDIUM_GRAY, FF969696",
            "DARK_SLATE_BLUE, FF003366",
            "SEA_GREEN, FF339966",
            "VERY_DARK_GREEN, FF003300",
            "DARK_OLIVE, FF333300",
            "BROWN, FF993300",
            "INDIGO, FF333399",
            "VERY_DARK_GRAY, FF333333",
            "SYSTEM_BACKGROUND, FFFFFFFF",
            "SYSTEM_FOREGROUND, FF000000"
    })
    void getSrgbTest(IndexedColor.Value givenValue, String expectedArgbValue) {
        IndexedColor color = new IndexedColor(givenValue);
        SrgbColor rgb = color.getSrgbColor();
        assertEquals(expectedArgbValue, rgb.getColorValue());
    }

    @Test
    @DisplayName("Test of the Equals method (multiple cases)")
    void equalsTest() {
        IndexedColor color1 = new IndexedColor(IndexedColor.Value.RED);
        IndexedColor color2 = new IndexedColor(IndexedColor.Value.RED);
        assertTrue(color1.equals(color2));

        IndexedColor color3 = new IndexedColor();
        IndexedColor color4 = new IndexedColor();
        assertTrue(color3.equals(color4));

        IndexedColor color5 = new IndexedColor(23);
        IndexedColor color6 = new IndexedColor(23);
        assertTrue(color5.equals(color6));

        IndexedColor color1b = new IndexedColor(10);
        assertTrue(color1.equals(color1b));

        IndexedColor colorDefault1 = new IndexedColor(IndexedColor.Value.SYSTEM_FOREGROUND);
        assertTrue(color3.equals(colorDefault1));
    }

    @Test
    @DisplayName("Test of the Equals method on inequality (multiple cases)")
    void equalsTest2() {
        IndexedColor color1 = new IndexedColor(IndexedColor.Value.BRIGHT_GREEN);
        IndexedColor color2 = new IndexedColor(IndexedColor.Value.BLUE);
        assertFalse(color1.equals(color2));

        Object object = new Object();
        assertFalse(color1.equals(object));

        IndexedColor color3 = null;
        assertFalse(color1.equals(color3));

        IndexedColor color4 = new IndexedColor(55);
        assertFalse(color1.equals(color4));
    }

    @Test
    @DisplayName("Test of the GetHashCode method (multiple cases)")
    void getHashCodeTest() {
        IndexedColor color1 = new IndexedColor(IndexedColor.Value.YELLOW);
        IndexedColor color2 = new IndexedColor(IndexedColor.Value.YELLOW);
        assertEquals(color1.hashCode(), color2.hashCode());

        IndexedColor color3 = new IndexedColor();
        IndexedColor color4 = new IndexedColor();
        assertEquals(color3.hashCode(), color4.hashCode());

        IndexedColor color5 = new IndexedColor(30);
        IndexedColor color6 = new IndexedColor(30);
        assertEquals(color5.hashCode(), color6.hashCode());
    }

    @Test
    @DisplayName("Test of the GetHashCode method on inequality (multiple cases)")
    void getHashCodeTest2() {
        IndexedColor color1 = new IndexedColor(IndexedColor.Value.MAGENTA);
        IndexedColor color2 = new IndexedColor(IndexedColor.Value.CYAN);
        assertNotEquals(color1.hashCode(), color2.hashCode());

        IndexedColor color3 = new IndexedColor(12);
        IndexedColor color4 = new IndexedColor(45);
        assertNotEquals(color3.hashCode(), color4.hashCode());

        IndexedColor color5 = new IndexedColor();
        IndexedColor color6 = new IndexedColor(IndexedColor.Value.SYSTEM_BACKGROUND);
        assertNotEquals(color5.hashCode(), color6.hashCode());
    }

    // C# ImplicitOperatorTest omitted: Java does not support user-defined conversion operators.

}
