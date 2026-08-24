/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.colors;

import java.util.Objects;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.internal.interfaces.TypedColor;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

/**
 * Class representing an indexed color from the legacy OOXML / Excel indexed color palette.
 */
public class IndexedColor implements TypedColor<IndexedColor.Value> {

    /**
     * Legacy OOXML / Excel indexed color palette.
     * <p>
     * This palette exists for backward compatibility with older Excel formats. Indices 0–7 are redundant with 8–15.
     * </p>
     */
    public enum Value {
        /** Black (duplicate of index 8). */
        BLACK_0(0),
        /** White (duplicate of index 9). */
        WHITE_1(1),
        /** Red (duplicate of index 10). */
        RED_2(2),
        /** Bright green (duplicate of index 11). */
        BRIGHT_GREEN_3(3),
        /** Blue (duplicate of index 12). */
        BLUE_4(4),
        /** Yellow (duplicate of index 13). */
        YELLOW_5(5),
        /** Magenta (duplicate of index 14). */
        MAGENTA_6(6),
        /** Cyan (duplicate of index 15). */
        CYAN_7(7),
        /** Black (#000000). */
        BLACK(8),
        /** White (#FFFFFF). */
        WHITE(9),
        /** Red (#FF0000). */
        RED(10),
        /** Bright green (#00FF00). */
        BRIGHT_GREEN(11),
        /** Blue (#0000FF). */
        BLUE(12),
        /** Yellow (#FFFF00). */
        YELLOW(13),
        /** Magenta / Fuchsia (#FF00FF). */
        MAGENTA(14),
        /** Cyan / Aqua (#00FFFF). */
        CYAN(15),
        /** Dark red / maroon (#800000). */
        DARK_RED(16),
        /** Dark green (#008000). */
        DARK_GREEN(17),
        /** Dark blue / navy (#000080). */
        DARK_BLUE(18),
        /** Olive (#808000). */
        OLIVE(19),
        /** Purple (#800080). */
        PURPLE(20),
        /** Teal (#008080). */
        TEAL(21),
        /** Light gray / silver (#C0C0C0). */
        LIGHT_GRAY(22),
        /** Medium gray (#808080). */
        GRAY(23),
        /** Light cornflower blue (#9999FF). */
        LIGHT_CORNFLOWER_BLUE(24),
        /** Dark rose (#993366). */
        DARK_ROSE(25),
        /** Light yellow (#FFFFCC). */
        LIGHT_YELLOW(26),
        /** Light cyan (#CCFFFF). */
        LIGHT_CYAN(27),
        /** Dark purple (#660066). */
        DARK_PURPLE(28),
        /** Salmon pink (#FF8080). */
        SALMON(29),
        /** Medium blue (#0066CC). */
        MEDIUM_BLUE(30),
        /** Light lavender blue (#CCCCFF). */
        LIGHT_LAVENDER(31),
        /** Dark navy blue (#000080). */
        NAVY(32),
        /** Strong magenta (#FF00FF). */
        STRONG_MAGENTA(33),
        /** Strong yellow (#FFFF00). */
        STRONG_YELLOW(34),
        /** Strong cyan (#00FFFF). */
        STRONG_CYAN(35),
        /** Dark violet (#800080). */
        DARK_VIOLET(36),
        /** Dark maroon (#800000). */
        DARK_MAROON(37),
        /** Dark teal (#008080). */
        DARK_TEAL(38),
        /** Pure blue (#0000FF). */
        PURE_BLUE(39),
        /** Sky blue (#00CCFF). */
        SKY_BLUE(40),
        /** Pale cyan (#CCFFFF). */
        PALE_CYAN(41),
        /** Light mint green (#CCFFCC). */
        LIGHT_MINT(42),
        /** Light pastel yellow (#FFFF99). */
        PASTEL_YELLOW(43),
        /** Light sky blue (#99CCFF). */
        LIGHT_SKY_BLUE(44),
        /** Rose pink (#FF99CC). */
        ROSE(45),
        /** Lavender (#CC99FF). */
        LAVENDER(46),
        /** Peach (#FFCC99). */
        PEACH(47),
        /** Royal blue (#3366FF). */
        ROYAL_BLUE(48),
        /** Turquoise (#33CCCC). */
        TURQUOISE(49),
        /** Light olive green (#99CC00). */
        LIGHT_OLIVE(50),
        /** Gold (#FFCC00). */
        GOLD(51),
        /** Orange (#FF9900). */
        ORANGE(52),
        /** Dark orange (#FF6600). */
        DARK_ORANGE(53),
        /** Blue gray (#666699). */
        BLUE_GRAY(54),
        /** Medium gray (#969696). */
        MEDIUM_GRAY(55),
        /** Dark slate blue (#003366). */
        DARK_SLATE_BLUE(56),
        /** Sea green (#339966). */
        SEA_GREEN(57),
        /** Very dark green (#003300). */
        VERY_DARK_GREEN(58),
        /** Dark olive (#333300). */
        DARK_OLIVE(59),
        /** Brown (#993300). */
        BROWN(60),
        /** Dark rose (duplicate of index 25). */
        DARK_ROSE_DUPLICATE(61),
        /** Indigo / dark blue-purple (#333399). */
        INDIGO(62),
        /** Very dark gray (#333333). */
        VERY_DARK_GRAY(63),
        /**
         * System foreground color.
         * <p>
         * The actual color is determined by the host operating system or theme.
         * </p>
         */
        SYSTEM_FOREGROUND(64),
        /**
         * System background color.
         * <p>
         * The actual color is determined by the host operating system or theme.
         * </p>
         */
        SYSTEM_BACKGROUND(65);

        public final int value;

        private Value(int value) {
            this.value = value;
        }

    }

    /**
     * Default indexed color (system foreground color)
     */
    public static final Value DEFAULT_INDEXED_COLOR = Value.SYSTEM_FOREGROUND;

    /**
     * Default ARGB value for system foreground color
     */
    public static final String DEFAULT_SYSTEM_FOREGROUND_COLOR_ARGB = "FF000000";
    /**
     * Default ARGB value for system background color
     */
    public static final String DEFAULT_SYSTEM_BACKGROUND_COLOR_ARGB = "FFFFFFFF";

    private Value colorValue;

    /**
     * Gets the value of the indexed color
     *
     * @return Value of the type {@code T}
     */
    @Override
    public Value getColorValue() {
        return colorValue;
    }

    /**
     * Sets the value of the indexed color
     *
     * @param colorValue Value of the type {@code T}
     */
    @Override
    public void setColorValue(Value colorValue) {
        this.colorValue = colorValue;
    }

    /**
     * Gets the string representation of the indexed color value
     *
     * @return String of the indexed color
     */
    @Override
    public String getStringValue() {
        return ParserUtils.toString(colorValue.value);
    }

    /**
     * Default constructor with default indexed color
     */
    public IndexedColor() {
        colorValue = DEFAULT_INDEXED_COLOR;
    }

    /**
     * Constructor with specified indexed color value
     *
     * @param color Indexed color
     */
    public IndexedColor(Value color) {
        colorValue = color;
    }

    /**
     * Constructor with specified indexed color index
     *
     * @param colorIndex Color index
     * @throws StyleException Throws a StyleException if the color index is out of range
     */
    public IndexedColor(int colorIndex) throws StyleException {
        if (colorIndex < 0 || colorIndex > 65) {
            throw new StyleException("Indexed color value must be between 0 and 65.");
        }
        colorValue = Value.values()[colorIndex];
    }

    /**
     * Gets the ARGB hex code representation of the indexed color
     *
     * @return ARGB value of the current color instance
     */
    private String getArgbValue() {
        return getArgbValue(colorValue);
    }

    /**
     * Gets the sRGB color representation of the indexed color
     *
     * @return sRGB color instance
     */
    public SrgbColor getSrgbColor() {
        return new SrgbColor(getArgbValue());
    }

    /**
     * Maps the indexed color value to its ARGB hex code representation
     *
     * @param indexedValue Enum value
     * @return ARGB value
     */
    public static String getArgbValue(Value indexedValue) {
        return switch (indexedValue) {
            // 0–7 (duplicates of 8–15)
            case Value.BLACK_0, Value.BLACK -> "FF000000";
            case Value.WHITE_1, Value.WHITE -> "FFFFFFFF";
            case Value.RED_2, Value.RED -> "FFFF0000";
            case Value.BRIGHT_GREEN_3, Value.BRIGHT_GREEN -> "FF00FF00";
            case Value.BLUE_4, Value.BLUE, Value.PURE_BLUE -> "FF0000FF";
            case Value.YELLOW_5, Value.YELLOW, Value.STRONG_YELLOW -> "FFFFFF00";
            case Value.MAGENTA_6, Value.MAGENTA, Value.STRONG_MAGENTA -> "FFFF00FF";
            case Value.CYAN_7, Value.CYAN, Value.STRONG_CYAN -> "FF00FFFF";

            // Extended palette
            case Value.DARK_RED, Value.DARK_MAROON -> "FF800000";
            case Value.DARK_GREEN -> "FF008000";
            case Value.DARK_BLUE, Value.NAVY -> "FF000080";
            case Value.OLIVE -> "FF808000";
            case Value.PURPLE, Value.DARK_VIOLET -> "FF800080";
            case Value.TEAL, Value.DARK_TEAL -> "FF008080";
            case Value.LIGHT_GRAY -> "FFC0C0C0";
            case Value.GRAY -> "FF808080";
            case Value.LIGHT_CORNFLOWER_BLUE -> "FF9999FF";
            case Value.DARK_ROSE, Value.DARK_ROSE_DUPLICATE -> "FF993366";
            case Value.LIGHT_YELLOW -> "FFFFFFCC";
            case Value.LIGHT_CYAN, Value.PALE_CYAN -> "FFCCFFFF";
            case Value.DARK_PURPLE -> "FF660066";
            case Value.SALMON -> "FFFF8080";
            case Value.MEDIUM_BLUE -> "FF0066CC";
            case Value.LIGHT_LAVENDER -> "FFCCCCFF";
            case Value.SKY_BLUE -> "FF00CCFF";
            case Value.LIGHT_MINT -> "FFCCFFCC";
            case Value.PASTEL_YELLOW -> "FFFFFF99";
            case Value.LIGHT_SKY_BLUE -> "FF99CCFF";
            case Value.ROSE -> "FFFF99CC";
            case Value.LAVENDER -> "FFCC99FF";
            case Value.PEACH -> "FFFFCC99";
            case Value.ROYAL_BLUE -> "FF3366FF";
            case Value.TURQUOISE -> "FF33CCCC";
            case Value.LIGHT_OLIVE -> "FF99CC00";
            case Value.GOLD -> "FFFFCC00";
            case Value.ORANGE -> "FFFF9900";
            case Value.DARK_ORANGE -> "FFFF6600";
            case Value.BLUE_GRAY -> "FF666699";
            case Value.MEDIUM_GRAY -> "FF969696";
            case Value.DARK_SLATE_BLUE -> "FF003366";
            case Value.SEA_GREEN -> "FF339966";
            case Value.VERY_DARK_GREEN -> "FF003300";
            case Value.DARK_OLIVE -> "FF333300";
            case Value.BROWN -> "FF993300";
            case Value.INDIGO -> "FF333399";
            case Value.VERY_DARK_GRAY -> "FF333333";
            case Value.SYSTEM_BACKGROUND ->
                // Excel default: white background
                    DEFAULT_SYSTEM_BACKGROUND_COLOR_ARGB;
            default ->
                // Excel default: black text
                    DEFAULT_SYSTEM_FOREGROUND_COLOR_ARGB;
        };
    }

    /**
     * Determines whether the specified object is equal to the current object
     *
     * @param o the reference object with which to compare.
     * @return True if both objects are equal
     */
    @Override
    public final boolean equals(Object o) {
        if (!(o instanceof IndexedColor that)) {
            return false;
        }

        return colorValue == that.colorValue;
    }

    /**
     * Gets the hash code of the instance
     *
     * @return Hash code
     */
    @Override
    public int hashCode() {
        return Objects.hashCode(colorValue);
    }
}
