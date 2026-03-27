//STATUS: REVISE ENUM
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.colors;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.interfaces.ITypedColor;

/**
 * Class representing an indexed color from the legacy OOXML / Excel indexed color palette.
 * <p>
 * Java equivalent of {@code NanoXLSX.Colors.IndexedColor} in the .NET reference.
 * </p>
 */
public class IndexedColor implements ITypedColor<IndexedColor.Value> {

    /**
     * Legacy OOXML / Excel indexed color palette (indices 0–65). Indices 0–7 are redundant with 8–15.
     */
    public enum Value {
        /** Black (duplicate of index 8, index 0) */
        Black0(0),
        /** White (duplicate of index 9, index 1) */
        White1(1),
        /** Red (duplicate of index 10, index 2) */
        Red2(2),
        /** Bright green (duplicate of index 11, index 3) */
        BrightGreen3(3),
        /** Blue (duplicate of index 12, index 4) */
        Blue4(4),
        /** Yellow (duplicate of index 13, index 5) */
        Yellow5(5),
        /** Magenta (duplicate of index 14, index 6) */
        Magenta6(6),
        /** Cyan (duplicate of index 15, index 7) */
        Cyan7(7),
        /** Black (#000000, index 8) */
        Black(8),
        /** White (#FFFFFF, index 9) */
        White(9),
        /** Red (#FF0000, index 10) */
        Red(10),
        /** Bright green (#00FF00, index 11) */
        BrightGreen(11),
        /** Blue (#0000FF, index 12) */
        Blue(12),
        /** Yellow (#FFFF00, index 13) */
        Yellow(13),
        /** Magenta / Fuchsia (#FF00FF, index 14) */
        Magenta(14),
        /** Cyan / Aqua (#00FFFF, index 15) */
        Cyan(15),
        /** Dark red / maroon (#800000, index 16) */
        DarkRed(16),
        /** Dark green (#008000, index 17) */
        DarkGreen(17),
        /** Dark blue / navy (#000080, index 18) */
        DarkBlue(18),
        /** Olive (#808000, index 19) */
        Olive(19),
        /** Purple (#800080, index 20) */
        Purple(20),
        /** Teal (#008080, index 21) */
        Teal(21),
        /** Light gray / silver (#C0C0C0, index 22) */
        LightGray(22),
        /** Medium gray (#808080, index 23) */
        Gray(23),
        /** Light cornflower blue (#9999FF, index 24) */
        LightCornflowerBlue(24),
        /** Dark rose (#993366, index 25) */
        DarkRose(25),
        /** Light yellow (#FFFFCC, index 26) */
        LightYellow(26),
        /** Light cyan (#CCFFFF, index 27) */
        LightCyan(27),
        /** Dark purple (#660066, index 28) */
        DarkPurple(28),
        /** Salmon pink (#FF8080, index 29) */
        Salmon(29),
        /** Medium blue (#0066CC, index 30) */
        MediumBlue(30),
        /** Light lavender blue (#CCCCFF, index 31) */
        LightLavender(31),
        /** Dark navy blue (#000080, index 32) */
        Navy(32),
        /** Strong magenta (#FF00FF, index 33) */
        StrongMagenta(33),
        /** Strong yellow (#FFFF00, index 34) */
        StrongYellow(34),
        /** Strong cyan (#00FFFF, index 35) */
        StrongCyan(35),
        /** Dark violet (#800080, index 36) */
        DarkViolet(36),
        /** Dark maroon (#800000, index 37) */
        DarkMaroon(37),
        /** Dark teal (#008080, index 38) */
        DarkTeal(38),
        /** Pure blue (#0000FF, index 39) */
        PureBlue(39),
        /** Sky blue (#00CCFF, index 40) */
        SkyBlue(40),
        /** Pale cyan (#CCFFFF, index 41) */
        PaleCyan(41),
        /** Light mint green (#CCFFCC, index 42) */
        LightMint(42),
        /** Light pastel yellow (#FFFF99, index 43) */
        PastelYellow(43),
        /** Light sky blue (#99CCFF, index 44) */
        LightSkyBlue(44),
        /** Rose pink (#FF99CC, index 45) */
        Rose(45),
        /** Lavender (#CC99FF, index 46) */
        Lavender(46),
        /** Peach (#FFCC99, index 47) */
        Peach(47),
        /** Royal blue (#3366FF, index 48) */
        RoyalBlue(48),
        /** Turquoise (#33CCCC, index 49) */
        Turquoise(49),
        /** Light olive green (#99CC00, index 50) */
        LightOlive(50),
        /** Gold (#FFCC00, index 51) */
        Gold(51),
        /** Orange (#FF9900, index 52) */
        Orange(52),
        /** Dark orange (#FF6600, index 53) */
        DarkOrange(53),
        /** Blue gray (#666699, index 54) */
        BlueGray(54),
        /** Medium gray (#969696, index 55) */
        MediumGray(55),
        /** Dark slate blue (#003366, index 56) */
        DarkSlateBlue(56),
        /** Sea green (#339966, index 57) */
        SeaGreen(57),
        /** Very dark green (#003300, index 58) */
        VeryDarkGreen(58),
        /** Dark olive (#333300, index 59) */
        DarkOlive(59),
        /** Brown (#993300, index 60) */
        Brown(60),
        /** Dark rose duplicate (index 61) */
        DarkRoseDuplicate(61),
        /** Indigo / dark blue-purple (#333399, index 62) */
        Indigo(62),
        /** Very dark gray (#333333, index 63) */
        VeryDarkGray(63),
        /** System foreground color (index 64). The actual color is determined by the host operating system or theme. */
        SystemForeground(64),
        /** System background color (index 65). The actual color is determined by the host operating system or theme. */
        SystemBackground(65);

        private final int code;

        Value(int code) {
            this.code = code;
        }
    }

    /** Default indexed color (system foreground) */
    public static final Value DEFAULT_INDEXED_COLOR = Value.SystemForeground;
    /** Default ARGB for system foreground (black) */
    public static final String DEFAULT_SYSTEM_FOREGROUND_COLOR_ARGB = "FF000000";
    /** Default ARGB for system background (white) */
    public static final String DEFAULT_SYSTEM_BACKGROUND_COLOR_ARGB = "FFFFFFFF";

    private Value colorValue;

    /**
     * Gets the indexed color enum value
     *
     * @return Color enum value
     */
    @Override
    public Value getColorValue() {
        return colorValue;
    }

    /**
     * Sets the indexed color enum value
     *
     * @param value Color enum value
     */
    @Override
    public void setColorValue(Value value) {
        this.colorValue = value;
    }

    /**
     * Gets the numeric index of the color as a string
     *
     * @return Numeric index string
     */
    @Override
    public String getStringValue() {
        return String.valueOf(colorValue.ordinal());
    }

    /**
     * Gets the ARGB hex code for the current indexed color
     *
     * @return ARGB hex string
     */
    public String getArgbValue() {
        return getArgbValue(colorValue);
    }

    /**
     * Gets an {@link SrgbColor} representation of this indexed color
     *
     * @return SrgbColor instance
     */
    public SrgbColor getSrgbColor() {
        return new SrgbColor(getArgbValue());
    }

    /**
     * Default constructor — uses {@link #DEFAULT_INDEXED_COLOR}
     */
    public IndexedColor() {
        this.colorValue = DEFAULT_INDEXED_COLOR;
    }

    /**
     * Constructor with specified indexed color value
     *
     * @param color Indexed color enum value
     */
    public IndexedColor(Value color) {
        this.colorValue = color;
    }

    /**
     * Constructor with a numeric color index (0–65)
     *
     * @param colorIndex Color palette index
     * @throws StyleException Thrown if the index is out of range (0–65)
     */
    public IndexedColor(int colorIndex) {
        if (colorIndex < 0 || colorIndex > 65) {
            throw new StyleException("Indexed color value must be between 0 and 65.");
        }
        this.colorValue = Value.values()[colorIndex];
    }

    /**
     * Maps an indexed color enum value to its ARGB hex representation
     *
     * @param indexedValue Enum value to map
     * @return ARGB hex string (8 characters)
     */
    public static String getArgbValue(Value indexedValue) {
        switch (indexedValue) {
            case Black0:
            case Black:
                return "FF000000";
            case White1:
            case White:
                return "FFFFFFFF";
            case Red2:
            case Red:
                return "FFFF0000";
            case BrightGreen3:
            case BrightGreen:
                return "FF00FF00";
            case Blue4:
            case Blue:
            case PureBlue:
                return "FF0000FF";
            case Yellow5:
            case Yellow:
            case StrongYellow:
                return "FFFFFF00";
            case Magenta6:
            case Magenta:
            case StrongMagenta:
                return "FFFF00FF";
            case Cyan7:
            case Cyan:
            case StrongCyan:
                return "FF00FFFF";
            case DarkRed:
            case DarkMaroon:
                return "FF800000";
            case DarkGreen:
                return "FF008000";
            case DarkBlue:
            case Navy:
                return "FF000080";
            case Olive:
                return "FF808000";
            case Purple:
            case DarkViolet:
                return "FF800080";
            case Teal:
            case DarkTeal:
                return "FF008080";
            case LightGray:
                return "FFC0C0C0";
            case Gray:
                return "FF808080";
            case LightCornflowerBlue:
                return "FF9999FF";
            case DarkRose:
            case DarkRoseDuplicate:
                return "FF993366";
            case LightYellow:
                return "FFFFFFCC";
            case LightCyan:
            case PaleCyan:
                return "FFCCFFFF";
            case DarkPurple:
                return "FF660066";
            case Salmon:
                return "FFFF8080";
            case MediumBlue:
                return "FF0066CC";
            case LightLavender:
                return "FFCCCCFF";
            case SkyBlue:
                return "FF00CCFF";
            case LightMint:
                return "FFCCFFCC";
            case PastelYellow:
                return "FFFFFF99";
            case LightSkyBlue:
                return "FF99CCFF";
            case Rose:
                return "FFFF99CC";
            case Lavender:
                return "FFCC99FF";
            case Peach:
                return "FFFFCC99";
            case RoyalBlue:
                return "FF3366FF";
            case Turquoise:
                return "FF33CCCC";
            case LightOlive:
                return "FF99CC00";
            case Gold:
                return "FFFFCC00";
            case Orange:
                return "FFFF9900";
            case DarkOrange:
                return "FFFF6600";
            case BlueGray:
                return "FF666699";
            case MediumGray:
                return "FF969696";
            case DarkSlateBlue:
                return "FF003366";
            case SeaGreen:
                return "FF339966";
            case VeryDarkGreen:
                return "FF003300";
            case DarkOlive:
                return "FF333300";
            case Brown:
                return "FF993300";
            case Indigo:
                return "FF333399";
            case VeryDarkGray:
                return "FF333333";
            case SystemBackground:
                return DEFAULT_SYSTEM_BACKGROUND_COLOR_ARGB;
            default:
                return DEFAULT_SYSTEM_FOREGROUND_COLOR_ARGB;
        }
    }

    /**
     * Determines whether the specified object is equal to the current object
     *
     * @param obj Other object to compare
     * @return True if both objects are equal
     */
    @Override
    public boolean equals(Object obj) {
        if (this == obj) return true;
        if (!(obj instanceof IndexedColor)) return false;
        return colorValue == ((IndexedColor) obj).colorValue;
    }

    /**
     * Gets the hash code of the instance
     *
     * @return Hash code
     */
    @Override
    public int hashCode() {
        return 800285905 + colorValue.hashCode();
    }
}
