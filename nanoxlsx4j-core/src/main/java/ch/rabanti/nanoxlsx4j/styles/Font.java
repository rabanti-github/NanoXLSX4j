/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import java.util.Objects;

import ch.rabanti.nanoxlsx4j.colors.Color;
import ch.rabanti.nanoxlsx4j.colors.IndexedColor;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

/** Class representing a Font entry used to define text formatting. */
public class Font extends AbstractStyle {

    /** Minimum possible font size. */
    public static final float MIN_FONT_SIZE = 1f;
    /** Maximum possible font size. */
    public static final float MAX_FONT_SIZE = 409f;
    /** Default font size. */
    public static final float DEFAULT_FONT_SIZE = 11f;
    /** Default major-scheme font name. */
    public static final String DEFAULT_MAJOR_FONT = "Calibri Light";
    /** Default minor-scheme font name. */
    public static final String DEFAULT_MINOR_FONT = "Calibri";
    /** Default font name. */
    public static final String DEFAULT_FONT_NAME = DEFAULT_MINOR_FONT;
    /** Default font family. */
    public static final FontFamilyValue DEFAULT_FONT_FAMILY = FontFamilyValue.SWISS;
    /** Default font scheme. */
    public static final SchemeValue DEFAULT_FONT_SCHEME = SchemeValue.MINOR;
    /** Default vertical text alignment. */
    public static final VerticalTextAlignValue DEFAULT_VERTICAL_ALIGN = VerticalTextAlignValue.NONE;

    /** Font scheme. */
    public enum SchemeValue {
        /** Major font scheme. */
        MAJOR,
        /** Minor font scheme. */
        MINOR,
        /** No font scheme. */
        NONE
    }

    /** Vertical alignment relative to the text baseline. */
    public enum VerticalTextAlignValue {
        /** Render text at the baseline. */
        BASELINE,
        /** Render text as subscript. */
        SUBSCRIPT,
        /** Render text as superscript. */
        SUPERSCRIPT,
        /** Render text normally. */
        NONE
    }

    /** Underline style. */
    public enum UnderlineValue {
        /** Single underline. */
        SINGLE,
        /** Double underline. */
        DOUBLE,
        /** Single accounting underline. */
        SINGLE_ACCOUNTING,
        /** Double accounting underline. */
        DOUBLE_ACCOUNTING,
        /** No underline. */
        NONE
    }

    /** Font charset identifiers defined by OOXML. */
    public enum CharsetValue {
        /** Application-defined charset. */
        APPLICATION_DEFINED(-1),
        /** ANSI charset. */
        ANSI(0),
        /** Default charset. */
        DEFAULT(1),
        /** Symbol charset. */
        SYMBOLS(2),
        /** Macintosh Roman charset. */
        MACINTOSH(77),
        /** Shift JIS charset. */
        JIS(128),
        /** Hangul charset. */
        HANGUL(129),
        /** Johab charset. */
        JOHAB(130),
        /** GBK charset. */
        GBK(134),
        /** Chinese Big Five charset. */
        BIG_5(136),
        /** Greek charset. */
        GREEK(161),
        /** Turkish charset. */
        TURKISH(162),
        /** Vietnamese charset. */
        VIETNAMESE(163),
        /** Hebrew charset. */
        HEBREW(177),
        /** Arabic charset. */
        ARABIC(178),
        /** Baltic charset. */
        BALTIC(186),
        /** Russian charset. */
        RUSSIAN(204),
        /** Thai charset. */
        THAI(222),
        /** Eastern European charset. */
        EASTERN_EUROPEAN(238),
        /** OEM charset. */
        OEM(255);

        private final int value;

        CharsetValue(int value) {
            this.value = value;
        }

        /**
         * Gets the numeric OOXML charset identifier.
         *
         * @return Charset identifier
         */
        public int getValue() {
            return value;
        }
    }

    /** Font-family identifiers defined by OOXML. */
    public enum FontFamilyValue {
        /** Family is not defined or not applicable. */
        NOT_APPLICABLE(0),
        /** Roman font family. */
        ROMAN(1),
        /** Swiss font family. */
        SWISS(2),
        /** Modern font family. */
        MODERN(3),
        /** Script font family. */
        SCRIPT(4),
        /** Decorative font family. */
        DECORATIVE(5),
        /** Reserved family identifier. */
        RESERVED_1(6),
        /** Reserved family identifier. */
        RESERVED_2(7),
        /** Reserved family identifier. */
        RESERVED_3(8),
        /** Reserved family identifier. */
        RESERVED_4(9),
        /** Reserved family identifier. */
        RESERVED_5(10),
        /** Reserved family identifier. */
        RESERVED_6(11),
        /** Reserved family identifier. */
        RESERVED_7(12),
        /** Reserved family identifier. */
        RESERVED_8(13),
        /** Reserved family identifier. */
        RESERVED_9(14);

        private final int value;

        FontFamilyValue(int value) {
            this.value = value;
        }

        /**
         * Gets the numeric OOXML family identifier.
         *
         * @return Family identifier
         */
        public int getValue() {
            return value;
        }
    }

    @AppendAnnotation
    private boolean bold;
    @AppendAnnotation
    private boolean italic;
    @AppendAnnotation
    private boolean strike;
    @AppendAnnotation
    private UnderlineValue underline = UnderlineValue.NONE;
    @AppendAnnotation
    private boolean outline;
    @AppendAnnotation
    private boolean shadow;
    @AppendAnnotation
    private boolean condense;
    @AppendAnnotation
    private boolean extend;
    @AppendAnnotation
    private CharsetValue charset = CharsetValue.DEFAULT;
    @AppendAnnotation
    private Color colorValue;
    @AppendAnnotation
    private FontFamilyValue family;
    @AppendAnnotation
    private String name = DEFAULT_FONT_NAME;
    @AppendAnnotation
    private SchemeValue scheme;
    @AppendAnnotation
    private float size;
    @AppendAnnotation
    private VerticalTextAlignValue verticalAlign;

    /** Creates a font with the reference implementation's defaults. */
    public Font() {
        size = DEFAULT_FONT_SIZE;
        setName(DEFAULT_FONT_NAME);
        family = DEFAULT_FONT_FAMILY;
        colorValue = Color.createNone();
        scheme = DEFAULT_FONT_SCHEME;
        verticalAlign = DEFAULT_VERTICAL_ALIGN;
    }

    /**
     * Gets whether the font is bold.
     * @return True if bold
     */
    public boolean isBold() {
        return bold;
    }

    /**
     * Sets whether the font is bold.
     * @param bold Whether the font is bold
     */
    public void setBold(boolean bold) {
        this.bold = bold;
    }

    /**
     * Gets whether the font is italic.
     * @return True if italic
     */
    public boolean isItalic() {
        return italic;
    }

    /**
     * Sets whether the font is italic.
     * @param italic Whether the font is italic
     */
    public void setItalic(boolean italic) {
        this.italic = italic;
    }

    /**
     * Gets whether the font is struck through.
     * @return True if struck through
     */
    public boolean isStrike() {
        return strike;
    }

    /**
     * Sets whether the font is struck through.
     * @param strike Whether the font is struck through
     */
    public void setStrike(boolean strike) {
        this.strike = strike;
    }

    /**
     * Gets the underline style.
     * @return Underline style
     */
    public UnderlineValue getUnderline() {
        return underline;
    }

    /**
     * Sets the underline style.
     * @param underline Underline style
     */
    public void setUnderline(UnderlineValue underline) {
        this.underline = Objects.requireNonNull(underline, "underline");
    }

    /**
     * Gets whether the font has an outline.
     * @return True if outlined
     */
    public boolean isOutline() {
        return outline;
    }

    /**
     * Sets whether the font has an outline.
     * @param outline Whether the font has an outline
     */
    public void setOutline(boolean outline) {
        this.outline = outline;
    }

    /**
     * Gets whether the font has a shadow.
     * @return True if shadowed
     */
    public boolean isShadow() {
        return shadow;
    }

    /**
     * Sets whether the font has a shadow.
     * @param shadow Whether the font has a shadow
     */
    public void setShadow(boolean shadow) {
        this.shadow = shadow;
    }

    /**
     * Gets whether the font is condensed.
     * @return True if condensed
     */
    public boolean isCondense() {
        return condense;
    }

    /**
     * Sets whether the font is condensed.
     * @param condense Whether the font is condensed
     */
    public void setCondense(boolean condense) {
        this.condense = condense;
    }

    /**
     * Gets whether the font is extended.
     * @return True if extended
     */
    public boolean isExtend() {
        return extend;
    }

    /**
     * Sets whether the font is extended.
     * @param extend Whether the font is extended
     */
    public void setExtend(boolean extend) {
        this.extend = extend;
    }

    /**
     * Gets the font charset.
     * @return Font charset
     */
    public CharsetValue getCharset() {
        return charset;
    }

    /**
     * Sets the font charset.
     * @param charset Font charset
     */
    public void setCharset(CharsetValue charset) {
        this.charset = Objects.requireNonNull(charset, "charset");
    }

    /**
     * Gets the compound font color.
     * @return Font color
     */
    public Color getColorValue() {
        return colorValue;
    }

    /**
     * Sets the compound font color.
     * @param colorValue Font color
     */
    public void setColorValue(Color colorValue) {
        this.colorValue = colorValue;
    }

    /**
     * Sets the font color from an RGB or ARGB value.
     * @param colorValue RGB or ARGB value
     */
    public void setColorValue(String colorValue) {
        setColorValue(Color.createRgb(colorValue));
    }

    /**
     * Sets the font color from an index from 0 through 65.
     * @param colorIndex Color index
     */
    public void setColorValue(int colorIndex) {
        setColorValue(Color.createIndexed(colorIndex));
    }

    /**
     * Sets the font color from an indexed-color enum value.
     * @param colorIndex Indexed color
     */
    public void setColorValue(IndexedColor.Value colorIndex) {
        setColorValue(Color.createIndexed(colorIndex));
    }

    /**
     * Gets the font family.
     * @return Font family
     */
    public FontFamilyValue getFamily() {
        return family;
    }

    /**
     * Sets the font family.
     * @param family Font family
     */
    public void setFamily(FontFamilyValue family) {
        this.family = Objects.requireNonNull(family, "family");
    }

    /**
     * Gets whether this instance equals a default font.
     * @return True if this is a default font
     */
    public boolean isDefaultFont() {
        return equals(new Font());
    }

    /**
     * Gets the font name.
     * @return Font name
     */
    public String getName() {
        return name;
    }

    /**
     * Sets the font name and derives the font scheme.
     *
     * @param name Font name
     * @throws StyleException if the name is null or empty
     */
    public void setName(String name) {
        this.name = name;
        validateFontScheme();
    }

    /**
     * Gets the font scheme.
     * @return Font scheme
     */
    public SchemeValue getScheme() {
        return scheme;
    }

    /**
     * Sets the font scheme.
     * @param scheme Font scheme
     */
    public void setScheme(SchemeValue scheme) {
        this.scheme = Objects.requireNonNull(scheme, "scheme");
    }

    /**
     * Gets the font size.
     * @return Font size
     */
    public float getSize() {
        return size;
    }

    /**
     * Sets the font size, clamped to the supported range.
     *
     * @param size Font size
     */
    public void setSize(float size) {
        if (size < MIN_FONT_SIZE) {
            this.size = MIN_FONT_SIZE;
        } else if (size > MAX_FONT_SIZE) {
            this.size = MAX_FONT_SIZE;
        } else {
            this.size = size;
        }
    }

    /**
     * Gets the vertical text alignment.
     * @return Vertical text alignment
     */
    public VerticalTextAlignValue getVerticalAlign() {
        return verticalAlign;
    }

    /**
     * Sets the vertical text alignment.
     * @param verticalAlign Vertical text alignment
     */
    public void setVerticalAlign(VerticalTextAlignValue verticalAlign) {
        this.verticalAlign = Objects.requireNonNull(verticalAlign, "verticalAlign");
    }

    private void validateFontScheme() {
        if (name == null || name.isEmpty()) {
            throw new StyleException("The font name was null or empty");
        }
        if (name.equals(DEFAULT_MINOR_FONT)) {
            scheme = SchemeValue.MINOR;
        } else if (name.equals(DEFAULT_MAJOR_FONT)) {
            scheme = SchemeValue.MAJOR;
        } else {
            scheme = SchemeValue.NONE;
        }
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("\"Font\": {\n");
        addPropertyAsJson(sb, "Bold", bold, false);
        addPropertyAsJson(sb, "Charset", charset, false);
        addPropertyAsJson(sb, "ColorValue", colorValue, false);
        addPropertyAsJson(sb, "VerticalAlign", verticalAlign, false);
        addPropertyAsJson(sb, "Family", family, false);
        addPropertyAsJson(sb, "Italic", italic, false);
        addPropertyAsJson(sb, "Name", name, false);
        addPropertyAsJson(sb, "Scheme", scheme, false);
        addPropertyAsJson(sb, "Size", size, false);
        addPropertyAsJson(sb, "Strike", strike, false);
        addPropertyAsJson(sb, "Underline", underline, false);
        addPropertyAsJson(sb, "Outline", outline, false);
        addPropertyAsJson(sb, "Shadow", shadow, false);
        addPropertyAsJson(sb, "Condense", condense, false);
        addPropertyAsJson(sb, "Extend", extend, false);
        addPropertyAsJson(sb, "HashCode", hashCode(), true);
        sb.append("\n}");
        return sb.toString();
    }

    /**
     * Copies this font without its internal ID.
     *
     * @return Dereferenced copy
     */
    @Override
    public AbstractStyle copy() {
        Font copy = new Font();
        copy.bold = bold;
        copy.charset = charset;
        copy.colorValue = colorValue;
        copy.verticalAlign = verticalAlign;
        copy.family = family;
        copy.italic = italic;
        copy.name = name;
        copy.scheme = scheme;
        copy.size = size;
        copy.strike = strike;
        copy.underline = underline;
        copy.outline = outline;
        copy.shadow = shadow;
        copy.condense = condense;
        copy.extend = extend;
        return copy;
    }

    @Override
    public int hashCode() {
        return Objects.hash(size, bold, charset, colorValue, family, italic, name, scheme, strike, underline, outline,
                shadow, condense, extend, verticalAlign);
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (!(obj instanceof Font other)) {
            return false;
        }
        return size == other.size
                && bold == other.bold
                && italic == other.italic
                && strike == other.strike
                && underline == other.underline
                && outline == other.outline
                && shadow == other.shadow
                && condense == other.condense
                && extend == other.extend
                && charset == other.charset
                && Objects.equals(colorValue, other.colorValue)
                && family == other.family
                && Objects.equals(name, other.name)
                && scheme == other.scheme
                && verticalAlign == other.verticalAlign;
    }

    /**
     * Copies this font without requiring a cast.
     *
     * @return Dereferenced copy
     */
    public Font copyFont() {
        return (Font) copy();
    }

    /**
     * Converts vertical text alignment to its OOXML name.
     *
     * @param align Alignment to convert
     * @return OOXML name, or an empty string for none
     */
    static String getVerticalTextAlignName(VerticalTextAlignValue align) {
        return switch (align) {
            case BASELINE -> "baseline";
            case SUBSCRIPT -> "subscript";
            case SUPERSCRIPT -> "superscript";
            case NONE -> "";
        };
    }

    /**
     * Parses an OOXML vertical text alignment name.
     *
     * @param name Alignment name
     * @return Parsed alignment, or {@link VerticalTextAlignValue#NONE} for an unknown value
     */
    static VerticalTextAlignValue getVerticalTextAlignEnum(String name) {
        if (name == null) {
            return VerticalTextAlignValue.NONE;
        }
        return switch (name) {
            case "baseline" -> VerticalTextAlignValue.BASELINE;
            case "subscript" -> VerticalTextAlignValue.SUBSCRIPT;
            case "superscript" -> VerticalTextAlignValue.SUPERSCRIPT;
            default -> VerticalTextAlignValue.NONE;
        };
    }

    /**
     * Converts an underline style to its OOXML name.
     *
     * @param underline Underline style
     * @return OOXML name, or an empty string for single/none
     */
    static String getUnderlineName(UnderlineValue underline) {
        return switch (underline) {
            case DOUBLE -> "double";
            case SINGLE_ACCOUNTING -> "singleAccounting";
            case DOUBLE_ACCOUNTING -> "doubleAccounting";
            case SINGLE, NONE -> "";
        };
    }

    /**
     * Parses an OOXML underline name.
     *
     * @param name Underline name
     * @return Parsed underline, or {@link UnderlineValue#NONE} for an unknown value
     */
    static UnderlineValue getUnderlineEnum(String name) {
        if (name == null) {
            return UnderlineValue.NONE;
        }
        return switch (name) {
            case "double" -> UnderlineValue.DOUBLE;
            case "singleAccounting" -> UnderlineValue.SINGLE_ACCOUNTING;
            case "doubleAccounting" -> UnderlineValue.DOUBLE_ACCOUNTING;
            default -> UnderlineValue.NONE;
        };
    }
}
