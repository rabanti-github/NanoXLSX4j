/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.colors.Color;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

/**
 * Class representing a Font entry. The Font entry is used to define text formatting
 *
 * @author Raphael Stoeckli
 */
public class Font extends AbstractStyle {
    // ### C O N S T A N T S ###

    /**
     * The default font name that is declared as Major Font (See {@link ch.rabanti.nanoxlsx4j.styles.Font.SchemeValue})
     */
    public static final String DEFAULT_MAJOR_FONT = "Calibri Light";

    /**
     * The default font name that is declared as Minor Font (See {@link ch.rabanti.nanoxlsx4j.styles.Font.SchemeValue})
     */
    public static final String DEFAULT_MINOR_FONT = "Calibri";

    /**
     * Default font family as constant
     */
    public static final String DEFAULT_FONT_NAME = DEFAULT_MINOR_FONT;

    /**
     * Default font scheme
     */
    public static final SchemeValue DEFAULT_FONT_SCHEME = SchemeValue.minor;

    /**
     * Maximum possible font size
     */
    public static final float MIN_FONT_SIZE = 1f;

    /**
     * Minimum possible font size
     */
    public static final float MAX_FONT_SIZE = 409f;

    /**
     * Default font size
     */
    public static final float DEFAULT_FONT_SIZE = 11f;

    /**
     * Default font family
     */
    public static final FontFamilyValue DEFAULT_FONT_FAMILY = FontFamilyValue.Swiss;

    /**
     * Default vertical alignment
     */
    public static final VerticalTextAlignValue DEFAULT_VERTICAL_ALIGN = VerticalTextAlignValue.none;


    // ### E N U M S ###

    /**
     * Enum for the vertical alignment of the text from baseline
     */
    public enum VerticalTextAlignValue {
        /**
         * Text will be rendered at the baseline and presented in the same size as surrounding text
         */
        baseline,
        /**
         * Text will be rendered as subscript
         */
        subscript,
        /**
         * Text will be rendered as superscript
         */
        superscript,
        /**
         * Text will be rendered normal
         */
        none,
    }

    /**
     * Enum for the font scheme
     */
    public enum SchemeValue {
        /**
         * Font scheme is major
         */
        major(1),
        /**
         * Font scheme is minor (default)
         */
        minor(2),
        /**
         * No Font scheme is used
         */
        none(0);

        private final int value;

        SchemeValue(int value) {
            this.value = value;
        }
    }

    /**
     * Enum for the style of the underline property of a stylized text
     */
    public enum UnderlineValue {
        /**
         * Text contains a single underline
         */
        u_single,
        /**
         * Text contains a double underline
         */
        u_double,
        /**
         * Text contains a single, accounting underline
         */
        singleAccounting,
        /**
         * Text contains a double, accounting underline
         */
        doubleAccounting,
        /**
         * Text contains no underline (default)
         */
        none,
    }

    /**
     * Enum for the charset definitions of a font, according to OOXML specs
     */
    public enum CharsetValue {
        /**
         * Application-defined (any other value than the defined enum values; can be ignored)
         */
        ApplicationDefined(-1),
        /**
         * Charset according to iso-8859-1
         */
        ANSI(0),
        /**
         * Default charset (not defined more specific)
         */
        Default(1),
        /**
         * Symbols from the private Unicode range U+FF00 to U+FFFF
         */
        Symbols(2),
        /**
         * Macintosh charset, Standard Roman
         */
        Macintosh(77),
        /**
         * Shift JIS charset (shift_jis)
         */
        JIS(128),
        /**
         * Hangul charset (ks_c_5601-1987)
         */
        Hangul(129),
        /**
         * Johab charset (KSC-5601-1992)
         */
        Johab(130),
        /**
         * GBK charset (GB-2312)
         */
        GBK(134),
        /**
         * Chinese Big Five charset
         */
        Big5(136),
        /**
         * Greek charset (windows-1253)
         */
        Greek(161),
        /**
         * Turkish charset (iso-8859-9)
         */
        Turkish(162),
        /**
         * Vietnamese charset (windows-1258)
         */
        Vietnamese(163),
        /**
         * Hebrew charset (windows-1255)
         */
        Hebrew(177),
        /**
         * Arabic charset (windows-1256)
         */
        Arabic(178),
        /**
         * Baltic charset (windows-1257)
         */
        Baltic(186),
        /**
         * Russian charset (windows-1251)
         */
        Russian(204),
        /**
         * Thai charset (windows-874)
         */
        Thai(222),
        /**
         * Eastern Europe charset (windows-1250)
         */
        EasternEuropean(238),
        /**
         * OEM characters, not defined by ECMA-376
         */
        OEM(255);

        private final int value;

        CharsetValue(int value) {
            this.value = value;
        }

        /**
         * Gets the numeric value of the charset
         *
         * @return Numeric value
         */
        public int getValue() {
            return value;
        }
    }

    /**
     * Enum for the font family, according to OOXML simple type definition
     */
    public enum FontFamilyValue {
        /**
         * The family is not defined or not applicable
         */
        NotApplicable(0),
        /**
         * The specified font implements a Roman font
         */
        Roman(1),
        /**
         * The specified font implements a Swiss font (default)
         */
        Swiss(2),
        /**
         * The specified font implements a Modern font
         */
        Modern(3),
        /**
         * The specified font implements a Script font
         */
        Script(4),
        /**
         * The specified font implements a Decorative font
         */
        Decorative(5),
        /**
         * Reserved (do not use)
         */
        Reserved1(6),
        /**
         * Reserved (do not use)
         */
        Reserved2(7),
        /**
         * Reserved (do not use)
         */
        Reserved3(8),
        /**
         * Reserved (do not use)
         */
        Reserved4(9),
        /**
         * Reserved (do not use)
         */
        Reserved5(10),
        /**
         * Reserved (do not use)
         */
        Reserved6(11),
        /**
         * Reserved (do not use)
         */
        Reserved7(12),
        /**
         * Reserved (do not use)
         */
        Reserved8(13),
        /**
         * Reserved (do not use)
         */
        Reserved9(14);

        private final int value;

        FontFamilyValue(int value) {
            this.value = value;
        }

        /**
         * Gets the numeric value of the font family
         *
         * @return Numeric value
         */
        public int getValue() {
            return value;
        }
    }

    // ### P R I V A T E F I E L D S ###
    private boolean bold;
    private boolean italic;
    private boolean strike;
    private boolean outline;
    private boolean shadow;
    private boolean condense;
    private boolean extend;
    private UnderlineValue underline;
    private float size;
    // OOXML: Chp.18.8.29
    private String name;
    // OOXML: Chp.18.8.18 and 18.18.94
    private FontFamilyValue family;
    // OOXML: Chp.18.8.3 and 20.1.6.2(p2839ff)
    private Color colorValue;
    private SchemeValue scheme;
    private VerticalTextAlignValue verticalAlign;
    // OOXML: Chp.19.2.1.13
    private CharsetValue charset;

    // ### G E T T E R S & S E T T E R S ###

    /**
     * Gets the bold parameter of the font
     *
     * @return If true, the font is bold
     */
    public boolean isBold() {
        return bold;
    }

    /**
     * Sets the bold parameter of the font
     *
     * @param bold If true, the font is bold
     */
    public void setBold(boolean bold) {
        this.bold = bold;
    }

    /**
     * Gets the italic parameter of the font
     *
     * @return If true, the font is italic
     */
    public boolean isItalic() {
        return italic;
    }

    /**
     * Sets the italic parameter of the font
     *
     * @param italic If true, the font is italic
     */
    public void setItalic(boolean italic) {
        this.italic = italic;
    }

    /**
     * Gets whether the font is struck through
     *
     * @return If true, the font is declared as strike-through
     */
    public boolean isStrike() {
        return strike;
    }

    /**
     * Sets whether the font is struck through
     *
     * @param strike If true, the font is declared as strike-through
     */
    public void setStrike(boolean strike) {
        this.strike = strike;
    }

    /**
     * Gets whether the font has an outline defined
     *
     * @return If true, an outline is rendered around the text
     */
    public boolean isOutline() {
        return outline;
    }

    /**
     * Sets whether the font has an outline defined
     *
     * @param outline If true, an outline is rendered around the text
     */
    public void setOutline(boolean outline) {
        this.outline = outline;
    }

    /**
     * Gets whether the font has a drop shadow
     *
     * @return If true, a shadow is rendered on the text
     */
    public boolean isShadow() {
        return shadow;
    }

    /**
     * Sets whether the font has a drop shadow
     *
     * @param shadow If true, a shadow is rendered on the text
     */
    public void setShadow(boolean shadow) {
        this.shadow = shadow;
    }

    /**
     * Gets whether the font is rendered condensed
     *
     * @return If true, characters are placed closer to each other
     */
    public boolean isCondense() {
        return condense;
    }

    /**
     * Sets whether the font is rendered condensed
     *
     * @param condense If true, characters are placed closer to each other
     */
    public void setCondense(boolean condense) {
        this.condense = condense;
    }

    /**
     * Gets whether the font is rendered extended
     *
     * @return If true, characters are placed more distant to each other
     */
    public boolean isExtend() {
        return extend;
    }

    /**
     * Sets whether the font is rendered extended
     *
     * @param extend If true, characters are placed more distant to each other
     */
    public void setExtend(boolean extend) {
        this.extend = extend;
    }

    /**
     * Gets the underline style of the font
     *
     * @return Underline value
     * @apiNote If set to {@link UnderlineValue#none} no underline will be applied (default)
     */
    public UnderlineValue getUnderline() {
        return underline;
    }

    /**
     * Sets the underline style of the font
     *
     * @param underline Underline value
     * @apiNote If set to {@link UnderlineValue#none} no underline will be applied (default)
     */
    public void setUnderline(UnderlineValue underline) {
        this.underline = underline;
    }

    /**
     * Gets the font size. Valid range is from 1.0 to 409.0
     *
     * @return Font size
     */
    public float getSize() {
        return size;
    }

    /**
     * Sets the Font size. Valid range is from 1 to 409
     *
     * @param size Font size
     */
    public void setSize(float size) {
        if (size < MIN_FONT_SIZE) {
            this.size = MIN_FONT_SIZE;
        }
        else if (size > MAX_FONT_SIZE) {
            this.size = MAX_FONT_SIZE;
        }
        else {
            this.size = size;
        }
    }

    /**
     * Gets the font name (Default is Calibri)
     *
     * @return Font name
     */
    public String getName() {
        return name;
    }

    /**
     * Sets the font name (Default is Calibri)
     *
     * @param name Font name
     * @throws StyleException thrown if the name is null or empty.
     * @apiNote Note that the font name is not validated whether it is a valid or existing font. The font name may not
     * exceed more than 31 characters
     */
    public void setName(String name) {
        this.name = name;
        validateFontScheme();
    }

    /**
     * Gets the font family
     *
     * @return Font family enum value
     */
    public FontFamilyValue getFamily() {
        return family;
    }

    /**
     * Sets the font family
     *
     * @param family Font family enum value
     */
    public void setFamily(FontFamilyValue family) {
        this.family = family;
    }

    /**
     * Gets the color of the font as a {@link Color} compound object
     *
     * @return Font color
     */
    public Color getColorValue() {
        return colorValue;
    }

    /**
     * Sets the color of the font from a {@link Color} compound object.
     * Use {@link Color#createNone()} to omit the color (default).
     *
     * @param colorValue Font color
     */
    public void setColorValue(Color colorValue) {
        this.colorValue = colorValue;
    }

    /**
     * Sets the color of the font from an ARGB or RGB hex string.
     * The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF.
     * To omit the color, use {@link #setColorValue(Color)} with {@link Color#createNone()}.
     *
     * @param colorValue Font color as ARGB hex string
     * @throws StyleException thrown if the passed ARGB value is not valid
     */
    public void setColorValue(String colorValue) {
        Fill.validateColor(colorValue, true, true);
        if (colorValue == null || colorValue.isEmpty()) {
            this.colorValue = Color.createNone();
        } else {
            this.colorValue = Color.createRgb(colorValue);
        }
    }

    /**
     * Gets the font scheme (Default is minor)
     *
     * @return Font scheme
     */
    public SchemeValue getScheme() {
        return scheme;
    }

    /**
     * Sets the Font scheme (Default is minor)
     *
     * @param scheme Font scheme
     */
    public void setScheme(SchemeValue scheme) {
        this.scheme = scheme;
    }

    /**
     * Gets the vertical text alignment of the font (Default is none)
     *
     * @return Vertical text alignment of the font
     */
    public VerticalTextAlignValue getVerticalAlign() {
        return verticalAlign;
    }

    /**
     * Sets the vertical text alignment of the font (Default is none)
     *
     * @param verticalAlign Vertical text alignment of the font
     */
    public void setVerticalAlign(VerticalTextAlignValue verticalAlign) {
        this.verticalAlign = verticalAlign;
    }

    /**
     * Gets the charset of the Font
     *
     * @return Charset enum value
     */
    public CharsetValue getCharset() {
        return charset;
    }

    /**
     * Sets the charset of the Font
     *
     * @param charset Charset enum value
     */
    public void setCharset(CharsetValue charset) {
        this.charset = charset;
    }

    /**
     * Gets whether this object is the default font
     *
     * @return In true the font is equals the default font
     */
    public boolean isDefaultFont() {
        Font temp = new Font();
        return this.equals(temp);
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Default constructor
     */
    public Font() {
        this.size = DEFAULT_FONT_SIZE;
        this.name = DEFAULT_FONT_NAME;
        this.family = DEFAULT_FONT_FAMILY;
        this.colorValue = Color.createNone();
        this.charset = CharsetValue.Default;
        this.scheme = DEFAULT_FONT_SCHEME;
        this.verticalAlign = DEFAULT_VERTICAL_ALIGN;
        this.underline = UnderlineValue.none;
    }

    // ### M E T H O D S ###

    /**
     * Validates the font name and sets the scheme automatically
     */
    private void validateFontScheme() {
        if ((name == null || name.isEmpty()) && !StyleRepository.getInstance().isImportInProgress()) {
            throw new StyleException("The font name was null or empty");
        }
        if (name.equals(DEFAULT_MINOR_FONT)) {
            scheme = SchemeValue.minor;
        }
        else if (name.equals(DEFAULT_MAJOR_FONT)) {
            scheme = SchemeValue.major;
        }
        else {
            scheme = SchemeValue.none;
        }
    }

    /**
     * Override toString method
     *
     * @return String of a class instance
     */
    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("\"Font\": {\n");
        addPropertyAsJson(sb, "Bold", bold);
        addPropertyAsJson(sb, "Charset", charset);
        addPropertyAsJson(sb, "ColorValue", colorValue);
        addPropertyAsJson(sb, "VerticalAlign", verticalAlign);
        addPropertyAsJson(sb, "Family", family);
        addPropertyAsJson(sb, "Italic", italic);
        addPropertyAsJson(sb, "Name", name);
        addPropertyAsJson(sb, "Scheme", scheme);
        addPropertyAsJson(sb, "Size", size);
        addPropertyAsJson(sb, "Strike", strike);
        addPropertyAsJson(sb, "Underline", underline);
        addPropertyAsJson(sb, "Outline", outline);
        addPropertyAsJson(sb, "Shadow", shadow);
        addPropertyAsJson(sb, "Condense", condense);
        addPropertyAsJson(sb, "Extend", extend);
        addPropertyAsJson(sb, "HashCode", this.hashCode(), true);
        sb.append("\n}");
        return sb.toString();
    }

    /**
     * Method to copy the current object to a new one with casting
     *
     * @return Copy of the current object without the internal ID
     */
    public Font copyFont() {
        return (Font) copy();
    }

    /**
     * Method to copy the current object to a new one
     *
     * @return Copy of the current object without the internal ID
     */
    @Override
    public Font copy() {
        Font copy = new Font();
        copy.setBold(this.bold);
        copy.setCharset(this.charset);
        copy.setColorValue(this.colorValue);
        copy.setVerticalAlign(this.verticalAlign);
        copy.setFamily(this.family);
        copy.setItalic(this.italic);
        copy.setName(this.name);
        copy.setScheme(this.scheme);
        copy.setSize(this.size);
        copy.setStrike(this.strike);
        copy.setUnderline(this.underline);
        copy.setOutline(this.outline);
        copy.setShadow(this.shadow);
        copy.setCondense(this.condense);
        copy.setExtend(this.extend);
        return copy;
    }

    /**
     * Override method to calculate the hash of this component
     *
     * @return Calculated hash as string
     * @implNote Note that autogenerated hashcode algorithms may cause collisions. Do not use 0 as fallback value for
     * every field
     */
    @Override
    public int hashCode() {
        int result = (size != +0.0f ? Float.floatToIntBits(size) : 1);
        result = 31 * result + (name != null ? name.hashCode() : 2);
        result = 31 * result + (family != null ? family.hashCode() : 4);
        result = 31 * result + (colorValue != null ? colorValue.hashCode() : 8);
        result = 31 * result + (scheme != null ? scheme.hashCode() : 16);
        result = 31 * result + (verticalAlign != null ? verticalAlign.hashCode() : 32);
        result = 31 * result + (bold ? 1 : 64);
        result = 31 * result + (italic ? 1 : 128);
        result = 31 * result + (underline != null ? underline.hashCode() : 256);
        result = 31 * result + (strike ? 1 : 512);
        result = 31 * result + (charset != null ? charset.hashCode() : 1024);
        result = 31 * result + (outline ? 1 : 2048);
        result = 31 * result + (shadow ? 1 : 4096);
        result = 31 * result + (condense ? 1 : 8192);
        result = 31 * result + (extend ? 1 : 16384);
        return result;
    }

    // ### S T A T I C M E T H O D S ###

    /**
     * Converts a VerticalTextAlignValue enum to its OOXML string representation
     *
     * @param align Enum value to convert
     * @return OOXML string or empty string for none/baseline
     */
    public static String getVerticalTextAlignName(VerticalTextAlignValue align) {
        switch (align) {
            case baseline: return "baseline";
            case subscript: return "subscript";
            case superscript: return "superscript";
            default: return "";
        }
    }

    /**
     * Converts an OOXML string to its corresponding VerticalTextAlignValue enum
     *
     * @param name OOXML string value
     * @return Corresponding enum, or {@link VerticalTextAlignValue#none} if not found
     */
    public static VerticalTextAlignValue getVerticalTextAlignEnum(String name) {
        if (name == null) return VerticalTextAlignValue.none;
        switch (name) {
            case "baseline": return VerticalTextAlignValue.baseline;
            case "subscript": return VerticalTextAlignValue.subscript;
            case "superscript": return VerticalTextAlignValue.superscript;
            default: return VerticalTextAlignValue.none;
        }
    }

    /**
     * Converts an UnderlineValue enum to its OOXML string representation
     *
     * @param underline Enum value to convert
     * @return OOXML string or empty string for single/none
     */
    public static String getUnderlineName(UnderlineValue underline) {
        switch (underline) {
            case u_double: return "double";
            case singleAccounting: return "singleAccounting";
            case doubleAccounting: return "doubleAccounting";
            default: return "";
        }
    }

    /**
     * Converts an OOXML string to its corresponding UnderlineValue enum
     *
     * @param name OOXML string value
     * @return Corresponding enum, or {@link UnderlineValue#none} if not found
     */
    public static UnderlineValue getUnderlineEnum(String name) {
        if (name == null) return UnderlineValue.none;
        switch (name) {
            case "double": return UnderlineValue.u_double;
            case "singleAccounting": return UnderlineValue.singleAccounting;
            case "doubleAccounting": return UnderlineValue.doubleAccounting;
            default: return UnderlineValue.none;
        }
    }

    /**
     * Converts a numeric string (from XML) to its corresponding FontFamilyValue enum
     *
     * @param numericValue Numeric string value
     * @return Corresponding enum, or {@link FontFamilyValue#NotApplicable} if not found
     */
    public static FontFamilyValue getFontFamilyEnum(String numericValue) {
        if (numericValue == null) return FontFamilyValue.NotApplicable;
        try {
            int val = Integer.parseInt(numericValue);
            for (FontFamilyValue f : FontFamilyValue.values()) {
                if (f.getValue() == val) return f;
            }
        }
        catch (NumberFormatException e) {
            // fall through
        }
        return FontFamilyValue.NotApplicable;
    }

    /**
     * Converts a numeric string (from XML) to its corresponding CharsetValue enum
     *
     * @param numericValue Numeric string value
     * @return Corresponding enum, or {@link CharsetValue#ApplicationDefined} if not found
     */
    public static CharsetValue getCharsetEnum(String numericValue) {
        if (numericValue == null) return CharsetValue.ApplicationDefined;
        try {
            int val = Integer.parseInt(numericValue);
            for (CharsetValue c : CharsetValue.values()) {
                if (c.getValue() == val) return c;
            }
        }
        catch (NumberFormatException e) {
            // fall through
        }
        return CharsetValue.ApplicationDefined;
    }

}
