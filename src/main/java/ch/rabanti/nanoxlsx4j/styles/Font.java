/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

/**
 * Class representing a Font entry. The Font entry is used to define text formatting
 *
 * @author Raphael Stoeckli
 */
public class Font extends AbstractStyle{
// ### E N U M S ###
    /**
     * Default font family as constant
     */
    public static final String DEFAULT_FONT_NAME = "Calibri";

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
    public static final String DEFAULT_FONT_FAMILY = "2";

    /**
     * Default font scheme
     */
    public static final SchemeValue DEFAULT_FONT_SCHEME = SchemeValue.minor;

    /**
     * Default vertical alignment
     */
    public static final VerticalAlignValue DEFAULT_VERTICAL_ALIGN = VerticalAlignValue.none;

    /**
     * Enum for the vertical alignment of the text from base line
     */
    public enum VerticalAlignValue {
        // baseline, // Maybe not used in Excel
        /**
         * Text will be rendered as subscript
         */
        subscript(1),
        /**
         * Text will be rendered as superscript
         */
        superscript(2),
        /**
         * Text will be rendered normal
         */
        none(0);

        private final int value;

        VerticalAlignValue(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }
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

        public int getValue() {
            return value;
        }
    }

    // ### P R I V A T E  F I E L D S ###
    private float size;
    private String name;
    private String family;
    private int colorTheme;
    private String colorValue;
    private SchemeValue scheme;
    private VerticalAlignValue verticalAlign;
    private boolean bold;
    private boolean italic;
    private boolean underline;
    private boolean doubleUnderline;
    private boolean strike;
    private String charset;

// ### G E T T E R S  &  S E T T E R S ###

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
        if (size < MIN_FONT_SIZE ) {
            this.size = MIN_FONT_SIZE;
        } else if (size > MAX_FONT_SIZE) {
            this.size = MAX_FONT_SIZE;
        } else {
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
     * @apiNote Note that the font name is not validated whether it is a valid or existing font
     */
    public void setName(String name) {
        if (name == null || name.isEmpty()){
            throw new StyleException("A general style exception occurred", "The font name was null or empty");
        }
        this.name = name;
    }

    /**
     * Gets the font family (Default is 2)
     *
     * @return Font family
     */
    public String getFamily() {
        return family;
    }

    /**
     * Sets the font family (Default is 2)
     *
     * @param family Font family
     */
    public void setFamily(String family) {
        this.family = family;
    }

    /**
     * Gets the font color theme (Default is 1)
     *
     * @return Font color theme
     */
    public int getColorTheme() {
        return colorTheme;
    }

    /**
     * Sets the font color theme (Default is 1)
     *
     * @param colorTheme Font color theme
     * @throws StyleException thrown if the number is below 1
     */
    public void setColorTheme(int colorTheme) {
        if (colorTheme < 1)
        {
            throw new StyleException( StyleException.GENERAL, "The color theme number " + colorTheme + " is invalid. Should be >0");
        }
        this.colorTheme = colorTheme;
    }

    /**
     * Gets the Font color (default is empty)
     *
     * @return Font color
     */
    public String getColorValue() {
        return colorValue;
    }

    /**
     * Sets the color code of the font color. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF.<br>
     * To omit the color, an empty string can be set. Empty is also default.
     *
     * @param colorValue Font color
     * @throws StyleException thrown if the passed ARGB value is not valid
     */
    public void setColorValue(String colorValue) {
        Fill.validateColor(colorValue, true, true);
        this.colorValue = colorValue;
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
     * Gets the alignment of the font (Default is none)
     *
     * @return Alignment of the font
     */
    public VerticalAlignValue getVerticalAlign() {
        return verticalAlign;
    }

    /**
     * Sets the Alignment of the font (Default is none)
     *
     * @param verticalAlign Alignment of the font
     */
    public void setVerticalAlign(VerticalAlignValue verticalAlign) {
        this.verticalAlign = verticalAlign;
    }

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
     * Gets the underline parameter of the font
     *
     * @return If true, the font as one underline
     */
    public boolean isUnderline() {
        return underline;
    }

    /**
     * Sets the underline parameter of the font
     *
     * @param underline If true, the font as one underline
     */
    public void setUnderline(boolean underline) {
        this.underline = underline;
    }

    /**
     * Gets the double-underline parameter of the font
     *
     * @return If true, the font ha a double underline
     */
    public boolean isDoubleUnderline() {
        return doubleUnderline;
    }

    /**
     * Sets the double-underline parameter of the font
     *
     * @param doubleUnderline If true, the font ha a double underline
     */
    public void setDoubleUnderline(boolean doubleUnderline) {
        this.doubleUnderline = doubleUnderline;
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
     * Gets the charset of the Font (Default is empty)
     *
     * @return Charset of the Font
     */
    public String getCharset() {
        return charset;
    }

    /**
     * Sets the charset of the Font (Default is empty)
     *
     * @param charset Charset of the Font
     */
    public void setCharset(String charset) {
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
        this.colorTheme = 1;
        this.colorValue = "";
        this.charset = "";
        this.scheme = DEFAULT_FONT_SCHEME;
        this.verticalAlign = DEFAULT_VERTICAL_ALIGN;
    }

    /**
     * Override toString method
     *
     * @return String of a class instance
     */
    @Override
    public String toString() {
        return "Font:" + this.hashCode();
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
        copy.setColorTheme(this.colorTheme);
        copy.setColorValue(this.colorValue);
        copy.setVerticalAlign(this.verticalAlign);
        copy.setDoubleUnderline(this.doubleUnderline);
        copy.setFamily(this.family);
        copy.setItalic(this.italic);
        copy.setName(this.name);
        copy.setScheme(this.scheme);
        copy.setSize(this.size);
        copy.setStrike(this.strike);
        copy.setUnderline(this.underline);
        return copy;
    }

    /**
     * Override method to calculate the hash of this component
     *
     * @return Calculated hash as string
     */
    @Override
    public int hashCode() {
        int p = 257;
        int r = 1;
        r *= p + (this.bold ? 0 : 1);
        r *= p + (this.italic ? 0 : 2);
        r *= p + (this.underline ? 0 : 4);
        r *= p + (this.doubleUnderline ? 0 : 8);
        r *= p + (this.strike ? 0 : 16);
        r *= p + this.colorTheme;
        r *= p + this.colorValue.hashCode();
        r *= p + this.family.hashCode();
        r *= p + this.name.hashCode();
        r *= p + this.scheme.getValue();
        r *= p + this.verticalAlign.value;
        r *= p + this.charset.hashCode();
        r *= p + Float.hashCode(this.size);
        return r;
    }

}
