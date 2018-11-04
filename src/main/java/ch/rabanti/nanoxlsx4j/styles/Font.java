/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import static ch.rabanti.nanoxlsx4j.styles.StyleManager.FONTPREFIX;

/**
 * Class representing a Font entry. The Font entry is used to define text formatting
 * @author Raphael Stoeckli
 */
public class Font extends AbstractStyle
{
// ### E N U M S ###
    /**
     * Default font family as constant
     */
    public static final String DEFAULTFONT = "Calibri";
    
    /**
     * Enum for the vertical alignment of the text from base line
     */
    public enum VerticalAlignValue
    {
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
        VerticalAlignValue(int value) { this.value = value; }
        public int getValue() { return value; }        
    }    
    
    /**
     * Enum for the font scheme
     */
    public enum SchemeValue
    {
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
        SchemeValue(int value) { this.value = value; }
        public int getValue() { return value; }  
    }    
    
// ### P R I V A T E  F I E L D S ###
    private int size;
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
     * Gets the font size. Valid range is from 8 to 75
     * @return Font size
     */
    public int getSize() {
        return size;
    }

    /**
     * Sets the Font size. Valid range is from 8 to 75
     * @param size Font size
     */
    public void setSize(int size) {
        if (size < 8) { this.size = 8; }
        else if (size > 75) { this.size = 72; }
        else { this.size = size; }
    }

    /**
     * Gets the font name (Default is Calibri)
     * @return Font name
     */
    public String getName() {
        return name;
    }

    /**
     * Sets the font name (Default is Calibri)
     * @param name Font name
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * Gets the font family (Default is 2)
     * @return Font family
     */
    public String getFamily() {
        return family;
    }

    /**
     * Sets the font family (Default is 2)
     * @param family Font family
     */
    public void setFamily(String family) {
        this.family = family;
    }

    /**
     * Gets the font color theme (Default is 1)
     * @return Font color theme
     */
    public int getColorTheme() {
        return colorTheme;
    }

    /**
     * Sets the font color theme (Default is 1)
     * @param colorTheme Font color theme
     */
    public void setColorTheme(int colorTheme) {
        this.colorTheme = colorTheme;
    }

    /**
     * Gets the Font color (default is empty)
     * @return Font color
     */
    public String getColorValue() {
        return colorValue;
    }

    /**
     * Sets the font color (default is empty)
     * @param colorValue Font color
     */
    public void setColorValue(String colorValue) {
        this.colorValue = colorValue;
    }

    /**
     * Gets the font scheme (Default is minor)
     * @return Font scheme
     */
    public SchemeValue getScheme() {
        return scheme;
    }

    /**
     * Sets the Font scheme (Default is minor)
     * @param scheme Font scheme
     */
    public void setScheme(SchemeValue scheme) {
        this.scheme = scheme;
    }

    /**
     * Gets the alignment of the font (Default is none)
     * @return Alignment of the font
     */
    public VerticalAlignValue getVerticalAlign() {
        return verticalAlign;
    }

    /**
     * Sets the Alignment of the font (Default is none)
     * @param verticalAlign Alignment of the font
     */
    public void setVerticalAlign(VerticalAlignValue verticalAlign) {
        this.verticalAlign = verticalAlign;
    }

    /**
     * Gets the bold parameter of the font
     * @return If true, the font is bold
     */
    public boolean isBold() {
        return bold;
    }

    /**
     * Sets the bold parameter of the font
     * @param bold If true, the font is bold
     */
    public void setBold(boolean bold) {
        this.bold = bold;
    }

    /**
     * Gets the italic parameter of the font
     * @return If true, the font is italic
     */
    public boolean isItalic() {
        return italic;
    }

    /**
     * Sets the italic parameter of the font
     * @param italic If true, the font is italic
     */
    public void setItalic(boolean italic) {
        this.italic = italic;
    }

    /**
     * Gets the underline parameter of the font
     * @return If true, the font as one underline
     */
    public boolean isUnderline() {
        return underline;
    }

    /**
     * Sets the underline parameter of the font
     * @param underline If true, the font as one underline
     */
    public void setUnderline(boolean underline) {
        this.underline = underline;
    }

    /**
     * Gets the double-underline parameter of the font
     * @return If true, the font ha a double underline
     */
    public boolean isDoubleUnderline() {
        return doubleUnderline;
    }

    /**
     * Sets the double-underline parameter of the font
     * @param doubleUnderline If true, the font ha a double underline
     */
    public void setDoubleUnderline(boolean doubleUnderline) {
        this.doubleUnderline = doubleUnderline;
    }

    /**
     * Gets whether the font is struck through
     * @return If true, the font is declared as strike-through
     */
    public boolean isStrike() {
        return strike;
    }

    /**
     * Sets whether the font is struck through
     * @param strike If true, the font is declared as strike-through
     */
    public void setStrike(boolean strike) {
        this.strike = strike;
    }

    /**
     * Gets the charset of the Font (Default is empty)
     * @return Charset of the Font
     */
    public String getCharset() {
        return charset;
    }

    /**
     * Sets the charset of the Font (Default is empty)
     * @param charset Charset of the Font
     */
    public void setCharset(String charset) {
        this.charset = charset;
    }
    
    /**
     * Gets whether this object is the default font
     * @return In true the font is equals the default font
     */
    public boolean isDefaultFont()
    {
        Font temp = new Font();
        return this.equals(temp); 
    }
    
// ### C O N S T R U C T O R S ###   
    /**
     * Default constructor
     */
    public Font()
    {
        this.size = 11;
        this.name = DEFAULTFONT;
        this.family = "2";
        this.colorTheme = 1;
        this.colorValue = "";
        this.charset = "";
        this.scheme = SchemeValue.minor;
        this.verticalAlign = VerticalAlignValue.none;
    }    
    
     /**
     * Override toString method
     * @return String of a class instance
     */
    @Override
    public String toString()
    {
        return this.getHash();
    }
    
    /**
     * Method to copy the current object to a new one
     * @return Copy of the current object without the internal ID
     */          
    @Override
    public Font copy()
    {
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
     * @return Calculated hash as string
     */
    @Override
    String calculateHash()
    {
        StringBuilder sb = new StringBuilder();
        sb.append(FONTPREFIX);
        castValue(this.bold, sb, ':');
        castValue(this.italic, sb, ':');
        castValue(this.underline, sb, ':');
        castValue(this.doubleUnderline, sb, ':');
        castValue(this.strike, sb, ':');
        castValue(this.colorTheme, sb, ':');
        castValue(this.colorValue, sb, ':');
        castValue(this.family, sb, ':');
        castValue(this.name, sb, ':');
        castValue(this.scheme.getValue(), sb, ':');
        castValue(this.verticalAlign.getValue(), sb, ':');
        castValue(this.charset, sb, ':');
        castValue(this.size, sb, null);
        return sb.toString();
    }
    
    
}
