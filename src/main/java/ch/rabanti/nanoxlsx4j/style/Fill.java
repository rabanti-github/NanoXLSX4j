/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.style;

import static ch.rabanti.nanoxlsx4j.style.StyleManager.FILLPREFIX;

/**
 *
 * @author Raphael Stoeckli
 */
public class Fill extends AbstractStyle
{
// ### C O N S T A N T S ###
    public static final String DEFAULTCOLOR = "FF000000";
// ### E N U M S ###
    /**
     * Enum for the type of the fill pattern
     */
    public enum PatternValue
    {
        /**
         * No pattern (default)
         */
        none(0),
        /**
        * Solid fill (for colors)
        */
        solid(1),
        /**
        * Dark gray fill
        */
        darkGray(2),
        /**
        * Medium gray fill
        */
        mediumGray(3),
        /**
        * Light gray fill
        */
        lightGray(4),
        /**
        * 6.25% gray fill
        */
        gray0625(5),
        /**
        * 12.5% gray fill
        */
        gray125(6);
        
        private final int value;
        PatternValue(int value) { this.value = value; }
        public int getValue() { return value; }
        
    }
    
    /**
     * Enum for the type of the fill
     */
    public enum FillType
    {
        /**
        * Color defines a pattern color
        */
        patternColor(0),
        /**
         * Color defines a solid fill color
        */
        fillColor(1);
      
        private final int value;
        FillType(int value) { this.value = value; }
        public int getValue() { return value; }
    }    
    
    /**
     * Gets the pattern name from the enum
     * @param pattern Enum to process
     * @return The valid value of the pattern as String
     */
    public static String getPatternName(PatternValue pattern)
    {
        String output;
        switch (pattern)
        {
            case none:
                output = "none";
                break;
            case solid:
                output = "solid";
                break;
            case darkGray:
                output = "darkGray";
                break;
            case mediumGray:
                output = "mediumGray";
                break;
            case lightGray:
                output = "lightGray";
                break;
            case gray0625:
                output = "gray0625";
                break;
            case gray125:
                output = "gray125";
                break;
            default:
                output = "none";
                break;
        }
        return output;
    }    

// ### P R I V A T E  F I E L D S ###
    public int indexedColor;
    public PatternValue patternFill;
    public String foregroundColor;
    public String backgroundColor;

// ### G E T T E R S  &  S E T T E R S ###
    /**
     * Gets the indexed color (Default is 64)
     * @return Indexed color
     */
    public int getIndexedColor() {
        return indexedColor;
    }

    /**
     * Sets the indexed color (Default is 64)
     * @param indexedColor Indexed color
     */
    public void setIndexedColor(int indexedColor) {
        this.indexedColor = indexedColor;
    }

    /**
     * Gets the pattern type of the fill (Default is none)
     * @return Pattern type of the fill
     */
    public PatternValue getPatternFill() {
        return patternFill;
    }

    /**
     * Sets the pattern type of the fill (Default is none)
     * @param patternFill Pattern type of the fill
     */
    public void setPatternFill(PatternValue patternFill) {
        this.patternFill = patternFill;
    }

    /**
     * Gets the foreground color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
     * @return Foreground color of the fill
     */
    public String getForegroundColor() {
        return foregroundColor;
    }

    /**
     * Sets the foreground color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
     * @param foregroundColor Foreground color of the fill
     */
    public void setForegroundColor(String foregroundColor) {
        this.foregroundColor = foregroundColor;
    }

    /**
     * Gets the Background color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
     * @return Background color of the fill
     */
    public String getBackgroundColor() {
        return backgroundColor;
    }

    /**
     * Sets the background color of the fill. The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF
     * @param backgroundColor Background color of the fill
     */
    public void setBackgroundColor(String backgroundColor) {
        this.backgroundColor = backgroundColor;
    }
    
// ### C O N S T R U C T O R S ###   
    /**
     * Default constructor
     */
    public Fill()
    {
        this.indexedColor = 64;
        this.patternFill = PatternValue.none;
        this.foregroundColor = DEFAULTCOLOR;
        this.backgroundColor = DEFAULTCOLOR;
    }  
    
    /**
     * Constructor with foreground and background color
     * @param foreground Foreground color of the fill
     * @param background Background color of the fill
     */
    public Fill(String foreground, String background)
    {
        this.backgroundColor = background;
        this.foregroundColor = foreground;
        this.indexedColor = 64;
        this.patternFill = PatternValue.solid;
    }
    
    /**
     * Constructor with color value and fill type
     * @param value Color value
     * @param fillType Fill type (fill or pattern)
     */
    public Fill(String value, FillType fillType)
    {
        if (fillType == FillType.fillColor)
        {
            this.backgroundColor = value;
            this.foregroundColor = DEFAULTCOLOR;
        }
        else
        {
            this.backgroundColor = DEFAULTCOLOR;
            this.foregroundColor = value;
        }
        this.indexedColor = 64;
        this.patternFill = PatternValue.solid;
    }
    
// ### M E T H O D S ###
    /**
     * Sets the color and the depending fill type
     * @param value Color value
     * @param fillType Fill type (fill or pattern)
     */
    public void SetColor(String value, FillType fillType)
    {
        if (fillType == FillType.fillColor)
        {
            this.foregroundColor = value;
            this.backgroundColor = DEFAULTCOLOR;
        }
        else
        {
            this.foregroundColor = DEFAULTCOLOR;
            this.backgroundColor = value;
        }
        this.patternFill = PatternValue.solid;
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
    public Fill copy()
    {
        Fill copy = new Fill();
        copy.setBackgroundColor(this.backgroundColor);
        copy.setForegroundColor(this.foregroundColor);
        copy.setIndexedColor(this.indexedColor);
        copy.setPatternFill(this.patternFill);
        return copy;
    }  
    
    /**
     * Override method to calculate the hash of this component
     * @return Calculated hash as string
     */
    @Override
    String calculateHash() {
        
        StringBuilder sb = new StringBuilder();
        sb.append(FILLPREFIX);        
        castValue(this.indexedColor, sb, ':');
        castValue(this.patternFill.getValue(), sb, ':');
        castValue(this.foregroundColor, sb, ':');
        castValue(this.backgroundColor, sb, null);
        return sb.toString();
    }    
    
    
}
