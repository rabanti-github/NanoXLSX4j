/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.style;

/**
 * Factory class with the most important predefined styles
 * @author Raphael stoeckli
 */
public final class BasicStyles {
    
// ### E N U M S ###
    /**
     * Enum with style selection
     */
    private enum StyleEnum
    {
        /**
        * Format text bold
        */
        bold,
        /**
        * Format text italic
        */
        italic,
        /**
        * Format text bold and italic
        */
        boldItalic,
        /**
        * Format text with an underline
        */
        underline,
        /**
        * Format text with a double underline
        */
        doubleUnderline,
        /**
        * Format text with a strike-through
        */
        strike,
        /**
        * Format number as date
        */
        dateFormat,
        /**
        * Rounds number as an integer
        */
        roundFormat,
        /**
        * Format cell with a thin border
        */
        borderFrame,
        /**
        * Format cell with a thin border and a thick bottom line as header cell
        */
        borderFrameHeader,
        /**
        * Special pattern fill style for compatibility purpose 
        */
        dottedFill_0_125,
        /**
         * Style to apply on merged cells
         */
        mergeCellStyle,
    }
// ### P R I V A T E   S T A T I C  F I E L D S ###      
     private static Style bold, italic, boldItalic, underline, doubleUnderline, strike, dateFormat, roundFormat, borderFrame, borderFrameHeader, dottedFill_0_125, mergeCellStyle;
     
// ### P U B L I C   S T A T I C  F I E L D S ### 
    /**
    * Gets the bold style
    * @return Style object
    */
    public static Style Bold()
    { return getStyle(StyleEnum.bold);}
    /**
    * Gets the italic style
    * @return Style object
    */
    public static Style Italic()
    { return getStyle(StyleEnum.italic);}
    /**
    * Gets the bold and italic style
    * @return Style object
    */
    public static Style BoldItalic()
    { return getStyle(StyleEnum.boldItalic);}
    /**
    * Gets the underline style
    * @return Style object
    */
    public static Style Underline()
    { return getStyle(StyleEnum.underline);}
    /**
    * Gets the double underline style
    * @return Style object
    */
    public static Style DoubleUnderline()
    { return getStyle(StyleEnum.doubleUnderline);}
    /**
    * Gets the strike style
    * @return Style object
    */
    public static Style Strike()
    { return getStyle(StyleEnum.strike);}
    /**
    * Gets the date format style
    * @return Style object
    */
    public static Style DateFormat()
    { return getStyle(StyleEnum.dateFormat);}
    /**
    * Gets the round format style
    * @return Style object
    */
    public static Style RoundFormat()
    { return getStyle(StyleEnum.roundFormat);}
    /**
    * Gets the border frame style
    * @return Style object
    */
    public static Style BorderFrame()
    { return getStyle(StyleEnum.borderFrame);}
    /**
    * Gets the border style for header cells
    * @return Style object
    */
    public static Style BorderFrameHeader()
    { return getStyle(StyleEnum.borderFrameHeader);}
    /**
    * Gets the special pattern fill style (for compatibility)
    * @return Style object
    */
    public static Style DottedFill_0_125()
    { return getStyle(StyleEnum.dottedFill_0_125);} 

    /**
     * Gets the style used when merging cells
     * @return Style object
     */
    public static Style MergeCellStyle()
    { return getStyle(StyleEnum.mergeCellStyle);} 
    
// ### S T A T I C   M E T H O D S ###
    /**
     * Method to maintain the styles and to create singleton instances
     * @param value Enum value to maintain
     * @return The style according to the passed enum value
     */
    private static Style getStyle(StyleEnum value)
    {
        Style s = null;
        switch (value)
        {
            case bold:
                if (bold == null)
                {
                    bold = new Style();
                    bold.getFont().setBold(true);
                }
                s = bold;
                break;
            case italic:
                if (italic == null)
                {
                    italic = new Style();
                    italic.getFont().setItalic(true);
                }
                s = italic;
                break;
            case boldItalic:
                if (boldItalic == null)
                {
                    boldItalic = new Style();
                    boldItalic.getFont().setItalic(true);
                    boldItalic.getFont().setBold(true);
                }
                s = boldItalic;
                break;
            case underline:
                if (underline == null)
                {
                    underline = new Style();
                    underline.getFont().setUnderline(true);
                }
                s = underline;
                break;
            case doubleUnderline:
                if (doubleUnderline == null)
                {
                    doubleUnderline = new Style();
                    doubleUnderline.getFont().setDoubleUnderline(true);
                }
                s = doubleUnderline;
                break;
            case strike:
                if (strike == null)
                {
                    strike = new Style();
                    strike.getFont().setStrike(true);
                }
                s = strike;
                break;
            case dateFormat:
                if (dateFormat == null)
                {
                    dateFormat = new Style();
                    dateFormat.getNumberFormat().setNumber(NumberFormat.FormatNumber.format_14);
                }
                s = dateFormat;
                break;
            case roundFormat:
                if (roundFormat == null)
                {
                    roundFormat = new Style();
                    roundFormat.getNumberFormat().setNumber(NumberFormat.FormatNumber.format_1);
                }
                s = roundFormat;
                break;
            case borderFrame:
                if (borderFrame == null)
                {
                    borderFrame = new Style();
                    borderFrame.getBorder().setTopStyle(Border.StyleValue.thin);
                    borderFrame.getBorder().setBottomStyle(Border.StyleValue.thin);
                    borderFrame.getBorder().setLeftStyle(Border.StyleValue.thin);
                    borderFrame.getBorder().setRightStyle(Border.StyleValue.thin);
                }
                s = borderFrame;
                break;
            case borderFrameHeader:
                if (borderFrameHeader == null)
                {
                    borderFrameHeader = new Style();
                    borderFrameHeader.getBorder().setTopStyle(Border.StyleValue.thin);
                    borderFrameHeader.getBorder().setBottomStyle(Border.StyleValue.medium);
                    borderFrameHeader.getBorder().setLeftStyle(Border.StyleValue.thin);
                    borderFrameHeader.getBorder().setRightStyle(Border.StyleValue.thin);
                    borderFrameHeader.getFont().setBold(true);
                }
                s = borderFrameHeader;
                break;
            case dottedFill_0_125:
                if (dottedFill_0_125 == null)
                {
                    dottedFill_0_125 = new Style();
                    dottedFill_0_125.getFill().setPatternFill(Fill.PatternValue.gray125);
                }
                s = dottedFill_0_125;
                break;
                
            case mergeCellStyle:
                if (mergeCellStyle == null)
                {
                    mergeCellStyle = new Style();
                    mergeCellStyle.getCellXf().setForceApplyAlignment(true);
                }
                s = mergeCellStyle;
                break;                
                
            default:
                break;
        }
        return s.copyStyle(); // Copy makes basic styles immutable
    }     
     
     
}
