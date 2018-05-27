/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.style;

import static ch.rabanti.nanoxlsx4j.style.StyleManager.NUMBERFORMATPREFIX;


/**
 * Class representing a NumberFormat entry. The NumberFormat entry is used to define cell formats like currency or date
 * @author Raphael Stoeckli
 */
public class NumberFormat extends AbstractStyle
{
// ### C O N S T A N T S ###
    /**
     * Start ID for custom number formats as constant
     */
    public static final int CUSTOMFORMAT_START_NUMBER = 124;
    
// ### E N U M S ###
    /**
     * Enum for the defined number formats
     */
    public enum FormatNumber
    {
        /**
        * No format / Default
        */
        none(0),
        /**
        * Format: 0
        */
        format_1(1),
        /**
        * Format: 0.00
        */
        format_2(2),
        /**
        * Format: #,##0
        */
        format_3(3),
        /**
        * Format: #,##0.00
        */
        format_4(4),
        /**
        * Format: $#,##0_);($#,##0)
        */
        format_5(5),
        /**
        * Format: $#,##0_);[Red]($#,##0)
        */
        format_6(6),
        /**
        * Format: $#,##0.00_);($#,##0.00)
        */
        format_7(7),
        /**
        * Format: $#,##0.00_);[Red]($#,##0.00)
        */
        format_8(8),
        /**
        * Format: 0%
        */
        format_9(9),
        /**
        * Format: 0.00%
        */
        format_10(10),
        /**
        * Format: 0.00E+00
        */
        format_11(11),
        /**
        * Format: # ?/?
        */
        format_12(12),
        /**
        * Format: # ??/??
        */
        format_13(13),
        /**
        * Format: m/d/yyyy
        */
        format_14(14),
        /**
        * Format: d-mmm-yy
        */
        format_15(15),
        /**
        * Format: d-mmm
        */
        format_16(16),
        /**
        * Format: mmm-yy
        */
        format_17(17),
        /**
        * Format: mm AM/PM
        */
        format_18(18),
        /**
        * Format: h:mm:ss AM/PM
        */
        format_19(19),
        /**
        * Format: h:mm
        */
        format_20(20),
        /**
        * Format: h:mm:ss
        */
        format_21(21),
        /**
        * Format: m/d/yyyy h:mm
        */
        format_22(22),
        /**
        * Format: #,##0_);(#,##0)
        */
        format_37(37),
        /**
        * Format: #,##0_);[Red](#,##0)
        */
        format_38(38),
        /**
        * Format: #,##0.00_);(#,##0.00)
        */
        format_39(39),
        /**
        * Format: #,##0.00_);[Red](#,##0.00)
        */
        format_40(40),
        /**
        * Format: mm:ss
        */
        format_45(45),
        /**
        * Format: [h]:mm:ss
        */
        format_46(46),
        /**
        * Format: mm:ss.0
        */
        format_47(47),
        /**
        * Format: ##0.0E+0
        */
        format_48(48),
        /**
        * Format: #
        */
        format_49(49),
        /**
        * Custom Format (ID 164 and higher)
        */
        custom(164);

        private final int numVal;
        /**
         * Enum constructor with numeric value
         * @param numVal Numeric value of the enum entry
         */
        FormatNumber(int numVal)
        {
            this.numVal = numVal;
        }

        /**
         * Gets the numeric value of the enum entry
         * @return Numeric value of the enum entry
         */
        public int getValue()
        {
            return numVal;
        }
    }
    
// ### P R I V A T E  F I E L D S ###
    private FormatNumber number;
    private int customFormatID;
    private String customFormatCode;

// ### G E T T E R S  &  S E T T E R S ###
    /**
     * Gets the format number. Set it to custom (164) in case of custom number formats
     * @return Format number
     */
    public FormatNumber getNumber() {
        return number;
    }

    /**
     * Sets the format number. Set it to custom (164) in case of custom number formats
     * @param number Format number
     */
    public void setNumber(FormatNumber number) {
        this.number = number;
    }

    /**
     * Gets the format number of the custom format. Must be higher or equal then predefined custom number (164)
     * @return Format number of the custom format
     */
    public int getCustomFormatID() {
        return customFormatID;
    }

    /**
     * Sets the format number of the custom format. Must be higher or equal then predefined custom number (164)
     * @param customFormatID Format number of the custom format
     */
    public void setCustomFormatID(int customFormatID) {
        this.customFormatID = customFormatID;
    }

    /**
     * Gets the custom format code in the notation of Excel
     * @return Custom format code
     */
    public String getCustomFormatCode() {
        return customFormatCode;
    }

    /**
     * Sets the custom format code in the notation of Excel
     * @param customFormatCode Custom format code
     */
    public void setCustomFormatCode(String customFormatCode) {
        this.customFormatCode = customFormatCode;
    }
    
    /**
     * Gets whether this object is a custom format
     * @return Returns true in case of a custom format (higher or equals 164)
     */
    public boolean isCustomFormat()
    {
        return number == FormatNumber.custom;
    }
 
// ### C O N S T R U C T O R S ###   
    /**
     * Default constructor
     */
    public NumberFormat()
    {
        this.number = FormatNumber.none;
        this.customFormatCode = "";
        this.customFormatID = CUSTOMFORMAT_START_NUMBER;
    }  
    
// ### M E T H O D S ### 
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
    public NumberFormat copy()
    {
        NumberFormat copy = new NumberFormat();
        copy.setCustomFormatCode(this.customFormatCode);
        copy.setCustomFormatID(this.customFormatID);
        copy.setNumber(this.number);
        return copy;
    }

     /**
     * Override method to calculate the hash of this component
     * @return Calculated hash as string
     */
    @Override
    String calculateHash() {
        StringBuilder sb = new StringBuilder();
        sb.append(NUMBERFORMATPREFIX);
        castValue(this.customFormatCode, sb, ':');
        castValue(this.customFormatID, sb, ':');
        castValue(this.number.getValue(), sb, null);
        return sb.toString();
    }
    
    
}
