/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import java.util.Arrays;
import java.util.Optional;

/**
 * Class representing a NumberFormat entry. The NumberFormat entry is used to define cell formats like currency or date
 *
 * @author Raphael Stoeckli
 */
public class NumberFormat extends AbstractStyle implements Modifiable {
// ### C O N S T A N T S ###
    /**
     * Start ID for custom number formats as constant
     */
    public static final int CUSTOMFORMAT_START_NUMBER = 164;

// ### E N U M S ###

    /**
     * Enum for the defined number formats
     */
    public enum FormatNumber {
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
         *
         * @param numVal Numeric value of the enum entry
         */
        FormatNumber(int numVal) {
            this.numVal = numVal;
        }

        /**
         * Gets the numeric value of the enum entry
         *
         * @return Numeric value of the enum entry
         */
        public int getValue() {
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
     *
     * @return Format number
     */
    public FormatNumber getNumber() {
        return number;
    }

    /**
     * Sets the format number. Set it to custom (164) in case of custom number formats
     *
     * @param number Format number
     */
    public void setNumber(FormatNumber number) {
        this.number = number;
    }

    /**
     * Gets the format number of the custom format. Must be higher or equal then predefined custom number (164)
     *
     * @return Format number of the custom format
     */
    public int getCustomFormatID() {
        return customFormatID;
    }

    /**
     * Sets the format number of the custom format. Must be higher or equal then predefined custom number (164)
     *
     * @param customFormatID Format number of the custom format
     */
    public void setCustomFormatID(int customFormatID) {
        this.customFormatID = customFormatID;
    }

    /**
     * Gets the custom format code in the notation of Excel
     *
     * @return Custom format code
     */
    public String getCustomFormatCode() {
        return customFormatCode;
    }

    /**
     * Sets the custom format code in the notation of Excel
     *
     * @param customFormatCode Custom format code
     */
    public void setCustomFormatCode(String customFormatCode) {
        this.customFormatCode = customFormatCode;
    }

    /**
     * Gets whether this object is a custom format
     *
     * @return Returns true in case of a custom format (higher or equals 164)
     */
    public boolean isCustomFormat() {
        return number == FormatNumber.custom;
    }

// ### C O N S T R U C T O R S ###

    /**
     * Default constructor
     */
    public NumberFormat() {
        this.number = FormatNumber.none;
        this.customFormatCode = "";
        this.customFormatID = CUSTOMFORMAT_START_NUMBER;
    }

// ### M E T H O D S ###

    /**
     * Override toString method
     *
     * @return String of a class instance
     */
    @Override
    public String toString() {
        return "NumberFormat:" + this.hashCode();
    }

    /**
     * Method to copy the current object to a new one
     *
     * @return Copy of the current object without the internal ID
     */
    @Override
    public NumberFormat copy() {
        NumberFormat copy = new NumberFormat();
        copy.setCustomFormatCode(this.customFormatCode);
        copy.setCustomFormatID(this.customFormatID);
        copy.setNumber(this.number);
        return copy;
    }

    /**
     * Override method to calculate the hash of this component
     *
     * @return Calculated hash as string
     */
    @Override
    public int hashCode() {
        int p = 251;
        int r = 1;
        r *= p + this.customFormatCode.hashCode();
        r *= p + this.customFormatID;
        r *= p + this.number.getValue();
        return r;
    }

    /**
     * Returns whether the current style component was modified (differs from new class instance)
     * @return True if the current object was modified, otherwise false
     */
    @Override
    public boolean isModified() {
        NumberFormat reference = new NumberFormat();
        if(areDifferent(this.customFormatCode, reference.customFormatCode)){
            return true;
        }
        if(areDifferent(this.customFormatID, reference.customFormatID)){
            return true;
        }
        return areDifferent(this.number.numVal, reference.number.numVal); // last statement returns either true or false (whole object not modified)
    }

    /**
     * Tries to parse registered format numbers. If the parsing fails, it is assumed that the number is a custom format number (164 or higher) and 'custom' is returned
     *
     * @param number Raw number to parse<
     * @return Format range. Will return 'invalid' if out of any range (e.g. negative value)
     */
    public static NumberFormatEvaluation tryParseFormatNumber(int number) {
        NumberFormatEvaluation result = new NumberFormatEvaluation();
        Optional<FormatNumber> enumValue = Arrays.stream(FormatNumber.values()).filter(x -> x.getValue() == number).findFirst();
        if (enumValue.isPresent()) {
            result.setFormat(enumValue.get());
            result.setRange(NumberFormatEvaluation.FormatRange.defined_format);
        } else {
            if (number < 0) {
                result.setFormat(FormatNumber.none);
                result.setRange(NumberFormatEvaluation.FormatRange.invalid);
            } else if (number > 0 && number < CUSTOMFORMAT_START_NUMBER) {
                result.setFormat(FormatNumber.none);
                result.setRange(NumberFormatEvaluation.FormatRange.undefined);
            } else {
                result.setFormat(FormatNumber.custom);
                result.setRange(NumberFormatEvaluation.FormatRange.custom_format);
            }
        }
        return result;
    }

    /**
     * Determines whether a defined style format number represents a date (or date and time).<br>
     * Note: Custom number formats (higher than 164), as well as not officially defined numbers (below 164) are currently not considered during the check and will return false
     *
     * @param number Format number to check
     * @return True if the format represents a date, otherwise false
     */
    public static boolean isDateFormat(FormatNumber number) {
        switch (number) {
            case format_14:
            case format_15:
            case format_16:
            case format_17:
            case format_22:
                return true;
            default:
                return false;
        }
    }

    /**
     * Determines whether a defined style format number represents a time).<br>
     * Note: Custom number formats (higher than 164), as well as not officially defined numbers (below 164) are currently not considered during the check and will return false
     *
     * @param number Format number to check
     * @return True if the format represents a time, otherwise false
     */
    public static boolean isTimeFormat(FormatNumber number) {
        switch (number) {
            case format_18:
            case format_19:
            case format_20:
            case format_21:
            case format_45:
            case format_46:
            case format_47:
                return true;
            default:
                return false;
        }
    }

// ### S U B - C L A S S E S ###

    /**
     * Class represents the evaluation of a number format (number) and its purpose
     */
    public static class NumberFormatEvaluation {

        /**
         * Range or validity of the format number
         */
        public enum FormatRange {
            /**
             * Format from 0 to 163 (with gaps)
             */
            defined_format,
            /**
             * Custom defined formats from 164 and higher
             */
            custom_format,
            /**
             * Probably invalid format numbers (e.g. negative value)
             */
            invalid,
            /**
             * Values between 0 and 164 that are not defined as enum value. This may be caused by changes of the OOXML specifications or Excel versions that have encoded loaded files
             */
            undefined
        }

        private FormatNumber format;
        private FormatRange range;

        /**
         * Gets the evaluated Format
         *
         * @return Number format
         */
        public FormatNumber getFormatNumber() {
            return format;
        }

        /**
         * Gets the range of the evaluated number format
         *
         * @return Format range
         */
        public FormatRange getRange() {
            return range;
        }

        /**
         * Sets the evaluated Format
         *
         * @param format Number format
         */
        public void setFormat(FormatNumber format) {
            this.format = format;
        }

        /**
         * Sets the range of the evaluated number format
         *
         * @param range Format range
         */
        public void setRange(FormatRange range) {
            this.range = range;
        }
    }

}
