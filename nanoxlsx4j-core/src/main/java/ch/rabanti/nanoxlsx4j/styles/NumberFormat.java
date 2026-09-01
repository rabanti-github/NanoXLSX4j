/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import java.util.Objects;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

/**
 * Class representing a NumberFormat entry. The NumberFormat entry is used to define cell formats such as currency or
 * date formats.
 */
public class NumberFormat extends AbstractStyle {

    /** Start ID for custom number formats. */
    public static final int CUSTOM_FORMAT_START_NUMBER = 164;

    /** Default number format. */
    public static final FormatNumber DEFAULT_NUMBER = FormatNumber.NONE;

    /** Enum for predefined number formats. */
    public enum FormatNumber {
        /** No format or default format. */
        NONE(0),
        /** Format: 0. */
        FORMAT_1(1),
        /** Format: 0.00. */
        FORMAT_2(2),
        /** Format: #,##0. */
        FORMAT_3(3),
        /** Format: #,##0.00. */
        FORMAT_4(4),
        /** Format: $#,##0_);($#,##0). */
        FORMAT_5(5),
        /** Format: $#,##0_);[Red]($#,##0). */
        FORMAT_6(6),
        /** Format: $#,##0.00_);($#,##0.00). */
        FORMAT_7(7),
        /** Format: $#,##0.00_);[Red]($#,##0.00). */
        FORMAT_8(8),
        /** Format: 0%. */
        FORMAT_9(9),
        /** Format: 0.00%. */
        FORMAT_10(10),
        /** Format: 0.00E+00. */
        FORMAT_11(11),
        /** Format: # ?/?. */
        FORMAT_12(12),
        /** Format: # ??/??. */
        FORMAT_13(13),
        /** Format: m/d/yyyy. */
        FORMAT_14(14),
        /** Format: d-mmm-yy. */
        FORMAT_15(15),
        /** Format: d-mmm. */
        FORMAT_16(16),
        /** Format: mmm-yy. */
        FORMAT_17(17),
        /** Format: mm AM/PM. */
        FORMAT_18(18),
        /** Format: h:mm:ss AM/PM. */
        FORMAT_19(19),
        /** Format: h:mm. */
        FORMAT_20(20),
        /** Format: h:mm:ss. */
        FORMAT_21(21),
        /** Format: m/d/yyyy h:mm. */
        FORMAT_22(22),
        /** Format: #,##0_);(#,##0). */
        FORMAT_37(37),
        /** Format: #,##0_);[Red](#,##0). */
        FORMAT_38(38),
        /** Format: #,##0.00_);(#,##0.00). */
        FORMAT_39(39),
        /** Format: #,##0.00_);[Red](#,##0.00). */
        FORMAT_40(40),
        /** Format: mm:ss. */
        FORMAT_45(45),
        /** Format: [h]:mm:ss. */
        FORMAT_46(46),
        /** Format: mm:ss.0. */
        FORMAT_47(47),
        /** Format: ##0.0E+0. */
        FORMAT_48(48),
        /** Format: #. */
        FORMAT_49(49),
        /** Custom format with ID 164 or higher. */
        CUSTOM(164);

        private final int value;

        FormatNumber(int value) {
            this.value = value;
        }

        /**
         * Gets the numeric OOXML format identifier.
         *
         * @return Numeric format identifier
         */
        public int getValue() {
            return value;
        }
    }

    /** Range or validity of a parsed format number. */
    public enum FormatRange {
        /** A declared format from 0 through 164. */
        DEFINED_FORMAT,
        /** A custom format above 164. */
        CUSTOM_FORMAT,
        /** An invalid format number, such as a negative value. */
        INVALID,
        /** An undeclared value between 0 and 164. */
        UNDEFINED
    }

    @AppendAnnotation
    private String customFormatCode;
    @AppendAnnotation
    private int customFormatId;
    @AppendAnnotation
    private FormatNumber number;

    /** Creates a number format with the default predefined and custom identifiers. */
    public NumberFormat() {
        number = DEFAULT_NUMBER;
        customFormatCode = "";
        customFormatId = CUSTOM_FORMAT_START_NUMBER;
    }

    /**
     * Gets the raw custom format code in Excel notation.
     *
     * @return Custom format code
     */
    public String getCustomFormatCode() {
        return customFormatCode;
    }

    /**
     * Sets the raw custom format code in Excel notation. The code is not escaped or unescaped automatically.
     *
     * @param customFormatCode Custom format code
     * @throws FormatException if the value is null or empty
     */
    public void setCustomFormatCode(String customFormatCode) {
        if (customFormatCode == null || customFormatCode.isEmpty()) {
            throw new FormatException("A custom format code cannot be null or empty");
        }
        this.customFormatCode = customFormatCode;
    }

    /**
     * Gets the identifier of the custom format.
     *
     * @return Custom format identifier
     */
    public int getCustomFormatId() {
        return customFormatId;
    }

    /**
     * Sets the identifier of the custom format.
     *
     * @param customFormatId Custom format identifier
     * @throws StyleException if the value is below {@value #CUSTOM_FORMAT_START_NUMBER}
     */
    public void setCustomFormatId(int customFormatId) {
        if (customFormatId < CUSTOM_FORMAT_START_NUMBER) {
            throw new StyleException("The number '" + customFormatId
                    + "' is not a valid custom format ID. Must be at least " + CUSTOM_FORMAT_START_NUMBER);
        }
        this.customFormatId = customFormatId;
    }

    /**
     * Gets whether this number format represents a custom format.
     *
     * @return True if the number is {@link FormatNumber#CUSTOM}
     */
    public boolean isCustomFormat() {
        return number == FormatNumber.CUSTOM;
    }

    /**
     * Gets the predefined format number.
     *
     * @return Format number
     */
    public FormatNumber getNumber() {
        return number;
    }

    /**
     * Sets the predefined format number.
     *
     * @param number Format number
     */
    public void setNumber(FormatNumber number) {
        this.number = Objects.requireNonNull(number, "number");
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("\"NumberFormat\": {\n");
        addPropertyAsJson(sb, "CustomFormatCode", customFormatCode, false);
        addPropertyAsJson(sb, "CustomFormatID", customFormatId, false);
        addPropertyAsJson(sb, "Number", number, false);
        addPropertyAsJson(sb, "HashCode", hashCode(), true);
        sb.append("\n}");
        return sb.toString();
    }

    /**
     * Copies this number format without its internal ID.
     *
     * @return Dereferenced copy
     */
    @Override
    public AbstractStyle copy() {
        NumberFormat copy = new NumberFormat();
        copy.customFormatCode = customFormatCode;
        copy.customFormatId = customFormatId;
        copy.number = number;
        return copy;
    }

    /**
     * Copies this number format without requiring a cast.
     *
     * @return Dereferenced copy
     */
    public NumberFormat copyNumberFormat() {
        return (NumberFormat) copy();
    }

    @Override
    public int hashCode() {
        return Objects.hash(customFormatCode, customFormatId, number);
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (!(obj instanceof NumberFormat other)) {
            return false;
        }
        return customFormatId == other.customFormatId
                && Objects.equals(customFormatCode, other.customFormatCode)
                && number == other.number;
    }

    /**
     * Determines whether a predefined format represents a date or a date and time.
     *
     * @param number Format number to check
     * @return True if the format represents a date
     */
    public static boolean isDateFormat(FormatNumber number) {
        return switch (number) {
            case FORMAT_14, FORMAT_15, FORMAT_16, FORMAT_17, FORMAT_22 -> true;
            default -> false;
        };
    }

    /**
     * Determines whether a predefined format represents a time.
     *
     * @param number Format number to check
     * @return True if the format represents a time
     */
    public static boolean isTimeFormat(FormatNumber number) {
        return switch (number) {
            case FORMAT_18, FORMAT_19, FORMAT_20, FORMAT_21, FORMAT_45, FORMAT_46, FORMAT_47 -> true;
            default -> false;
        };
    }

    /**
     * Parses a raw format number and classifies its range.
     *
     * @param number Raw format number
     * @return Parsed format number and its range classification
     */
    public static NumberFormatEvaluation tryParseFormatNumber(int number) {
        for (FormatNumber formatNumber : FormatNumber.values()) {
            if (formatNumber.getValue() == number) {
                return new NumberFormatEvaluation(FormatRange.DEFINED_FORMAT, formatNumber);
            }
        }
        if (number < 0) {
            return new NumberFormatEvaluation(FormatRange.INVALID, FormatNumber.NONE);
        } else if (number > 0 && number < CUSTOM_FORMAT_START_NUMBER) {
            return new NumberFormatEvaluation(FormatRange.UNDEFINED, FormatNumber.NONE);
        }
        return new NumberFormatEvaluation(FormatRange.CUSTOM_FORMAT, FormatNumber.CUSTOM);
    }

    /**
     * Parsed number format and its range classification.
     *
     * @param range        Range classification
     * @param formatNumber Parsed format number
     */
    public record NumberFormatEvaluation(FormatRange range, FormatNumber formatNumber) {
        /**
         * Creates an immutable format-number evaluation.
         *
         * @param range        Range classification
         * @param formatNumber Parsed format number
         */
        public NumberFormatEvaluation {
            Objects.requireNonNull(range, "range");
            Objects.requireNonNull(formatNumber, "formatNumber");
        }
    }
}
