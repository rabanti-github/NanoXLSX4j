/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.colors.AutoColor;
import ch.rabanti.nanoxlsx4j.colors.Color;
import ch.rabanti.nanoxlsx4j.colors.IndexedColor;
import ch.rabanti.nanoxlsx4j.colors.SrgbColor;
import ch.rabanti.nanoxlsx4j.colors.SystemColor;
import ch.rabanti.nanoxlsx4j.colors.ThemeColor;
import ch.rabanti.nanoxlsx4j.interfaces.IColor;
import ch.rabanti.nanoxlsx4j.utils.Validators;

/**
 * Class representing a Fill (background) entry. The Fill entry is used to define background colors and fill patterns
 *
 * @author Raphael Stoeckli
 */
public class Fill extends AbstractStyle {
    // ### C O N S T A N T S ###
    /**
     * Default Color (foreground or background) as {@link Color} instance
     */
    public static final Color DEFAULT_COLOR = Color.createRgb(SrgbColor.DEFAULT_SRGB_COLOR);

    /**
     * Default indexed color as {@link Color} instance (SystemForeground, index 64)
     */
    public static final Color DEFAULT_INDEXED_COLOR = Color.createIndexed(IndexedColor.DEFAULT_INDEXED_COLOR);

    /**
     * Default pattern
     */
    public static final PatternValue DEFAULT_PATTERN_FILL = PatternValue.none;

    // ### E N U M S ###

    /**
     * Enum for the type of the fill pattern
     */
    public enum PatternValue {
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

        PatternValue(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }

    }

    /**
     * Enum for the type of the fill
     */
    public enum FillType {
        /**
         * Color defines a pattern color
         */
        patternColor(0),
        /**
         * Color defines a solid fill color
         */
        fillColor(1);

        private final int value;

        FillType(int value) {
            this.value = value;
        }

        public int getValue() {
            return value;
        }
    }

    // ### P R I V A T E F I E L D S ###
    private Color foregroundColor;
    private Color backgroundColor;
    private PatternValue patternFill;

    // ### G E T T E R S & S E T T E R S ###

    /**
     * Gets the pattern type of the fill (Default is none)
     *
     * @return Pattern type of the fill
     */
    public PatternValue getPatternFill() {
        return patternFill;
    }

    /**
     * Sets the pattern type of the fill (Default is none)
     *
     * @param patternFill Pattern type of the fill
     */
    public void setPatternFill(PatternValue patternFill) {
        this.patternFill = patternFill;
    }

    /**
     * Gets the foreground color of the fill as a {@link Color} compound object
     *
     * @return Foreground color
     */
    public Color getForegroundColor() {
        return foregroundColor;
    }

    /**
     * Sets the foreground color of the fill from a {@link Color} compound object.
     * If the pattern is currently none, it is automatically set to solid.
     *
     * @param foregroundColor Foreground color
     */
    public void setForegroundColor(Color foregroundColor) {
        this.foregroundColor = foregroundColor;
        if (this.patternFill == PatternValue.none) {
            this.patternFill = PatternValue.solid;
        }
    }

    /**
     * Sets the foreground color of the fill from an ARGB or RGB hex string.
     * The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF.
     * If the pattern is currently none, it is automatically set to solid.
     *
     * @param foregroundColor Foreground color as ARGB hex string
     */
    public void setForegroundColor(String foregroundColor) {
        validateColor(foregroundColor, true);
        setForegroundColor(Color.createRgb(foregroundColor));
    }

    /**
     * Gets the background color of the fill as a {@link Color} compound object
     *
     * @return Background color
     */
    public Color getBackgroundColor() {
        return backgroundColor;
    }

    /**
     * Sets the background color of the fill from a {@link Color} compound object.
     * If the pattern is currently none, it is automatically set to solid.
     *
     * @param backgroundColor Background color
     */
    public void setBackgroundColor(Color backgroundColor) {
        this.backgroundColor = backgroundColor;
        if (this.patternFill == PatternValue.none) {
            this.patternFill = PatternValue.solid;
        }
    }

    /**
     * Sets the background color of the fill from an ARGB or RGB hex string.
     * The value is expressed as hex string with the format AARRGGBB. AA (Alpha) is usually FF.
     * If the pattern is currently none, it is automatically set to solid.
     *
     * @param backgroundColor Background color as ARGB hex string
     */
    public void setBackgroundColor(String backgroundColor) {
        validateColor(backgroundColor, true);
        setBackgroundColor(Color.createRgb(backgroundColor));
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Default constructor
     */
    public Fill() {
        this.patternFill = DEFAULT_PATTERN_FILL;
        this.foregroundColor = DEFAULT_COLOR;
        this.backgroundColor = DEFAULT_COLOR;
    }

    /**
     * Constructor with foreground and background color as ARGB strings
     *
     * @param foreground Foreground color of the fill (ARGB hex string)
     * @param background Background color of the fill (ARGB hex string)
     */
    public Fill(String foreground, String background) {
        this.backgroundColor = Color.createRgb(background);
        this.foregroundColor = Color.createRgb(foreground);
        this.patternFill = PatternValue.solid;
    }

    /**
     * Constructor with color value and fill type (ARGB string)
     *
     * @param value    Color value (ARGB hex string)
     * @param fillType Fill type (fill or pattern)
     */
    public Fill(String value, FillType fillType) {
        if (fillType == FillType.fillColor) {
            this.backgroundColor = DEFAULT_COLOR;
            this.foregroundColor = Color.createRgb(value);
        }
        else {
            this.backgroundColor = Color.createRgb(value);
            this.foregroundColor = DEFAULT_COLOR;
        }
        this.patternFill = PatternValue.solid;
    }

    // ### M E T H O D S ###

    /**
     * Sets the color and the depending fill type from an ARGB string
     *
     * @param value    Color value (ARGB hex string)
     * @param fillType Fill type (fill or pattern)
     */
    public void setColor(String value, FillType fillType) {
        if (fillType == FillType.fillColor) {
            this.backgroundColor = DEFAULT_COLOR;
            this.foregroundColor = Color.createRgb(value);
        }
        else {
            this.backgroundColor = Color.createRgb(value);
            this.foregroundColor = DEFAULT_COLOR;
        }
        this.patternFill = PatternValue.solid;
    }

    /**
     * Sets the color depending on fill type from a {@link Color} compound object
     *
     * @param value    Color value
     * @param fillType Fill type (fill or pattern)
     */
    public void setColor(Color value, FillType fillType) {
        if (fillType == FillType.fillColor) {
            this.backgroundColor = DEFAULT_COLOR;
            this.foregroundColor = value;
        }
        else {
            this.backgroundColor = value;
            this.foregroundColor = DEFAULT_COLOR;
        }
        this.patternFill = PatternValue.solid;
    }

    /**
     * Sets the color depending on fill type from an {@link IColor} component
     *
     * @param value    Color component
     * @param fillType Fill type (fill or pattern)
     */
    public void setColor(IColor value, FillType fillType) {
        setColor(getColorByComponent(value), fillType);
    }

    /**
     * Override toString method
     *
     * @return String of a class instance
     */
    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("\"Fill\": {\n");
        addPropertyAsJson(sb, "BackgroundColor", backgroundColor);
        addPropertyAsJson(sb, "ForegroundColor", foregroundColor);
        addPropertyAsJson(sb, "PatternFill", patternFill);
        addPropertyAsJson(sb, "HashCode", this.hashCode(), true);
        sb.append("\n}");
        return sb.toString();
    }

    /**
     * Method to copy the current object to a new one with casting
     *
     * @return Copy of the current object without the internal ID
     */
    public Fill copyFill() {
        return (Fill) copy();
    }

    /**
     * Method to copy the current object to a new one
     *
     * @return Copy of the current object without the internal ID
     */
    @Override
    public Fill copy() {
        Fill copy = new Fill();
        copy.setBackgroundColor(this.backgroundColor);
        copy.setForegroundColor(this.foregroundColor);
        copy.setPatternFill(this.patternFill);
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
        int result = (patternFill != null ? patternFill.hashCode() : 1);
        result = 31 * result + (foregroundColor != null ? foregroundColor.hashCode() : 2);
        result = 31 * result + (backgroundColor != null ? backgroundColor.hashCode() : 4);
        return result;
    }

    // ### S T A T I C F U N C T I O N S ###

    /**
     * Gets the pattern name from the enum
     *
     * @param pattern Enum to process
     * @return The valid value of the pattern as String
     */
    public static String getPatternName(PatternValue pattern) {
        String output;
        switch (pattern) {
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

    /**
     * Converts an OOXML string to its corresponding PatternValue enum
     *
     * @param name OOXML pattern type string
     * @return Corresponding enum, or {@link PatternValue#none} if not found
     */
    public static PatternValue getPatternEnum(String name) {
        if (name == null) return PatternValue.none;
        switch (name) {
            case "none": return PatternValue.none;
            case "solid": return PatternValue.solid;
            case "darkGray": return PatternValue.darkGray;
            case "mediumGray": return PatternValue.mediumGray;
            case "lightGray": return PatternValue.lightGray;
            case "gray0625": return PatternValue.gray0625;
            case "gray125": return PatternValue.gray125;
            default: return PatternValue.none;
        }
    }

    /**
     * Validates the passed string, whether it is a valid RGB value that can be used for Fills or Fonts
     *
     * @param hexCode  Hex string to check
     * @param useAlpha If true, two additional characters (total 8) are expected as alpha value
     * @throws StyleException thrown if an invalid hex value is passed
     */
    public static void validateColor(String hexCode, boolean useAlpha) {
        Validators.validateColor(hexCode, useAlpha, false);
    }

    /**
     * Validates the passed string, whether it is a valid RGB value that can be used for Fills or Fonts
     *
     * @param hexCode    Hex string to check
     * @param useAlpha   If true, two additional characters (total 8) are expected as alpha value
     * @param allowEmpty Optional parameter that allows null or empty as valid values
     * @throws StyleException thrown if an invalid hex value is passed
     */
    public static void validateColor(String hexCode, boolean useAlpha, boolean allowEmpty) {
        Validators.validateColor(hexCode, useAlpha, allowEmpty);
    }

    /**
     * Creates a {@link Color} compound object from an {@link IColor} component
     *
     * @param component Color component
     * @return Color instance
     */
    private static Color getColorByComponent(IColor component) {
        if (component instanceof SrgbColor) {
            return Color.createRgb((SrgbColor) component);
        }
        else if (component instanceof IndexedColor) {
            return Color.createIndexed((IndexedColor) component);
        }
        else if (component instanceof ThemeColor) {
            return Color.createTheme((ThemeColor) component);
        }
        else if (component instanceof SystemColor) {
            return Color.createSystem((SystemColor) component);
        }
        else if (component instanceof AutoColor) {
            return Color.createAuto();
        }
        else {
            return Color.createNone();
        }
    }

}
