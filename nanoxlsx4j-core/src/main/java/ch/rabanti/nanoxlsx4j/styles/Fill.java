/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import java.util.Objects;

import ch.rabanti.nanoxlsx4j.colors.AutoColor;
import ch.rabanti.nanoxlsx4j.colors.Color;
import ch.rabanti.nanoxlsx4j.colors.IndexedColor;
import ch.rabanti.nanoxlsx4j.colors.SrgbColor;
import ch.rabanti.nanoxlsx4j.colors.SystemColor;
import ch.rabanti.nanoxlsx4j.colors.ThemeColor;
import ch.rabanti.nanoxlsx4j.internal.interfaces.BaseColor;

/**
 * Class representing a Fill (background) entry. The Fill entry is used to define background colors and fill patterns.
 */
public class Fill extends AbstractStyle {

    /** Default color for the foreground or background. */
    public static final Color DEFAULT_COLOR = Color.createRgb("FF000000");

    /** Default indexed color. */
    public static final Color DEFAULT_INDEXED_COLOR = Color.createIndexed(IndexedColor.DEFAULT_INDEXED_COLOR);

    /** Default fill pattern. */
    public static final PatternValue DEFAULT_PATTERN_FILL = PatternValue.NONE;

    /** Enum for the semantic role of a color in the fill. */
    public enum FillType {
        /** Color defines a pattern color. */
        PATTERN_COLOR,
        /** Color defines a solid fill color. */
        FILL_COLOR
    }

    /** Enum for the supported fill patterns. */
    public enum PatternValue {
        /** No pattern. */
        NONE,
        /** Solid fill. */
        SOLID,
        /** Dark gray fill. */
        DARK_GRAY,
        /** Medium gray fill. */
        MEDIUM_GRAY,
        /** Light gray fill. */
        LIGHT_GRAY,
        /** 6.25 percent gray fill. */
        GRAY_0625,
        /** 12.5 percent gray fill. */
        GRAY_125
    }

    @AppendAnnotation
    private Color backgroundColor;
    @AppendAnnotation
    private Color foregroundColor;
    @AppendAnnotation
    private PatternValue patternFill;

    /** Creates a fill with the default colors and no pattern. */
    public Fill() {
        patternFill = DEFAULT_PATTERN_FILL;
        foregroundColor = DEFAULT_COLOR;
        backgroundColor = DEFAULT_COLOR;
    }

    /**
     * Creates a solid fill from foreground and background RGB or ARGB values.
     *
     * @param foreground Foreground color
     * @param background Background color
     */
    public Fill(String foreground, String background) {
        setBackgroundColor(Color.createRgb(background));
        setForegroundColor(Color.createRgb(foreground));
        patternFill = PatternValue.SOLID;
    }

    /**
     * Creates a solid fill from an RGB or ARGB value and its semantic role.
     *
     * @param value    Color value
     * @param fillType Fill or pattern color
     */
    public Fill(String value, FillType fillType) {
        setColor(value, fillType);
    }

    /**
     * Creates a solid foreground fill from an RGB or ARGB value.
     *
     * @param value Color value
     */
    public Fill(String value) {
        this(value, FillType.FILL_COLOR);
    }

    /**
     * Creates a solid foreground fill from an indexed color.
     *
     * @param index Indexed color value
     */
    public Fill(IndexedColor.Value index) {
        this();
        setForegroundColor(Color.createIndexed(index));
        patternFill = PatternValue.SOLID;
    }

    /**
     * Creates a solid foreground fill from a numeric color index.
     *
     * @param index Color index from 0 to 65
     */
    public Fill(int index) {
        this();
        setForegroundColor(Color.createIndexed(index));
        patternFill = PatternValue.SOLID;
    }

    /**
     * Gets the background color of the fill.
     *
     * @return Background color of the fill
     */
    public Color getBackgroundColor() {
        return backgroundColor;
    }

    /**
     * Sets the background color and activates a solid pattern when no pattern is currently selected.
     *
     * @param backgroundColor Background color
     */
    public void setBackgroundColor(Color backgroundColor) {
        this.backgroundColor = backgroundColor;
        activateSolidPattern();
    }

    /**
     * Sets the background color from an RGB or ARGB value.
     *
     * @param backgroundColor Background color value
     */
    public void setBackgroundColor(String backgroundColor) {
        setBackgroundColor(Color.createRgb(backgroundColor));
    }

    /**
     * Gets the foreground color of the fill.
     *
     * @return Foreground color of the fill
     */
    public Color getForegroundColor() {
        return foregroundColor;
    }

    /**
     * Sets the foreground color and activates a solid pattern when no pattern is currently selected.
     *
     * @param foregroundColor Foreground color
     */
    public void setForegroundColor(Color foregroundColor) {
        this.foregroundColor = foregroundColor;
        activateSolidPattern();
    }

    /**
     * Sets the foreground color from an RGB or ARGB value.
     *
     * @param foregroundColor Foreground color value
     */
    public void setForegroundColor(String foregroundColor) {
        setForegroundColor(Color.createRgb(foregroundColor));
    }

    /**
     * Gets the pattern type of the fill.
     *
     * @return Pattern type of the fill
     */
    public PatternValue getPatternFill() {
        return patternFill;
    }

    /**
     * Sets the pattern type of the fill.
     *
     * @param patternFill Pattern type of the fill
     */
    public void setPatternFill(PatternValue patternFill) {
        this.patternFill = Objects.requireNonNull(patternFill, "patternFill");
    }

    private void activateSolidPattern() {
        if (patternFill == PatternValue.NONE) {
            patternFill = PatternValue.SOLID;
        }
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("\"Fill\": {\n");
        addPropertyAsJson(sb, "BackgroundColor", backgroundColor, false);
        addPropertyAsJson(sb, "ForegroundColor", foregroundColor, false);
        addPropertyAsJson(sb, "PatternFill", patternFill, false);
        addPropertyAsJson(sb, "HashCode", hashCode(), true);
        sb.append("\n}");
        return sb.toString();
    }

    /**
     * Copies this fill without its internal ID.
     *
     * @return Dereferenced copy
     */
    @Override
    public AbstractStyle copy() {
        Fill copy = new Fill();
        copy.backgroundColor = backgroundColor;
        copy.foregroundColor = foregroundColor;
        copy.patternFill = patternFill;
        return copy;
    }

    @Override
    public int hashCode() {
        return Objects.hash(backgroundColor, foregroundColor, patternFill);
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (!(obj instanceof Fill other)) {
            return false;
        }
        return Objects.equals(backgroundColor, other.backgroundColor)
                && Objects.equals(foregroundColor, other.foregroundColor)
                && patternFill == other.patternFill;
    }

    /**
     * Copies this fill without requiring a cast.
     *
     * @return Dereferenced copy
     */
    public Fill copyFill() {
        return (Fill) copy();
    }

    /**
     * Sets an RGB or ARGB color according to its semantic role.
     *
     * @param value    Color value
     * @param fillType Fill or pattern color
     */
    public void setColor(String value, FillType fillType) {
        setColor(Color.createRgb(value), fillType);
    }

    /**
     * Sets a compound color according to its semantic role.
     *
     * @param value    Color value
     * @param fillType Fill or pattern color
     */
    public void setColor(Color value, FillType fillType) {
        Objects.requireNonNull(fillType, "fillType");
        if (fillType == FillType.FILL_COLOR) {
            backgroundColor = DEFAULT_COLOR;
            setForegroundColor(value);
        } else {
            setBackgroundColor(value);
            foregroundColor = DEFAULT_COLOR;
        }
        patternFill = PatternValue.SOLID;
    }

    /**
     * Sets a color component according to its semantic role.
     *
     * @param value    Color component
     * @param fillType Fill or pattern color
     */
    public void setColor(BaseColor value, FillType fillType) {
        setColor(getColorByComponent(value), fillType);
    }

    /**
     * Gets the OOXML pattern name from the enum.
     *
     * @param pattern Pattern to process
     * @return OOXML pattern name
     */
    static String getPatternName(PatternValue pattern) {
        return switch (pattern) {
            case SOLID -> "solid";
            case DARK_GRAY -> "darkGray";
            case MEDIUM_GRAY -> "mediumGray";
            case LIGHT_GRAY -> "lightGray";
            case GRAY_0625 -> "gray0625";
            case GRAY_125 -> "gray125";
            case NONE -> "none";
        };
    }

    /**
     * Parses an OOXML pattern name.
     *
     * @param name Pattern name
     * @return Corresponding pattern, or {@link PatternValue#NONE} for an unknown value
     */
    static PatternValue getPatternEnum(String name) {
        if (name == null) {
            return PatternValue.NONE;
        }
        return switch (name) {
            case "solid" -> PatternValue.SOLID;
            case "darkGray" -> PatternValue.DARK_GRAY;
            case "mediumGray" -> PatternValue.MEDIUM_GRAY;
            case "lightGray" -> PatternValue.LIGHT_GRAY;
            case "gray0625" -> PatternValue.GRAY_0625;
            case "gray125" -> PatternValue.GRAY_125;
            default -> PatternValue.NONE;
        };
    }

    private static Color getColorByComponent(BaseColor component) {
        if (component instanceof SrgbColor srgbColor) {
            return Color.createRgb(srgbColor);
        } else if (component instanceof IndexedColor indexedColor) {
            return Color.createIndexed(indexedColor);
        } else if (component instanceof ThemeColor themeColor) {
            return Color.createTheme(themeColor);
        } else if (component instanceof SystemColor systemColor) {
            return Color.createSystem(systemColor);
        } else if (component instanceof AutoColor || component == null) {
            return Color.createAuto();
        }
        return Color.createAuto();
    }
}
