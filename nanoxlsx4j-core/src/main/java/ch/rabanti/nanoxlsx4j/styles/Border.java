/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import java.util.Objects;

import ch.rabanti.nanoxlsx4j.utils.ParserUtils;
import ch.rabanti.nanoxlsx4j.utils.Validators;

/**
 * Class representing a Border entry. The Border entry is used to define frames and cell borders
 */
public class Border extends AbstractStyle {

    /** Default border style as constant. */
    public static final StyleValue DEFAULT_BORDER_STYLE = StyleValue.NONE;

    /** Default border color as constant. */
    public static final String DEFAULT_BORDER_COLOR = "";

    @AppendAnnotation
    private String bottomColor;
    @AppendAnnotation
    private StyleValue bottomStyle;
    @AppendAnnotation
    private String diagonalColor;
    @AppendAnnotation
    private boolean diagonalDown;
    @AppendAnnotation
    private StyleValue diagonalStyle;
    @AppendAnnotation
    private boolean diagonalUp;
    @AppendAnnotation
    private String leftColor;
    @AppendAnnotation
    private StyleValue leftStyle;
    @AppendAnnotation
    private String rightColor;
    @AppendAnnotation
    private StyleValue rightStyle;
    @AppendAnnotation
    private String topColor;
    @AppendAnnotation
    private StyleValue topStyle;

    /** Enum for the border style. */
    public enum StyleValue {
        /** No border. */
        NONE,
        /** Hair border. */
        HAIR,
        /** Dotted border. */
        DOTTED,
        /** Dashed border with double dots. */
        DASH_DOT_DOT,
        /** Dash-dotted border. */
        DASH_DOT,
        /** Dashed border. */
        DASHED,
        /** Thin border. */
        THIN,
        /** Medium-dashed border with double dots. */
        MEDIUM_DASH_DOT_DOT,
        /** Slant dash-dotted border. */
        SLANT_DASH_DOT,
        /** Medium dash-dotted border. */
        MEDIUM_DASH_DOT,
        /** Medium dashed border. */
        MEDIUM_DASHED,
        /** Medium border. */
        MEDIUM,
        /** Thick border. */
        THICK,
        /** Double border. */
        DOUBLE
    }

    /** Creates a border with the reference implementation's default values. */
    public Border() {
        setBottomColor(DEFAULT_BORDER_COLOR);
        setTopColor(DEFAULT_BORDER_COLOR);
        setLeftColor(DEFAULT_BORDER_COLOR);
        setRightColor(DEFAULT_BORDER_COLOR);
        setDiagonalColor(DEFAULT_BORDER_COLOR);
        leftStyle = DEFAULT_BORDER_STYLE;
        rightStyle = DEFAULT_BORDER_STYLE;
        topStyle = DEFAULT_BORDER_STYLE;
        bottomStyle = DEFAULT_BORDER_STYLE;
        diagonalStyle = DEFAULT_BORDER_STYLE;
    }

    /** @return Color code of the bottom border. */
    public String getBottomColor() {
        return bottomColor;
    }

    /**
     * Sets the bottom border color.
     *
     * @param bottomColor ARGB color code, or null/empty
     */
    public void setBottomColor(String bottomColor) {
        Validators.validateColor(bottomColor, true, true);
        this.bottomColor = ParserUtils.toUpper(bottomColor);
    }

    /** @return Style of the bottom border. */
    public StyleValue getBottomStyle() {
        return bottomStyle;
    }

    /** @param bottomStyle Style of the bottom border. */
    public void setBottomStyle(StyleValue bottomStyle) {
        this.bottomStyle = Objects.requireNonNull(bottomStyle, "bottomStyle");
    }

    /** @return Color code of the diagonal lines. */
    public String getDiagonalColor() {
        return diagonalColor;
    }

    /** @param diagonalColor ARGB color code, or null/empty. */
    public void setDiagonalColor(String diagonalColor) {
        Validators.validateColor(diagonalColor, true, true);
        this.diagonalColor = ParserUtils.toUpper(diagonalColor);
    }

    /** @return True if the downwards diagonal line is used. */
    public boolean isDiagonalDown() {
        return diagonalDown;
    }

    /** @param diagonalDown Whether the downwards diagonal line is used. */
    public void setDiagonalDown(boolean diagonalDown) {
        this.diagonalDown = diagonalDown;
    }

    /** @return Style of the diagonal lines. */
    public StyleValue getDiagonalStyle() {
        return diagonalStyle;
    }

    /** @param diagonalStyle Style of the diagonal lines. */
    public void setDiagonalStyle(StyleValue diagonalStyle) {
        this.diagonalStyle = Objects.requireNonNull(diagonalStyle, "diagonalStyle");
    }

    /** @return True if the upwards diagonal line is used. */
    public boolean isDiagonalUp() {
        return diagonalUp;
    }

    /** @param diagonalUp Whether the upwards diagonal line is used. */
    public void setDiagonalUp(boolean diagonalUp) {
        this.diagonalUp = diagonalUp;
    }

    /** @return Color code of the left border. */
    public String getLeftColor() {
        return leftColor;
    }

    /** @param leftColor ARGB color code, or null/empty. */
    public void setLeftColor(String leftColor) {
        Validators.validateColor(leftColor, true, true);
        this.leftColor = ParserUtils.toUpper(leftColor);
    }

    /** @return Style of the left border. */
    public StyleValue getLeftStyle() {
        return leftStyle;
    }

    /** @param leftStyle Style of the left border. */
    public void setLeftStyle(StyleValue leftStyle) {
        this.leftStyle = Objects.requireNonNull(leftStyle, "leftStyle");
    }

    /** @return Color code of the right border. */
    public String getRightColor() {
        return rightColor;
    }

    /** @param rightColor ARGB color code, or null/empty. */
    public void setRightColor(String rightColor) {
        Validators.validateColor(rightColor, true, true);
        this.rightColor = ParserUtils.toUpper(rightColor);
    }

    /** @return Style of the right border. */
    public StyleValue getRightStyle() {
        return rightStyle;
    }

    /** @param rightStyle Style of the right border. */
    public void setRightStyle(StyleValue rightStyle) {
        this.rightStyle = Objects.requireNonNull(rightStyle, "rightStyle");
    }

    /** @return Color code of the top border. */
    public String getTopColor() {
        return topColor;
    }

    /** @param topColor ARGB color code, or null/empty. */
    public void setTopColor(String topColor) {
        Validators.validateColor(topColor, true, true);
        this.topColor = ParserUtils.toUpper(topColor);
    }

    /** @return Style of the top border. */
    public StyleValue getTopStyle() {
        return topStyle;
    }

    /** @param topStyle Style of the top border. */
    public void setTopStyle(StyleValue topStyle) {
        this.topStyle = Objects.requireNonNull(topStyle, "topStyle");
    }

    @Override
    public int hashCode() {
        return Objects.hash(bottomColor, bottomStyle, diagonalColor, diagonalDown, diagonalUp, diagonalStyle,
                leftColor, leftStyle, rightColor, rightStyle, topColor, topStyle);
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (!(obj instanceof Border other)) {
            return false;
        }
        return diagonalDown == other.diagonalDown
                && diagonalUp == other.diagonalUp
                && Objects.equals(bottomColor, other.bottomColor)
                && bottomStyle == other.bottomStyle
                && Objects.equals(diagonalColor, other.diagonalColor)
                && diagonalStyle == other.diagonalStyle
                && Objects.equals(leftColor, other.leftColor)
                && leftStyle == other.leftStyle
                && Objects.equals(rightColor, other.rightColor)
                && rightStyle == other.rightStyle
                && Objects.equals(topColor, other.topColor)
                && topStyle == other.topStyle;
    }

    /**
     * Copies this border without its internal ID.
     *
     * @return Dereferenced copy
     */
    @Override
    public AbstractStyle copy() {
        Border copy = new Border();
        copy.bottomColor = bottomColor;
        copy.bottomStyle = bottomStyle;
        copy.diagonalColor = diagonalColor;
        copy.diagonalDown = diagonalDown;
        copy.diagonalStyle = diagonalStyle;
        copy.diagonalUp = diagonalUp;
        copy.leftColor = leftColor;
        copy.leftStyle = leftStyle;
        copy.rightColor = rightColor;
        copy.rightStyle = rightStyle;
        copy.topColor = topColor;
        copy.topStyle = topStyle;
        return copy;
    }

    /** @return Dereferenced copy without casting. */
    public Border copyBorder() {
        return (Border) copy();
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("\"Border\": {\n");
        addPropertyAsJson(sb, "BottomStyle", bottomStyle, false);
        addPropertyAsJson(sb, "DiagonalColor", diagonalColor, false);
        addPropertyAsJson(sb, "DiagonalDown", diagonalDown, false);
        addPropertyAsJson(sb, "DiagonalStyle", diagonalStyle, false);
        addPropertyAsJson(sb, "DiagonalUp", diagonalUp, false);
        addPropertyAsJson(sb, "LeftColor", leftColor, false);
        addPropertyAsJson(sb, "LeftStyle", leftStyle, false);
        addPropertyAsJson(sb, "RightColor", rightColor, false);
        addPropertyAsJson(sb, "RightStyle", rightStyle, false);
        addPropertyAsJson(sb, "TopColor", topColor, false);
        addPropertyAsJson(sb, "TopStyle", topStyle, false);
        addPropertyAsJson(sb, "HashCode", hashCode(), true);
        sb.append("\n}");
        return sb.toString();
    }

    /** @return True if this border contains only default values. */
    boolean isEmpty() {
        return Objects.equals(bottomColor, DEFAULT_BORDER_COLOR)
                && Objects.equals(topColor, DEFAULT_BORDER_COLOR)
                && Objects.equals(leftColor, DEFAULT_BORDER_COLOR)
                && Objects.equals(rightColor, DEFAULT_BORDER_COLOR)
                && Objects.equals(diagonalColor, DEFAULT_BORDER_COLOR)
                && leftStyle == DEFAULT_BORDER_STYLE
                && rightStyle == DEFAULT_BORDER_STYLE
                && topStyle == DEFAULT_BORDER_STYLE
                && bottomStyle == DEFAULT_BORDER_STYLE
                && diagonalStyle == DEFAULT_BORDER_STYLE
                && !diagonalDown
                && !diagonalUp;
    }

    /**
     * Gets the OOXML border style name from the enum.
     *
     * @param style Enum to process
     * @return Valid OOXML border style name
     */
    static String getStyleName(StyleValue style) {
        return switch (style) {
            case NONE -> "";
            case HAIR -> "hair";
            case DOTTED -> "dotted";
            case DASH_DOT_DOT -> "dashDotDot";
            case DASH_DOT -> "dashDot";
            case DASHED -> "dashed";
            case THIN -> "thin";
            case MEDIUM_DASH_DOT_DOT -> "mediumDashDotDot";
            case SLANT_DASH_DOT -> "slantDashDot";
            case MEDIUM_DASH_DOT -> "mediumDashDot";
            case MEDIUM_DASHED -> "mediumDashed";
            case MEDIUM -> "medium";
            case THICK -> "thick";
            case DOUBLE -> "double";
        };
    }

    /**
     * Parses an OOXML border style name.
     *
     * @param styleName String to parse
     * @return Corresponding style, or {@link StyleValue#NONE} for an unknown value
     */
    static StyleValue getStyleEnum(String styleName) {
        if (styleName == null) {
            return StyleValue.NONE;
        }
        return switch (styleName) {
            case "hair" -> StyleValue.HAIR;
            case "dotted" -> StyleValue.DOTTED;
            case "dashDotDot" -> StyleValue.DASH_DOT_DOT;
            case "dashDot" -> StyleValue.DASH_DOT;
            case "dashed" -> StyleValue.DASHED;
            case "thin" -> StyleValue.THIN;
            case "mediumDashDotDot" -> StyleValue.MEDIUM_DASH_DOT_DOT;
            case "slantDashDot" -> StyleValue.SLANT_DASH_DOT;
            case "mediumDashDot" -> StyleValue.MEDIUM_DASH_DOT;
            case "mediumDashed" -> StyleValue.MEDIUM_DASHED;
            case "medium" -> StyleValue.MEDIUM;
            case "thick" -> StyleValue.THICK;
            case "double" -> StyleValue.DOUBLE;
            default -> StyleValue.NONE;
        };
    }
}
