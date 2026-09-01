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
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

/**
 * Class representing an XF entry. The XF entry references other style instances such as borders or fills and defines
 * the positioning of cell content.
 */
public class CellXf extends AbstractStyle {

    /** Default horizontal alignment. */
    public static final HorizontalAlignValue DEFAULT_HORIZONTAL_ALIGNMENT = HorizontalAlignValue.NONE;

    /** Default text-break option. */
    public static final TextBreakValue DEFAULT_ALIGNMENT = TextBreakValue.NONE;

    /** Default text direction. */
    public static final TextDirectionValue DEFAULT_TEXT_DIRECTION = TextDirectionValue.HORIZONTAL;

    /** Default vertical alignment. */
    public static final VerticalAlignValue DEFAULT_VERTICAL_ALIGNMENT = VerticalAlignValue.NONE;

    /** Enum for horizontal cell alignment. */
    public enum HorizontalAlignValue {
        /** Align content left. */
        LEFT,
        /** Center content. */
        CENTER,
        /** Align content right. */
        RIGHT,
        /** Fill the cell with the content. */
        FILL,
        /** Justify content. */
        JUSTIFY,
        /** Use general alignment. */
        GENERAL,
        /** Use center-continuous alignment. */
        CENTER_CONTINUOUS,
        /** Use distributed alignment. */
        DISTRIBUTED,
        /** Do not apply a horizontal alignment. */
        NONE
    }

    /** Enum for text-break options. */
    public enum TextBreakValue {
        /** Wrap text. */
        WRAP_TEXT,
        /** Shrink text to fit the cell. */
        SHRINK_TO_FIT,
        /** Allow text to overflow the cell. */
        NONE
    }

    /** Enum for the general text direction. */
    public enum TextDirectionValue {
        /** Horizontal text direction. */
        HORIZONTAL,
        /** Vertical text direction. */
        VERTICAL
    }

    /** Enum for vertical cell alignment. */
    public enum VerticalAlignValue {
        /** Align content at the bottom. */
        BOTTOM,
        /** Align content at the top. */
        TOP,
        /** Center content. */
        CENTER,
        /** Justify content. */
        JUSTIFY,
        /** Use distributed alignment. */
        DISTRIBUTED,
        /** Do not apply a vertical alignment. */
        NONE
    }

    @AppendAnnotation
    private boolean forceApplyAlignment;
    @AppendAnnotation
    private boolean hidden;
    @AppendAnnotation
    private HorizontalAlignValue horizontalAlign;
    @AppendAnnotation
    private boolean locked;
    @AppendAnnotation
    private TextBreakValue alignment;
    @AppendAnnotation
    private TextDirectionValue textDirection;
    @AppendAnnotation
    private int textRotation;
    @AppendAnnotation
    private VerticalAlignValue verticalAlign;
    @AppendAnnotation
    private int indent;

    /** Creates an XF entry with Excel's default alignment and protection values. */
    public CellXf() {
        horizontalAlign = DEFAULT_HORIZONTAL_ALIGNMENT;
        alignment = DEFAULT_ALIGNMENT;
        textDirection = DEFAULT_TEXT_DIRECTION;
        verticalAlign = DEFAULT_VERTICAL_ALIGNMENT;
        locked = true;
        textRotation = 0;
        indent = 0;
    }

    /**
     * Gets whether the apply-alignment attribute is forced.
     *
     * @return True if apply-alignment is forced
     */
    public boolean isForceApplyAlignment() {
        return forceApplyAlignment;
    }

    /**
     * Sets whether the apply-alignment attribute is forced.
     *
     * @param forceApplyAlignment Whether apply-alignment is forced
     */
    public void setForceApplyAlignment(boolean forceApplyAlignment) {
        this.forceApplyAlignment = forceApplyAlignment;
    }

    /**
     * Gets whether the hidden protection attribute is enabled.
     *
     * @return True if hidden protection is enabled
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Sets whether the hidden protection attribute is enabled.
     *
     * @param hidden Whether hidden protection is enabled
     */
    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }

    /**
     * Gets the horizontal alignment.
     *
     * @return Horizontal alignment
     */
    public HorizontalAlignValue getHorizontalAlign() {
        return horizontalAlign;
    }

    /**
     * Sets the horizontal alignment.
     *
     * @param horizontalAlign Horizontal alignment
     */
    public void setHorizontalAlign(HorizontalAlignValue horizontalAlign) {
        this.horizontalAlign = Objects.requireNonNull(horizontalAlign, "horizontalAlign");
    }

    /**
     * Gets whether locked protection is enabled.
     *
     * @return True if locked protection is enabled
     */
    public boolean isLocked() {
        return locked;
    }

    /**
     * Sets whether locked protection is enabled.
     *
     * @param locked Whether locked protection is enabled
     */
    public void setLocked(boolean locked) {
        this.locked = locked;
    }

    /**
     * Gets the text-break option.
     *
     * @return Text-break option
     */
    public TextBreakValue getAlignment() {
        return alignment;
    }

    /**
     * Sets the text-break option.
     *
     * @param alignment Text-break option
     */
    public void setAlignment(TextBreakValue alignment) {
        this.alignment = Objects.requireNonNull(alignment, "alignment");
    }

    /**
     * Gets the direction of the text within the cell.
     *
     * @return Text direction
     */
    public TextDirectionValue getTextDirection() {
        return textDirection;
    }

    /**
     * Sets the direction of the text within the cell and recalculates its internal rotation.
     *
     * @param textDirection Text direction
     * @throws FormatException if the current rotation is outside the supported range
     */
    public void setTextDirection(TextDirectionValue textDirection) {
        this.textDirection = Objects.requireNonNull(textDirection, "textDirection");
        calculateInternalRotation();
    }

    /**
     * Gets the text rotation in degrees.
     *
     * @return Text rotation from -90 through 90, or 255 for vertical text
     */
    public int getTextRotation() {
        return textRotation;
    }

    /**
     * Sets the text rotation in degrees and resets the direction to horizontal.
     *
     * @param textRotation Text rotation from -90 through 90
     * @throws FormatException if the rotation is outside the supported range
     */
    public void setTextRotation(int textRotation) {
        this.textRotation = textRotation;
        setTextDirection(TextDirectionValue.HORIZONTAL);
        calculateInternalRotation();
    }

    /**
     * Gets the vertical alignment.
     *
     * @return Vertical alignment
     */
    public VerticalAlignValue getVerticalAlign() {
        return verticalAlign;
    }

    /**
     * Sets the vertical alignment.
     *
     * @param verticalAlign Vertical alignment
     */
    public void setVerticalAlign(VerticalAlignValue verticalAlign) {
        this.verticalAlign = Objects.requireNonNull(verticalAlign, "verticalAlign");
    }

    /**
     * Gets the indentation used with left, right, or distributed alignment.
     *
     * @return Indentation level
     */
    public int getIndent() {
        return indent;
    }

    /**
     * Sets the indentation used with left, right, or distributed alignment.
     *
     * @param indent Indentation level
     * @throws StyleException if the indentation is negative
     */
    public void setIndent(int indent) {
        if (indent < 0) {
            throw new StyleException("The indent value '" + indent + "' is not valid. It must be >= 0");
        }
        this.indent = indent;
    }

    /**
     * Calculates the OOXML rotation value represented by the direction and rotation properties.
     *
     * @return Internal OOXML rotation value
     * @throws FormatException if the rotation is outside the range -90 through 90
     */
    int calculateInternalRotation() {
        if (textRotation < -90 || textRotation > 90) {
            throw new FormatException("The rotation value (" + ParserUtils.toString(textRotation)
                    + "°) is out of range. Range is form -90° to +90°");
        }
        if (textDirection == TextDirectionValue.VERTICAL) {
            textRotation = 255;
            return textRotation;
        }
        if (textRotation >= 0) {
            return textRotation;
        }
        return 90 - textRotation;
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("\"StyleXF\": {\n");
        addPropertyAsJson(sb, "HorizontalAlign", horizontalAlign, false);
        addPropertyAsJson(sb, "Alignment", alignment, false);
        addPropertyAsJson(sb, "TextDirection", textDirection, false);
        addPropertyAsJson(sb, "TextRotation", textRotation, false);
        addPropertyAsJson(sb, "VerticalAlign", verticalAlign, false);
        addPropertyAsJson(sb, "ForceApplyAlignment", forceApplyAlignment, false);
        addPropertyAsJson(sb, "Locked", locked, false);
        addPropertyAsJson(sb, "Hidden", hidden, false);
        addPropertyAsJson(sb, "Indent", indent, false);
        addPropertyAsJson(sb, "HashCode", hashCode(), true);
        sb.append("\n}");
        return sb.toString();
    }

    @Override
    public int hashCode() {
        return Objects.hash(forceApplyAlignment, hidden, horizontalAlign, locked, alignment, textDirection,
                textRotation, verticalAlign, indent);
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (!(obj instanceof CellXf other)) {
            return false;
        }
        return forceApplyAlignment == other.forceApplyAlignment
                && hidden == other.hidden
                && horizontalAlign == other.horizontalAlign
                && locked == other.locked
                && alignment == other.alignment
                && textDirection == other.textDirection
                && textRotation == other.textRotation
                && verticalAlign == other.verticalAlign
                && indent == other.indent;
    }

    /**
     * Copies this XF entry without its internal ID.
     *
     * @return Dereferenced copy
     */
    @Override
    public AbstractStyle copy() {
        CellXf copy = new CellXf();
        copy.horizontalAlign = horizontalAlign;
        copy.alignment = alignment;
        copy.textDirection = textDirection;
        copy.textRotation = textRotation;
        copy.verticalAlign = verticalAlign;
        copy.forceApplyAlignment = forceApplyAlignment;
        copy.locked = locked;
        copy.hidden = hidden;
        copy.indent = indent;
        return copy;
    }

    /**
     * Copies this XF entry without requiring a cast.
     *
     * @return Dereferenced copy
     */
    public CellXf copyCellXf() {
        return (CellXf) copy();
    }

    /**
     * Converts a horizontal alignment to its OOXML string representation.
     *
     * @param align Alignment to convert
     * @return OOXML alignment name, or an empty string for none
     */
    static String getHorizontalAlignName(HorizontalAlignValue align) {
        return switch (align) {
            case LEFT -> "left";
            case CENTER -> "center";
            case RIGHT -> "right";
            case FILL -> "fill";
            case JUSTIFY -> "justify";
            case GENERAL -> "general";
            case CENTER_CONTINUOUS -> "centerContinuous";
            case DISTRIBUTED -> "distributed";
            case NONE -> "";
        };
    }

    /**
     * Parses an OOXML horizontal alignment name.
     *
     * @param name Alignment name
     * @return Parsed alignment, or {@link HorizontalAlignValue#NONE} for an unknown value
     */
    static HorizontalAlignValue getHorizontalAlignEnum(String name) {
        if (name == null) {
            return HorizontalAlignValue.NONE;
        }
        return switch (name) {
            case "left" -> HorizontalAlignValue.LEFT;
            case "center" -> HorizontalAlignValue.CENTER;
            case "right" -> HorizontalAlignValue.RIGHT;
            case "fill" -> HorizontalAlignValue.FILL;
            case "justify" -> HorizontalAlignValue.JUSTIFY;
            case "general" -> HorizontalAlignValue.GENERAL;
            case "centerContinuous" -> HorizontalAlignValue.CENTER_CONTINUOUS;
            case "distributed" -> HorizontalAlignValue.DISTRIBUTED;
            default -> HorizontalAlignValue.NONE;
        };
    }

    /**
     * Converts a vertical alignment to its OOXML string representation.
     *
     * @param align Alignment to convert
     * @return OOXML alignment name, or an empty string for none
     */
    static String getVerticalAlignName(VerticalAlignValue align) {
        return switch (align) {
            case BOTTOM -> "bottom";
            case TOP -> "top";
            case CENTER -> "center";
            case JUSTIFY -> "justify";
            case DISTRIBUTED -> "distributed";
            case NONE -> "";
        };
    }

    /**
     * Parses an OOXML vertical alignment name.
     *
     * @param name Alignment name
     * @return Parsed alignment, or {@link VerticalAlignValue#NONE} for an unknown value
     */
    static VerticalAlignValue getVerticalAlignEnum(String name) {
        if (name == null) {
            return VerticalAlignValue.NONE;
        }
        return switch (name) {
            case "bottom" -> VerticalAlignValue.BOTTOM;
            case "top" -> VerticalAlignValue.TOP;
            case "center" -> VerticalAlignValue.CENTER;
            case "justify" -> VerticalAlignValue.JUSTIFY;
            case "distributed" -> VerticalAlignValue.DISTRIBUTED;
            default -> VerticalAlignValue.NONE;
        };
    }
}
