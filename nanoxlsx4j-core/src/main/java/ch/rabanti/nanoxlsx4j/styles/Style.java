/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import java.util.Optional;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

/**
 * Container for the components that form a cell style.
 * <p>
 * The component instances contain the actual formatting information. A style initializes every component with its
 * default value and can selectively append the non-default properties of another style or component.
 */
public class Style extends AbstractStyle {

    @AppendAnnotation(nestedProperty = true)
    private Border currentBorder;

    @AppendAnnotation(nestedProperty = true)
    private CellXf currentCellXf;

    @AppendAnnotation(nestedProperty = true)
    private Fill currentFill;

    @AppendAnnotation(nestedProperty = true)
    private Font currentFont;

    @AppendAnnotation(nestedProperty = true)
    private NumberFormat currentNumberFormat;

    @AppendAnnotation(ignore = true)
    private String name;

    @AppendAnnotation(ignore = true)
    private boolean internalStyle;

    /**
     * Creates a style containing default component instances. Its informal name is derived from its initial hash.
     */
    public Style() {
        initializeComponents();
        name = Integer.toString(hashCode());
    }

    /**
     * Creates a style containing default component instances.
     *
     * @param name informal style name
     */
    public Style(String name) {
        initializeComponents();
        this.name = name;
    }

    /**
     * Creates a style with a forced serialization order.
     *
     * @param name          informal style name
     * @param forcedOrder   style position used for sorting
     * @param internalStyle whether this is a system-internal style that is not meant to be altered
     */
    public Style(String name, int forcedOrder, boolean internalStyle) {
        initializeComponents();
        this.name = name;
        setInternalId(Optional.of(forcedOrder));
        this.internalStyle = internalStyle;
    }

    private void initializeComponents() {
        currentBorder = new Border();
        currentCellXf = new CellXf();
        currentFill = new Fill();
        currentFont = new Font();
        currentNumberFormat = new NumberFormat();
    }

    /** @return current border component */
    public Border getCurrentBorder() {
        return currentBorder;
    }

    /** @param currentBorder border component */
    public void setCurrentBorder(Border currentBorder) {
        this.currentBorder = currentBorder;
    }

    /** @return current cell-format component */
    public CellXf getCurrentCellXf() {
        return currentCellXf;
    }

    /** @param currentCellXf cell-format component */
    public void setCurrentCellXf(CellXf currentCellXf) {
        this.currentCellXf = currentCellXf;
    }

    /** @return current fill component */
    public Fill getCurrentFill() {
        return currentFill;
    }

    /** @param currentFill fill component */
    public void setCurrentFill(Fill currentFill) {
        this.currentFill = currentFill;
    }

    /** @return current font component */
    public Font getCurrentFont() {
        return currentFont;
    }

    /** @param currentFont font component */
    public void setCurrentFont(Font currentFont) {
        this.currentFont = currentFont;
    }

    /** @return current number-format component */
    public NumberFormat getCurrentNumberFormat() {
        return currentNumberFormat;
    }

    /** @param currentNumberFormat number-format component */
    public void setCurrentNumberFormat(NumberFormat currentNumberFormat) {
        this.currentNumberFormat = currentNumberFormat;
    }

    /**
     * Gets the informal style name. It is not used as an identifier when workbook styles are collected.
     *
     * @return informal style name
     */
    public String getName() {
        return name;
    }

    /**
     * Sets the informal style name. It is not used as an identifier when workbook styles are collected.
     *
     * @param name informal style name
     */
    public void setName(String name) {
        this.name = name;
    }

    /** @return whether this is a system-internal style */
    public boolean isInternalStyle() {
        return internalStyle;
    }

    /**
     * Appends the non-default properties of a style component or compound style to this style.
     *
     * @param styleToAppend style or component to append
     * @return this style, for method chaining
     */
    public Style append(AbstractStyle styleToAppend) {
        if (styleToAppend == null) {
            return this;
        }

        Class<?> type = styleToAppend.getClass();
        if (type == Border.class) {
            currentBorder.copyFields((Border) styleToAppend, new Border());
        } else if (type == CellXf.class) {
            currentCellXf.copyFields((CellXf) styleToAppend, new CellXf());
        } else if (type == Fill.class) {
            currentFill.copyFields((Fill) styleToAppend, new Fill());
        } else if (type == Font.class) {
            currentFont.copyFields((Font) styleToAppend, new Font());
        } else if (type == NumberFormat.class) {
            currentNumberFormat.copyFields((NumberFormat) styleToAppend, new NumberFormat());
        } else if (type == Style.class) {
            Style style = (Style) styleToAppend;
            currentBorder.copyFields(style.currentBorder, new Border());
            currentCellXf.copyFields(style.currentCellXf, new CellXf());
            currentFill.copyFields(style.currentFill, new Fill());
            currentFont.copyFields(style.currentFont, new Font());
            currentNumberFormat.copyFields(style.currentNumberFormat, new NumberFormat());
        }
        return this;
    }

    /**
     * Creates a deep copy of this compound style without its informal name, internal flag, or internal ID.
     *
     * @return dereferenced copy
     * @throws StyleException if a component reference is missing
     */
    @Override
    public AbstractStyle copy() {
        ensureComponentsPresent("The style could not be copied because one or more components are missing as references");

        Style copy = new Style();
        copy.currentBorder = currentBorder.copyBorder();
        copy.currentCellXf = currentCellXf.copyCellXf();
        copy.currentFill = currentFill.copyFill();
        copy.currentFont = currentFont.copyFont();
        copy.currentNumberFormat = currentNumberFormat.copyNumberFormat();
        return copy;
    }

    /** @return dereferenced style copy without requiring a cast */
    public Style copyStyle() {
        return (Style) copy();
    }

    /**
     * Calculates a compound hash from all style components. The informal name and internal metadata are excluded.
     *
     * @return compound style hash
     * @throws StyleException if a component reference is missing
     */
    @Override
    public int hashCode() {
        ensureComponentsPresent(
                "The hash of the style could not be created because one or more components are missing as references");

        int prime = 241;
        int result = 1;
        result *= prime + currentBorder.hashCode();
        result *= prime + currentCellXf.hashCode();
        result *= prime + currentFill.hashCode();
        result *= prime + currentFont.hashCode();
        result *= prime + currentNumberFormat.hashCode();
        return result;
    }

    private void ensureComponentsPresent(String message) {
        if (currentBorder == null || currentCellXf == null || currentFill == null || currentFont == null
                || currentNumberFormat == null) {
            throw new StyleException(message);
        }
    }

    /** @return debug representation containing this style and all its components */
    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("{\n\"Style\": {\n");
        addPropertyAsJson(sb, "Name", name, false);
        addPropertyAsJson(sb, "HashCode", hashCode(), false);
        sb.append(currentBorder).append(",\n");
        sb.append(currentCellXf).append(",\n");
        sb.append(currentFill).append(",\n");
        sb.append(currentFont).append(",\n");
        sb.append(currentNumberFormat).append("\n}\n}");
        return sb.toString();
    }
}
