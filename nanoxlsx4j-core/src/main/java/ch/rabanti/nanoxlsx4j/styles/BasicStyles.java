/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.utils.ParserUtils;
import ch.rabanti.nanoxlsx4j.utils.Validators;

/**
 * Factory for commonly used styles.
 * <p>
 * Every method returns a new style. Callers can therefore alter or combine the result without changing styles
 * returned by later calls.
 */
public final class BasicStyles {

    private BasicStyles() {
        // Do not instantiate
    }

    /** @return a style with a bold font */
    public static Style getBold() {
        Style style = new Style();
        style.getCurrentFont().setBold(true);
        return style;
    }

    /** @return a style with an italic font */
    public static Style getItalic() {
        Style style = new Style();
        style.getCurrentFont().setItalic(true);
        return style;
    }

    /** @return a style with a bold and italic font */
    public static Style getBoldItalic() {
        Style style = new Style();
        style.getCurrentFont().setBold(true);
        style.getCurrentFont().setItalic(true);
        return style;
    }

    /** @return a style with a single-underlined font */
    public static Style getUnderline() {
        Style style = new Style();
        style.getCurrentFont().setUnderline(Font.UnderlineValue.SINGLE);
        return style;
    }

    /** @return a style with a double-underlined font */
    public static Style getDoubleUnderline() {
        Style style = new Style();
        style.getCurrentFont().setUnderline(Font.UnderlineValue.DOUBLE);
        return style;
    }

    /** @return a style with a struck-through font */
    public static Style getStrike() {
        Style style = new Style();
        style.getCurrentFont().setStrike(true);
        return style;
    }

    /** @return a style using Excel's built-in date format 14 */
    public static Style getDateFormat() {
        Style style = new Style();
        style.getCurrentNumberFormat().setNumber(NumberFormat.FormatNumber.FORMAT_14);
        return style;
    }

    /** @return a style using Excel's built-in time format 21 */
    public static Style getTimeFormat() {
        Style style = new Style();
        style.getCurrentNumberFormat().setNumber(NumberFormat.FormatNumber.FORMAT_21);
        return style;
    }

    /** @return a style that displays a number rounded to an integer */
    public static Style getRoundFormat() {
        Style style = new Style();
        style.getCurrentNumberFormat().setNumber(NumberFormat.FormatNumber.FORMAT_1);
        return style;
    }

    /** @return a style with a thin border on all four sides */
    public static Style getBorderFrame() {
        Style style = new Style();
        style.getCurrentBorder().setTopStyle(Border.StyleValue.THIN);
        style.getCurrentBorder().setBottomStyle(Border.StyleValue.THIN);
        style.getCurrentBorder().setLeftStyle(Border.StyleValue.THIN);
        style.getCurrentBorder().setRightStyle(Border.StyleValue.THIN);
        return style;
    }

    /** @return a bold header style with thin side/top borders and a medium bottom border */
    public static Style getBorderFrameHeader() {
        Style style = new Style();
        style.getCurrentBorder().setTopStyle(Border.StyleValue.THIN);
        style.getCurrentBorder().setBottomStyle(Border.StyleValue.MEDIUM);
        style.getCurrentBorder().setLeftStyle(Border.StyleValue.THIN);
        style.getCurrentBorder().setRightStyle(Border.StyleValue.THIN);
        style.getCurrentFont().setBold(true);
        return style;
    }

    /**
     * Gets the 12.5 percent gray pattern used for compatibility purposes.
     *
     * @return a style using the {@link Fill.PatternValue#GRAY_125} fill pattern
     */
    public static Style getDottedFill0125() {
        Style style = new Style();
        style.getCurrentFill().setPatternFill(Fill.PatternValue.GRAY_125);
        return style;
    }

    /** @return a style that forces alignment to be applied to merged cells */
    public static Style getMergeCellStyle() {
        Style style = new Style();
        style.getCurrentCellXf().setForceApplyAlignment(true);
        return style;
    }

    /**
     * Creates a style with an opaque RGB font color.
     *
     * @param rgb six-character RGB value, without an alpha component
     * @return style with the requested font color
     * @throws ch.rabanti.nanoxlsx4j.exceptions.StyleException if {@code rgb} is not a valid RGB value
     */
    public static Style colorizedText(String rgb) {
        Validators.validateColor(rgb, false);
        Style style = new Style();
        style.getCurrentFont().setColorValue(ParserUtils.toUpper("FF" + rgb));
        return style;
    }

    /**
     * Creates a style with an opaque solid RGB background.
     *
     * @param rgb six-character RGB value, without an alpha component
     * @return style with the requested background color
     * @throws ch.rabanti.nanoxlsx4j.exceptions.StyleException if {@code rgb} is not a valid RGB value
     */
    public static Style colorizedBackground(String rgb) {
        Validators.validateColor(rgb, false);
        Style style = new Style();
        style.getCurrentFill().setColor(ParserUtils.toUpper("FF" + rgb), Fill.FillType.FILL_COLOR);
        return style;
    }

    /**
     * Creates a style with the requested font name and default font size.
     *
     * @param fontName font name
     * @return style with the requested font
     */
    public static Style font(String fontName) {
        return font(fontName, Font.DEFAULT_FONT_SIZE, false, false);
    }

    /**
     * Creates a style with the requested font name and size.
     *
     * @param fontName font name
     * @param fontSize font size in points; values outside the supported range are clamped by {@link Font}
     * @return style with the requested font
     */
    public static Style font(String fontName, float fontSize) {
        return font(fontName, fontSize, false, false);
    }

    /**
     * Creates a style with the requested font name, size, and bold state.
     *
     * @param fontName font name
     * @param fontSize font size in points; values outside the supported range are clamped by {@link Font}
     * @param bold     whether the font is bold
     * @return style with the requested font
     */
    public static Style font(String fontName, float fontSize, boolean bold) {
        return font(fontName, fontSize, bold, false);
    }

    /**
     * Creates a style with the requested font properties.
     * <p>
     * NanoXLSX4j cannot validate whether a font is installed or whether it supplies bold and italic variants. Excel
     * may use a fallback font when the requested font is unavailable.
     *
     * @param fontName font name
     * @param fontSize font size in points; values outside the supported range are clamped by {@link Font}
     * @param bold     whether the font is bold
     * @param italic   whether the font is italic
     * @return style with the requested font
     */
    public static Style font(String fontName, float fontSize, boolean bold, boolean italic) {
        Style style = new Style();
        style.getCurrentFont().setName(fontName);
        style.getCurrentFont().setSize(fontSize);
        style.getCurrentFont().setBold(bold);
        style.getCurrentFont().setItalic(italic);
        return style;
    }
}
