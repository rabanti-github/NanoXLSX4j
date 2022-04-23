/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.styles.AbstractStyle;
import ch.rabanti.nanoxlsx4j.styles.Border;
import ch.rabanti.nanoxlsx4j.styles.CellXf;
import ch.rabanti.nanoxlsx4j.styles.Fill;
import ch.rabanti.nanoxlsx4j.styles.Font;
import ch.rabanti.nanoxlsx4j.styles.NumberFormat;
import ch.rabanti.nanoxlsx4j.styles.Style;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/**
 * Class representing a collection of pre-processed styles and their components. This class is internally used and should not be used otherwise.
 *
 * @author Raphael Stoeckli
 */
public class StyleReaderContainer {

    // ### P R I V A T E  F I E L D S ###

    private List<CellXf> cellXfs = new ArrayList<>();
    private List<NumberFormat> numberFormats = new ArrayList<>();
    private List<Style> styles = new ArrayList<>();
    private List<Border> borders = new ArrayList<>();
    private List<Fill> fills = new ArrayList<>();
    private List<Font> fonts = new ArrayList<>();

    // ### M E T H O D S ###

    /**
     * Gets the number of resolved styles
     *
     * @return Number of style entries in the container
     */
    public int getStyleCount() {
        return this.styles.size();
    }

    /**
     * Adds a style component and determines the appropriate type of it automatically
     *
     * @param component Style component to add to the collections
     * @throws StyleException Thrown if an unknown style component type was passed
     */
    public void addStyleComponent(AbstractStyle component) {
        if (component instanceof CellXf) {
            this.cellXfs.add((CellXf) component);
        } else if (component instanceof NumberFormat) {
            this.numberFormats.add((NumberFormat) component);
        } else if (component instanceof Style) {
            this.styles.add((Style) component);
        } else if (component instanceof Border) {
            this.borders.add((Border) component);
        } else if (component instanceof Fill) {
            this.fills.add((Fill) component);
        } else if (component instanceof Font) {
            this.fonts.add((Font) component);
        } else {
            throw new StyleException("The style definition of the type '" + component.getClass().getName() + "' is unknown or not implemented yet");
        }
    }

    /**
     * Returns a whole style by its index
     *
     * @param index            Index of the style
     * @param returnNullOnFail If true, null will be returned if the style could not be retrieved. Otherwise an exception will be thrown
     * @return Style object or null if parameter returnNullOnFail was set to true and the component could not be retrieved
     * @throws StyleException Thrown StyleException if the style was not found and the parameter returnNullOnFail was set to false
     */
    public Style getStyle(String index, boolean returnNullOnFail) {
        int number;
        try {
            number = Integer.parseInt(index);
            return (Style) getComponent(Style.class, number, returnNullOnFail);
        } catch (Exception ex) {
            if (returnNullOnFail) {
                return null;
            } else {
                throw new StyleException("The style definition could not be retrieved, because of the invalid style index '" + index + "'");
            }
        }
    }

    /**
     * Returns a whole style by its index. It also returns information about the type of the style, regarding dates or times
     *
     * @param index            Index of the style
     * @param returnNullOnFail If true, null will be returned if the style could not be retrieved. Otherwise an exception will be thrown
     * @return Style object or null if parameter returnNullOnFail was set to true and the component could not be retrieved
     * @throws StyleException Thrown StyleException if the style was not found and the parameter returnNullOnFail was set to false
     */
    public StyleResult evaluateDateTimeStyle(int index, boolean returnNullOnFail) {
        Style style = (Style) getComponent(Style.class, index, returnNullOnFail);
        StyleResult result = new StyleResult(style);
        if (style != null) {
            result.setDateStyle(NumberFormat.isDateFormat(style.getNumberFormat().getNumber()));
            result.setTimeStyle(NumberFormat.isTimeFormat(style.getNumberFormat().getNumber()));
        } else {
            result.setDateStyle(false);
            result.setTimeStyle(false);
        }
        return result;
    }

    /**
     * Returns a cell XF component by its index.<br>
     * Note: The method is currently not used but prepared for usage when the style reader is fully implemented
     *
     * @param index            Internal index of the style component
     * @param returnNullOnFail If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown
     * @return Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved
     * @throws StyleException Thrown if the component was not found and the parameter returnNullOnFail was set to false
     */
    public CellXf getCellXF(int index, boolean returnNullOnFail) {
        return (CellXf) getComponent(CellXf.class, index, returnNullOnFail);
    }

    /**
     * Returns a number format component by its index.<br>
     *
     * @param index            Internal index of the style component
     * @param returnNullOnFail If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown
     * @return Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved
     * @throws StyleException Thrown if the component was not found and the parameter returnNullOnFail was set to false
     */
    public NumberFormat getNumberFormat(int index, boolean returnNullOnFail) {
        return (NumberFormat) getComponent(NumberFormat.class, index, returnNullOnFail);
    }

    /**
     * Returns a border component by its index.<br>
     * Note: The method is currently not used but prepared for usage when the style reader is fully implemented
     *
     * @param index            Internal index of the style component
     * @param returnNullOnFail If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown
     * @return Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved
     * @throws StyleException Thrown if the component was not found and the parameter returnNullOnFail was set to false
     */
    public Border getBorder(int index, boolean returnNullOnFail) {
        return (Border) getComponent(Border.class, index, returnNullOnFail);
    }

    /**
     * Returns a fill component by its index.<br>
     * Note: The method is currently not used but prepared for usage when the style reader is fully implemented
     *
     * @param index            Internal index of the style component
     * @param returnNullOnFail If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown
     * @return Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved
     * @throws StyleException Thrown if the component was not found and the parameter returnNullOnFail was set to false
     */
    public Fill getFill(int index, boolean returnNullOnFail) {
        return (Fill) getComponent(Fill.class, index, returnNullOnFail);
    }

    /**
     * Returns a font component by its index.<br>
     * Note: The method is currently not used but prepared for usage when the style reader is fully implemented
     *
     * @param index            Internal index of the style component
     * @param returnNullOnFail If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown
     * @return Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved
     * @throws StyleException Thrown if the component was not found and the parameter returnNullOnFail was set to false
     */
    public Font getFont(int index, boolean returnNullOnFail) {
        return (Font) getComponent(Font.class, index, returnNullOnFail);
    }

    /**
     * Gets the next internal id of a style
     *
     * @return Next id of styles (collected in this class)
     */
    public int getNextStyleId() {
        return this.styles.size();
    }

    /**
     * Gets the next internal id of a cell XF component.<br>
     * The method is currently not used but prepared for usage when the style reader is fully implemented
     *
     * @return Next id of the component type (collected in this class)
     */
    public int getNextCellXFId() {
        return this.cellXfs.size();
    }

    /**
     * Gets the next internal id of a number format component
     *
     * @return Next id of the component type (collected in this class)
     */
    public int getNextNumberFormatId() {
        return this.numberFormats.size();
    }

    /**
     * Gets the next internal id of a border component.<br>
     * The method is currently not used but prepared for usage when the style reader is fully implemented
     *
     * @return Next id of the component type (collected in this class)
     */
    public int getNextBorderId() {
        return this.borders.size();
    }

    /**
     * Gets the next internal id of a fill component.<br>
     * The method is currently not used but prepared for usage when the style reader is fully implemented
     *
     * @return Next id of the component type (collected in this class)
     */
    public int getNextFillId() {
        return this.fills.size();
    }

    /**
     * Gets the next internal id of a font component.<br>
     * The method is currently not used but prepared for usage when the style reader is fully implemented
     *
     * @return Next id of the component type (collected in this class)
     */
    public int getNextFontId() {
        return this.fonts.size() ;
    }

    /**
     * Internal method to retrieve style components
     *
     * @param cls              Class of the style component
     * @param index            Internal index of the style components
     * @param returnNullOnFail If true, null will be returned if the component could not be retrieved. Otherwise an exception will be thrown
     * @return Style component or null if parameter returnNullOnFail was set to true and the component could not be retrieved
     * @throws StyleException Thrown if the component was not found and the parameter returnNullOnFail was set to false
     */
    private <T> AbstractStyle getComponent(T cls, int index, boolean returnNullOnFail) {
        try {
            if (cls.equals(CellXf.class)) {
                return this.cellXfs.get(index);
            } else if (cls.equals(NumberFormat.class)) {
                //Number format entries are handles differently, since identified by 'numFmtId'. Other components are identified by its entry index
                Optional<NumberFormat> result = numberFormats.stream().filter(x -> x.getInternalID() == index).findFirst();
                if (result.isPresent()) {
                    return result.get();
                } else {
                    throw new StyleException("The number format with the numFmtId: " + index + " was not found");
                }
            } else if (cls.equals(Style.class)) {
                return this.styles.get(index);
            } else if (cls.equals(Border.class)) {
                return this.borders.get(index);
            } else if (cls.equals(Fill.class)) {
                return this.fills.get(index);
            } else if (cls.equals(Font.class)) {
                return this.fonts.get(index);
            } else {
                throw new StyleException("The style definition of the type '" + cls.getClass().getName() + "' is unknown or not implemented yet");
            }
        } catch (Exception ex) {
            if (returnNullOnFail) {
                return null;
            } else {
                throw new StyleException("The style definition could not be retrieved. Please see inner exception:", ex);
            }
        }
    }

    // ### S U B - C L A S S E S ###

    /**
     * Result class regarding date and time styles
     */
    public static class StyleResult {

        private boolean isDateStyle;
        private boolean isTimeStyle;
        private Style result;

        /**
         * Gets whether the style is to describe a date
         *
         * @return True if a date style
         */
        public boolean isDateStyle() {
            return isDateStyle;
        }

        /**
         * Sets whether the style is to describe a date
         *
         * @param dateStyle True if a date style
         */
        public void setDateStyle(boolean dateStyle) {
            isDateStyle = dateStyle;
        }

        /**
         * Gets whether the style is to describe a time
         *
         * @return True if a time style
         */
        public boolean isTimeStyle() {
            return isTimeStyle;
        }

        /**
         * Sets whether the style is to describe a time
         *
         * @param timeStyle True if a time style
         */
        public void setTimeStyle(boolean timeStyle) {
            isTimeStyle = timeStyle;
        }

        /**
         * Gets the style as result
         *
         * @return Style component
         */
        public Style getResult() {
            return result;
        }

        /**
         * Sets the style as result
         *
         * @param result Style component
         */
        public void setResult(Style result) {
            this.result = result;
        }

        /**
         * Constructor with definition of the result
         *
         * @param style Style component
         */
        public StyleResult(Style style) {
            this.result = style;
        }

    }

}
