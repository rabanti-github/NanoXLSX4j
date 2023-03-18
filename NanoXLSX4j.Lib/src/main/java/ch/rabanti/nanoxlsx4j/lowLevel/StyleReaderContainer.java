/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.lowLevel;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.styles.AbstractStyle;
import ch.rabanti.nanoxlsx4j.styles.Border;
import ch.rabanti.nanoxlsx4j.styles.CellXf;
import ch.rabanti.nanoxlsx4j.styles.Fill;
import ch.rabanti.nanoxlsx4j.styles.Font;
import ch.rabanti.nanoxlsx4j.styles.NumberFormat;
import ch.rabanti.nanoxlsx4j.styles.Style;

/**
 * Class representing a collection of pre-processed styles and their components.
 * This class is internally used and should not be used otherwise.
 *
 * @author Raphael Stoeckli
 */
public class StyleReaderContainer {

	// ### P R I V A T E F I E L D S ###

	private final List<CellXf> cellXfs = new ArrayList<>();
	private final List<NumberFormat> numberFormats = new ArrayList<>();
	private final List<Style> styles = new ArrayList<>();
	private final List<Border> borders = new ArrayList<>();
	private final List<Fill> fills = new ArrayList<>();
	private final List<Font> fonts = new ArrayList<>();
	private final List<String> mruColors = new ArrayList<>();

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
	 * Adds a style component and determines the appropriate type of it
	 * automatically
	 *
	 * @param component
	 *            Style component to add to the collections
	 */
	public void addStyleComponent(AbstractStyle component) {
		if (component instanceof CellXf) {
			this.cellXfs.add((CellXf) component);
		}
		else if (component instanceof NumberFormat) {
			this.numberFormats.add((NumberFormat) component);
		}
		else if (component instanceof Style) {
			this.styles.add((Style) component);
		}
		else if (component instanceof Border) {
			this.borders.add((Border) component);
		}
		else if (component instanceof Fill) {
			this.fills.add((Fill) component);
		}
		else if (component instanceof Font) {
			this.fonts.add((Font) component);
		}
	}

	/**
	 * Returns a whole style by its index
	 *
	 * @param index
	 *            Index of the style
	 * @return Style object or null if the component could not be retrieved
	 */
	public Style getStyle(String index) {
		int number;
		try {
			number = Integer.parseInt(index);
			return (Style) getComponent(Style.class, number);
		}
		catch (Exception ex) {
			return null;
		}
	}

	/**
	 * Returns a whole style by its index. It also returns information about the
	 * type of the style, regarding dates or times
	 *
	 * @param index
	 *            Index of the style
	 * @return Style object or null if the component could not be retrieved
	 */
	public StyleResult evaluateDateTimeStyle(int index) {
		Style style = (Style) getComponent(Style.class, index);
		StyleResult result = new StyleResult(style);
		if (style != null) {
			result.setDateStyle(NumberFormat.isDateFormat(style.getNumberFormat().getNumber()));
			result.setTimeStyle(NumberFormat.isTimeFormat(style.getNumberFormat().getNumber()));
		}
		return result;
	}

	/**
	 * Returns a number format component by its index
	 *
	 * @param index
	 *            Internal index of the style component
	 * @return Style component or null if the component could not be retrieved
	 * @throws StyleException
	 *             Thrown if the component was not found
	 */
	public NumberFormat getNumberFormat(int index) {
		return (NumberFormat) getComponent(NumberFormat.class, index);
	}

	/**
	 * Returns a border component by its index
	 *
	 * @param index
	 *            Internal index of the style component
	 * @return Style component or null if the component could not be retrieved
	 */
	public Border getBorder(int index) {
		return (Border) getComponent(Border.class, index);
	}

	/**
	 * Returns a fill component by its index
	 *
	 * @param index
	 *            Internal index of the style component
	 * @return Style component or null if the component could not be retrieved
	 */
	public Fill getFill(int index) {
		return (Fill) getComponent(Fill.class, index);
	}

	/**
	 * Returns a font component by its index
	 *
	 * @param index
	 *            Internal index of the style component
	 * @return Style component or null if the component could not be retrieved
	 */
	public Font getFont(int index) {
		return (Font) getComponent(Font.class, index);
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
	 * The method is currently not used but prepared for usage when the style reader
	 * is fully implemented
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
	 * The method is currently not used but prepared for usage when the style reader
	 * is fully implemented
	 *
	 * @return Next id of the component type (collected in this class)
	 */
	public int getNextBorderId() {
		return this.borders.size();
	}

	/**
	 * Gets the next internal id of a fill component.<br>
	 * The method is currently not used but prepared for usage when the style reader
	 * is fully implemented
	 *
	 * @return Next id of the component type (collected in this class)
	 */
	public int getNextFillId() {
		return this.fills.size();
	}

	/**
	 * Gets the next internal id of a font component.<br>
	 * The method is currently not used but prepared for usage when the style reader
	 * is fully implemented
	 *
	 * @return Next id of the component type (collected in this class)
	 */
	public int getNextFontId() {
		return this.fonts.size();
	}

	/**
	 * Internal method to retrieve style components.<br>
	 * Note: CellXF is not handled, since retrieved in the style reader in a
	 * different way
	 *
	 * @param cls
	 *            Class of the style component
	 * @param index
	 *            Internal index of the style components
	 * @return Style component or null if the component could not be retrieved
	 */
	private <T> AbstractStyle getComponent(T cls, int index) {
		try {
			if (cls.equals(NumberFormat.class)) {
				// Number format entries are handles differently, since identified by
				// 'numFmtId'. Other components are identified by its entry index
				Optional<NumberFormat> result = numberFormats.stream().filter(x -> x.getInternalID() == index).findFirst();
				if (result.isPresent()) {
					return result.get();
				}
				else {
					throw new StyleException("The number format with the numFmtId: " + index + " was not found");
				}
			}
			else if (cls.equals(Style.class)) {
				return this.styles.get(index);
			}
			else if (cls.equals(Border.class)) {
				return this.borders.get(index);
			}
			else if (cls.equals(Fill.class)) {
				return this.fills.get(index);
			}
			else { // must be font (CellXF is not handled here)
				return this.fonts.get(index);
			}
		}
		catch (Exception ex) {
			// ignore
		}
		return null;
	}

	/**
	 * Adds a color value to the MRU list
	 *
	 * @param value
	 *            ARGB value
	 */
	void addMruColor(String value) {
		this.mruColors.add(value);
	}

	/**
	 * Gets the MRU colors as list
	 *
	 * @return ARGB values
	 */
	List<String> getMruColors() {
		return this.mruColors;
	}

	// ### S U B - C L A S S E S ###

	/**
	 * Result class regarding date and time styles
	 */
	public static class StyleResult {

		private boolean isDateStyle = false;
		private boolean isTimeStyle = false;
		private final Style result;

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
		 * @param dateStyle
		 *            True if a date style
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
		 * @param timeStyle
		 *            True if a time style
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
		 * Constructor with definition of the result
		 *
		 * @param style
		 *            Style component
		 */
		public StyleResult(Style style) {
			this.result = style;
		}

	}

}
