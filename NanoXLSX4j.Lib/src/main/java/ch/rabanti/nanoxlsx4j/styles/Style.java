/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

/**
 * Class representing a style which consists of several components
 *
 * @author Raphael Stoeckli
 */
public class Style extends AbstractStyle {

	// ### P R I V A T E F I E L D S ###
	@AppendAnnotation(nestedProperty = true)
	private Border borderRef;
	@AppendAnnotation(nestedProperty = true)
	private CellXf cellXfRef;
	@AppendAnnotation(nestedProperty = true)
	private Fill fillRef;
	@AppendAnnotation(nestedProperty = true)
	private Font fontRef;
	@AppendAnnotation(nestedProperty = true)
	private NumberFormat numberFormatRef;
	@AppendAnnotation(ignore = true)
	private String name;
	@AppendAnnotation(ignore = true)
	private boolean internalStyle;

	// ### G E T T E R S & S E T T E R S ###

	/**
	 * Gets the border component of the style
	 *
	 * @return Border of the style
	 */
	public Border getBorder() {
		return borderRef;
	}

	/**
	 * Gets the cellXf component of the style
	 *
	 * @return CellXf of the style
	 */
	public CellXf getCellXf() {
		return cellXfRef;
	}

	/**
	 * Gets the fill component of the style
	 *
	 * @return Fill of the style
	 */
	public Fill getFill() {
		return fillRef;
	}

	/**
	 * Gets the font component of the style
	 *
	 * @return Font of the style
	 */
	public Font getFont() {
		return fontRef;
	}

	/**
	 * Gets the number format component of the style
	 *
	 * @return Number format of the style
	 */
	public NumberFormat getNumberFormat() {
		return numberFormatRef;
	}

	/**
	 * Sets the border component of the style
	 *
	 * @param borderRef
	 *            Border of the style
	 */
	public void setBorder(Border borderRef) {
		this.borderRef = borderRef;
	}

	/**
	 * Sets the cellXf component of the style
	 *
	 * @param cellXfRef
	 *            CellXf of the style
	 */
	public void setCellXf(CellXf cellXfRef) {
		this.cellXfRef = cellXfRef;
	}

	/**
	 * Sets the fill component of the style
	 *
	 * @param fillRef
	 *            Fill of the style
	 */
	public void setFill(Fill fillRef) {
		this.fillRef = fillRef;
	}

	/**
	 * Sets the font component of the style
	 *
	 * @param fontRef
	 *            Font of the style
	 */
	public void setFont(Font fontRef) {
		this.fontRef = fontRef;
	}

	/**
	 * Sets the number format component of the style
	 *
	 * @param numberFormatRef
	 *            Number format of the style
	 */
	public void setNumberFormat(NumberFormat numberFormatRef) {
		this.numberFormatRef = numberFormatRef;
	}

	/**
	 * Gets the informal name of the style.
	 *
	 * @return Name
	 */
	public String getName() {
		return name;
	}

	/**
	 * Sets the name of the style. If not defined, the automatically calculated hash
	 * will be used as name
	 *
	 * @param name
	 *            Name
	 * @apiNote The name is informal and not considered as an identifier, when
	 *          collecting all styles for a workbook
	 */
	public void setName(String name) {
		this.name = name;
	}

	/**
	 * Gets whether the style is system internal
	 *
	 * @return If true, the style is an internal style. Such styles are not meant to
	 *         be altered
	 */
	public boolean isInternalStyle() {
		return internalStyle;
	}

	// ### C O N S T R U C T O R S ###

	/**
	 * Default constructor
	 */
	public Style() {
		this.borderRef = new Border();
		this.cellXfRef = new CellXf();
		this.fillRef = new Fill();
		this.fontRef = new Font();
		this.numberFormatRef = new NumberFormat();
		this.name = Integer.toString(this.hashCode());
	}

	/**
	 * Constructor with parameters
	 *
	 * @param name
	 *            Name of the style
	 */
	public Style(String name) {
		this.borderRef = new Border();
		this.cellXfRef = new CellXf();
		this.fillRef = new Fill();
		this.fontRef = new Font();
		this.numberFormatRef = new NumberFormat();
		this.name = name;
	}

	/**
	 * Constructor with parameters (internal use)
	 *
	 * @param name
	 *            Name of the style
	 * @param forcedOrder
	 *            Number of the style for sorting purpose. The style will be placed
	 *            at this position (internal use only)
	 * @param internal
	 *            If true, the style is marked as internal
	 */
	public Style(String name, int forcedOrder, boolean internal) {
		this.borderRef = new Border();
		this.cellXfRef = new CellXf();
		this.fillRef = new Fill();
		this.fontRef = new Font();
		this.numberFormatRef = new NumberFormat();
		this.name = name;
		this.setInternalID(forcedOrder);
		this.internalStyle = internal;
	}

	// ### M E T H O D S ###

	/**
	 * Appends the specified style parts to the current one. The parts can be
	 * instances of sub-classes like Border or CellXf or a Style instance. Only the
	 * altered properties of the specified style or style part that differs from a
	 * new / untouched style instance will be appended. This enables method chaining
	 *
	 * @param styleToAppend
	 *            The style to append or a sub-class of Style
	 * @return Current style with appended style parts
	 */
	public Style append(AbstractStyle styleToAppend) {
		if (styleToAppend == null) {
			return this;
		}
		if (styleToAppend instanceof Border) {
			this.getBorder().copyFields(styleToAppend, new Border());
		}
		else if (styleToAppend instanceof CellXf) {
			this.getCellXf().copyFields(styleToAppend, new CellXf());
		}
		else if (styleToAppend instanceof Fill) {
			this.getFill().copyFields(styleToAppend, new Fill());
		}
		else if (styleToAppend instanceof Font) {
			this.getFont().copyFields(styleToAppend, new Font());
		}
		else if (styleToAppend instanceof NumberFormat) {
			this.getNumberFormat().copyFields(styleToAppend, new NumberFormat());
		}
		else if (styleToAppend instanceof Style) {
			this.getBorder().copyFields(((Style) styleToAppend).getBorder(), new Border());
			this.getCellXf().copyFields(((Style) styleToAppend).getCellXf(), new CellXf());
			this.getFill().copyFields(((Style) styleToAppend).getFill(), new Fill());
			this.getFont().copyFields(((Style) styleToAppend).getFont(), new Font());
			this.getNumberFormat().copyFields(((Style) styleToAppend).getNumberFormat(), new NumberFormat());
		}
		return this;
	}

	/**
	 * Override toString method
	 *
	 * @return String of a class instance
	 */
	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder();
		sb.append("{\n\"Style\": {\n");
		addPropertyAsJson(sb, "Name", name);
		addPropertyAsJson(sb, "HashCode", this.hashCode());
		sb.append(borderRef.toString()).append(",\n");
		sb.append(cellXfRef.toString()).append(",\n");
		sb.append(fillRef.toString()).append(",\n");
		sb.append(fontRef.toString()).append(",\n");
		sb.append(numberFormatRef.toString()).append("\n}\n}");
		return sb.toString();
	}

	/**
	 * Override method to calculate the hash of this component
	 *
	 * @return Calculated hash as string
	 */
	@Override
	public int hashCode() {
		if (borderRef == null || cellXfRef == null || fillRef == null || fontRef == null || numberFormatRef == null) {
			throw new StyleException("The hash of the style could not be created because one or more components are missing as references");
		}
		int p = 241;
		int r = 1;
		r *= p + this.borderRef.hashCode();
		r *= p + this.cellXfRef.hashCode();
		r *= p + this.fillRef.hashCode();
		r *= p + this.fontRef.hashCode();
		r *= p + this.numberFormatRef.hashCode();
		return r;
	}

	/**
	 * Method to copy the current object to a new one
	 *
	 * @return Copy of the current object without the internal ID
	 */
	@Override
	public AbstractStyle copy() {
		if (borderRef == null || cellXfRef == null || fillRef == null || fontRef == null || numberFormatRef == null) {
			throw new StyleException("The style could not be copied because one or more components are missing as references");
		}
		Style copy = new Style();
		copy.setBorder(this.borderRef.copy());
		copy.setCellXf(this.cellXfRef.copy());
		copy.setFill(this.fillRef.copy());
		copy.setFont(this.fontRef.copy());
		copy.setNumberFormat(this.numberFormatRef.copy());
		return copy;
	}

	/**
	 * Copies the current object as cast style
	 *
	 * @return Copy of the current object without the internal ID
	 */
	public Style copyStyle() {
		return (Style) this.copy();
	}

}