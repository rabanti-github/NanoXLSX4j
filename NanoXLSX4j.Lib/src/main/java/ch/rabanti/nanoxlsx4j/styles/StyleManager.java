/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Map;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Workbook;

/**
 * Class representing a style manager to maintain all styles and its components
 * of a workbook.<br>
 * This class is only internally used to compose the style environment right
 * before saving an XLSX file
 *
 * @author Raphael Stockeli
 */
public class StyleManager {

	// ### P R I V A T E F I E L D S ###
	private final ArrayList<AbstractStyle> borders;
	private final ArrayList<AbstractStyle> cellXfs;
	private final ArrayList<AbstractStyle> fills;
	private final ArrayList<AbstractStyle> fonts;
	private final ArrayList<AbstractStyle> numberFormats;
	private final ArrayList<AbstractStyle> styles;

	// ### C O N S T R U C T O R S ###

	/**
	 * Default constructor
	 */
	public StyleManager() {
		this.borders = new ArrayList<>();
		this.cellXfs = new ArrayList<>();
		this.fills = new ArrayList<>();
		this.fonts = new ArrayList<>();
		this.numberFormats = new ArrayList<>();
		this.styles = new ArrayList<>();
	}

	// ### M E T H O D S ###

	/**
	 * Gets a component by its hash
	 *
	 * @param list
	 *            List to check
	 * @param hash
	 *            Hash of the component
	 * @return Determined component. If not found, null will be returned
	 */
	private AbstractStyle getComponentByHash(ArrayList<AbstractStyle> list, int hash) {
		int len = list.size();
		for (int i = 0; i < len; i++) {
			if (list.get(i).hashCode() == hash) {
				return list.get(i);
			}
		}
		return null;
	}

	/**
	 * Gets all borders of the style manager
	 *
	 * @return Array of borders
	 */
	public Border[] getBorders() {
		return this.borders.toArray(new Border[this.borders.size()]);
	}

	/**
	 * Gets the number of borders in the style manager
	 *
	 * @return Number of stored borders
	 */
	public int getBorderStyleNumber() {
		return this.borders.size();
	}

	/**
	 * Gets all fills of the style manager
	 *
	 * @return Array of fills
	 */
	public Fill[] getFills() {
		return this.fills.toArray(new Fill[this.fills.size()]);
	}

	/**
	 * Gets the number of fills in the style manager
	 *
	 * @return Number of stored fills
	 */
	public int getFillStyleNumber() {
		return this.fills.size();
	}

	/**
	 * Gets all fonts of the style manager
	 *
	 * @return Array of fonts
	 */
	public Font[] getFonts() {
		return this.fonts.toArray(new Font[this.fonts.size()]);
	}

	/**
	 * Gets the number of fonts in the style manager
	 *
	 * @return Number of stored fonts
	 */
	public int getFontStyleNumber() {
		return this.fonts.size();
	}

	/**
	 * Gets all number formats of the style manager
	 *
	 * @return Array of number formats
	 */
	public NumberFormat[] getNumberFormats() {
		return this.numberFormats.toArray(new NumberFormat[this.numberFormats.size()]);
	}

	/**
	 * Gets all styles of the style manager
	 *
	 * @return Array of styles
	 */
	public Style[] getStyles() {
		return this.styles.toArray(new Style[this.styles.size()]);

	}

	/**
	 * Gets the number of styles in the style manager
	 *
	 * @return Number of stored styles
	 */
	public int getStyleNumber() {
		return this.styles.size();
	}

	/**
	 * Adds a style component to the manager
	 *
	 * @param style
	 *            Style to add
	 * @return Added or determined style in the manager
	 */
	public Style addStyle(Style style) {
		int hash = addStyleComponent(style);
		return (Style) this.getComponentByHash(this.styles, hash);
	}

	/**
	 * Adds a style component to the manager with an ID
	 *
	 * @param style
	 *            Component to add
	 * @param id
	 *            Id of the component
	 * @return Hash of the added or determined component
	 */
	private int addStyleComponent(AbstractStyle style, Integer id) {
		style.setInternalID(id);
		return addStyleComponent(style);
	}

	/**
	 * Adds a style component to the manager
	 *
	 * @param style
	 *            Component to add
	 * @return Hash of the added or determined component
	 */
	private int addStyleComponent(AbstractStyle style) {
		int hash = style.hashCode();
		if (style instanceof Border) {
			if (this.getComponentByHash(this.borders, hash) == null) {
				this.borders.add(style);
			}
			reorganize(borders);
		}
		else if (style instanceof CellXf) {
			if (this.getComponentByHash(this.cellXfs, hash) == null) {
				this.cellXfs.add(style);
			}
			reorganize(cellXfs);
		}
		else if (style instanceof Fill) {
			if (this.getComponentByHash(this.fills, hash) == null) {
				this.fills.add(style);
			}
			reorganize(fills);
		}
		else if (style instanceof Font) {
			if (this.getComponentByHash(this.fonts, hash) == null) {
				this.fonts.add(style);
			}
			reorganize(fonts);
		}
		else if (style instanceof NumberFormat) {
			if (this.getComponentByHash(this.numberFormats, hash) == null) {
				this.numberFormats.add(style);
			}
			reorganize(numberFormats);
		}
		else if (style instanceof Style) {
			Style s = (Style) style;
			if (this.getComponentByHash(this.styles, hash) == null) {
				Integer id;
				if (s.getInternalID() == null) {
					id = Integer.MAX_VALUE;
					s.setInternalID(id);
				}
				else {
					id = s.getInternalID();
				}
				int temp = this.addStyleComponent(s.getBorder(), id);
				s.setBorder((Border) this.getComponentByHash(this.borders, temp));
				temp = this.addStyleComponent(s.getCellXf(), id);
				s.setCellXf((CellXf) this.getComponentByHash(this.cellXfs, temp));
				temp = this.addStyleComponent(s.getFill(), id);
				s.setFill((Fill) this.getComponentByHash(this.fills, temp));
				temp = this.addStyleComponent(s.getFont(), id);
				s.setFont((Font) this.getComponentByHash(this.fonts, temp));
				temp = this.addStyleComponent(s.getNumberFormat(), id);
				s.setNumberFormat((NumberFormat) this.getComponentByHash(this.numberFormats, temp));
				this.styles.add(s);
			}
			reorganize(styles);
			hash = s.hashCode();
		}
		return hash;
	}

	/**
	 * Method to gather all styles of the cells in all worksheets
	 *
	 * @param workbook
	 *            Workbook to get all cells with possible style definitions
	 * @return StyleManager object, to be processed by the save methods
	 */
	public static StyleManager getManagedStyles(Workbook workbook) {
		StyleManager styleManager = new StyleManager();
		styleManager.addStyle(new Style("default", 0, true));
		Style borderStyle = new Style("default_border_style", 1, true);
		borderStyle.setBorder(BasicStyles.DottedFill_0_125().getBorder());
		borderStyle.setFill(BasicStyles.DottedFill_0_125().getFill());
		styleManager.addStyle(borderStyle);

		for (int i = 0; i < workbook.getWorksheets().size(); i++) {
			for (Map.Entry<String, Cell> cell : workbook.getWorksheets().get(i).getCells().entrySet()) {
				if (cell.getValue().getCellStyle() != null) {
					Style resolvedStyle = styleManager.addStyle(cell.getValue().getCellStyle());
					workbook.getWorksheets().get(i).getCells().get(cell.getKey()).setStyle(resolvedStyle, true);
				}
			}
		}
		return styleManager;
	}

	/**
	 * Method to reorganize / reorder a list of style components
	 *
	 * @param list
	 *            List to reorganize
	 */
	private void reorganize(ArrayList<AbstractStyle> list) {
		int len = list.size();
		Collections.sort(list);
		int id = 0;
		for (int i = 0; i < len; i++) {
			list.get(i).setInternalID(id);
			id++;
		}
	}
}