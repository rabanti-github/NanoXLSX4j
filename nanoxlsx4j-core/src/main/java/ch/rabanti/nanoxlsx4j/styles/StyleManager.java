/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Column;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.annotations.InternalApi;

/**
 * Class representing a style manager to maintain all styles and its components of a workbook.<br> This class is only
 * internally used to compose the style environment right before saving an XLSX file.
 */
@InternalApi
class StyleManager {

    private final List<AbstractStyle> borders;
    private final List<AbstractStyle> cellXfs;
    private final List<AbstractStyle> fills;
    private final List<AbstractStyle> fonts;
    private final List<AbstractStyle> numberFormats;
    private final List<AbstractStyle> styles;

    /**
     * Default constructor
     */
    @InternalApi
    public StyleManager() {
        this.borders = new ArrayList<>();
        this.cellXfs = new ArrayList<>();
        this.fills = new ArrayList<>();
        this.fonts = new ArrayList<>();
        this.numberFormats = new ArrayList<>();
        this.styles = new ArrayList<>();
    }

    /**
     * Gets a component by its hash
     *
     * @param list List to check
     * @param hash Hash of the component
     * @return Determined component. If not found, null will be returned
     */
    private AbstractStyle getComponentByHash(List<AbstractStyle> list, int hash) {
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
    @InternalApi
    public Border[] getBorders() {
        return this.borders.toArray(new Border[this.borders.size()]);
    }

    /**
     * Gets the number of borders in the style manager
     *
     * @return Number of stored borders
     */
    @InternalApi
    public int getBorderStyleNumber() {
        return this.borders.size();
    }

    /**
     * Gets all fills of the style manager
     *
     * @return Array of fills
     */
    @InternalApi
    public Fill[] getFills() {
        return this.fills.toArray(new Fill[this.fills.size()]);
    }

    /**
     * Gets the number of fills in the style manager
     *
     * @return Number of stored fills
     */
    @InternalApi
    public int getFillStyleNumber() {
        return this.fills.size();
    }

    /**
     * Gets all fonts of the style manager
     *
     * @return Array of fonts
     */
    @InternalApi
    public Font[] getFonts() {
        return this.fonts.toArray(new Font[this.fonts.size()]);
    }

    /**
     * Gets the number of fonts in the style manager
     *
     * @return Number of stored fonts
     */
    @InternalApi
    public int getFontStyleNumber() {
        return this.fonts.size();
    }

    /**
     * Gets the number of number formats in the style manager
     *
     * @return Number of stored number formats
     */
    @InternalApi
    public int getNumberFormatStyleNumber() {
        return this.numberFormats.size();
    }

    /**
     * Gets all number formats of the style manager
     *
     * @return Array of number formats
     */
    @InternalApi
    public NumberFormat[] getNumberFormats() {
        return this.numberFormats.toArray(new NumberFormat[this.numberFormats.size()]);
    }

    /**
     * Gets all styles of the style manager
     *
     * @return Array of styles
     */
    @InternalApi
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
     * @param style Style to add
     * @return Added or determined style in the manager
     */
    @InternalApi
    public Style addStyle(Style style) {
        int hash = addStyleComponent(style);
        return (Style) this.getComponentByHash(this.styles, hash);
    }

    /**
     * Adds a style component to the manager with an ID
     *
     * @param style Component to add
     * @param id    ID of the component
     * @return Hash of the added or determined component
     */
    private int addStyleComponent(AbstractStyle style, int id) {
        style.setInternalId(Optional.of(id));
        return addStyleComponent(style);
    }

    /**
     * Adds a style component to the manager
     *
     * @param style Component to add
     * @return Hash of the added or determined component
     */
    private int addStyleComponent(AbstractStyle style) {
        int hash = style.hashCode();
        if (style.getClass() == Border.class) {
            if (this.getComponentByHash(this.borders, hash) == null) {
                this.borders.add(style);
            }
            reorganize(borders);
        } else if (style.getClass() == CellXf.class) {
            if (this.getComponentByHash(this.cellXfs, hash) == null) {
                this.cellXfs.add(style);
            }
            reorganize(cellXfs);
        } else if (style.getClass() == Fill.class) {
            if (this.getComponentByHash(this.fills, hash) == null) {
                this.fills.add(style);
            }
            reorganize(fills);
        } else if (style.getClass() == Font.class) {
            if (this.getComponentByHash(this.fonts, hash) == null) {
                this.fonts.add(style);
            }
            reorganize(fonts);
        } else if (style.getClass() == NumberFormat.class) {
            if (this.getComponentByHash(this.numberFormats, hash) == null) {
                this.numberFormats.add(style);
            }
            reorganize(numberFormats);
        } else if (style.getClass() == Style.class) {
            Style s = (Style) style;
            if (this.getComponentByHash(this.styles, hash) == null) {
                Optional<Integer> internalId = s.getInternalId();
                int id;
                if (internalId.isEmpty()) {
                    id = Integer.MAX_VALUE;
                    s.setInternalId(Optional.of(id));
                } else {
                    id = internalId.get();
                }
                int temp = this.addStyleComponent(s.getCurrentBorder(), id);
                s.setCurrentBorder((Border) this.getComponentByHash(this.borders, temp));
                temp = this.addStyleComponent(s.getCurrentCellXf(), id);
                s.setCurrentCellXf((CellXf) this.getComponentByHash(this.cellXfs, temp));
                temp = this.addStyleComponent(s.getCurrentFill(), id);
                s.setCurrentFill((Fill) this.getComponentByHash(this.fills, temp));
                temp = this.addStyleComponent(s.getCurrentFont(), id);
                s.setCurrentFont((Font) this.getComponentByHash(this.fonts, temp));
                temp = this.addStyleComponent(s.getCurrentNumberFormat(), id);
                s.setCurrentNumberFormat((NumberFormat) this.getComponentByHash(this.numberFormats, temp));
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
     * @param workbook Workbook to get all cells with possible style definitions
     * @return StyleManager object, to be processed by the save methods
     */
    public static StyleManager getManagedStyles(Workbook workbook) {
        StyleManager styleManager = new StyleManager();
        styleManager.addStyle(new Style("default", 0, true));
        Style borderStyle = new Style("default_border_style", 1, true);
        borderStyle.setCurrentBorder(BasicStyles.getDottedFill0125().getCurrentBorder());
        borderStyle.setCurrentFill(BasicStyles.getDottedFill0125().getCurrentFill());
        styleManager.addStyle(borderStyle);

        for (Worksheet worksheet : workbook.getWorksheets()) {
            for (Cell cell : worksheet.getCellValues()) {
                if (cell.getCellStyle() != null) {
                    Style resolvedStyle = styleManager.addStyle(cell.getCellStyle());
                    cell.setStyleInternal(resolvedStyle, true);
                }
            }
            for (Map.Entry<Integer, Column> column : worksheet.getColumns().entrySet()) {
                if (column.getValue().getDefaultColumnStyle() != null) {
                    Style resolvedStyle = styleManager.addStyle(column.getValue().getDefaultColumnStyle());
                    column.getValue().setDefaultColumnStyleInternal(resolvedStyle, true);
                }
            }
        }
        return styleManager;
    }

    /**
     * Method to reorganize / reorder a list of style components
     *
     * @param list List to reorganize
     */
    private void reorganize(List<AbstractStyle> list) {
        int len = list.size();
        Collections.sort(list);
        int id = 0;
        for (int i = 0; i < len; i++) {
            list.get(i).setInternalId(Optional.of(id));
            id++;
        }
    }
}
