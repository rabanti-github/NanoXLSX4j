/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * Class representing a worksheet of a workbook
 */
public class Worksheet {

    // ### C O N S T A N T S ###
    /**
     * Threshold, using when floats are compared
     */
    private static final float FLOAT_THRESHOLD = 0.0001f;
    /**
     * Maximum number of characters a worksheet name can have
     */
    public static final int MAX_WORKSHEET_NAME_LENGTH = 31;
    /**
     * Default column width as constant
     */
    public static final float DEFAULT_COLUMN_WIDTH = 10f;
    /**
     * Default row height as constant
     */
    public static final float DEFAULT_ROW_HEIGHT = 15f;
    /**
     * Maximum column number (zero-based)
     */
    public static final int MAX_COLUMN_NUMBER = 16383;
    /**
     * Maximum column width as constant
     */
    public static final float MAX_COLUMN_WIDTH = 255f;
    /**
     * Maximum row number (zero-based)
     */
    public static final int MAX_ROW_NUMBER = 1048575;
    /**
     * Maximum row height as constant
     */
    public static final float MAX_ROW_HEIGHT = 409.5f;
    /**
     * Minimum column number (zero-based)
     */
    public static final int MIN_COLUMN_NUMBER = 0;
    /**
     * Minimum column width as constant
     */
    public static final float MIN_COLUMN_WIDTH = 0f;
    /**
     * Minimum row number (zero-based)
     */
    public static final int MIN_ROW_NUMBER = 0;
    /**
     * Minimum row height as constant
     */
    public static final float MIN_ROW_HEIGHT = 0f;

    /**
     * Automatic zoom factor of a worksheet
     */
    public static final int AUTO_ZOOM_FACTOR = 0;
    /**
     * Minimum zoom factor of a worksheet. If set to this value, the zoom is set to automatic
     */
    public static final int MIN_ZOOM_FACTOR = 10;
    /**
     * Maximum zoom factor of a worksheet
     */
    public static final int MAX_ZOOM_FACTOR = 400;

    // ### E N U M S ###

    /**
     * Enum to define the direction when using AddNextCell method
     */
    public enum CellDirection {
        /**
         * The next cell will be on the same row (A1,B1,C1...)
         */
        ColumnToColumn,
        /**
         * The next cell will be on the same column (A1,A2,A3...)
         */
        RowToRow,
        /**
         * The address of the next cell will be not changed when adding a cell (for
         * manual definition of cell addresses)
         */
        Disabled
    }

    /**
     * Enum to define the possible protection types when protecting a worksheet
     */
    public enum SheetProtectionValue {
        // sheet, // Is always on 1 if protected
        /**
         * If selected, the user can edit objects if the worksheets is protected
         */
        objects,
        /**
         * If selected, the user can edit scenarios if the worksheets is protected
         */
        scenarios,
        /**
         * If selected, the user can format cells if the worksheets is protected
         */
        formatCells,
        /**
         * If selected, the user can format columns if the worksheets is protected
         */
        formatColumns,
        /**
         * If selected, the user can format rows if the worksheets is protected
         */
        formatRows,
        /**
         * If selected, the user can insert columns if the worksheets is protected
         */
        insertColumns,
        /**
         * If selected, the user can insert rows if the worksheets is protected
         */
        insertRows,
        /**
         * If selected, the user can insert hyperlinks if the worksheets is protected
         */
        insertHyperlinks,
        /**
         * If selected, the user can delete columns if the worksheets is protected
         */
        deleteColumns,
        /**
         * If selected, the user can delete rows if the worksheets is protected
         */
        deleteRows,
        /**
         * If selected, the user can select locked cells if the worksheets is protected
         */
        selectLockedCells,
        /**
         * If selected, the user can sort cells if the worksheets is protected
         */
        sort,
        /**
         * If selected, the user can use auto filters if the worksheets is protected
         */
        autoFilter,
        /**
         * If selected, the user can use pivot tables if the worksheets is protected
         */
        pivotTables,
        /**
         * If selected, the user can select unlocked cells if the worksheets is
         * protected
         */
        selectUnlockedCells
    }

    /**
     * Enum to define the pane position or active pane in a slip worksheet
     */
    public enum WorksheetPane {
        /**
         * The pane is located in the bottom right of the split worksheet
         */
        bottomRight,
        /**
         * The pane is located in the top right of the split worksheet
         */
        topRight,
        /**
         * The pane is located in the bottom left of the split worksheet
         */
        bottomLeft,
        /**
         * The pane is located in the top left of the split worksheet
         */
        topLeft
    }

    /**
     * Enum to define how a worksheet is displayed in the spreadsheet application (Excel)
     */
    public enum SheetViewType {
        /**
         * The worksheet is displayed without pagination (default)
         */
        normal,
        /**
         * The worksheet is displayed with indicators where the page would break if it were printed
         */
        pageBreakPreview,
        /**
         * The worksheet is displayed like it would be printed
         */
        pageLayout
    }

    // ### P R I V A T E F I E L D S ###
    private Style activeStyle;
    private Range autoFilterRange;
    private Map<String, Cell> cells;
    private Map<Integer, Column> columns;
    private CellDirection currentCellDirection;
    private int currentColumnNumber;
    private int currentRowNumber;
    private float defaultColumnWidth;
    private float defaultRowHeight;
    private Map<Integer, Boolean> hiddenRows;
    private Map<String, Range> mergedCells;
    private Map<Integer, Float> rowHeights;
    private List<Range> selectedCells;
    private int sheetID;
    private String sheetName;
    private String sheetProtectionPassword = null;
    private String sheetProtectionPasswordHash = null;
    private List<SheetProtectionValue> sheetProtectionValues;
    private boolean useSheetProtection;
    private boolean useActiveStyle;
    private WorksheetPane activePane = null;
    private boolean hidden;
    private Workbook workbookReference;
    private Float paneSplitTopHeight;
    private Float paneSplitLeftWidth;
    private Boolean freezeSplitPanes;
    private Address paneSplitTopLeftCell;
    private Address paneSplitAddress;
    private boolean showGridLines;
    private boolean showRowColumnHeaders;
    private boolean showRuler;
    private SheetViewType viewType;
    private Map<SheetViewType, Integer> zoomFactor;

    // ### G E T T E R S & S E T T E R S ###

    /**
     * Sets the column auto filter within the defined column range
     *
     * @param range Range to apply auto filter on. The range could be 'A1:C10' for
     *              instance. The end row will be recalculated automatically when
     *              saving the file
     * @throws RangeException  Thrown if the passed range out of range
     * @throws FormatException Thrown if the passed range is malformed
     */
    public void setAutoFilterRange(String range) {
        this.autoFilterRange = Cell.resolveCellRange(range);
        recalculateAutoFilter();
        recalculateColumns();
    }

    /**
     * Gets the range of the auto filter. If null, no auto filters are applied
     *
     * @return Range of auto filter
     */
    public Range getAutoFilterRange() {
        return autoFilterRange;
    }

    /**
     * Gets the cells of the worksheet as map with the cell address as key and the
     * cell object as value
     *
     * @return List of Cell objects
     */
    public Map<String, Cell> getCells() {
        return cells;
    }

    /**
     * Gets all columns with non-standard properties, like auto filter applied or a
     * special width as map with the zero-based column index as key and the column
     * object as value
     *
     * @return map of columns
     */
    public Map<Integer, Column> getColumns() {
        return columns;
    }

    /**
     * Gets the direction when using AddNextCell method
     *
     * @return Cell direction
     */
    public CellDirection getCurrentCellDirection() {
        return currentCellDirection;
    }

    /**
     * Sets the direction when using AddNextCell method
     *
     * @param currentCellDirection Cell direction
     */
    public void setCurrentCellDirection(CellDirection currentCellDirection) {
        this.currentCellDirection = currentCellDirection;
    }

    /**
     * Sets the current column number (zero based)
     *
     * @param columnNumber Column number (zero based)
     * @throws RangeException Thrown if the number is out of the valid range. Range is from 0
     *                        to 16383 (16384 columns)
     */
    public void setCurrentColumnNumber(int columnNumber) {
        Cell.validateColumnNumber(columnNumber);
        this.currentColumnNumber = columnNumber;
    }

    /**
     * Sets the current row number (zero based)
     *
     * @param rowNumber Row number (zero based)
     * @throws RangeException Thrown if the number is out of the valid range. Range is from 0
     *                        to 1048575 (1048576 rows)
     */
    public void setCurrentRowNumber(int rowNumber) {
        Cell.validateRowNumber(rowNumber);
        this.currentRowNumber = rowNumber;
    }

    /**
     * Gets the current column number (zero based)
     *
     * @return Column number (zero-based)
     */
    public int getCurrentColumnNumber() {
        return this.currentColumnNumber;
    }

    /**
     * Gets the current row number (zero based)
     *
     * @return Row number (zero-based)
     */
    public int getCurrentRowNumber() {
        return this.currentRowNumber;
    }

    /**
     * Gets the default column width
     *
     * @return Default column width
     */
    public float getDefaultColumnWidth() {
        return defaultColumnWidth;
    }

    /**
     * Sets the default column width
     *
     * @param defaultColumnWidth Default column width
     * @throws RangeException Throws a RangeException if the passed width is out of range (set)
     */
    public void setDefaultColumnWidth(float defaultColumnWidth) {
        if (defaultColumnWidth < MIN_COLUMN_WIDTH || defaultColumnWidth > MAX_COLUMN_WIDTH) {
            throw new RangeException("The passed default row height is out of range (" + MIN_COLUMN_WIDTH + " to " + MAX_COLUMN_WIDTH + ")");
        }
        this.defaultColumnWidth = defaultColumnWidth;
    }

    /**
     * Gets the default Row height
     *
     * @return Default Row height
     */
    public float getDefaultRowHeight() {
        return defaultRowHeight;
    }

    /**
     * Sets the default Row height
     *
     * @param defaultRowHeight Default Row height
     * @throws RangeException Throws a RangeException if the passed height is out of range
     *                        (set)
     */
    public void setDefaultRowHeight(float defaultRowHeight) {
        if (defaultRowHeight < MIN_ROW_HEIGHT || defaultRowHeight > MAX_ROW_HEIGHT) {
            throw new RangeException("The passed default row height is out of range (" + MIN_ROW_HEIGHT + " to " + MAX_ROW_HEIGHT + ")");
        }
        this.defaultRowHeight = defaultRowHeight;
    }

    /**
     * Gets the hidden rows as map with the zero-based row number as key and a
     * boolean as value. True indicates hidden, false visible. Entries with the
     * value false are not affecting the worksheet. These entries can be removed<br>
     *
     * @return Map with hidden rows
     */
    public Map<Integer, Boolean> getHiddenRows() {
        return hiddenRows;
    }

    /**
     * Gets defined row heights as map with the zero-based row number as key and the
     * height (float from 0 to 409.5) as value
     *
     * @return Map of row heights
     */
    public Map<Integer, Float> getRowHeights() {
        return rowHeights;
    }

    /**
     * Gets the merged cells (only references) as map with the cell address as key
     * and the range object as value
     *
     * @return Hashmap with merged cell references
     */
    public Map<String, Range> getMergedCells() {
        return mergedCells;
    }

    /**
     * Gets either null (if no cells are selected), or the first defined range of
     * selected cells
     *
     * @return First cell range of the selected cells
     * @deprecated This method is a deprecated subset of the function
     * SelectedCellRanges. SelectedCellRanges will get this function
     * name in a future version. Therefore, the type will change
     */
    public Range getSelectedCells() {
        if (selectedCells.isEmpty()) {
            return null;
        } else {
            return selectedCells.get(0);
        }
    }

    /**
     * Gets all ranges of selected cells of this worksheet. An empty list is
     * returned if no cells are selected
     *
     * @return All ranges of the selected cells
     */
    public List<Range> getSelectedCellRanges() {
        return selectedCells;
    }

    /**
     * Sets a single range of selected cells on this worksheet. All existing ranges
     * will be removed range
     *
     * @param range Range as string to set as single cell range for selected cells, or
     *              null to remove the selected cells
     * @throws RangeException  Thrown if the passed range out of range
     * @throws FormatException Thrown if the passed range is malformed
     * @deprecated This method is a deprecated subset of the function
     * AddSelectedCells. It will be removed in a future version
     */
    public void setSelectedCells(String range) {
        removeSelectedCells();
        if (range == null) {
            this.selectedCells.clear();
        } else {
            setSelectedCells(new Range(range));
        }
    }

    /**
     * Sets a single range of selected cells on this worksheet. All existing ranges
     * will be removed. Null will remove all selected cells
     *
     * @param range Range to set as single cell range for selected cells
     * @deprecated This method is a deprecated subset of the function
     * AddSelectedCells. It will be removed in a future version
     */
    public void setSelectedCells(Range range) {
        removeSelectedCells();
        addSelectedCells(range);
    }

    /**
     * Sets the selected cells on this worksheet. If both addresses are null, the
     * selected cell range is removed
     *
     * @param startAddress Start address of the range to set as single cell range for
     *                     selected cells
     * @param endAddress   End address of the range to set as single cell range for selected
     *                     cells
     * @throws RangeException Thrown if either the start address or end address is null
     * @deprecated This method is a deprecated subset of the function
     * AddSelectedCells. It will be removed in a future version
     */
    public void setSelectedCells(Address startAddress, Address endAddress) {
        if (startAddress == null && endAddress != null || startAddress != null && endAddress == null) {
            throw new RangeException("Either the start or end address is null (invalid range)");
        }
        if (startAddress == null) {
            removeSelectedCells();
        } else {
            setSelectedCells(new Range(startAddress, endAddress));
        }
    }

    /**
     * Adds a range to the selected cells on this worksheet
     *
     * @param range Cell range to be added as selected cells
     */
    public void addSelectedCells(Range range) {
        selectedCells.add(range);
    }

    /**
     * Adds a range to the selected cells on this worksheet
     *
     * @param startAddress Start address of the range to add
     * @param endAddress   End address of the range to add
     */
    public void addSelectedCells(Address startAddress, Address endAddress) {
        selectedCells.add(new Range(startAddress, endAddress));
    }

    /**
     * Adds a range to the selected cells on this worksheet. Null or empty as value
     * will be ignored
     *
     * @param range Cell range to add as selected cells
     */
    public void addSelectedCells(String range) {
        if (!Helper.isNullOrEmpty(range)) {
            selectedCells.add(Cell.resolveCellRange(range));
        }
    }

    /**
     * Gets the internal ID of the worksheet
     *
     * @return Worksheet ID
     */
    public int getSheetID() {
        return sheetID;
    }

    /**
     * Sets the internal ID of the worksheet
     *
     * @param sheetID Worksheet ID
     */
    public void setSheetID(int sheetID) {
        if (sheetID < 1) {
            throw new FormatException("The ID " + sheetID + " is invalid. Worksheet IDs must be >0");
        }
        this.sheetID = sheetID;
    }

    /**
     * Gets the name of the sheet
     *
     * @return Name of the sheet
     */
    public String getSheetName() {
        return sheetName;
    }

    /**
     * Gets whether the worksheet is protected
     *
     * @return If true, the worksheet is protected
     */
    public boolean isUseSheetProtection() {
        return useSheetProtection;
    }

    /**
     * Sets whether the worksheet is protected
     *
     * @param useSheetProtection If true, the worksheet is protected
     */
    public void setUseSheetProtection(boolean useSheetProtection) {
        this.useSheetProtection = useSheetProtection;
    }

    /**
     * Gets the password used for sheet protection
     *
     * @return Password (UTF-8)
     * @apiNote If a workbook with password protected worksheets is loaded, only the
     * {@link Worksheet#getSheetProtectionPasswordHash()} is loaded. The
     * password itself cannot be recovered. Use the
     * {@link Worksheet#getSheetProtectionPasswordHash()} getter to check
     * whether there is a password set
     */
    public String getSheetProtectionPassword() {
        return sheetProtectionPassword;
    }

    /**
     * Sets the encryption hash of the password, defined with
     * {@link Worksheet#setSheetProtectionPassword(String)}. This method is usually
     * used when reading a workbook.
     *
     * @param hash Hash, either loaded from a workbook or generated by
     *             {@link Helper#generatePasswordHash(String)}
     * @apiNote Do not use this method to set a password. Use
     * {@link Worksheet#setSheetProtectionPassword(String)} instead
     */
    public void setSheetProtectionPasswordHash(String hash) {
        this.sheetProtectionPasswordHash = hash;
    }

    /**
     * gets the encryption hash of the password, defined with
     * {@link Worksheet#setSheetProtectionPassword(String)}. The value will be null,
     * if no password is defined
     *
     * @return Encrypted password as String
     */
    public String getSheetProtectionPasswordHash() {
        return sheetProtectionPasswordHash;
    }

    /**
     * Sets or removes the password for worksheet protection. If set,
     * UseSheetProtection will be also set to true
     *
     * @param password Password (UTF-8) to protect the worksheet. If the password is null
     *                 or empty, no password will be used
     */
    public void setSheetProtectionPassword(String password) {
        if (Helper.isNullOrEmpty(password)) {
            this.sheetProtectionPassword = null;
            this.sheetProtectionPasswordHash = null;
            this.useSheetProtection = false;
        } else {
            this.sheetProtectionPassword = password;
            this.sheetProtectionPasswordHash = Helper.generatePasswordHash(password);
            this.useSheetProtection = true;
        }
    }

    /**
     * Gets the list of SheetProtectionValues. These values define the allowed
     * actions if the worksheet is protected
     *
     * @return List of SheetProtectionValues
     */
    public List<SheetProtectionValue> getSheetProtectionValues() {
        return sheetProtectionValues;
    }

    /**
     * Gets the Reference to the parent Workbook
     *
     * @return Workbook reference
     */
    public Workbook getWorkbookReference() {
        return workbookReference;
    }

    /**
     * Sets the Reference to the parent Workbook
     *
     * @param workbookReference Workbook reference
     */
    public void setWorkbookReference(Workbook workbookReference) {
        this.workbookReference = workbookReference;
        if (workbookReference != null) {
            workbookReference.validateWorksheets();
        }
    }

    /**
     * Gets whether the worksheet is hidden
     *
     * @return If true, the worksheet is not listed in the worksheet tabs of the
     * workbook
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Sets whether the worksheet is hidden<br>
     * If the worksheet is not part of a workbook, or the only one in the workbook,
     * an exception will be thrown If the worksheet is the selected one, and
     * attempted to set hidden, an exception will be thrown. Define another selected
     * worksheet prior to this call, in this case.
     *
     * @param hidden If true, the worksheet is not listed as tab in the workbook's
     *               worksheet selection
     */
    public void setHidden(boolean hidden) {
        this.hidden = hidden;
        if (hidden && workbookReference != null) {
            workbookReference.validateWorksheets();
        }
    }

    /**
     * Gets the height of the upper, horizontal split pane, measured from the top of
     * the window.<br>
     * The value is nullable. If null, no horizontal split of the worksheet is
     * applied.<br>
     * The value is only applicable to split the worksheet into panes, but not to
     * freeze them.<br>
     * See also {@link #getPaneSplitAddress()}
     *
     * @return Height of the top pane until the split line appears
     * @apiNote Note: This value will be modified to the Excel-internal
     * representation, calculated by
     * {@link Helper#getInternalPaneSplitHeight(float)}
     */
    public Float getPaneSplitTopHeight() {
        return paneSplitTopHeight;
    }

    /**
     * Gets the width of the left, vertical split pane, measured from the left of
     * the window.<br>
     * The value is nullable. If null, no vertical split of the worksheet is
     * applied<br>
     * The value is only applicable to split the worksheet into panes, but not to
     * freeze them.<br>
     * See also: {@link #getPaneSplitAddress()}
     *
     * @return Width form the left border until the split line appears
     * @apiNote Note: This value will be modified to the Excel-internal
     * representation, calculated by
     * {@link Helper#getInternalColumnWidth(float, float, float)}}
     */
    public Float getPaneSplitLeftWidth() {
        return paneSplitLeftWidth;
    }

    /**
     * Gets whether split panes are frozen.<br>
     * The value is nullable. If null, no freezing is applied. This property also
     * does not apply if {@link #getPaneSplitAddress()} is null
     *
     * @return True if panes are frozen
     */
    public Boolean getFreezeSplitPanes() {
        return freezeSplitPanes;
    }

    /**
     * Gets the Top Left cell address of the bottom right pane if applicable and
     * splitting is applied.<br>
     * The column is only relevant for vertical split, whereas the row component is
     * only relevant for a horizontal split.<br>
     * The value is nullable. If null, no splitting was defined.
     *
     * @return Address of the top Left cell address of the bottom right pane
     */
    public Address getPaneSplitTopLeftCell() {
        return paneSplitTopLeftCell;
    }

    /**
     * Gets the split address for frozen panes or if pane split was defined in
     * number of columns and / or rows.<br>
     * For vertical splits, only the column component is considered. For horizontal
     * splits, only the row component is considered.<br>
     * The value is nullable. If null, no frozen panes or split by columns / rows
     * are applied to the worksheet. However, splitting can still be applied, if the
     * value is defined in characters.<br>
     * See also: {@link #getPaneSplitLeftWidth()} and
     * {@link #getPaneSplitTopHeight()} for splitting in characters (without
     * freezing)
     *
     * @return Address where the panes splits the worksheet apart
     */
    public Address getPaneSplitAddress() {
        return paneSplitAddress;
    }

    /**
     * Gets the active Pane is splitting is applied.<br>
     * The value is nullable. If null, no splitting was defined
     *
     * @return Active pane if defined
     */
    public WorksheetPane getActivePane() {
        return activePane;
    }

    /**
     * Gets the active Style of the worksheet. If null, no style is defined as
     * active
     *
     * @return Active style of the worksheet
     */
    public Style getActiveStyle() {
        return activeStyle;
    }

    /**
     * Gets whether grid lines are visible on the current worksheet. Default is true
     *
     * @return True if grid lines are visible
     */
    public boolean isShowingGridLines() {
        return showGridLines;
    }

    /**
     * Sets whether grid lines are visible on the current worksheet. Default is true
     *
     * @param showGridLines True if grid lines are visible
     */
    public void setShowingGridLines(boolean showGridLines) {
        this.showGridLines = showGridLines;
    }

    /**
     * Gets whether the column and row headers are visible on the current worksheet. Default is true
     *
     * @return True if column and row headers are visible
     */
    public boolean isShowingRowColumnHeaders() {
        return showRowColumnHeaders;
    }

    /**
     * Sets whether the column and row headers are visible on the current worksheet. Default is true	 * @param showRowColumnHeaders
     *
     * @param showRowColumnHeaders True if column and row headers are visible
     */
    public void setShowingRowColumnHeaders(boolean showRowColumnHeaders) {
        this.showRowColumnHeaders = showRowColumnHeaders;
    }

    /**
     * Gets whether a ruler is displayed over the column headers. This value only applies if {@link Worksheet#setViewType(SheetViewType)} is set to {@link Worksheet.SheetViewType#pageLayout}. Default is true
     *
     * @return True if rules are visible
     */
    public boolean isShowingRuler() {
        return showRuler;
    }

    /**
     * Sets whether a ruler is displayed over the column headers. This value only applies if {@link Worksheet#setViewType(SheetViewType)} is set to {@link Worksheet.SheetViewType#pageLayout}. Default is true
     *
     * @param showRuler True if rulers are visible
     */
    public void setShowingRuler(boolean showRuler) {
        this.showRuler = showRuler;
    }

    /**
     * Gets how the current worksheet is displayed in the spreadsheet application (Excel)
     *
     * @return View type of the current worksheet
     */
    public SheetViewType getViewType() {
        return viewType;
    }

    /**
     * Sets how the current worksheet is displayed in the spreadsheet application (Excel)
     *
     * @param viewType View type of the current worksheet
     */
    public void setViewType(SheetViewType viewType) {
        this.viewType = viewType;
        setZoomFactor(viewType, 100);
    }

    /**
     * Gets the zoom factor of the {@link Worksheet#setViewType(SheetViewType)} of the current worksheet. If {@link Worksheet#AUTO_ZOOM_FACTOR}, the zoom factor is set to automatic
     *
     * @return Map of the zoom factors depending on the view type
     */
    public int getZoomFactor() {
        return zoomFactor.get(viewType);
    }

    /**
     * Sets the zoom factor of the {@link Worksheet#setViewType(SheetViewType)} of the current worksheet. If {@link Worksheet#AUTO_ZOOM_FACTOR}, the zoom factor is set to automatic
     *
     * @param zoomFactor Map of the zoom factors depending on the view type
     * @throws WorksheetException Thrown if the zoom factor is not {@link Worksheet#AUTO_ZOOM_FACTOR} or below {@link Worksheet#MIN_ZOOM_FACTOR} or above {@link Worksheet#MAX_ZOOM_FACTOR}
     * @apiNote It is possible to add further zoom factors for inactive view types, using the function {@link Worksheet#setZoomFactor(SheetViewType, int)}
     */
    public void setZoomFactor(int zoomFactor) {
        this.setZoomFactor(viewType, zoomFactor);
    }

    /**
     * Gets all defined zoom factors per {@link Worksheet.SheetViewType} of the current worksheet. Use {@link Worksheet#setZoomFactor(SheetViewType, int)} to define the values
     *
     * @return Map of defined zoom factors of the current worksheet
     */
    public Map<SheetViewType, Integer> getZoomFactors() {
        return  zoomFactor;
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Default constructor. A worksheet created with this constructor cannot be used
     * to assign styles to a cell. This will cause an exception unless a reference
     * to the workbook was set
     *
     * @throws StyleException Thrown if a style is added to the worksheet without w workbook
     *                        reference
     */
    public Worksheet() {
        init();
    }

    /**
     * Constructor with worksheet name
     *
     * @param name Name of the new worksheet
     * @throws FormatException Thrown if the name contains illegal characters or is too long
     * @apiNote Note that the worksheet name is not checked against other worksheets
     * with this operation. This is later performed when the worksheet is
     * added to the workbook
     */
    public Worksheet(String name) {
        init();
        setSheetName(name);
    }

    /**
     * Constructor with workbook reference, name and sheet ID
     *
     * @param name      Name of the worksheet
     * @param id        ID of the worksheet (for internal use)
     * @param reference Reference to the parent Workbook
     * @throws FormatException Thrown if the name contains illegal characters or is too long
     */
    public Worksheet(String name, int id, Workbook reference) {
        init();
        setSheetName(name);
        this.setSheetID(id);
        this.workbookReference = reference;
    }

    // ### M E T H O D S ###

    // ### M E T H O D S - A D D N E X T C E L L ###

    /**
     * Adds an object to the next cell position. If the type of the value does not
     * match with one of the supported data types, it will be cast to a String. A
     * prepared object of the type Cell will not be cast but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int,
     * double, float, long, Date, boolean. All other types will be cast into a
     * String using the default toString() method
     *
     * @param value Unspecified value to insert
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCell(Object value) {
        addNextCell(castValue(value, this.currentColumnNumber, this.currentRowNumber), true, null);
    }

    /**
     * Adds an object to the next cell position. If the type of the value does not
     * match with one of the supported data types, it will be cast to a String.A
     * prepared object of the type Cell will not be cast but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int,
     * double, float, long, Date, boolean. All other types will be cast into a
     * String using the default toString() method
     *
     * @param value Unspecified value to insert
     * @param style Style object to apply on this cell
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCell(Object value, Style style) {
        addNextCell(castValue(value, this.currentColumnNumber, this.currentRowNumber), true, style);
    }

    /**
     * Method to insert a generic cell to the next cell position. If the cell object
     * already has a style definition, and a style or active style is defined, the
     * cell style will be merged, otherwise just set
     *
     * @param cell        Cell object to insert
     * @param incremental If true, the address value (row or column) will be incremented,
     *                    otherwise not
     * @param style       If not null, the defined style will be applied to the cell,
     *                    otherwise no style or the default style will be applied
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    private void addNextCell(Cell cell, boolean incremental, Style style) {
        // date and time styles are already defined by the passed cell object
        if (style != null || (activeStyle != null && useActiveStyle)) {

            if (cell.getCellStyle() == null && useActiveStyle) {
                cell.setStyle(activeStyle);
            } else if (cell.getCellStyle() == null && style != null) {
                cell.setStyle(style);
            } else if (cell.getCellStyle() != null && useActiveStyle) {
                Style mixedStyle = (Style) cell.getCellStyle().copy();
                mixedStyle.append(activeStyle);
                cell.setStyle(mixedStyle);
            } else if (cell.getCellStyle() != null && style != null) {
                Style mixedStyle = (Style) cell.getCellStyle().copy();
                mixedStyle.append(style);
                cell.setStyle(mixedStyle);
            }
        }
        String address = cell.getCellAddress();
        this.cells.put(address, cell);
        if (incremental) {
            if (this.getCurrentCellDirection() == CellDirection.ColumnToColumn) {
                this.currentColumnNumber++;
            } else if (this.getCurrentCellDirection() == CellDirection.RowToRow) {
                this.currentRowNumber++;
            }
            // else = disabled
        } else {
            if (this.getCurrentCellDirection() == CellDirection.ColumnToColumn) {
                this.currentColumnNumber = cell.getColumnNumber() + 1;
                this.currentRowNumber = cell.getRowNumber();
            } else if (this.getCurrentCellDirection() == CellDirection.RowToRow) {
                this.currentColumnNumber = cell.getColumnNumber();
                this.currentRowNumber = cell.getRowNumber() + 1;
            }
            // else = disabled
        }
    }

    // ### M E T H O D S - A D D C E L L ###

    /**
     * Adds an object to the defined cell address. If the type of the value does not
     * match with one of the supported data types, it will be cast to a String. A
     * prepared object of the type Cell will not be cast but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int,
     * double, float, long, Date, boolean. All other types will be cast into a
     * String using the default toString() method
     *
     * @param value        Unspecified value to insert
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @throws StyleException Thrown if the active style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, int columnNumber, int rowNumber) {
        addNextCell(castValue(value, columnNumber, rowNumber), false, null);
    }

    /**
     * Adds an object to the defined cell address. If the type of the value does not
     * match with one of the supported data types, it will be cast to a String. A
     * prepared object of the type Cell will not be cast but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int,
     * double, float, long, Date, boolean. All other types will be cast into a
     * String using the default toString() method
     *
     * @param value        Unspecified value to insert
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @param style        Style to apply on the cell
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, int columnNumber, int rowNumber, Style style) {
        addNextCell(castValue(value, columnNumber, rowNumber), false, style);

    }

    /**
     * Adds an object to the defined cell address. If the type of the value does not
     * match with one of the supported data types, it will be cast to a String. A
     * prepared object of the type Cell will not be cast but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int,
     * double, float, long, Date, boolean. All other types will be cast into a
     * String using the default toString() method
     *
     * @param value   Unspecified value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException  Thrown if the default style was malformed
     * @throws RangeException  Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, String address) {
        Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row);
    }

    /**
     * Adds an object to the defined cell address. If the type of the value does not
     * match with one of the supported data types, it will be cast to a String. A
     * prepared object of the type Cell will not be cast but adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int,
     * double, float, long, Date, boolean. All other types will be cast into a
     * String using the default toString() method
     *
     * @param value   Unspecified value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @param style   Style to apply on the cell
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException  Thrown if the default style was malformed
     * @throws RangeException  Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, String address, Style style) {
        Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row, style);
    }

    // ### M E T H O D S - A D D C E L L F O R M U L A ###

    /**
     * Adds a cell formula as string to the defined cell address
     *
     * @param formula Formula to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException  Thrown if the default style was malformed
     * @throws RangeException  Thrown if the next cell is out of range (on row or column)
     */
    public void addCellFormula(String formula, String address) {
        Address adr = Cell.resolveCellCoordinate(address);
        Cell c = new Cell(formula, Cell.CellType.FORMULA, adr.Column, adr.Row);
        addNextCell(c, false, null);
    }

    /**
     * Adds a cell formula as string to the defined cell address
     *
     * @param formula Formula to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @param style   Style to apply on the cell
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException  Thrown if the default style was malformed
     * @throws RangeException  Thrown if the next cell is out of range (on row or column)
     */
    public void addCellFormula(String formula, String address, Style style) {
        Address adr = Cell.resolveCellCoordinate(address);
        Cell c = new Cell(formula, Cell.CellType.FORMULA, adr.Column, adr.Row);
        addNextCell(c, false, style);
    }

    /**
     * Adds a cell formula as string to the defined cell address
     *
     * @param formula      Formula to insert
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellFormula(String formula, int columnNumber, int rowNumber) {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, columnNumber, rowNumber);
        addNextCell(c, false, null);
    }

    /**
     * Adds a cell formula as string to the defined cell address
     *
     * @param formula      Formula to insert
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @param style        Style to apply on the cell
     * @throws StyleException Thrown if the passed style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellFormula(String formula, int columnNumber, int rowNumber, Style style) {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, columnNumber, rowNumber);
        addNextCell(c, false, style);
    }

    /**
     * Adds a formula as string to the next cell position
     *
     * @param formula Formula to insert
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCellFormula(String formula) {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true, null);
    }

    /**
     * Adds a formula as string to the next cell position
     *
     * @param formula Formula to insert
     * @param style   Style to apply on the cell
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCellFormula(String formula, Style style) {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true, style);
    }

    // ### M E T H O D S - A D D C E L L R A N G E ###

    /**
     * Adds a list of object values to a defined cell range. If the type of the a
     * particular value does not match with one of the supported data types, it will
     * be cast to a String. A prepared object of the type Cell will not be cast but
     * adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int,
     * double, float, long, Date, boolean. All other types will be cast into a
     * String using the default toString() method
     *
     * @param values       List of unspecified objects to insert
     * @param startAddress Start address
     * @param endAddress   End address
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellRange(List<Object> values, Address startAddress, Address endAddress) {
        addCellRangeInternal(values, startAddress, endAddress, null);
    }

    /**
     * Adds a list of object values to a defined cell range. If the type of the a
     * particular value does not match with one of the supported data types, it will
     * be cast to a String. A prepared object of the type Cell will not be cast but
     * adjusted<br>
     * Recognized are the following data types: Cell (prepared object), String, int,
     * double, float, long, Date, boolean. All other types will be cast into a
     * String using the default toString() method
     *
     * @param values       List of unspecified objects to insert
     * @param startAddress Start address
     * @param endAddress   End address
     * @param style        Style to apply on the all cells of the range
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellRange(List<Object> values, Address startAddress, Address endAddress, Style style) {
        addCellRangeInternal(values, startAddress, endAddress, style);
    }

    /**
     * Adds a list of object values to a defined cell range. If the type of the a
     * particular value does not match with one of the supported data types, it will
     * be cast to a String. A prepared object of the type Cell will not be cast but
     * adjusted<br>
     * The data types in the passed list can be mixed. Recognized are the following
     * data types: Cell (prepared object), String, int, double, float, long, Date,
     * boolean. All other types will be cast into a String using the default
     * toString() method
     *
     * @param values    List of unspecified objects to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException  Thrown if the default style was malformed
     * @throws RangeException  Thrown if the next cell is out of range (on row or column)
     */
    public void addCellRange(List<Object> values, String cellRange) {
        Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress, null);
    }

    /**
     * Adds a list of object values to a defined cell range. If the type of the a
     * particular value does not match with one of the supported data types, it will
     * be cast to a String. A prepared object of the type Cell will not be cast but
     * adjusted<br>
     * The data types in the passed list can be mixed. Recognized are the following
     * data types: Cell (prepared object), String, int, double, float, long, Date,
     * boolean. All other types will be cast into a String using the default
     * toString() method
     *
     * @param values    List of unspecified objects to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @param style     Style to apply on the all cells of the range
     * @throws FormatException Thrown if the passed address is malformed
     * @throws StyleException  Thrown if the default style was malformed
     * @throws RangeException  Thrown if the next cell is out of range (on row or column)
     */
    public void addCellRange(List<Object> values, String cellRange, Style style) {
        Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress, style);
    }

    /**
     * Internal function to add a generic list of value to the defined cell range
     *
     * @param <T>          Data type of the generic value list
     * @param values       List of values
     * @param startAddress Start address
     * @param endAddress   End address
     * @param style        Style to apply on the all cells of the range. If null, no style or
     *                     the default style will be applied
     * @throws StyleException Thrown if the default style was malformed
     * @throws RangeException Thrown if the next cell is out of range (on row or column)
     */
    private <T> void addCellRangeInternal(List<T> values, Address startAddress, Address endAddress, Style style) {
        List<Address> addresses = Cell.getCellRange(startAddress, endAddress);
        if (values.size() != addresses.size()) {
            throw new RangeException(
                    "The number of passed values (" + values.size() + ") differs from the number of cells within the range (" + addresses.size() + ")");
        }
        List<Cell> list = Cell.convertArray(values);
        int len = values.size();
        for (int i = 0; i < len; i++) {
            list.get(i).setRowNumber(addresses.get(i).Row);
            list.get(i).setColumnNumber(addresses.get(i).Column);
            addNextCell(list.get(i), false, style);
        }
    }

    // ### M E T H O D S - R E M O V E C E L L ###

    /**
     * Removes a previous inserted cell at the defined address
     *
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @return Returns true if the cell could be removed (existed), otherwise false
     * (did not exist)
     * @throws RangeException Thrown if the resolved cell address is out of range
     */
    public boolean removeCell(int columnNumber, int rowNumber) {
        String address = Cell.resolveCellAddress(columnNumber, rowNumber);
        if (this.cells.containsKey(address)) {
            this.cells.remove(address);
            return true;
        } else {
            return false;
        }
    }

    /**
     * Removes a previous inserted cell at the defined address
     *
     * @param address Cell address in the format A1 - XFD1048576
     * @return Returns true if the cell could be removed (existed), otherwise false
     * (did not exist)
     * @throws RangeException  Thrown if the resolved cell address is out of range
     * @throws FormatException Thrown if the passed address is malformed
     */
    public boolean removeCell(String address) {
        Address adr = Cell.resolveCellCoordinate(address);
        return removeCell(adr.Column, adr.Row);
    }

    // ### M E T H O D S - S E T S T Y L E ###

    /**
     * Sets the passed style on the passed cell range. If cells are already
     * existing, the style will be added or replaced. Otherwise, an empty cell will
     * be added with the assigned style. If the passed style is null, all styles
     * will be removed on existing cells and no additional (empty) cells are added
     * to the worksheet
     *
     * @param cellRange Cell range to apply the style
     * @param style     Style to apply or null to clear the range
     * @throws RangeException Throws a {@link RangeException} if the range is invalid
     * @implNote This method may invalidate an existing date and time value since
     * dates and times are defined by specific style. The result of a
     * redefinition will be a number, instead of a date or time
     */
    public void setStyle(Range cellRange, Style style) {
        if (cellRange == null) {
            throw new RangeException("No cell range was defined");
        }
        List<Address> addresses = cellRange.resolveEnclosedAddresses();
        for (Address address : addresses) {
            String key = address.getAddress();
            if (this.cells.containsKey(key)) {
                if (style == null) {
                    cells.get(key).removeStyle();
                } else {
                    cells.get(key).setStyle(style);
                }
            } else {
                if (style != null) {
                    addCell(null, address.Column, address.Row, style);
                }
            }
        }
    }

    /**
     * Sets the passed style on the passed cell range, derived from a start and end
     * address. If cells are already existing, the style will be added or replaced.
     * Otherwise, an empty cell will be added with the assigned style. If the passed
     * style is null, all styles will be removed on existing cells and no additional
     * (empty) cells are added to the worksheet
     *
     * @param startAddress Start address of the cell range
     * @param endAddress   End address of the cell range
     * @param style        Style to apply or null to clear the range
     * @implNote This method may invalidate an existing date and time value since
     * dates and times are defined by specific style. The result of a
     * redefinition will be a number, instead of a date or time
     */
    public void setStyle(Address startAddress, Address endAddress, Style style) {
        setStyle(new Range(startAddress, endAddress), style);
    }

    /**
     * Sets the passed style on the passed (singular) cell address. If the cell is
     * already existing, the style will be added or replaced. Otherwise, an empty
     * cell will be added with the assigned style. If the passed style is null, all
     * styles will be removed on existing cells and no additional (empty) cells are
     * added to the worksheet
     *
     * @param address Cell address to apply the style
     * @param style   Style to apply or null to clear the range
     * @implNote This method may invalidate an existing date and time value since
     * dates and times are defined by specific style. The result of a
     * redefinition will be a number, instead of a date or time
     */
    public void setStyle(Address address, Style style) {
        setStyle(address, address, style);
    }

    /**
     * Sets the passed style on the passed address expression. Such an expression
     * may be a single cell or a cell range. If the cell is already existing, the
     * style will be added or replaced. Otherwise, an empty cell will be added with
     * the assigned style. If the passed style is null, all styles will be removed
     * on existing cells and no additional (empty) cells are added to the worksheet
     *
     * @param addressExpression Expression of a cell address or range of addresses
     * @param style             Style to apply or null to clear the range
     * @implNote This method may invalidate an existing date and time value since
     * dates and times are defined by specific style. The result of a
     * redefinition will be a number, instead of a date or time
     */
    public void setStyle(String addressExpression, Style style) {
        Cell.AddressScope scope = Cell.getAddressScope(addressExpression);
        if (scope == Cell.AddressScope.SingleAddress) {
            Address address = new Address(addressExpression);
            setStyle(address, style);
        } else if (scope == Cell.AddressScope.Range) {
            Range range = new Range(addressExpression);
            setStyle(range, style);
        } else {
            throw new FormatException("The passed address'" + addressExpression + "' is neither a cell address, nor a range");
        }
    }

    // ### C O M M O N M E T H O D S ###

    /**
     * Method to add allowed actions if the worksheet is protected. If one or more
     * values are added, UseSheetProtection will be set to true
     *
     * @param typeOfProtection Allowed action on the worksheet or cells
     * @apiNote If {@link SheetProtectionValue#selectLockedCells} is added,
     * {@link SheetProtectionValue#selectUnlockedCells} is added
     * automatically
     */
    public void addAllowedActionOnSheetProtection(SheetProtectionValue typeOfProtection) {
        if (typeOfProtection == null) {
            return;
        }
        if (!this.sheetProtectionValues.contains(typeOfProtection)) {
            if (typeOfProtection == SheetProtectionValue.selectLockedCells && !this.sheetProtectionValues.contains(SheetProtectionValue.selectUnlockedCells)) {
                this.sheetProtectionValues.add(SheetProtectionValue.selectUnlockedCells);
            }
            this.sheetProtectionValues.add(typeOfProtection);
            this.setUseSheetProtection(true);
        }
    }

    /**
     * Sets the defined column as hidden
     *
     * @param columnNumber Column number to hide on the worksheet
     * @throws RangeException Thrown if the passed row number was out of range
     */
    public void addHiddenColumn(int columnNumber) {
        setColumnHiddenState(columnNumber, true);
    }

    /**
     * Sets the defined column as hidden
     *
     * @param columnAddress Column address to hide on the worksheet
     * @throws RangeException Thrown if the passed row number was out of range
     */
    public void addHiddenColumn(String columnAddress) {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnHiddenState(columnNumber, true);
    }

    /**
     * Sets the defined row as hidden
     *
     * @param rowNumber Row number to hide on the worksheet
     * @throws RangeException Thrown if the passed column number was out of range
     */
    public void addHiddenRow(int rowNumber) {
        setRowHiddenState(rowNumber, true);
    }

    /**
     * Method to cast a value or align an object of the type Cell to the context of
     * the worksheet
     *
     * @param value  Unspecified value or object of the type Cell
     * @param column Column index
     * @param row    Row index
     * @return Cell object
     */
    private Cell castValue(Object value, int column, int row) {
        Cell c;
        if (value instanceof Cell) {
            c = (Cell) value;
            c.setCellAddress2(new Address(column, row));
        } else {
            c = new Cell(value, Cell.CellType.DEFAULT, column, row);
        }
        return c;
    }

    /**
     * Clears the active style of the worksheet. All later added cells will contain
     * no style unless another active style is set
     */
    public void clearActiveStyle() {
        this.useActiveStyle = false;
        this.activeStyle = null;
    }

    /**
     * Gets the cell of the specified address
     *
     * @param address Address of the cell
     * @return Cell object
     * @throws WorksheetException Throws a WorksheetException if the cell was null or not found on
     *                            the cell table of this worksheet
     */
    public Cell getCell(Address address) {
        if (address == null) {
            throw new WorksheetException("No address to get was provided");
        }
        return getCell(address.getAddress());
    }

    /**
     * Gets the cell of the specified address as String
     *
     * @param address Address string of the cell
     * @return Cell object
     * @throws WorksheetException Throws a WorksheetException if the cell was not found on the cell
     *                            table of this worksheet
     */
    public Cell getCell(String address) {
        if (!this.cells.containsKey(address)) {
            throw new WorksheetException("The cell with the address " + address + " does not exist in this worksheet");
        }
        return this.cells.get(address);
    }

    /**
     * Gets the cell of the specified column and row number (zero-based)
     *
     * @param columnNumber Column number of the cell (zero-based)
     * @param rowNumber    Row number of the cell (zero-based)
     * @return Cell object
     * @throws WorksheetException Throws a WorksheetException if the cell was not found on the cell
     *                            table of this worksheet
     */
    public Cell getCell(int columnNumber, int rowNumber) {
        return getCell(new Address(columnNumber, rowNumber));
    }

    /**
     * Gets whether the specified address exists in the worksheet. Existing means
     * that a value was stored at the address
     *
     * @param address Address to check
     * @return True if the cell exists, otherwise false
     */
    public boolean hasCell(Address address) {
        return this.cells.containsKey(address.getAddress());
    }

    /**
     * Gets whether the specified address exists in the worksheet. Existing means
     * that a value was stored at the address
     *
     * @param columnNumber Column number of the cell to check (zero-based)
     * @param rowNumber    Row number of the cell to check (zero-based)
     * @return True if the cell exists, otherwise false
     * @throws RangeException A RangeException is thrown if the column or row number is invalid
     */
    public boolean hasCell(int columnNumber, int rowNumber) {
        return hasCell(new Address(columnNumber, rowNumber));
    }

    /**
     * Resets the defined column, if existing. The corresponding instance will be
     * removed from {@link WorksheetPane#getColumns()}
     *
     * @param columnNumber Column number to reset (zero-based)
     * @apiNote If the column is inside an autoFilter-Range, the column cannot be
     * entirely removed from {@link #getColumns()}. The hidden state will
     * be set to false and width to default, in this case.
     */
    public void resetColumn(int columnNumber) {
        if (columns.containsKey(columnNumber) && !columns.get(columnNumber).hasAutoFilter()) // AutoFilters cannot have gaps
        {
            columns.remove(columnNumber);
        } else if (columns.containsKey(columnNumber)) {
            columns.get(columnNumber).setHidden(false);
            columns.get(columnNumber).setWidth(DEFAULT_COLUMN_WIDTH);
        }
    }

    /**
     * Gets the first existing column number in the current worksheet (zero-based)
     *
     * @return Zero-based column number. In case of an empty worksheet, -1 will be
     * returned
     * @apiNote getFirstColumnNumber() will not return the first column with data in
     * any case. If there is a formatted but empty cell (or many) before
     * the first cell with data. getFirstColumnNumber() will return the
     * column number of this empty cell. Use
     * {@link Worksheet#getFirstDataColumnNumber()} in this case.
     */
    public int getFirstColumnNumber() {
        return getBoundaryNumber(false, true);
    }

    /**
     * Gets the first existing column number with data in the current worksheet
     * (zero-based)
     *
     * @return Zero-based column number. In case of an empty worksheet, -1 will be
     * returned
     * @apiNote getFirstDataColumnNumber() will ignore formatted but empty cells
     * before the first column with data. If you want the first defined
     * column, use {@link Worksheet#getFirstColumnNumber()} instead.
     */
    public int getFirstDataColumnNumber() {
        return getBoundaryDataNumber(false, true, true);
    }

    /**
     * Gets the first existing row number in the current worksheet (zero-based)
     *
     * @return Zero-based row number. In case of an empty worksheet, -1 will be
     * returned
     * @apiNote getLastColumnNumber() will not return the first column with data in
     * any case. If there is a formatted but empty cell (or many) before
     * the first cell with data. getFirstRowNumber() will return the column
     * number of this empty cell. Use
     * {@link Worksheet#getFirstDataRowNumber()} in this case.
     */
    public int getFirstRowNumber() {
        return getBoundaryNumber(true, true);
    }

    /**
     * Gets the first existing row number with data in the current worksheet
     * (zero-based)
     *
     * @return Zero-based row number. In case of an empty worksheet, -1 will be
     * returned
     * @apiNote getFirstDataRowNumber() will ignore formatted but empty cells before
     * the first row with data. If you want the first defined row, use
     * {@link Worksheet#getFirstRowNumber()} instead.
     */
    public int getFirstDataRowNumber() {
        return getBoundaryDataNumber(true, true, true);
    }

    /**
     * Gets the last existing column number in the current worksheet (zero-based)
     *
     * @return Zero-based column number. In case of an empty worksheet, -1 will be
     * returned
     * @apiNote getLastColumnNumber() will not return the last column with data in
     * any case. If there is a formatted but empty cell (or many) after the
     * last cell with data. getLastColumnNumber() will return the column
     * number of this empty cell. Use
     * {@link Worksheet#getLastDataColumnNumber()} in this case.
     */
    public int getLastColumnNumber() {
        return getBoundaryNumber(false, false);
    }

    /**
     * Gets the last existing column number with data in the current worksheet
     * (zero-based)
     *
     * @return Zero-based column number. in case of an empty worksheet, -1 will be
     * returned
     * @apiNote getLastDataColumnNumber() will ignore formatted but empty cells
     * after the last column with data. If you want the last defined
     * column, use {@link Worksheet#getLastColumnNumber()} instead.
     */
    public int getLastDataColumnNumber() {
        return getBoundaryDataNumber(false, false, true);
    }

    /**
     * Gets the last existing row number in the current worksheet (zero-based)
     *
     * @return Zero-based row number. In case of an empty worksheet, -1 will be
     * returned
     * @apiNote getLastRowNumber() will not return the last row with data in any
     * case. If there is a formatted but empty cell (or many) after the
     * last cell with data. getLastRowNumber() will return the row number
     * of this empty cell. Use {@link Worksheet#getLastDataRowNumber()} in
     * this case.
     */
    public int getLastRowNumber() {
        return getBoundaryNumber(true, false);
    }

    /**
     * Gets the last existing row number with data in the current worksheet
     * (zero-based)
     *
     * @return Zero-based row number. in case of an empty worksheet, -1 will be
     * returned
     * @apiNote getLastDataRowNumber() will ignore formatted but empty cells after
     * the last row with data. If you want the last defined row, use
     * {@link Worksheet#getLastRowNumber()} instead.
     */
    public int getLastDataRowNumber() {
        return getBoundaryDataNumber(true, false, true);
    }

    /**
     * Gets the last existing cell in the current worksheet (bottom right)
     *
     * @return Nullable Cell Address. If no cell address could be determined, null
     * will be returned
     * @apiNote getLastCellAddress() will not return the last cell with data in any
     * case. If there is a formatted (or with definitions of hidden states,
     * AutoFilters, heights or widths) but empty cell (or many) beyond the
     * last cell with data, getLastCellAddress() will return the address of
     * this empty cell. Use {@link Worksheet#getLastDataCellAddress()} in
     * this case.
     */
    public Address getLastCellAddress() {
        int lastRow = getLastRowNumber();
        int lastColumn = getLastColumnNumber();
        if (lastRow < 0 || lastColumn < 0) {
            return null;
        }
        return new Address(lastColumn, lastRow);
    }

    /**
     * Gets the last existing cell in the current worksheet (bottom right)
     *
     * @return Nullable Cell Address. If no cell address could be determined, null
     * will be returned
     * @apiNote getLastDataCellAddress() will ignore formatted (or with definitions
     * of hidden states, AutoFilters, heights or widths) but empty cells
     * beyond the last cell with data. If you want the last defined cell,
     * use {@link Worksheet#getLastCellAddress()} instead.
     */
    public Address getLastDataCellAddress() {
        int lastRow = getLastDataRowNumber();
        int lastColumn = getLastDataColumnNumber();
        if (lastRow < 0 || lastColumn < 0) {
            return null;
        }
        return new Address(lastColumn, lastRow);
    }

    /**
     * Gets the first existing cell in the current worksheet (bottom right)
     *
     * @return Nullable Cell Address. If no cell address could be determined, null
     * will be returned
     * @apiNote getFirstCellAddress() will not return the first cell with data in
     * any case. If there is a formated but empty cell (or many) before the
     * first cell with data, GetLastCellAddress() will return the address
     * of this empty cell. Use {@link Worksheet#getFirstDataCellAddress()}
     * in this case.
     */
    public Address getFirstCellAddress() {
        int firstRow = getFirstRowNumber();
        int firstColumn = getFirstColumnNumber();
        if (firstRow < 0 || firstColumn < 0) {
            return null;
        }
        return new Address(firstColumn, firstRow);
    }

    /**
     * Gets the first existing cell with data in the current worksheet (bottom
     * right)
     *
     * @return Nullable Cell Address. If no cell address could be determined, null
     * will be returned
     * @apiNote getFirstDataCellAddress() will ignore formatted but empty cells
     * before the first cell with data. If you want the first defined cell,
     * use {@link Worksheet#getFirstCellAddress()} instead.
     */
    public Address getFirstDataCellAddress() {
        int firstRow = getFirstDataRowNumber();
        int firstColumn = getLastDataColumnNumber();
        if (firstRow < 0 || firstColumn < 0) {
            return null;
        }
        return new Address(firstColumn, firstRow);
    }

    /**
     * Gets either the minimum or maximum row or column number, considering only
     * calls with data
     *
     * @param row         If true, the min or max row is returned, otherwise the column
     * @param min         If true, the min value of the row or column is defined, otherwise
     *                    the max value
     * @param ignoreEmpty If true, empty cell values are ignored, otherwise considered
     *                    without checking the content
     * @return Min or max number, or -1 if not defined
     */
    private int getBoundaryDataNumber(boolean row, boolean min, boolean ignoreEmpty) {
        if (cells.isEmpty()) {
            return -1;
        }
        if (!ignoreEmpty) {
            if (row && min) {
                return cells.values().stream().min(Comparator.comparingInt(Cell::getRowNumber)).get().getRowNumber();
            } else if (row) {
                return cells.values().stream().max(Comparator.comparingInt(Cell::getRowNumber)).get().getRowNumber();
            } else if (min) {
                return cells.values().stream().min(Comparator.comparingInt(Cell::getColumnNumber)).get().getColumnNumber();
            } else {
                return cells.values().stream().max(Comparator.comparingInt(Cell::getColumnNumber)).get().getColumnNumber();
            }
        }
        List<Cell> nonEmptyCells = cells.values().stream().filter(x -> x.getValue() != null).collect(Collectors.toList());
        if (nonEmptyCells.isEmpty()) {
            return -1;
        }
        if (row && min) {
            return nonEmptyCells.stream().filter(x -> x.getValue().toString() != "").min(Comparator.comparingInt(Cell::getRowNumber)).get().getRowNumber();
        } else if (row) {
            return nonEmptyCells.stream().filter(x -> x.getValue().toString() != "").max(Comparator.comparingInt(Cell::getRowNumber)).get().getRowNumber();
        } else if (min) {
            return nonEmptyCells.stream().filter(x -> x.getValue().toString() != "").max(Comparator.comparingInt(Cell::getColumnNumber)).get()
                    .getColumnNumber();
        } else {
            return nonEmptyCells.stream().filter(x -> x.getValue().toString() != "").min(Comparator.comparingInt(Cell::getColumnNumber)).get()
                    .getColumnNumber();
        }
    }

    /**
     * Gets either the minimum or maximum row or column number, considering all
     * available data
     *
     * @param row If true, the min or max row is returned, otherwise the column
     * @param min If true, the min value of the row or column is defined, otherwise
     *            the max value
     * @return Min or max number, or -1 if not defined
     */
    private int getBoundaryNumber(boolean row, boolean min) {
        int cellBoundary = getBoundaryDataNumber(row, min, false);
        if (row) {
            int heightBoundary = -1;
            if (!rowHeights.isEmpty()) {
                heightBoundary = min ? rowHeights.keySet().stream().min(Comparator.comparingInt(a -> a)).get() : rowHeights.keySet().stream()
                        .max(Comparator.comparingInt(a -> a)).get();
            }
            int hiddenBoundary = -1;
            if (!hiddenRows.isEmpty()) {
                hiddenBoundary = min ? hiddenRows.keySet().stream().min(Comparator.comparingInt(a -> a)).get() : hiddenRows.keySet().stream()
                        .max(Comparator.comparingInt(a -> a)).get();
            }
            return min ? getMinRow(cellBoundary, heightBoundary, hiddenBoundary) : getMaxRow(cellBoundary, heightBoundary, hiddenBoundary);
        } else {
            int columnDefBoundary = -1;
            if (!columns.isEmpty()) {
                columnDefBoundary = min ? columns.keySet().stream().min(Comparator.comparingInt(a -> a)).get() : columns.keySet().stream()
                        .max(Comparator.comparingInt(a -> a)).get();
            }
            if (min) {
                return cellBoundary >= 0 && cellBoundary < columnDefBoundary ? cellBoundary : columnDefBoundary;
            } else {
                return cellBoundary >= 0 && cellBoundary > columnDefBoundary ? cellBoundary : columnDefBoundary;
            }
        }
    }

    /**
     * Gets the maximum row coordinate either from cell data, height definitions or
     * hidden rows
     *
     * @param cellBoundary   Row number of max cell data
     * @param heightBoundary Row number of max defined row height
     * @param hiddenBoundary Row number of max defined hidden row
     * @return Max row number or -1 if nothing valid defined
     */
    private int getMaxRow(int cellBoundary, int heightBoundary, int hiddenBoundary) {
        int highest = -1;
        if (cellBoundary >= 0) {
            highest = cellBoundary;
        }
        if (heightBoundary >= 0 && heightBoundary > highest) {
            highest = heightBoundary;
        }
        if (hiddenBoundary >= 0 && hiddenBoundary > highest) {
            highest = hiddenBoundary;
        }
        return highest;
    }

    /**
     * Gets the minimum row coordinate either from cell data, height definitions or
     * hidden rows
     *
     * @param cellBoundary   Row number of min cell data
     * @param heightBoundary Row number of min defined row height
     * @param hiddenBoundary Row number of min defined hidden row
     * @return Min row number or -1 if nothing valid defined
     */
    private int getMinRow(int cellBoundary, int heightBoundary, int hiddenBoundary) {
        int lowest = Integer.MAX_VALUE;
        if (cellBoundary >= 0) {
            lowest = cellBoundary;
        }
        if (heightBoundary >= 0 && heightBoundary < lowest) {
            lowest = heightBoundary;
        }
        if (hiddenBoundary >= 0 && hiddenBoundary < lowest) {
            lowest = hiddenBoundary;
        }
        return lowest == Integer.MAX_VALUE ? -1 : lowest;
    }

    /**
     * Gets a row as list of cell objects
     *
     * @param rowNumber Row number (zero-based)
     * @return List of cell objects. If the row doesn't exist, an empty list is
     * returned
     */
    public List<Cell> getRow(int rowNumber) {
        List<Cell> list = new ArrayList<>();
        for (Map.Entry<String, Cell> cell : cells.entrySet()) {
            if (cell.getValue().getRowNumber() == rowNumber) {
                list.add(cell.getValue());
            }
        }
        list.sort(Comparator.comparingInt(Cell::getColumnNumber));
        return list;
    }

    /**
     * Gets a column as list of cell objects
     *
     * @param columnAddress Column address
     * @return List of cell objects. If the column doesn't exist, an empty list is
     * returned
     * @throws RangeException is thrown if the address is not valid
     */
    public List<Cell> getColumn(String columnAddress) {
        int column = Cell.resolveColumn(columnAddress);
        return getColumn(column);
    }

    /**
     * Gets a column as list of cell objects
     *
     * @param columnNumber Column number (zero-based)
     * @return List of cell objects. If the column doesn't exist, an empty list is
     * returned
     */
    public List<Cell> getColumn(int columnNumber) {
        List<Cell> list = new ArrayList<>();
        for (Map.Entry<String, Cell> cell : cells.entrySet()) {
            if (cell.getValue().getColumnNumber() == columnNumber) {
                list.add(cell.getValue());
            }
        }
        list.sort(Comparator.comparingInt(Cell::getRowNumber));
        return list;
    }

    /**
     * Set the current cell address
     *
     * @param columnAddress Column number (zero based)
     * @param width         Row number (zero based)
     * @throws RangeException Thrown if the address is out of the valid range. Range is from 0
     *                        to 16383 (16384 columns)
     */
    public void setColumnWidth(String columnAddress, float width) {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnWidth(columnNumber, width);
    }

    /**
     * Set the current cell address
     *
     * @param address Cell address in the format A1 - XFD1048576
     * @throws RangeException  Thrown if the address is out of the valid range. Range is for
     *                         rows from 0 to 1048575 (1048576 rows) and for columns from 0 to
     *                         16383 (16384 columns)
     * @throws FormatException Thrown if the passed address is malformed
     */
    public void setCurrentCellAddress(String address) {
        Address adr = Cell.resolveCellCoordinate(address);
        setCurrentCellAddress(adr.Column, adr.Row);
    }

    /**
     * Sets the name of the sheet
     *
     * @param sheetName Name of the sheet
     * @throws FormatException Thrown if the name contains illegal characters or is longer than
     *                         31 characters
     */
    public void setSheetName(String sheetName) {
        if (Helper.isNullOrEmpty(sheetName)) {
            throw new FormatException("The sheet name must be between 1 and " + MAX_WORKSHEET_NAME_LENGTH + " characters");
        }
        if (sheetName.length() > MAX_WORKSHEET_NAME_LENGTH) {
            throw new FormatException("The sheet name must be between 1 and " + MAX_WORKSHEET_NAME_LENGTH + " characters");
        }
        if (sheetName.matches(".*[\\[\\]\\*\\?/\\\\].*")) {
            throw new FormatException("The sheet name must must not contain the characters [  ]  * ? / \\ ");
        }
        this.sheetName = sheetName;
    }

    /**
     * Sets the name of the sheet
     *
     * @param sheetName Name of the sheet
     * @param sanitize  If true, the filename will be sanitized automatically according to
     *                  the specifications of Excel
     * @throws WorksheetException Thrown if no workbook is referenced. This information is
     *                            necessary to determine whether the name already exists
     */
    public void setSheetName(String sheetName, boolean sanitize) {
        if (sanitize) {
            this.sheetName = ""; // Empty name (temporary) to prevent conflicts during sanitizing
            this.sheetName = Worksheet.sanitizeWorksheetName(sheetName, this.workbookReference);
        } else {
            setSheetName(sheetName);
        }
    }

    /**
     * Sets the horizontal split of the worksheet into two panes. The measurement in
     * characters cannot be used to freeze panes
     *
     * @param topPaneHeight Height (similar to row height) from top of the worksheet to the
     *                      split line in characters
     * @param topLeftCell   Top Left cell address of the bottom right pane (if applicable).
     *                      Only the row component is important in a horizontal split
     * @param activePane    Active pane in the split window (can be null) (can be null)
     */
    public void setHorizontalSplit(float topPaneHeight, Address topLeftCell, WorksheetPane activePane) {
        setSplit(null, topPaneHeight, topLeftCell, activePane);
    }

    /**
     * Sets the horizontal split of the worksheet into two panes. The measurement in
     * rows can be used to split and freeze panes
     *
     * @param numberOfRowsFromTop Number of rows from top of the worksheet to the split line. The
     *                            particular row heights are considered
     * @param freeze              If true, all panes are frozen, otherwise remains movable
     * @param topLeftCell         Top Left cell address of the bottom right pane (if applicable).
     *                            Only the row component is important in a horizontal split
     * @param activePane          Active pane in the split window (can be null)
     * @throws WorksheetException WorksheetException Thrown if the row number of the top left cell
     *                            is smaller the split panes number of rows from top, if freeze is
     *                            applied
     */
    public void setHorizontalSplit(int numberOfRowsFromTop, boolean freeze, Address topLeftCell, WorksheetPane activePane) {
        setSplit(null, numberOfRowsFromTop, freeze, topLeftCell, activePane);
    }

    /**
     * Sets the vertical split of the worksheet into two panes. The measurement in
     * columns can be used to split and freeze panes
     *
     * @param numberOfColumnsFromLeft Number of columns from left of the worksheet to the split line.
     *                                The particular column widths are considered
     * @param freeze                  If true, all panes are frozen, otherwise remains movable
     * @param topLeftCell             Top Left cell address of the bottom right pane (if applicable).
     *                                Only the column component is important in a vertical split
     * @param activePane              Active pane in the split window (can be null)
     * @throws WorksheetException Thrown if the column number of the top left cell is smaller the
     *                            split panes number of columns from left, if freeze is applied
     */
    public void setVerticalSplit(int numberOfColumnsFromLeft, boolean freeze, Address topLeftCell, WorksheetPane activePane) {
        setSplit(numberOfColumnsFromLeft, null, freeze, topLeftCell, activePane);
    }

    /**
     * Sets the vertical split of the worksheet into two panes. The measurement in
     * characters cannot be used to freeze panes
     *
     * @param leftPaneWidth Width (similar to column width) from left of the worksheet to the
     *                      split line in characters
     * @param topLeftCell   Top Left cell address of the bottom right pane (if applicable).
     *                      Only the column component is important in a vertical split
     * @param activePane    Active pane in the split window (can be null)
     */
    public void setVerticalSplit(float leftPaneWidth, Address topLeftCell, WorksheetPane activePane) {
        setSplit(leftPaneWidth, null, topLeftCell, activePane);
    }

    /**
     * Sets the horizontal and vertical split of the worksheet into four panes. The
     * measurement in rows and columns can be used to split and freeze panes
     *
     * @param numberOfColumnsFromLeft Number of columns from left of the worksheet to the split line.
     *                                The particular column widths are considered.<br>
     *                                The parameter is nullable. If left null, the method acts identical
     *                                to
     *                                {@link #setHorizontalSplit(int, boolean, Address, WorksheetPane)}
     * @param numberOfRowsFromTop     Number of rows from top of the worksheet to the split line. The
     *                                particular row heights are considered.<br>
     *                                The parameter is nullable. If left null, the method acts identical
     *                                to {@link #setVerticalSplit(int, boolean, Address, WorksheetPane)}
     * @param freeze                  If true, all panes are frozen, otherwise remains movable
     * @param topLeftCell             Top Left cell address of the bottom right pane (if applicable)
     * @param activePane              Active pane in the split window (can be null)
     * @throws WorksheetException Thrown if the address of the top left cell is smaller the split
     *                            panes address, if freeze is applied
     */
    public void setSplit(Integer numberOfColumnsFromLeft, Integer numberOfRowsFromTop, boolean freeze, Address topLeftCell, WorksheetPane activePane) {
        if (freeze) {
            if (numberOfColumnsFromLeft != null && topLeftCell.Column < numberOfColumnsFromLeft) {
                throw new WorksheetException("The column number " +
                                                     topLeftCell.Column +
                                                     " is not valid for a frozen, vertical split with the split pane column number " +
                                                     numberOfColumnsFromLeft);
            }
            if (numberOfRowsFromTop != null && topLeftCell.Row < numberOfRowsFromTop) {
                throw new WorksheetException("The row number " +
                                                     topLeftCell.Row +
                                                     " is not valid for a frozen, horizontal split height the split pane row number " +
                                                     numberOfRowsFromTop);
            }
        }
        this.paneSplitLeftWidth = null;
        this.paneSplitTopHeight = null;
        this.freezeSplitPanes = freeze;
        int row = numberOfRowsFromTop != null ? numberOfRowsFromTop : 0;
        int column = numberOfColumnsFromLeft != null ? numberOfColumnsFromLeft : 0;
        this.paneSplitAddress = new Address(column, row);
        this.paneSplitTopLeftCell = topLeftCell;
        this.activePane = activePane;
    }

    /**
     * Sets the horizontal and vertical split of the worksheet into four panes. The
     * measurement in characters cannot be used to freeze panes
     *
     * @param leftPaneWidth Width (similar to column width) from left of the worksheet to the
     *                      split line in characters.<br>
     *                      The parameter is nullable. If left null, the method acts identical
     *                      to {@link #setHorizontalSplit(float, Address, WorksheetPane)}
     * @param topPaneHeight Height (similar to row height) from top of the worksheet to the
     *                      split line in characters.<br>
     *                      The parameter is nullable. If left null, the method acts identical
     *                      to {@link #setVerticalSplit(float, Address, WorksheetPane)}
     * @param topLeftCell   Top Left cell address of the bottom right pane (if applicable)
     * @param activePane    Active pane in the split window (can be null)
     */
    public void setSplit(Float leftPaneWidth, Float topPaneHeight, Address topLeftCell, WorksheetPane activePane) {
        this.paneSplitLeftWidth = leftPaneWidth;
        this.paneSplitTopHeight = topPaneHeight;
        this.freezeSplitPanes = null;
        this.paneSplitAddress = null;
        this.paneSplitTopLeftCell = topLeftCell;
        this.activePane = activePane;
    }

    /**
     * Resets splitting of the worksheet into panes, as well as their freezing
     */
    public void resetSplit() {
        this.paneSplitLeftWidth = null;
        this.paneSplitTopHeight = null;
        this.freezeSplitPanes = null;
        this.paneSplitAddress = null;
        this.paneSplitTopLeftCell = null;
        this.activePane = null;
    }

    /**
     * Moves the current position to the next column
     */
    public void goToNextColumn() {
        this.currentColumnNumber++;
        this.currentRowNumber = 0;
        Cell.validateColumnNumber(currentColumnNumber);
    }

    /**
     * Moves the current position to the next column with the number of cells to
     * move
     *
     * @param numberOfColumns Number of columns to move
     * @apiNote The value can also be negative. However, resulting column numbers
     * below 0 or above 16383 will cause an exception
     */
    public void goToNextColumn(int numberOfColumns) {
        currentColumnNumber += numberOfColumns;
        currentRowNumber = 0;
        Cell.validateColumnNumber(currentColumnNumber);
    }

    /**
     * Moves the current position to the next column with the number of cells to
     * move and the option to keep the row position
     *
     * @param numberOfColumns Number of columns to move
     * @param keepRowPosition If true, the row position is preserved, otherwise set to 0
     * @apiNote The value can also be negative. However, resulting column numbers
     * below 0 or above 16383 will cause an exception
     */
    public void goToNextColumn(int numberOfColumns, boolean keepRowPosition) {
        currentColumnNumber += numberOfColumns;
        if (!keepRowPosition) {
            currentRowNumber = 0;
        }
        Cell.validateColumnNumber(currentColumnNumber);
    }

    /**
     * Moves the current position to the next row (use for a new line)
     */
    public void goToNextRow() {
        this.currentRowNumber++;
        this.currentColumnNumber = 0;
        Cell.validateRowNumber(currentRowNumber);
    }

    /**
     * Moves the current position to the next row with the number of cells to move
     * (use for a new line)
     *
     * @param numberOfRows Number of rows to move
     * @apiNote The value can also be negative. However, resulting row numbers below
     * 0 or above 1048575 will cause an exception
     */
    public void goToNextRow(int numberOfRows) {
        currentRowNumber += numberOfRows;
        currentColumnNumber = 0;
        Cell.validateRowNumber(currentRowNumber);
    }

    /**
     * Moves the current position to the next row with the number of cells to move
     * and the option to keep the row position (use for a new line)
     *
     * @param numberOfRows       Number of rows to move
     * @param keepColumnPosition If true, the column position is preserved, otherwise set to 0
     * @apiNote The value can also be negative. However, resulting row numbers below
     * 0 or above 1048575 will cause an exception
     */
    public void goToNextRow(int numberOfRows, boolean keepColumnPosition) {
        currentRowNumber += numberOfRows;
        if (!keepColumnPosition) {
            currentColumnNumber = 0;
        }
        Cell.validateRowNumber(currentRowNumber);
    }

    /**
     * Init method for constructors
     */
    private void init() {
        this.currentCellDirection = CellDirection.ColumnToColumn;
        this.cells = new HashMap<>();
        this.currentRowNumber = 0;
        this.currentColumnNumber = 0;
        this.defaultColumnWidth = DEFAULT_COLUMN_WIDTH;
        this.defaultRowHeight = DEFAULT_ROW_HEIGHT;
        this.rowHeights = new HashMap<>();
        this.activeStyle = null;
        this.workbookReference = null;
        this.mergedCells = new HashMap<>();
        this.sheetProtectionValues = new ArrayList<>();
        this.hiddenRows = new HashMap<>();
        this.columns = new HashMap<>();
        this.selectedCells = new ArrayList<>();
        this.viewType = SheetViewType.normal;
        this.zoomFactor = new HashMap<>();
        this.zoomFactor.put(viewType, 100);
        this.showGridLines = true;
        this.showRowColumnHeaders = true;
        this.showRuler = true;
    }

    /**
     * Merges the defined cell range
     *
     * @param cellRange Range to merge
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     */
    public String mergeCells(Range cellRange) {
        return mergeCells(cellRange.StartAddress, cellRange.EndAddress);
    }

    /**
     * Merges the defined cell range
     *
     * @param cellRange Range to merge (e.g. 'A1:B12')
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     * @throws FormatException Thrown if the passed address is malformed
     */
    public String mergeCells(String cellRange) {
        Range range = Cell.resolveCellRange(cellRange);
        return mergeCells(range.StartAddress, range.EndAddress);
    }

    /**
     * Merges the defined cell range
     *
     * @param startAddress Start address of the merged cell range
     * @param endAddress   End address of the merged cell range
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     * @throws RangeException Thrown if one of the passed cell addresses is out of range or if
     *                        one or more cell addresses are already occupied in another merge
     *                        range
     */
    public String mergeCells(Address startAddress, Address endAddress) {
        String key = startAddress.toString() + ":" + endAddress.toString();
        Range value = new Range(startAddress, endAddress);
        List<Address> cells = value.resolveEnclosedAddresses();
        for (Map.Entry<String, Range> item : mergedCells.entrySet()) {
            if (item.getValue().resolveEnclosedAddresses().stream().anyMatch(cells::contains)) {
                throw new RangeException(
                        "The passed range: " + value + " contains cells that are already in the defined merge range: " + item.getKey());
            }
        }
        this.mergedCells.put(key, value);
        return key;
    }

    /**
     * Method to recalculate the auto filter (columns) of this worksheet. This is an
     * internal method. There is no need to use it. It must be public to require
     * access from the XlsXWriter class
     */
    public void recalculateAutoFilter() {
        if (this.autoFilterRange == null) {
            return;
        }
        int start = this.autoFilterRange.StartAddress.Column;
        int end = this.autoFilterRange.EndAddress.Column;
        int endRow = 0;
        for (Map.Entry<String, Cell> item : this.getCells().entrySet()) {
            if (item.getValue().getColumnNumber() < start || item.getValue().getColumnNumber() > end) {
                continue;
            }
            if (item.getValue().getRowNumber() > endRow) {
                endRow = item.getValue().getRowNumber();
            }
        }
        Column c;
        for (int i = start; i <= end; i++) {
            if (!this.columns.containsKey(i)) {
                c = new Column(i);
                c.setAutoFilter(true);
                this.columns.put(i, c);
            } else {
                this.getColumns().get(i).setAutoFilter(true);
            }
        }
        this.autoFilterRange = new Range(new Address(start, 0), new Address(end, endRow));
    }

    /**
     * Method to recalculate the collection of columns of this worksheet. This is an
     * internal method. There is no need to use it. It must be public to require
     * access from the XlsXWriter class
     */
    public void recalculateColumns() {
        ArrayList<Integer> columnsToDelete = new ArrayList<>();
        for (Map.Entry<Integer, Column> col : this.getColumns().entrySet()) {
            if (!col.getValue().hasAutoFilter() &&
                    !col.getValue().isHidden() &&
                    Math.abs(col.getValue().getWidth() - DEFAULT_COLUMN_WIDTH) <= FLOAT_THRESHOLD) {
                columnsToDelete.add(col.getKey());
            }
            if (!col.getValue().hasAutoFilter() &&
                    !col.getValue().isHidden() &&
                    Math.abs(col.getValue().getWidth() - DEFAULT_COLUMN_WIDTH) <= FLOAT_THRESHOLD) {
                columnsToDelete.add(col.getKey());
            }
        }
        for (Iterator<Integer> index = columnsToDelete.iterator(); index.hasNext(); ) {
            this.columns.remove(index.next());
        }
    }

    /**
     * Method to resolve all merged cells of the worksheet. Only the value of the
     * very first cell of the merged cells range will be visible. The other values
     * are still present (set to EMPTY) but will not be stored in the worksheet.<br>
     * This is an internal method. There is no need to use it.
     *
     * @throws StyleException Thrown if an unreferenced style was in the style sheet
     * @throws RangeException Thrown if the cell range was not found
     */
    public void resolveMergedCells() {
        Style mergeStyle = BasicStyles.MergeCellStyle();
        int pos;
        for (Map.Entry<String, Range> range : getMergedCells().entrySet()) {
            pos = 0;
            List<Address> addresses = Cell.getCellRange(range.getValue().StartAddress, range.getValue().EndAddress);
            for (Address address : addresses) {
                Cell cell = null;
                if (!cells.containsKey(address.getAddress())) {
                    cell = new Cell();
                    cell.setDataType(Cell.CellType.EMPTY);
                    cell.setRowNumber(address.Row);
                    cell.setColumnNumber(address.Column);
                    addCell(cell, cell.getColumnNumber(), cell.getRowNumber());
                } else {
                    cell = getCells().get(address.getAddress());
                }
                if (pos != 0) {
                    cell.setDataType(Cell.CellType.EMPTY);
                    if (cell.getCellStyle() == null) {
                        cell.setStyle(mergeStyle);
                    } else {
                        Style mixedMergeStyle = cell.getCellStyle();
                        // TODO: There should be a better possibility to identify particular style
                        // elements that deviates
                        mixedMergeStyle.getCellXf().setForceApplyAlignment(mergeStyle.getCellXf().isForceApplyAlignment());
                        cell.setStyle(mixedMergeStyle);
                    }
                }
                pos++;
            }
        }

    }

    /**
     * Removes auto filters from the worksheet
     */
    public void removeAutoFilter() {
        this.autoFilterRange = null;
    }

    /**
     * Sets a previously defined, hidden column as visible again
     *
     * @param columnNumber Column number to make visible again
     * @throws RangeException Thrown if the passed row number was out of range
     */
    public void removeHiddenColumn(int columnNumber) {
        setColumnHiddenState(columnNumber, false);
    }

    /**
     * Sets a previously defined, hidden column as visible again
     *
     * @param columnAddress Column address to make visible again
     * @throws RangeException Thrown if the passed row number was out of range
     */
    public void removeHiddenColumn(String columnAddress) {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnHiddenState(columnNumber, false);
    }

    /**
     * Sets a previously defined, hidden row as visible again
     *
     * @param rowNumber Row number to hide on the worksheet
     * @throws RangeException Thrown if the passed column number was out of range
     */
    public void removeHiddenRow(int rowNumber) {
        setRowHiddenState(rowNumber, false);
    }

    /**
     * Removes the defined merged cell range
     *
     * @param range Cell range to remove the merging
     * @throws RangeException  Thrown if the passed cell range was not merged earlier
     * @throws FormatException Thrown if the passed address is malformed
     */
    public void removeMergedCells(String range) {
        if (range != null) {
            range = range.toUpperCase();
        }
        if (range == null || !this.mergedCells.containsKey(range)) {
            throw new RangeException("The cell range " + range + " was not found in the list of merged cell ranges");
        } else {
            List<Address> addresses = Cell.getCellRange(range);
            Cell cell;
            for (int i = 0; i < addresses.size(); i++) {
                if (this.cells.containsKey(addresses.get(i).toString())) {
                    cell = this.cells.get(addresses.get(i).toString());
                    if (BasicStyles.MergeCellStyle().equals(cell.getCellStyle())) {
                        cell.removeStyle();
                    }
                    cell.resolveCellType(); // resets the type
                }
            }
            this.mergedCells.remove(range);
        }
    }

    /**
     * Removes the cell selection of this worksheet
     */
    public void removeSelectedCells() {
        this.selectedCells.clear();
    }

    /**
     * Removes the defined, non-standard row height
     *
     * @param rowNumber Row number (zero-based)
     */
    public void removeRowHeight(int rowNumber) {
        rowHeights.remove(rowNumber);
    }

    /**
     * Removes an allowed action on the current worksheet or its cells
     *
     * @param value Allowed action on the worksheet or cells
     */
    public void removeAllowedActionOnSheetProtection(SheetProtectionValue value) {
        sheetProtectionValues.remove(value);
    }

    /**
     * Sets the active style of the worksheet. This style will be assigned to all
     * later added cells
     *
     * @param style Style to set as active style
     * @throws StyleException Thrown if the worksheet has no workbook referenced when trying to
     *                        set the active style
     */
    public void setActiveStyle(Style style) {
        this.useActiveStyle = style != null;
        this.activeStyle = style;
    }

    /**
     * Sets the column auto filter within the defined column range
     *
     * @param startColumn Column number with the first appearance of an auto filter drop
     *                    down
     * @param endColumn   Column number with the last appearance of an auto filter drop down
     * @throws RangeException Thrown if one of the passed column numbers are out of range
     */
    public void setAutoFilter(int startColumn, int endColumn) {
        String start = Cell.resolveCellAddress(startColumn, 0);
        String end = Cell.resolveCellAddress(endColumn, 0);
        if (endColumn < startColumn) {
            setAutoFilterRange(end + ":" + start);
        } else {
            setAutoFilterRange(start + ":" + end);
        }
    }

    /**
     * Sets the column auto filter within the defined column range
     *
     * @param range Range to apply auto filter on. The range could be 'A1:C10' for
     *              instance. The end row will be recalculated automatically when
     *              saving the file
     * @throws RangeException  Throws a RangeException if the passed range out of range
     * @throws FormatException Throws an FormatException if the passed range is malformed
     */
    public void setAutoFilter(String range) {
        autoFilterRange = Cell.resolveCellRange(range);
        recalculateAutoFilter();
        recalculateColumns();
    }

    /**
     * Sets the defined column as hidden or visible
     *
     * @param columnNumber Column number to hide on the worksheet
     * @param state        If true, the column will be hidden, otherwise be visible
     * @throws RangeException Thrown if the passed row number was out of range
     */
    private void setColumnHiddenState(int columnNumber, boolean state) {
        Cell.validateColumnNumber(columnNumber);
        if (columns.containsKey(columnNumber)) {
            columns.get(columnNumber).setHidden(state);
        } else if (state) {
            Column c = new Column(columnNumber);
            c.setHidden(true);
            columns.put(columnNumber, c);
        }
        if (!columns.get(columnNumber).isHidden() &&
                Math.abs(columns.get(columnNumber).getWidth() - DEFAULT_COLUMN_WIDTH) <= FLOAT_THRESHOLD &&
                !columns.get(columnNumber).hasAutoFilter()) {
            columns.remove(columnNumber);
        }
    }

    /**
     * Sets the width of the passed column number (zero-based)
     *
     * @param columnNumber Column number (zero-based, from 0 to 16383)
     * @param width        Width from 0 to 255.0
     * @throws RangeException Thrown if the address is out of the valid range. Range is from 0
     *                        to 16383 (16384 columns)
     */
    public void setColumnWidth(int columnNumber, float width) {
        Cell.validateColumnNumber(columnNumber);
        if (width < MIN_COLUMN_WIDTH || width > MAX_COLUMN_WIDTH) {
            throw new RangeException(
                    "The column width (" + width + ") is out of range. Range is from " + MIN_COLUMN_WIDTH + " to " + MAX_COLUMN_WIDTH + " (chars).");
        }
        if (this.columns.containsKey(columnNumber)) {
            this.columns.get(columnNumber).setWidth(width);
        } else {
            Column c = new Column(columnNumber);
            c.setWidth(width);
            this.columns.put(columnNumber, c);
        }
    }

    /**
     * Set the current cell address
     *
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @throws RangeException Thrown if the address is out of the valid range. Range is for
     *                        rows from 0 to 1048575 (1048576 rows) and for columns from 0 to
     *                        16383 (16384 columns)
     */
    public void setCurrentCellAddress(int columnNumber, int rowNumber) {
        setCurrentColumnNumber(columnNumber);
        setCurrentRowNumber(rowNumber);
    }

    /**
     * Sets the height of the passed row number (zero-based)
     *
     * @param rowNumber Row number (zero-based, 0 to 1048575)
     * @param height    Height from 0 to 409.5
     * @throws RangeException Thrown if the address is out of the valid range. Range is from 0
     *                        to 1048575 (1048576 rows)
     */
    public void setRowHeight(int rowNumber, float height) {
        Cell.validateRowNumber(rowNumber);
        if (height < 0 || height > 409.5) {
            throw new RangeException("The row height (" + height + ") is out of range. Range is from 0 to 409.5 (equals 546px).");
        }
        this.rowHeights.put(rowNumber, height);
    }

    /**
     * Sets the defined row as hidden or visible
     *
     * @param rowNumber Row number to hide on the worksheet
     * @param state     If true, the row will be hidden, otherwise visible
     * @throws RangeException Thrown if the passed row number was out of range
     */
    private void setRowHiddenState(int rowNumber, boolean state) {
        Cell.validateRowNumber(rowNumber);
        if (this.hiddenRows.containsKey(rowNumber)) {
            if (state) {
                this.hiddenRows.put(rowNumber, state);
            } else {
                this.hiddenRows.remove(rowNumber);
            }
        } else if (state) {
            this.hiddenRows.put(rowNumber, state);
        }
    }

    /**
     * Creates a (dereferenced) deep copy of this worksheet
     *
     * @return Copy of this worksheet
     * @apiNote Not considered in the copy are the internal ID, the worksheet name
     * and the workbook reference. Since styles are managed in a shared
     * repository, no dereferencing is applied (Styles are not
     * deep-copied).<br> Use
     * {@link Workbook#copyWorksheetTo(Worksheet, String, Workbook)}} or
     * {@link Workbook#copyWorksheetIntoThis(Worksheet, String)} to add a
     * copy of worksheet to a workbook. These methods will set the internal
     * ID, name and workbook reference.
     */
    public Worksheet copy() {
        Worksheet copy = new Worksheet();
        for (Map.Entry<String, Cell> cell : this.cells.entrySet()) {
            copy.addCell(cell.getValue().copy(), cell.getKey());
        }
        copy.activePane = this.activePane;
        copy.activeStyle = this.activeStyle;
        if (this.autoFilterRange != null) {
            copy.autoFilterRange = this.autoFilterRange.copy();
        }
        for (Map.Entry<Integer, Column> column : this.columns.entrySet()) {
            copy.columns.put(column.getKey(), column.getValue().copy());
        }
        copy.currentCellDirection = this.currentCellDirection;
        copy.currentColumnNumber = this.currentColumnNumber;
        copy.currentRowNumber = this.currentRowNumber;
        copy.defaultColumnWidth = this.defaultColumnWidth;
        copy.defaultRowHeight = this.defaultRowHeight;
        copy.freezeSplitPanes = this.freezeSplitPanes;
        copy.hidden = this.hidden;
        for (Map.Entry<Integer, Boolean> row : this.hiddenRows.entrySet()) {
            copy.hiddenRows.put(row.getKey(), row.getValue());
        }
        for (Map.Entry<String, Range> cell : this.mergedCells.entrySet()) {
            copy.mergedCells.put(cell.getKey(), cell.getValue().copy());
        }
        if (this.paneSplitAddress != null) {
            copy.paneSplitAddress = this.paneSplitAddress.copy();
        }
        copy.paneSplitLeftWidth = this.paneSplitLeftWidth;
        copy.paneSplitTopHeight = this.paneSplitTopHeight;
        if (this.paneSplitTopLeftCell != null) {
            copy.paneSplitTopLeftCell = this.paneSplitTopLeftCell.copy();
        }
        for (Map.Entry<Integer, Float> row : this.rowHeights.entrySet()) {
            copy.rowHeights.put(row.getKey(), row.getValue());
        }
        if (this.selectedCells.size() > 0) {
            for (Range selectedCellRange : this.selectedCells) {
                copy.addSelectedCells(selectedCellRange);
            }
        }
        copy.sheetProtectionPassword = this.sheetProtectionPassword;
        copy.sheetProtectionPasswordHash = this.sheetProtectionPasswordHash;
        copy.sheetProtectionValues.addAll(this.sheetProtectionValues);
        copy.useActiveStyle = this.useActiveStyle;
        copy.useSheetProtection = this.useSheetProtection;
        copy.showGridLines = this.showGridLines;
        copy.showRowColumnHeaders = this.showRowColumnHeaders;
        copy.showRuler = this.showRuler;
        copy.viewType = this.viewType;
        copy.zoomFactor.clear();
        for (Map.Entry<SheetViewType, Integer> zoomFactor : this.zoomFactor.entrySet()) {
            copy.setZoomFactor(zoomFactor.getKey(), zoomFactor.getValue());
        }
        return copy;
    }

    /**
     * Sets a zoom factor for a given {@link Worksheet.SheetViewType}. If {@link Worksheet#AUTO_ZOOM_FACTOR}, the zoom factor is set to automatic
     *
     * @param sheetViewType Sheet view type to apply the zoom factor on
     * @param zoomFactor    Zoom factor in percent
     * @throws WorksheetException Thrown if the zoom factor is not {@link Worksheet#AUTO_ZOOM_FACTOR} or below {@link Worksheet#MIN_ZOOM_FACTOR} or above {@link Worksheet#MAX_ZOOM_FACTOR}
     * @apiNote This factor is not the currently set factor. use the setter {@link Worksheet#setZoomFactor(SheetViewType, int)} to set the factor for the current {@link Worksheet.SheetViewType}
     */
    public void setZoomFactor(SheetViewType sheetViewType, int zoomFactor) {
        if (zoomFactor != AUTO_ZOOM_FACTOR && (zoomFactor < MIN_ZOOM_FACTOR || zoomFactor > MAX_ZOOM_FACTOR)) {
            throw new WorksheetException("The zoom factor " + zoomFactor + " is not valid. Valid are values between " + MIN_ZOOM_FACTOR + " and " + MAX_ZOOM_FACTOR + ", or " + AUTO_ZOOM_FACTOR + " (automatic)");
        }
        this.zoomFactor.put(sheetViewType, zoomFactor);
    }

    // ### S T A T I C M E T H O D S

    /**
     * Sanitizes a worksheet name
     *
     * @param input    Name to sanitize
     * @param workbook Workbook reference
     * @return Name of the sanitized worksheet
     * @throws WorksheetException thrown if the workbook reference is null, since all worksheets
     *                            have to be considered during sanitation
     */
    public static String sanitizeWorksheetName(String input, Workbook workbook) {
        if (input == null || input.isEmpty()) {
            input = "Sheet1";
        }
        int len;
        if (input.length() > MAX_WORKSHEET_NAME_LENGTH) {
            len = MAX_WORKSHEET_NAME_LENGTH;
        } else {
            len = input.length();
        }
        StringBuilder sb = new StringBuilder(MAX_WORKSHEET_NAME_LENGTH);
        char c;
        for (int i = 0; i < len; i++) {
            c = input.charAt(i);
            if (c == '[' || c == ']' || c == '*' || c == '?' || c == '\\' || c == '/') {
                sb.append('_');
            } else {
                sb.append(c);
            }
        }
        return getUnusedWorksheetName(sb.toString(), workbook);
    }

    /**
     * Determines the next unused worksheet name in the passed workbook
     *
     * @param name     Original name to start the check
     * @param workbook Workbook to look for existing worksheets
     * @return Not yet used worksheet name
     * @throws WorksheetException thrown if the workbook reference is null, since all worksheets
     *                            have to be considered during sanitation
     * @implNote The 'rare' case where 10^31 Worksheets exists (leads to a crash) is
     * deliberately not handled, since such a number of sheets would
     * consume at least one quintillion bytes of RAM... what is vastly out
     * of the 64 bit range
     */
    private static String getUnusedWorksheetName(String name, Workbook workbook) {
        if (workbook == null) {
            throw new WorksheetException("The workbook reference is null");
        }
        if (!worksheetExists(name, workbook)) {
            return name;
        }
        Pattern pattern = Pattern.compile("^(.*?)(\\d{1,31})$"); //
        Matcher matcher = pattern.matcher(name);
        String prefix = name;
        int number = 1;
        if (matcher.matches() && matcher.groupCount() > 1) {
            prefix = matcher.group(1);
            try {
                number = Integer.parseInt(matcher.group(2));
            } catch (Exception ex) {
                number = 0;
                // if this failed, the start number is 0 (parsed number was >max. int32)
            }
        }
        while (true) {
            String numberString = Integer.toString(number);
            if (numberString.length() + prefix.length() > MAX_WORKSHEET_NAME_LENGTH) {
                int endIndex = prefix.length() - (numberString.length() + prefix.length() - MAX_WORKSHEET_NAME_LENGTH);
                prefix = prefix.substring(0, endIndex);
            }
            String newName = prefix + numberString;
            if (!worksheetExists(newName, workbook)) {
                return newName;
            }
            number++;
        }
    }

    /**
     * Checks whether a worksheet with the given name exists
     *
     * @param name     Name to check
     * @param workbook Workbook reference
     * @return True if the name exits, otherwise false
     */
    private static boolean worksheetExists(String name, Workbook workbook) {
        int len = workbook.getWorksheets().size();
        for (int i = 0; i < len; i++) {
            if (workbook.getWorksheets().get(i).getSheetName().equals(name)) {
                return true;
            }
        }
        return false;
    }

}
