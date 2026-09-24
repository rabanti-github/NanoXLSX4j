/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Predicate;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.internal.CellKey;
import ch.rabanti.nanoxlsx4j.internal.FeatureSet;
import ch.rabanti.nanoxlsx4j.internal.StringKeyedCellView;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Password;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Style;
import ch.rabanti.nanoxlsx4j.utils.Comparators;
import ch.rabanti.nanoxlsx4j.utils.DataUtils;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;
import ch.rabanti.nanoxlsx4j.utils.Validators;

public class Worksheet {

// Constants
    /**
     * Maximum number of characters a worksheet name can have
     */
    public static final int MAX_WORKSHEET_NAME_LENGTH = 31;
    /**
     * Default column width as constant
     */
    public static final float DEFAULT_WORKSHEET_COLUMN_WIDTH = 10f;
    /**
     * Default row height as constant
     */
    public static final float DEFAULT_WORKSHEET_ROW_HEIGHT = 15f;
    /**
     * Maximum column number (zero-based) as constant
     */
    public static final int MAX_COLUM_NUMBER = 16383;
    /**
     * Minimum column number (zero-based) as constant
     */
    public static final int MIN_COLUM_NUMBER = 0;
    /**
     * Minimum column width as constant
     */
    public static final float MIN_COLUMN_WIDTH = 0f;
    /**
     * Minimum row height as constant
     */
    public static final float MIN_ROW_HEIGHT = 0f;
    /**
     * Maximum column width as constant
     */
    public static final float MAX_COLUMN_WIDTH = 255f;
    /**
     * Maximum row number (zero-based) as constant
     */
    public static final int MAX_ROW_NUMBER = 1048575;
    /**
     * Minimum row number (zero-based) as constant
     */
    public static final int MIN_ROW_NUMBER = 0;
    /**
     * Maximum row height as constant
     */
    public static final float MAX_ROW_HEIGHT = 409.5f;
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

//enums

    /**
     * Enum to define the direction when using AddNextCell method
     */
    public enum CellDirection {
        /**
         * The next cell will be on the same row (A1,B1,C1...)
         */
        COLUMN_TO_COLUMN,
        /**
         * The next cell will be on the same column (A1,A2,A3...)
         */
        ROW_TO_ROW,
        /**
         * The address of the next cell will be not changed when adding a cell (for manual definition of cell
         * addresses)
         */
        DISABLED
    }

    /**
     * Enum to define the possible protection types when protecting a worksheet
     */
    public enum SheetProtectionValue {
        // sheet, // Is always on 1 if protected
        /**
         * If selected, the user can edit objects if the worksheets is protected
         */
        OBJECTS,
        /**
         * If selected, the user can edit scenarios if the worksheets is protected
         */
        SCENARIOS,
        /**
         * If selected, the user can format cells if the worksheets is protected
         */
        FORMAT_CELLS,
        /**
         * If selected, the user can format columns if the worksheets is protected
         */
        FORMAT_COLUMNS,
        /**
         * If selected, the user can format rows if the worksheets is protected
         */
        FORMAT_ROWS,
        /**
         * If selected, the user can insert columns if the worksheets is protected
         */
        INSERT_COLUMNS,
        /**
         * If selected, the user can insert rows if the worksheets is protected
         */
        INSERT_ROWS,
        /**
         * If selected, the user can insert hyperlinks if the worksheets is protected
         */
        INSERT_HYPERLINKS,
        /**
         * If selected, the user can delete columns if the worksheets is protected
         */
        DELETE_COLUMNS,
        /**
         * If selected, the user can delete rows if the worksheets is protected
         */
        DELETE_ROWS,
        /**
         * If selected, the user can select locked cells if the worksheets is protected
         */
        SELECT_LOCKED_CELLS,
        /**
         * If selected, the user can sort cells if the worksheets is protected
         */
        SORT,
        /**
         * If selected, the user can use auto filters if the worksheets is protected
         */
        AUTO_FILTER,
        /**
         * If selected, the user can use pivot tables if the worksheets is protected
         */
        PIVOT_TABLES,
        /**
         * If selected, the user can select unlocked cells if the worksheets is protected
         */
        SELECT_UNLOCKED_CELLS
    }

    /**
     * Enum to define the pane position or active pane in a slip worksheet
     */
    public enum WorksheetPane {
        /**
         * The pane is located in the bottom right of the split worksheet
         */
        BOTTOM_RIGHT,
        /**
         * The pane is located in the top right of the split worksheet
         */
        TOP_RIGHT,
        /**
         * The pane is located in the bottom left of the split worksheet
         */
        BOTTOM_LEFT,
        /**
         * The pane is located in the top left of the split worksheet
         */
        TOP_LEFT,
    }

    /**
     * Enum to define how a worksheet is displayed in the spreadsheet application (Excel)
     */
    public enum SheetViewType {
        /**
         * The worksheet is displayed without pagination (default)
         */
        NORMAL,
        /**
         * The worksheet is displayed with indicators where the page would break if it were printed
         */
        PAGE_BREAK_PREVIEW,
        /**
         * The worksheet is displayed like it would be printed
         */
        PAGE_LAYOUT
    }

    //privateField
    private Style activeStyle;
    private Range autoFilterRange;
    // Let's assume a default of 1000 cells per worksheet, which is a good starting point for most use cases. This
    // can be adjusted if needed, but it will save some resizing operations on the dictionary in the average case;
    private final Map<CellKey, Cell> cells;
    private final StringKeyedCellView cellsStringView;
    private final Map<Integer, Column> columns;
    private String sheetName;
    private int currentRowNumber;
    private int currentColumnNumber;
    private float defaultRowHeight;
    private float defaultColumnWidth;

    private final Map<Integer, Float> rowHeights;
    private final Map<Integer, Boolean> hiddenRows;
    private final Map<String, Range> mergedCells;
    private final List<SheetProtectionValue> sheetProtectionValues;
    private boolean useActiveStyle;
    private boolean hidden;
    private Workbook workbookReference;
    private Password sheetProtectionPassword;
    private List<Range> selectedCells;
    private Boolean freezeSplitPanes;
    private Float paneSplitLeftWidth;
    private Float paneSplitTopHeight;
    private Address paneSplitTopLeftCell;
    private Address paneSplitAddress;
    private WorksheetPane activePane;
    private int sheetId;
    private SheetViewType viewType;
    private final Map<SheetViewType, Integer> zoomFactors;
    private CellDirection currentCellDirection;
    private boolean showGridLines = true;
    private boolean showRowColumnHeaders = true;
    private boolean showRuler = true;
    private boolean sheetProtection;
    private FeatureSet features = new FeatureSet();

    // Note: This is a live view of cells. It does not double the memory footprint
    private final Collection<Cell> cellValues;

//getters&setters

    /**
     * Gets the range of the auto-filter. Wrapped to Optional to provide empty as value. If empty, no auto-filter is
     * applied
     *
     * @return optional auto-filter range
     */
    public Optional<Range> getAutoFilterRange() {
        return Optional.ofNullable(autoFilterRange);
    }

    /**
     * Gets the cells of the worksheet as read-only map with the cell address string as key and the cell object as
     * value. Use {@link #getCellValues()} for iteration-heavy code paths to avoid per-cell address-string allocation.
     *
     * @return Map of cells
     */
    public Map<String, Cell> getCells() {
        return cellsStringView;
    }

    /**
     * Gets all cells of the worksheet as a collection. Preferred over {@link #getCells()} in performance-critical paths
     * (Writer, StyleManager) because no address string is allocated per cell.
     *
     * @return Cell values
     */
    public Collection<Cell> getCellValues() {
        return cellValues;
    }

    /**
     * Gets all columns with non-standard properties, like auto filter applied or a special width as map with the
     * zero-based column index as key and the column object as value
     *
     * @return Map of columns
     */
    public Map<Integer, Column> getColumns() {
        return columns;
    }

    /**
     * Gets the direction when using addNextCell method
     *
     * @return Current cell direction
     */
    public CellDirection getCurrentCellDirection() {
        return currentCellDirection;
    }

    /**
     * Sets the direction when using addNextCell method
     *
     * @param currentCellDirection Current cell direction
     */
    public void setCurrentCellDirection(CellDirection currentCellDirection) {
        this.currentCellDirection = currentCellDirection;
    }

    /**
     * Gets the default column width
     *
     * @return Column width
     */
    public float getDefaultColumnWidth() {
        return defaultColumnWidth;
    }

    /**
     * Sets the default column width
     *
     * @param defaultColumnWidth Column width
     */
    public void setDefaultColumnWidth(float defaultColumnWidth) {
        if (defaultColumnWidth < MIN_COLUMN_WIDTH || defaultColumnWidth > MAX_COLUMN_WIDTH) {
            throw new RangeException(
                    "The passed default column width is out of range (" + MIN_COLUMN_WIDTH + " to " + MAX_COLUMN_WIDTH +
                            ")");
        }
        this.defaultColumnWidth = defaultColumnWidth;
    }

    /**
     * Gets the default Row height
     *
     * @return Row height
     */
    public float getDefaultRowHeight() {
        return defaultRowHeight;
    }

    /**
     * sets the default Row height
     *
     * @param defaultRowHeight Row height
     */
    public void setDefaultRowHeight(float defaultRowHeight) {
        if (defaultRowHeight < MIN_ROW_HEIGHT || defaultRowHeight > MAX_ROW_HEIGHT) {
            throw new RangeException(
                    "The passed default row height is out of range (" + MIN_ROW_HEIGHT + " to " + MAX_ROW_HEIGHT + ")");
        }
        this.defaultRowHeight = defaultRowHeight;
    }

    /**
     * Gets the hidden rows as map with the zero-based row number as key and a boolean as value. True indicates hidden,
     * false visible.
     *
     * <p>Remarks: Entries with the value false are not affecting the worksheet. These entries can be removed</p>
     *
     * @return Map of hidden rows
     */
    public Map<Integer, Boolean> getHiddenRows() {
        return hiddenRows;
    }

    /**
     * Gets defined row heights as map with the zero-based row number as key and the height (float from 0 to 409.5) as
     * value
     *
     * @return Row heights
     */
    public Map<Integer, Float> getRowHeights() {
        return rowHeights;
    }

    /**
     * Gets the merged cells (only references) as map with the cell address as key and the range object as value
     *
     * @return Merged cells
     */
    public Map<String, Range> getMergedCells() {
        return mergedCells;
    }

    /**
     * Gets the cell ranges of selected cells of this worksheet. Returns an empty list if no cells are selected
     *
     * @return Selected cells
     */
    public List<Range> getSelectedCells() {
        return selectedCells;
    }

    /**
     * Gets the internal ID of the worksheet
     *
     * @return Sheet ID
     */
    public int getSheetId() {
        return sheetId;
    }

    /**
     * Sets the internal ID of the worksheet
     *
     * @param sheetId Sheet ID
     * @throws FormatException Thrown if the set sheet ID is smaller than 1
     */
    public void setSheetId(int sheetId) {
        if (sheetId < 1) {
            throw new FormatException("The ID " + sheetId + " is invalid. Worksheet IDs must be >0");
        }
        this.sheetId = sheetId;
    }

    /**
     * Gets the name of the worksheet
     *
     * @return Name of the worksheet
     */
    public String getSheetName() {
        return sheetName;
    }

    //Note: The explicit C# method SetSheetName was collapsed to the setter in Java

    /**
     * Validates and sets the worksheet name
     *
     * @param sheetName Name of the worksheet
     * @throws FormatException if the worksheet name is too long (max. 31) or contains illegal characters [  ]  * ? / \
     */
    public void setSheetName(String sheetName) {
        Validators.validateWorksheetName(sheetName);
        this.sheetName = sheetName;
    }

    // Moved from method section in C# to getters

    /**
     * Gets the current column number (zero based)
     *
     * @return Column number (zero-based)
     */
    public int getCurrentColumnNumber() {
        return currentColumnNumber;
    }

    // Moved from method section in C# to getters

    /**
     * Gets the current row number (zero based)
     *
     * @return Row number (zero-based)
     */
    public int getCurrentRowNumber() {
        return currentRowNumber;
    }

    /// <summary>
    /// Password instance of the worksheet protection. If a password was set, the pain text representation and the hash
    /// can be read from the instance
    /// </summary>
    /// \remark <remarks>The password of this property is stored in plain text at runtime but not stored to a worksheet.
    /// The plain text password cannot be recovered when loading a workbook. The hash is retrieved and can be
    /// reused</remarks>

    public Password getSheetProtectionPassword() {
        return sheetProtectionPassword;
    }

    /**
     * Internal setter of the sheet rotection password
     *
     * @param sheetProtectionPassword Sheet protection password
     */
    void setSheetProtectionPassword(Password sheetProtectionPassword) {
        this.sheetProtectionPassword = sheetProtectionPassword;
    }

    /**
     * Gets the list of SheetProtectionValues. These values define the allowed actions if the worksheet is protected
     */
    public List<SheetProtectionValue> getSheetProtectionValues() {
        return sheetProtectionValues;
    }

    // UseSheetProtection in C#

    /**
     * Gets whether the worksheet is protected
     *
     * @return If true, protection is enabled, otherwise false
     */
    public boolean getSheetProtection() {
        return sheetProtection;
    }

    // UseSheetProtection in C#

    /**
     * Sets whether the worksheet is protected. If true, protection is enabled
     *
     * @param sheetProtection If true, protection is enabled, otherwise false
     */
    public void setSheetProtection(boolean sheetProtection) {
        this.sheetProtection = sheetProtection;
    }

    /**
     * Gets the reference to the parent workbook
     *
     * @return Parent worksheet
     */
    public Workbook getWorkbookReference() {
        return workbookReference;
    }

    /**
     * Gets the reference to the parent workbook
     *
     * @param workbookReference Parent worksheet
     */
    public void setWorkbookReference(Workbook workbookReference) {
        this.workbookReference = workbookReference;
        if (workbookReference != null) {
            workbookReference.validateWorksheets();
        }
    }

    /**
     * Gets whether the worksheet is hidden. If true, the worksheet is not listed in the worksheet tabs of the
     * workbook.<br /> If the worksheet is not part of a workbook, or the only one in the workbook, an exception will be
     * thrown.<br /> If the worksheet is the selected one, and attempted to set hidden, an exception will be thrown.
     * Define another selected worksheet prior to this call, in this case.
     *
     * @return True if hidden, otherwise false
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Sets whether the worksheet is hidden. If true, the worksheet is not listed in the worksheet tabs of the
     * workbook.<br /> If the worksheet is not part of a workbook, or the only one in the workbook, an exception will be
     * thrown.<br /> If the worksheet is the selected one, and attempted to set hidden, an exception will be thrown.
     * Define another selected worksheet prior to this call, in this case.
     *
     * @param hidden True if hidden, otherwise false
     */
    public void setHidden(boolean hidden) {
        this.hidden = hidden;
        if (hidden && workbookReference != null) {
            workbookReference.validateWorksheets();
        }
    }

    /**
     * Gets the height of the upper, horizontal split pane, measured from the top of the window.<br /> The value is
     * optional. If empty, no horizontal split of the worksheet is applied.<br /> The value is only applicable to split
     * the worksheet into panes, but not to freeze them.<br /> See also: {@link #getPaneSplitAddress()}
     *
     * <p>Remarks: Note: This value will be modified to the Excel-internal representation,
     * calculated by {@link DataUtils#getInternalPaneSplitHeight(float)}.
     * </p>
     *
     * @return Optional pane split top height
     */
    public Optional<Float> getPaneSplitTopHeight() {
        return Optional.ofNullable(paneSplitTopHeight);
    }

    /**
     * Gets the width of the left, vertical split pane, measured from the left of the window.<br /> The value is
     * optional. If empty, no vertical split of the worksheet is applied<br /> The value is only applicable to split the
     * worksheet into panes, but not to freeze them.<br /> See also: {@link #getPaneSplitAddress()}
     *
     * <p>Remarks: Note: This value will be modified to the Excel-internal representation,
     * calculated by {@link DataUtils#getPaneSplitWidth(float, float, float)}.
     * </p>
     *
     * @return Optional pane split left width
     */
    public Optional<Float> getPaneSplitLeftWidth() {
        return Optional.ofNullable(paneSplitLeftWidth);
    }

    /**
     * Gets whether split panes are frozen.<br /> The value is optional. If empty, no freezing is applied. This getter
     * also does not apply if {@link #getPaneSplitAddress()} is empty
     *
     * @return Optional indicator whether split panes are frozen
     */
    public Optional<Boolean> getFreezeSplitPanes() {
        return Optional.ofNullable(freezeSplitPanes);
    }

    /**
     * Gets the Top Left cell address of the bottom right pane if applicable and splitting is applied.<br /> The column
     * is only relevant for vertical split, whereas the row component is only relevant for a horizontal split.<br /> The
     * value is optional. If empty, no splitting was defined.
     *
     * @return Optional pane split top left cell address
     */
    public Optional<Address> getPaneSplitTopLeftCell() {
        return Optional.ofNullable(paneSplitTopLeftCell);
    }

    /**
     * Gets the split address for frozen panes or if pane split was defined in number of columns and / or rows.<br />
     * For vertical splits, only the column component is considered. For horizontal splits, only the row component is
     * considered.<br /> The value is optional. If empty, no frozen panes or split by columns / rows are applied to the
     * worksheet. However, splitting can still be applied, if the value is defined in characters.<br /> See also:
     * {@link #getPaneSplitLeftWidth()} and {@link #getPaneSplitTopHeight()} for splitting in characters (without
     * freezing)
     *
     * @return Optional pane split address
     */
    public Optional<Address> getPaneSplitAddress() {
        return Optional.ofNullable(paneSplitAddress);
    }

    /// <summary>
    /// Gets the active Pane is splitting is applied.<br /> The value is optional. If empty, no splitting was defined
    /// </summary>
    ///
    /// @return Optional active pane
    public Optional<WorksheetPane> getActivePane() {
        return Optional.ofNullable(activePane);
    }

    /**
     * Gets the active Style of the worksheet. If null, no style is defined as active
     *
     * @return Active style (nullable)
     */
    public Style getActiveStyle() {
        return activeStyle;
    }

    /**
     * Gets whether grid lines are visible on the current worksheet. Default is true
     *
     * @return True if grid lines are shown, otherwise false
     */
    public boolean isShowGridLines() {
        return showGridLines;
    }

    /**
     * Sets whether grid lines are visible on the current worksheet.
     *
     * @param showGridLines True if grid lines are shown, otherwise false
     */
    public void setShowGridLines(boolean showGridLines) {
        this.showGridLines = showGridLines;
    }

    /**
     * Gets whether the column and row headers are visible on the current worksheet. Default is true
     *
     * @return True if column headers are shown, otherwise false
     */
    public boolean isShowRowColumnHeaders() {
        return showRowColumnHeaders;
    }

    /**
     * Sets whether the column and row headers are visible on the current worksheet. Default is true
     *
     * @param showRowColumnHeaders True if column headers are shown, otherwise false
     */
    public void setShowRowColumnHeaders(boolean showRowColumnHeaders) {
        this.showRowColumnHeaders = showRowColumnHeaders;
    }

    /**
     * Gets whether a ruler is displayed over the column headers. This value only applies if {@link #getViewType()} is
     * set to {@link SheetViewType#PAGE_LAYOUT}. Default is true
     *
     * @return True if the ruler is show, otherwise false
     */
    public boolean isShowRuler() {
        return showRuler;
    }

    /**
     * Sets whether a ruler is displayed over the column headers. This value only applies if {@link #getViewType()} is
     * set to {@link SheetViewType#PAGE_LAYOUT}. Default is true
     *
     * @param showRuler True if the ruler is show, otherwise false
     */
    public void setShowRuler(boolean showRuler) {
        this.showRuler = showRuler;
    }

    /**
     * Gets how the current worksheet is displayed in the spreadsheet application (Excel)
     *
     * @return View type
     */
    public SheetViewType getViewType() {
        return viewType;
    }

    /**
     * Sets how the current worksheet is displayed in the spreadsheet application (Excel)
     *
     * @param viewType View type
     */
    public void setViewType(SheetViewType viewType) {
        this.viewType = viewType;
        setZoomFactor(viewType, 100);
    }

    /**
     * Gets the zoom factor of the {@link #getViewType()} of the current worksheet. If {@link #AUTO_ZOOM_FACTOR}, the
     * zoom factor is set to automatic
     *
     * <p>Remarks: It is possible to add further zoom factors for inactive view types, using the function
     * {@link #setZoomFactor(SheetViewType, int)}</p>
     *
     * @return Zoom factor
     **/
    public int getZoomFactor() {
        return zoomFactors.get(viewType);
    }

    /**
     * Sets the zoom factor of the {@link #getViewType()} of the current worksheet. If {@link #AUTO_ZOOM_FACTOR}, the
     * zoom factor is set to automatic
     *
     * <p>Remarks: It is possible to add further zoom factors for inactive view types, using the function
     * {@link #setZoomFactor(SheetViewType, int)}</p>
     *
     * @param zoomFactor Zoom factor
     * @throws WorksheetException Throws a WorksheetException if the zoom factor is not {@link #AUTO_ZOOM_FACTOR} or
     *                            below {@link #MIN_ZOOM_FACTOR} or above {@link #MAX_ZOOM_FACTOR}
     */
    public void setZoomFactor(int zoomFactor) {
        setZoomFactor(viewType, zoomFactor);
    }

    /**
     * Gets all defined zoom factors per {@link SheetViewType} of the current worksheet. Use
     * {@link #setZoomFactor(SheetViewType, int)} to define the values
     *
     * @return Map of zoom factors
     */
    public Map<SheetViewType, Integer> getZoomFactors() {
        return zoomFactors;
    }

    /**
     * Internal feature set for cascading feature detection (consider in  {@link #copy()} but not in equals, getHashCode
     * etc.)
     *
     * @return Feature set
     */
    FeatureSet getFeatures() {
        return this.features;
    }

    // constructors

    /**
     * Default Constructor
     */
    public Worksheet() {
        currentCellDirection = CellDirection.COLUMN_TO_COLUMN;
        cells = new HashMap<>(
                1000); // Let's assume a default of 1000 cells per worksheet, which is a good starting point for most
        // use cases. This can be adjusted if needed, but it will save some resizing operations on the
        // dictionary in the average case
        cellsStringView = new StringKeyedCellView(cells);
        cellValues = Collections.unmodifiableCollection(cells.values());
        currentRowNumber = 0;
        currentColumnNumber = 0;
        defaultColumnWidth = DEFAULT_WORKSHEET_COLUMN_WIDTH;
        defaultRowHeight = DEFAULT_WORKSHEET_ROW_HEIGHT;
        rowHeights = new HashMap<>();
        mergedCells = new HashMap<>();
        selectedCells = new ArrayList<>();
        sheetProtectionValues = new ArrayList<SheetProtectionValue>();
        hiddenRows = new HashMap<>();
        columns = new HashMap<>();
        activeStyle = null;
        workbookReference = null;
        viewType = SheetViewType.NORMAL;
        zoomFactors = new HashMap<>();
        zoomFactors.put(viewType, 100);
        showGridLines = true;
        showRowColumnHeaders = true;
        showRuler = true;
        sheetProtectionPassword = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION);
    }

    /**
     * Constructor with worksheet name
     *
     * <p>Remarks: Note that the worksheet name is not checked and fully sanitized against other worksheets with this
     * operation. This is later performed when the worksheet is added to the workbook</p>
     */
    public Worksheet(String name) {
        this();
        setSheetName(name);
    }

    /**
     * Constructor with name and sheet ID
     *
     * @param name      Name of the worksheet
     * @param id        ID of the worksheet (for internal use)
     * @param reference Reference to the parent Workbook
     */
    public Worksheet(String name, int id, Workbook reference) {
        this();
        setSheetName(name);
        setSheetId(id);
        workbookReference = reference;
    }

//methods AddNextCell

    /**
     * Adds an object to the next cell position. If the type of the value does not match with one of the supported data
     * types, it will be cast to a String. A prepared object of the type Cell will not be cast but adjusted. The
     * direction of the next cell depends on the current cell direction (default is
     * {@link CellDirection#COLUMN_TO_COLUMN}).
     *
     * <p>Remarks: Recognized are the following data types: Cell (prepared object), String, int, double, float, long,
     * short, BigDecimal, byte, Date, Duration, boolean. All other types will be cast into a string using the default
     * ToString() method
     * </p>
     *
     * @param value Unspecified value to insert
     * @throws RangeException Throws a RangeException if the next cell is out of range (on row or column)
     */
    public void addNextCell(Object value) {
        addNextCell(castValue(value, currentColumnNumber, currentRowNumber), true, null);
    }

    /**
     * Adds an object to the next cell position. If the type of the value does not match with one of the supported data
     * types, it will be cast to a String. A prepared object of the type Cell will not be cast but adjusted. The
     * direction of the next cell depends on the current cell direction (default is
     * {@link CellDirection#COLUMN_TO_COLUMN}).
     *
     * <p>Remarks: Recognized are the following data types: Cell (prepared object), String, int, double, float, long,
     * short, BigDecimal, byte, Date, Duration, boolean. All other types will be cast into a string using the default
     * ToString() method
     * </p>
     *
     * @param value Unspecified value to insert
     * @param style Style object to apply on this cell
     * @throws RangeException Throws a RangeException if the next cell is out of range (on row or column)
     * @throws StyleException Throws a StyleException if the default style was malformed
     */
    public void addNextCell(Object value, Style style) {
        addNextCell(castValue(value, currentColumnNumber, currentRowNumber), true, style);
    }

    /**
     * Method to insert a generic cell to the next cell position. The direction of the next cell depends on the current
     * cell direction (default is {@link CellDirection#COLUMN_TO_COLUMN}).
     *
     * <p>Remarks: Recognized are the following data types: String, int, double, float, long, short, BigDecimal, byte,
     * Date, Duration, boolean. All other types will be cast into a string using the default toString() method.<br /> If
     * the cell object already has a style definition, and a style or active style is defined, the cell style will be
     * merged, otherwise just set
     * </p>
     *
     * @param cell        Cell object to insert
     * @param incremental If true, the address value (row or column) will be incremented, otherwise not
     * @param style       If not null, the defined style will be applied to the cell, otherwise no style or the default
     *                    style will be applied
     * @throws StyleException Throws a StyleException if the default style was malformed
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
        CellKey cellKey = new CellKey(cell.getColumnNumber(), cell.getRowNumber());
        if (cells.containsKey(cellKey)) {
            Cell previousCell = cells.get(cellKey);
            previousCell.unbindFeatures();
        }
        cells.put(cellKey, cell);
        cell.bindFeatures(features);
        if (incremental) {
            if (currentCellDirection == CellDirection.COLUMN_TO_COLUMN) {
                currentColumnNumber++;
            } else if (currentCellDirection == CellDirection.ROW_TO_ROW) {
                currentRowNumber++;
            } else {
                // disabled / no-op
            }
        } else {
            if (currentCellDirection == CellDirection.COLUMN_TO_COLUMN) {
                currentColumnNumber = cell.getColumnNumber() + 1;
                currentRowNumber = cell.getRowNumber();
            } else if (currentCellDirection == CellDirection.ROW_TO_ROW) {
                currentColumnNumber = cell.getColumnNumber();
                currentRowNumber = cell.getRowNumber() + 1;
            } else {
                // disabled / no-op
            }
        }
    }

    /**
     * Method to cast a value or align an object of the type Cell to the context of the worksheet
     *
     * @param value  Unspecified value or object of the type Cell
     * @param column Column index
     * @param row    Row index
     * @return Cell object
     */
    private static Cell castValue(Object value, int column, int row) {
        Cell c;
        if (value instanceof Cell) {
            c = (Cell) value;
            c.setCellAddress2(new Address(column, row));
        } else {
            c = new Cell(value, Cell.CellType.DEFAULT, column, row);
        }
        return c;
    }

//methods AddNextCell

    /**
     * Adds an object to the defined cell address. If the type of the value does not match with one of the supported
     * data types, it will be cast to a String. A prepared object of the type Cell will not be cast but adjusted
     *
     * <p>Remarks: Recognized are the following data types: Cell (prepared object), String, int, double, float, long,
     * short, BigDecimal, byte, Date, Duration, boolean. All other types will be cast into a string using the default
     * ToString() method
     * </p>
     *
     * @param value        Unspecified value to insert
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @throws RangeException Throws a RangeException if the passed cell address is out of range
     */
    public void addCell(Object value, int columnNumber, int rowNumber) {
        addNextCell(castValue(value, columnNumber, rowNumber), false, null);
    }

    /**
     * Adds an object to the defined cell address. If the type of the value does not match with one of the supported
     * data types, it will be cast to a String. A prepared object of the type Cell will not be cast but adjusted
     *
     * <p>Remarks: Recognized are the following data types: Cell (prepared object), String, int, double, float, long,
     * short, BigDecimal, byte, Date, Duration, boolean. All other types will be cast into a string using the default
     * ToString() method
     * </p>
     *
     * @param value        Unspecified value to insert
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @param style        Style to apply on the cell
     * @throws StyleException Throws a StyleException if the passed style is malformed
     * @throws RangeException Throws a RangeException if the passed cell address is out of range
     */
    public void addCell(Object value, int columnNumber, int rowNumber, Style style) {
        addNextCell(castValue(value, columnNumber, rowNumber), false, style);
    }

    /**
     * Adds an object to the defined cell address. If the type of the value does not match with one of the supported
     * data types, it will be cast to a String. A prepared object of the type Cell will not be cast but adjusted
     *
     * <p>Remarks: Recognized are the following data types: Cell (prepared object), String, int, double, float, long,
     * short, BigDecimal, byte, Date, Duration, boolean. All other types will be cast into a string using the default
     * ToString() method
     * </p>
     *
     * @param value   Unspecified value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws RangeException  Throws a RangeException if the passed cell address is out of range
     * @throws FormatException Throws a FormatException if the passed cell address is malformed
     */
    public void addCell(Object value, String address) {
        int column;
        int row;
        Address addr = Cell.resolveCellCoordinate(address);
        addCell(value, addr.column(), addr.row());
    }

    /**
     * Adds an object to the defined cell address. If the type of the value does not match with one of the supported
     * data types, it will be cast to a String. A prepared object of the type Cell will not be cast but adjusted
     *
     * <p>Remarks: Recognized are the following data types: Cell (prepared object), String, int, double, float, long,
     * short, BigDecimal, byte, Date, Duration, boolean. All other types will be cast into a string using the default
     * ToString() method
     * </p>
     *
     * @param value   Unspecified value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @param style   Style to apply on the cell
     * @throws StyleException  Throws a StyleException if the passed style is malformed
     * @throws RangeException  Throws a RangeException if the passed cell address is out of range
     * @throws FormatException Throws a FormatException if the passed cell address is malformed
     */
    public void addCell(Object value, String address, Style style) {
        int column;
        int row;
        Address addr = Cell.resolveCellCoordinate(address);
        addCell(value, addr.column(), addr.row(), style);
    }

// methods AddCellFormula

    /**
     * Adds a formula of the type {@link FormulaData.FormulaType#NORMAL} as string expression to the defined cell
     * address
     *
     * @param formula Formula expression to insert (without leading equal sign)
     * @param address Cell address in the format A1 - XFD1048576
     * @throws RangeException  Throws a RangeException if the passed cell address is out of range
     * @throws FormatException Throws a FormatException if the passed cell address is malformed
     */
    public void addCellFormula(String formula, String address) {
        int column;
        int row;
        Address addr = Cell.resolveCellCoordinate(address);
        Cell c = new Cell(formula, Cell.CellType.FORMULA, addr.column(), addr.row());
        addNextCell(c, false, null);
    }

    /**
     * Adds a formula of the type {@link FormulaData.FormulaType#NORMAL} as string expression to the defined cell
     * address
     *
     * @param formula Formula expression to insert (without leading equal sign)
     * @param address Cell address in the format A1 - XFD1048576
     * @param style   Style to apply on the cell
     * @throws StyleException  Throws a StyleException if the passed style was malformed
     * @throws RangeException  Throws a RangeException if the passed cell address is out of range
     * @throws FormatException Throws a FormatException if the passed cell address is malformed
     */
    public void addCellFormula(String formula, String address, Style style) {
        int column;
        int row;
        Address addr = Cell.resolveCellCoordinate(address);
        Cell c = new Cell(formula, Cell.CellType.FORMULA, addr.column(), addr.row());
        addNextCell(c, false, style);
    }

    /**
     * Adds a formula of the type  {@link FormulaData.FormulaType#NORMAL} as string expression to the defined cell
     * address
     *
     * @param formula      Formula expression to insert (without leading equal sign)
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @throws RangeException Throws a RangeException if the passed cell address is out of range
     */
    public void addCellFormula(String formula, int columnNumber, int rowNumber) {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, columnNumber, rowNumber);
        addNextCell(c, false, null);
    }

    /**
     * Adds a formula of the type  {@link FormulaData.FormulaType#NORMAL} as string expression to the defined cell
     * address
     *
     * @param formula      Formula expression to insert (without leading equal sign)
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @param style        Style to apply on the cell
     * @throws RangeException Throws a RangeException if the passed cell address is out of range
     */
    public void addCellFormula(String formula, int columnNumber, int rowNumber, Style style) {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, columnNumber, rowNumber);
        addNextCell(c, false, style);
    }

    /**
     * Adds a formula of the type  {@link FormulaData.FormulaType#NORMAL} as string expression to the next cell
     * position. The object {@link Cell#getFormula()} is created automatically.
     *
     * @param formula Formula expression to insert (without leading equal sign)
     * @throws RangeException Trows a RangeException if the next cell is out of range (on row or column)
     */
    public void addNextCellFormula(String formula) {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, currentColumnNumber, currentRowNumber);
        addNextCell(c, true, null);
    }

    /**
     * Adds a formula of the type  {@link FormulaData.FormulaType#NORMAL} as string expression to the next cell
     * position. The object {@link Cell#getFormula()} is created automatically.
     *
     * @param formula Formula expression to insert (without leading equal sign)
     * @param style   Style to apply on the cell
     * @throws RangeException Trows a RangeException if the next cell is out of range (on row or column)
     */
    public void addNextCellFormula(String formula, Style style) {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, currentColumnNumber, currentRowNumber);
        addNextCell(c, true, style);
    }

// methods AddCellReference

    /**
     * Adds a cell whose content is a reference to a {@link DefinedName} in the workbook (either workbook-scoped or
     * worksheet-scoped). The cell will have {@link Cell.CellType#FORMULA} as data type and
     * {@link FormulaData#getDefinedNameReference()} in {@link Cell#getFormula()} as value. The cell value will be
     * {@link DefinedName#getName()} (formula expression). If the defined name has the type
     * {@link DefinedName.NameType#RANGE} (array), the transposed cells of that array will be added / replaced and set
     * to {@link Cell.CellType#FORMULA} with a {@link FormulaData} object, and set Value in
     * {@link FormulaData#getMasterCellAddress()}. This might overwrite existing cells.
     *
     * <p>Remarks: To remove a cell reference, use the {@link #removeCell(int, int)} method. If a values is set by
     * {@link Cell#getValue()} it will be overwritten.</p>
     *
     * @param definedName  Defined name to reference. Must not be null.
     * @param columnNumber Column number (zero-based).
     * @param rowNumber    Row number (zero-based).
     * @return Returns an immutable list of cell addresses that are affected by the defined name. For
     * {@link FormulaData.FormulaType#ARRAY}, the list contains all addresses of the array elements and that of the
     * master cell as first element. In all other cases, the list will just contain the passed address of this method
     * @throws WorksheetException Thrown if {@code definedName} is null.
     * @throws RangeException     Thrown if the passed cell coordinate is out of range.
     */
    public List<Address> addCellReference(DefinedName definedName, int columnNumber, int rowNumber) {
        return addCellReference(definedName, columnNumber, rowNumber, null);
    }

    /**
     * Adds a cell whose content is a reference to a {@link DefinedName} in the workbook (either workbook-scoped or
     * worksheet-scoped). The cell will have {@link Cell.CellType#FORMULA} as data type and
     * {@link FormulaData#getDefinedNameReference()} in {@link Cell#getFormula()} as value. The cell value will be
     * {@link DefinedName#getName()} (formula expression). If the defined name has the type
     * {@link DefinedName.NameType#RANGE} (array), the transposed cells of that array will be added / replaced and set
     * to {@link Cell.CellType#FORMULA} with a {@link FormulaData} object, and set Value in
     * {@link FormulaData#getMasterCellAddress()}. This might overwrite existing cells.
     *
     * <p>Remarks: To remove a cell reference, use the {@link #removeCell(int, int)} method. If a values is set by
     * {@link Cell#getValue()} it will be overwritten.</p>
     *
     * @param definedName  Defined name to reference. Must not be null.
     * @param columnNumber Column number (zero-based).
     * @param rowNumber    Row number (zero-based).
     * @param cachedValue  Optional cached value that will be shown as long as the cell is not be refreshed. The value
     *                     will be ignored if the defined name type is {@link DefinedName.NameType#CONSTANT}
     * @return Returns an immutable list of cell addresses that are affected by the defined name. For
     * {@link FormulaData.FormulaType#ARRAY}, the list contains all addresses of the array elements and that of the
     * master cell as first element. In all other cases, the list will just contain the passed address of this method
     * @throws WorksheetException Thrown if {@code definedName} is null.
     * @throws RangeException     Thrown if the passed cell coordinate is out of range.
     */
    public List<Address> addCellReference(
            DefinedName definedName, int columnNumber, int rowNumber, Object cachedValue) {
        if (definedName == null) {
            throw new WorksheetException("The defined name to set as cell reference must not be null.");
        }
        Cell c = new Cell(definedName.getName(), Cell.CellType.FORMULA, columnNumber, rowNumber);
        Optional<Range> arrayRange = c.setReference(definedName, cachedValue);
        if (arrayRange.isPresent()) {
            c.getFormula().setFormulaRange(arrayRange.get().toString());
        }
        addNextCell(c, false, null);
        List<Address> list = new ArrayList<>();
        list.add(new Address(columnNumber, rowNumber));
        if (arrayRange.isPresent()) {
            List<Address> addedCells = addDefinedNameArrayCells(c, arrayRange.get(), null);
            list.addAll(addedCells);
        }
        return List.copyOf(list);
    }

    /**
     * Adds a cell whose content is a reference to a {@link DefinedName} in the workbook (either workbook-scoped or
     * worksheet-scoped) with a style. The cell will have {@link Cell.CellType#FORMULA} as data type and
     * {@link FormulaData#getDefinedNameReference()} in {@link Cell#getFormula()} as value. The cell value will be
     * {@link DefinedName#getName()} (formula expression). If the defined name has the type
     * {@link DefinedName.NameType#RANGE} (array), the transposed cells of that array will be added / replaced and set
     * to {@link Cell.CellType#FORMULA} with a {@link FormulaData} object, and set Value in
     * {@link FormulaData#getMasterCellAddress()}. The style will be applied to these cells as well.This might overwrite
     * existing cells.
     *
     * @param definedName  Defined name to reference. Must not be null.
     * @param columnNumber Column number (zero-based).
     * @param rowNumber    Row number (zero-based).
     * @param style        Style to apply on the cell.
     * @return Returns an immutable list of cell addresses that are affected by the defined name. For
     * {@link FormulaData.FormulaType#ARRAY}, the list contains all addresses of the array elements and that of the
     * master cell as first element. In all other cases, the list will just contain the passed address of this method
     * @throws WorksheetException Thrown if {@code definedName} is null.
     * @throws RangeException     Thrown if the passed cell coordinate is out of range.
     * @throws StyleException     Thrown if the passed style is malformed.
     */
    public List<Address> addCellReference(DefinedName definedName, int columnNumber, int rowNumber, Style style) {
        return addCellReference(definedName, columnNumber, rowNumber, style, null);
    }

    /**
     * Adds a cell whose content is a reference to a {@link DefinedName} in the workbook (either workbook-scoped or
     * worksheet-scoped) with a style. The cell will have {@link Cell.CellType#FORMULA} as data type and
     * {@link FormulaData#getDefinedNameReference()} in {@link Cell#getFormula()} as value. The cell value will be
     * {@link DefinedName#getName()} (formula expression). If the defined name has the type
     * {@link DefinedName.NameType#RANGE} (array), the transposed cells of that array will be added / replaced and set
     * to {@link Cell.CellType#FORMULA} with a {@link FormulaData} object, and set Value in
     * {@link FormulaData#getMasterCellAddress()}. The style will be applied to these cells as well.This might overwrite
     * existing cells.
     *
     * @param definedName  Defined name to reference. Must not be null.
     * @param columnNumber Column number (zero-based).
     * @param rowNumber    Row number (zero-based).
     * @param style        Style to apply on the cell.
     * @param cachedValue  Optional cached value that will be shown as long as the cell is not be refreshed. The value
     *                     will be ignored if the defined name type is {@link DefinedName.NameType#CONSTANT}
     * @return Returns an immutable list of cell addresses that are affected by the defined name. For
     * {@link FormulaData.FormulaType#ARRAY}, the list contains all addresses of the array elements and that of the
     * master cell as first element. In all other cases, the list will just contain the passed address of this method
     * @throws WorksheetException Thrown if {@code definedName} is null.
     * @throws RangeException     Thrown if the passed cell coordinate is out of range.
     * @throws StyleException     Thrown if the passed style is malformed.
     */
    public List<Address> addCellReference(
            DefinedName definedName, int columnNumber, int rowNumber, Style style, Object cachedValue) {
        if (definedName == null) {
            throw new WorksheetException("The defined name to set as cell reference must not be null.");
        }
        Cell c = new Cell(definedName.getName(), Cell.CellType.FORMULA, columnNumber, rowNumber);
        Optional<Range> arrayRange = c.setReference(definedName, cachedValue);
        if (arrayRange.isPresent()) {
            c.getFormula().setFormulaRange(arrayRange.get().toString());
        }
        addNextCell(c, false, style);
        List<Address> list = new ArrayList<>();
        list.add(new Address(columnNumber, rowNumber));
        if (arrayRange.isPresent()) {
            List<Address> addedCells = addDefinedNameArrayCells(c, arrayRange.get(), style);
            list.addAll(addedCells);
        }
        return List.copyOf(list);
    }

    /**
     * Adds a cell whose content is a reference to a {@link DefinedName} in the workbook (either workbook-scoped or
     * worksheet-scoped), addressed by string. The cell will have {@link Cell.CellType#FORMULA} as data type and
     * {@link FormulaData#getDefinedNameReference()} in {@link Cell#getFormula()} as value. The cell value will be
     * {@link DefinedName#getName()} (formula expression). If the defined name has the type
     * {@link DefinedName.NameType#RANGE} (array), the transposed cells of that array will be added / replaced and set
     * to {@link Cell.CellType#FORMULA} with a {@link FormulaData} object, and set Value in
     * {@link FormulaData#getMasterCellAddress()}. This might overwrite existing cells.
     *
     * @param definedName Defined name to reference. Must not be null.
     * @param address     Cell address in the format A1 - XFD1048576.
     * @return Returns an immutable list of cell addresses that are affected by the defined name. For
     * {@link FormulaData.FormulaType#ARRAY}, the list contains all addresses of the array elements and that of the
     * master cell as first element. In all other cases, the list will just contain the passed address of this method
     * @throws WorksheetException Thrown if {@code definedName} is null.
     * @throws RangeException     Thrown if the passed cell address is out of range.
     * @throws FormatException    Thrown if the passed cell address is malformed.
     */
    public List<Address> addCellReference(DefinedName definedName, String address) {
        return addCellReference(definedName, address, null);
    }

    /**
     * Adds a cell whose content is a reference to a {@link DefinedName} in the workbook (either workbook-scoped or
     * worksheet-scoped), addressed by string. The cell will have {@link Cell.CellType#FORMULA} as data type and
     * {@link FormulaData#getDefinedNameReference()} in {@link Cell#getFormula()} as value. The cell value will be
     * {@link DefinedName#getName()} (formula expression). If the defined name has the type
     * {@link DefinedName.NameType#RANGE} (array), the transposed cells of that array will be added / replaced and set
     * to {@link Cell.CellType#FORMULA} with a {@link FormulaData} object, and set Value in
     * {@link FormulaData#getMasterCellAddress()}. This might overwrite existing cells.
     *
     * @param definedName Defined name to reference. Must not be null.
     * @param address     Cell address in the format A1 - XFD1048576.
     * @param cachedValue Optional cached value that will be shown as long as the cell is not be refreshed. The value
     *                    will be ignored if the defined name type is {@link DefinedName.NameType#CONSTANT}
     * @return Returns an immutable list of cell addresses that are affected by the defined name. For
     * {@link FormulaData.FormulaType#ARRAY}, the list contains all addresses of the array elements and that of the
     * master cell as first element. In all other cases, the list will just contain the passed address of this method
     * @throws WorksheetException Thrown if {@code definedName} is null.
     * @throws RangeException     Thrown if the passed cell address is out of range.
     * @throws FormatException    Thrown if the passed cell address is malformed.
     */
    public List<Address> addCellReference(DefinedName definedName, String address, Object cachedValue) {
        int column;
        int row;
        Address addr = Cell.resolveCellCoordinate(address);
        return addCellReference(definedName, addr.column(), addr.row(), cachedValue);
    }

    /**
     * Adds a cell whose content is a reference to a {@link DefinedName} in the workbook (either workbook-scoped or
     * worksheet-scoped), addressed by string and with a style. The cell will have {@link Cell.CellType#FORMULA} as data
     * type and {@link FormulaData#getDefinedNameReference()} in {@link Cell#getFormula()} as value. The cell value will
     * be {@link DefinedName#getName()} (formula expression). If the defined name has the type
     * {@link DefinedName.NameType#RANGE} (array), the transposed cells of that array will be added / replaced and set
     * to {@link Cell.CellType#FORMULA} with a {@link FormulaData} object, and set Value in
     * {@link FormulaData#getMasterCellAddress()}. This might overwrite existing cells.
     *
     * @param definedName Defined name to reference. Must not be null.
     * @param address     Cell address in the format A1 - XFD1048576.
     * @param style       Style to apply on the cell.
     * @return Returns an immutable list of cell addresses that are affected by the defined name. For
     * {@link FormulaData.FormulaType#ARRAY}, the list contains all addresses of the array elements and that of the
     * master cell as first element. In all other cases, the list will just contain the passed address of this method
     * @throws WorksheetException Thrown if {@code definedName} is null.
     * @throws RangeException     Thrown if the passed cell address is out of range.
     * @throws FormatException    Thrown if the passed cell address is malformed.
     * @throws StyleException     Thrown if the passed style is malformed.
     */
    public List<Address> addCellReference(DefinedName definedName, String address, Style style) {
        return addCellReference(definedName, address, style, null);
    }

    /**
     * Adds a cell whose content is a reference to a {@link DefinedName} in the workbook (either workbook-scoped or
     * worksheet-scoped), addressed by string and with a style. The cell will have {@link DefinedName#getName()} as data
     * type and {@link FormulaData#getDefinedNameReference()} in {@link Cell#getFormula()} as value. The cell value will
     * be {@link DefinedName#getName()} (formula expression). If the defined name has the type
     * {@link DefinedName.NameType#RANGE} (array), the transposed cells of that array will be added / replaced and set
     * to {@link Cell.CellType#FORMULA} with a {@link FormulaData} object, and set Value in
     * {@link FormulaData#getMasterCellAddress()}. This might overwrite existing cells.
     *
     * @param definedName Defined name to reference. Must not be null.
     * @param address     Cell address in the format A1 - XFD1048576.
     * @param style       Style to apply on the cell.
     * @param cachedValue Optional cached value that will be shown as long as the cell is not be refreshed. The value
     *                    will be ignored if the defined name type is {@link DefinedName.NameType#CONSTANT}
     * @return Returns an immutable list of cell addresses that are affected by the defined name. For
     * {@link FormulaData.FormulaType#ARRAY}, the list contains all addresses of the array elements and that of the
     * master cell as first element. In all other cases, the list will just contain the passed address of this method
     * @throws WorksheetException Thrown if {@code definedName} is null.
     * @throws RangeException     Thrown if the passed cell address is out of range.
     * @throws FormatException    Thrown if the passed cell address is malformed.
     * @throws StyleException     Thrown if the passed style is malformed.
     */
    public List<Address> addCellReference(DefinedName definedName, String address, Style style, Object cachedValue) {
        int column;
        int row;
        Address addr = Cell.resolveCellCoordinate(address);
        return addCellReference(definedName, addr.column(), addr.row(), style, cachedValue);
    }

    /**
     * Adds array cells derived from a defined name object of a master cell, with array as type
     *
     * @param masterCell Master cell object
     * @param arrayRange Range of the array to insert
     * @param style      Optional style that will be applied to the array cells
     * @return Returns the list of all array cell addresses without the master cell
     */
    List<Address> addDefinedNameArrayCells(Cell masterCell, Range arrayRange, Style style) {
        List<Address> addresses = arrayRange.resolveEnclosedAddresses();
        List<Address> addedAddresses = new ArrayList<>();
        for (Address address : addresses) {
            if (address.row() == masterCell.getRowNumber() && address.column() == masterCell.getColumnNumber()) {
                continue; // Skip master cell
            }
            Cell arrayRefCell = new Cell(null, Cell.CellType.FORMULA);
            arrayRefCell.getFormula().setMasterCellAddress(masterCell.getCellAddress());
            arrayRefCell.getFormula().setFormulaRange(masterCell.getFormula().getFormulaRange());
            arrayRefCell.getFormula().setType(FormulaData.FormulaType.ARRAY);
            addCell(arrayRefCell, address.column(), address.row(), style);
            addedAddresses.add(address);
        }
        return List.copyOf(addedAddresses);
    }
// methods AddCellRange

    /**
     * Adds a list of object values to a defined cell range. If the type of the particular value does not match with one
     * of the supported data types, it will be cast to a String. Prepared objects of the type Cell will not be cast but
     * adjusted
     *
     * <p>Remarks: The data types in the passed list can be mixed. Recognized are the following data types: Cell
     * (prepared object), String, int, double, float, long, short, BigDecimal, byte, Date, Duration, boolean. All other
     * types will be cast into a string using the default toString() method
     * </p>
     *
     * @param values       List of unspecified objects to insert
     * @param startAddress Start address
     * @param endAddress   End address
     * @throws RangeException Throws a RangeException if the number of cells resolved from the range differs from the
     *                        number of passed values
     */
    public void addCellRange(List<Object> values, Address startAddress, Address endAddress) {
        addCellRangeInternal(values, startAddress, endAddress, null);
    }

    /**
     * Adds a list of object values to a defined cell range. If the type of the particular value does not match with one
     * of the supported data types, it will be cast to a String. Prepared objects of the type Cell will not be cast but
     * adjusted
     *
     * <p>Remarks: The data types in the passed list can be mixed. Recognized are the following data types: Cell
     * (prepared object), String, int, double, float, long, short, BigDecimal, byte, Date, Duration, boolean. All other
     * types will be cast into a string using the default toString() method
     * </p>
     *
     * @param values       List of unspecified objects to insert
     * @param startAddress Start address
     * @param endAddress   End address
     * @param style        Style to apply on the all cells of the range
     * @throws RangeException Throws a RangeException if the number of cells resolved from the range differs from the
     *                        number of passed values
     * @throws StyleException Throws a StyleException if the passed style is malformed
     */
    public void addCellRange(List<Object> values, Address startAddress, Address endAddress, Style style) {
        addCellRangeInternal(values, startAddress, endAddress, style);
    }

    /**
     * Adds a list of object values to a defined cell range. If the type of the particular value does not match with one
     * of the supported data types, it will be cast to a String. Prepared objects of the type Cell will not be cast but
     * adjusted
     *
     * <p>Remarks: The data types in the passed list can be mixed. Recognized are the following data types: Cell
     * (prepared object), String, int, double, float, long, short, BigDecimal, byte, Date, Duration, boolean. All other
     * types will be cast into a string using the default toString() method
     * </p>
     *
     * @param values    List of unspecified objects to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws RangeException  Throws a RangeException if the number of cells resolved from the range differs from the
     *                         number of passed values
     * @throws FormatException Throws a FormatException if the passed cell range is malformed
     */
    public void addCellRange(List<Object> values, String cellRange) {
        Range range = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, range.startAddress(), range.endAddress(), null);
    }

    /**
     * Adds a list of object values to a defined cell range. If the type of the particular value does not match with one
     * of the supported data types, it will be cast to a String. Prepared objects of the type Cell will not be cast but
     * adjusted
     *
     * <p>Remarks: The data types in the passed list can be mixed. Recognized are the following data types: Cell
     * (prepared object), String, int, double, float, long, short, BigDecimal, byte, Date, Duration, boolean. All other
     * types will be cast into a string using the default toString() method
     * </p>
     *
     * @param values    List of unspecified objects to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @param style     Style to apply on the all cells of the range
     * @throws RangeException  Throws a RangeException if the number of cells resolved from the range differs from the
     *                         number of passed values
     * @throws StyleException  Throws a StyleException if the passed style is malformed
     * @throws FormatException Throws a FormatException if the passed cell range is malformed
     */
    public void addCellRange(List<Object> values, String cellRange, Style style) {
        Range range = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, range.startAddress(), range.endAddress(), style);
    }

    /**
     * Adds a list of object values to a defined cell range. If the type of the particular value does not match with one
     * of the supported data types, it will be cast to a String. Prepared objects of the type Cell will not be cast but
     * adjusted
     *
     * <p>Remarks: The data types in the passed list can be mixed. Recognized are the following data types: Cell
     * (prepared object), String, int, double, float, long, short, BigDecimal, byte, Date, Duration, boolean. All other
     * types will be cast into a string using the default toString() method
     * </p>
     *
     * @param values    List of unspecified objects to insert
     * @param cellRange Cell range object
     * @throws RangeException Throws a RangeException if the number of cells resolved from the range differs from the
     *                        number of passed values
     */
    public void addCellRange(List<Object> values, Range cellRange) {
        addCellRangeInternal(values, cellRange.startAddress(), cellRange.endAddress(), null);
    }

    /**
     * Adds a list of object values to a defined cell range. If the type of the particular value does not match with one
     * of the supported data types, it will be cast to a String. Prepared objects of the type Cell will not be cast but
     * adjusted
     *
     * <p>Remarks: The data types in the passed list can be mixed. Recognized are the following data types: Cell
     * (prepared object), String, int, double, float, long, short, BigDecimal, byte, Date, Duration, boolean. All other
     * types will be cast into a string using the default toString() method
     * </p>
     *
     * @param values    List of unspecified objects to insert
     * @param cellRange Cell range object
     * @param style     Style to apply on the all cells of the range
     * @throws RangeException Throws a RangeException if the number of cells resolved from the range differs from the
     *                        number of passed values
     * @throws StyleException Throws a StyleException if the passed style is malformed
     */
    public void addCellRange(List<Object> values, Range cellRange, Style style) {
        addCellRangeInternal(values, cellRange.startAddress(), cellRange.endAddress(), style);
    }

    /**
     * Internal function to add a generic list of value to the defined cell range
     *
     * <p>Remarks: The data types in the passed list can be mixed. Recognized are the following data types: Cell
     * (prepared object), String, int, double, float, long, short, BigDecimal, byte, Date, Duration, boolean. All other
     * types will be cast into a string using the default toString() method
     * </p>
     *
     * @param <T>          Data type of the generic value list
     * @param values       List of values
     * @param startAddress Start address
     * @param endAddress   End address
     * @param style        Style to apply on the all cells of the range
     * @throws RangeException Throws a RangeException if the passes list is null or the number of cells differs from the
     *                        number of passed values
     */
    private <T> void addCellRangeInternal(List<T> values, Address startAddress, Address endAddress, Style style) {
        if (values == null) {
            throw new RangeException("The passed value list cannot be null");
        }
        List<Address> addresses = Cell.getCellRange(startAddress, endAddress);
        if (values.size() != addresses.size()) {
            throw new RangeException("The number of passed values (" + values.size() +
                    ") differs from the number of cells within the range (" + addresses.size() + ")");
        }
        List<Cell> list = Cell.convertArray(values);
        int len = values.size();
        for (int i = 0; i < len; i++) {
            list.get(i).setRowNumber(addresses.get(i).row());
            list.get(i).setColumnNumber(addresses.get(i).column());
            addNextCell(list.get(i), false, style);
        }
    }

// methods RemoveCell

    /**
     * Removes a previous inserted cell at the defined address
     *
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @return Returns true if the cell could be removed (existed), otherwise false (did not exist)
     * @throws RangeException Throws a RangeException if the passed cell address is out of range
     */
    public boolean removeCell(int columnNumber, int rowNumber) {
        CellKey key = new CellKey(columnNumber, rowNumber);
        if (cells.containsKey(key)) {
            cells.get(key).unbindFeatures(); // Decrease counter of features (if applicable)
        }
        Cell removed = cells.remove(key);
        return removed != null;
    }

    /**
     * Removes a previous inserted cell at the defined address
     *
     * @param address Cell address in the format A1 - XFD1048576
     * @return Returns true if the cell could be removed (existed), otherwise false (did not exist)
     * @throws RangeException  Throws a RangeException if the passed cell address is out of range
     * @throws FormatException Throws a FormatException if the passed cell address is malformed
     */
    public boolean removeCell(String address) {
        int row;
        int column;
        Address addr = Cell.resolveCellCoordinate(address);
        return removeCell(addr.column(), addr.row());
    }

// methods setStyle

    /**
     * Sets the passed style on the passed cell range. If cells are already existing, the style will be added or
     * replaced. Otherwise, an empty cell will be added with the assigned style. If the passed style is null, all styles
     * will be removed on existing cells and no additional (empty) cells are added to the worksheet
     *
     * <p>Remarks: Note: This method may invalidate an existing date or time value since dates and times are defined by
     * specific style. The result of a redefinition will be a number, instead of a date or time</p>
     *
     * @param cellRange Cell range to apply the style
     * @param style     Style to apply
     */
    public void setStyle(Range cellRange, Style style) {
        List<Address> addresses = cellRange.resolveEnclosedAddresses();
        for (Address address : addresses) {
            CellKey key = new CellKey(address.column(), address.row());
            if (cells.containsKey(key)) {
                Cell existing = cells.get(key);
                if (style == null) {
                    existing.removeStyle();
                } else {
                    existing.setStyle(style);
                }
            } else {
                if (style != null) {
                    addCell(null, address.column(), address.row(), style);
                }
            }
        }
    }

    /**
     * Sets the passed style on the passed cell range, derived from a start and end address. If cells are already
     * existing, the style will be added or replaced. Otherwise, an empty cell will be added with the assigned style. If
     * the passed style is null, all styles will be removed on existing cells and no additional (empty) cells are added
     * to the worksheet
     *
     * <p>Remarks: Note: This method may invalidate an existing date or time value since dates and times are defined by
     * specific style. The result of a redefinition will be a number, instead of a date or time</p>
     *
     * @param startAddress Start address of the cell range
     * @param endAddress   End address of the cell range
     * @param style        Style to apply or null to clear the range
     */
    public void setStyle(Address startAddress, Address endAddress, Style style) {
        setStyle(new Range(startAddress, endAddress), style);
    }

    /**
     * Sets the passed style on the passed (singular) cell address. If the cell is already existing, the style will be
     * added or replaced. Otherwise, an empty cell will be added with the assigned style. If the passed style is null,
     * all styles will be removed on existing cells and no additional (empty) cells are added to the worksheet
     *
     * <p>Remarks: Note: This method may invalidate an existing date or time value since dates and times are defined by
     * specific style. The result of a redefinition will be a number, instead of a date or time</p>
     *
     * @param address Cell address to apply the style
     * @param style   Style to apply or null to clear the range
     */
    public void setStyle(Address address, Style style) {
        setStyle(address, address, style);
    }

    /**
     * Sets the passed style on the passed address expression. Such an expression may be a single cell or a cell range.
     * If the cell is already existing, the style will be added or replaced. Otherwise, an empty cell will be added with
     * the assigned style. If the passed style is null, all styles will be removed on existing cells and no additional
     * (empty) cells are added to the worksheet
     *
     * <p>Remarks: Note: This method may invalidate an existing date or time value since dates and times are defined by
     * specific style. The result of a redefinition will be a number, instead of a date or time</p>
     *
     * @param addressExpression Expression of a cell address or range of addresses
     * @param style             Style to apply or null to clear the range
     */
    public void setStyle(String addressExpression, Style style) {
        Cell.AddressScope scope = Cell.getAddressScope(addressExpression);
        if (scope == Cell.AddressScope.SINGLE_ADDRESS) {
            Address address = new Address(addressExpression);
            setStyle(address, style);
        } else if (scope == Cell.AddressScope.RANGE) {
            Range range = new Range(addressExpression);
            setStyle(range, style);
        } else {
            throw new FormatException(
                    "The passed address'" + addressExpression + "' is neither a cell address, nor a range");
        }
    }

// boundary Functions

    /**
     * Gets the first existing column number in the current worksheet (zero-based)
     *
     * <p>Remarks: getFirstColumnNumber() will not return the first column with data in any case. If there is a
     * formatted but empty cell (or many) before the first cell with data, getFirstColumnNumber() will return the column
     * number of this empty cell. Use {@link #getFirstDataColumnNumber()} in this case.
     * </p>
     *
     * @return Zero-based column number. In case of an empty worksheet, -1 will be returned
     */
    public int getFirstColumnNumber() {
        return getBoundaryNumber(false, true);
    }

    /**
     * Gets the first existing column number with data in the current worksheet (zero-based)
     *
     * <p>Remarks: getFirstDataColumnNumber() will ignore formatted but empty cells before the first column with data.
     * If you want the first defined column, use {@link #getFirstColumnNumber()} instead.
     * </p>
     *
     * @return Zero-based column number. In case of an empty worksheet, -1 will be returned
     */
    public int getFirstDataColumnNumber() {
        return getBoundaryDataNumber(false, true, true);
    }

    /**
     * Gets the first existing row number in the current worksheet (zero-based)
     *
     * <p>Remarks: getFirstRowNumber() will not return the first row with data in any case. If there is a formatted but
     * empty cell (or many) before the first cell with data, getFirstRowNumber() will return the row number of this
     * empty cell. Use {@link #getFirstDataRowNumber()} in this case.
     * </p>
     *
     * @return Zero-based row number. In case of an empty worksheet, -1 will be returned
     */
    public int getFirstRowNumber() {
        return getBoundaryNumber(true, true);
    }

    /**
     * Gets the first existing row number with data in the current worksheet (zero-based)
     *
     * <p>Remarks: getFirstDataRowNumber() will ignore formatted but empty cells before the first row with data.
     * If you want the first defined row, use {@link #getFirstRowNumber()} instead.
     * </p>
     *
     * @return Zero-based row number. In case of an empty worksheet, -1 will be returned
     */
    public int getFirstDataRowNumber() {
        return getBoundaryDataNumber(true, true, true);
    }

    /**
     * Gets the last existing column number in the current worksheet (zero-based)
     *
     * <p>Remarks: getLastColumnNumber() will not return the last column with data in any case. If there is a formatted
     * (or with the definition of AutoFilter, column width or hidden state) but empty cell (or many) after the last cell
     * with data, getLastColumnNumber() will return the column number of this empty cell. Use
     * {@link #getLastDataColumnNumber()} in this case.
     * </p>
     *
     * @return Zero-based column number. In case of an empty worksheet, -1 will be returned
     */
    public int getLastColumnNumber() {
        return getBoundaryNumber(false, false);
    }

    /**
     * Gets the last existing column number with data in the current worksheet (zero-based)
     *
     * <p>Remarks: getLastDataColumnNumber() will ignore formatted (or with the definition of AutoFilter, column width
     * or hidden state) but empty cells after the last column with data. If you want the last defined column, use
     * {@link #getLastColumnNumber()} instead.
     * </p>
     *
     * @return Zero-based column number. in case of an empty worksheet, -1 will be returned
     */
    public int getLastDataColumnNumber() {
        return getBoundaryDataNumber(false, false, true);
    }

    /**
     * Gets the last existing row number in the current worksheet (zero-based)
     *
     * <p>Remarks: getLastRowNumber() will not return the last row with data in any case. If there is a formatted (or
     * with the definition of row height or hidden state) but empty cell (or many) after the last cell with data,
     * getLastRowNumber() will return the row number of this empty cell. Use {@link #getLastDataRowNumber()} in this
     * case.
     * </p>
     *
     * @return Zero-based row number. In case of an empty worksheet, -1 will be returned
     */
    public int getLastRowNumber() {
        return getBoundaryNumber(true, false);
    }

    /**
     * Gets the last existing row number with data in the current worksheet (zero-based)
     *
     * <p>Remarks: getLastDataColumnNumber() will ignore formatted (or with the definition of row height or hidden
     * state) but empty cells after the last column with data. If you want the last defined column, use
     * {@link #getLastRowNumber()} instead.
     * </p>
     *
     * @return Zero-based row number. in case of an empty worksheet, -1 will be returned
     */
    public int getLastDataRowNumber() {
        return getBoundaryDataNumber(true, false, true);
    }

    /**
     * Gets the last existing cell in the current worksheet (bottom right)
     *
     * <p>Remarks: GetLastCellAddress() will not return the last cell with data in any case. If there is a formatted
     * (or with definitions of hidden states, AutoFilters, heights or widths) but empty cell (or many) after the last
     * cell with data, getLastCellAddress() will return the address of this empty cell. Use
     * {@link #getLastDataCellAddress()} in this case.
     * </p>
     *
     * @return Optional Cell Address. If no cell address could be determined, empty will be returned
     */
    public Optional<Address> getLastCellAddress() {
        int lastRow = getLastRowNumber();
        int lastColumn = getLastColumnNumber();
        if (lastRow < 0 || lastColumn < 0) {
            return Optional.empty();
        }
        return Optional.of(new Address(lastColumn, lastRow));
    }

    /**
     * Gets the last existing cell with data in the current worksheet (bottom right)
     *
     * <p>Remarks: GetLastDataCellAddress() will ignore formatted (or with definitions of hidden states, AutoFilters,
     * heights or widths) but empty cells after the last cell with data. If you want the last defined cell, use
     * {@link #getLastCellAddress()} instead.
     * </p>
     *
     * @return optional Cell Address. If no cell address could be determined, empty will be returned
     */
    public Optional<Address> getLastDataCellAddress() {
        int lastRow = getLastDataRowNumber();
        int lastColumn = getLastDataColumnNumber();
        if (lastRow < 0 || lastColumn < 0) {
            return Optional.empty();
        }
        return Optional.of(new Address(lastColumn, lastRow));
    }

    /**
     * Gets the first existing cell in the current worksheet (bottom right)
     *
     * <p>Remarks: GetFirstCellAddress() will not return the first cell with data in any case. If there is a formatted
     * but empty cell (or many) before the first cell with data, GetLastCellAddress() will return the address of this
     * empty cell. Use {@link #getFirstDataCellAddress()} in this case.
     * </p>
     *
     * @return Optional Cell Address. If no cell address could be determined, empty will be returned
     */
    public Optional<Address> getFirstCellAddress() {
        int firstRow = getFirstRowNumber();
        int firstColumn = getFirstColumnNumber();
        if (firstRow < 0 || firstColumn < 0) {
            return Optional.empty();
        }
        return Optional.of(new Address(firstColumn, firstRow));
    }

    /// <summary>
    ///  Gets the first existing cell with data in the current worksheet (bottom right)
    /// </summary>
    /// <returns>Optional Cell Address. If no cell address could be determined, empty will be returned</returns>
    /// \remark <remarks>GetFirstDataCellAddress() will ignore formatted but empty cells before the first cell with
    /// data. If you want the first defined cell, use <see cref="getFirstCellAddress()"/> instead.</remarks>
    public Optional<Address> getFirstDataCellAddress() {
        int firstRow = getFirstDataRowNumber();
        int firstColumn = getFirstDataColumnNumber();
        if (firstRow < 0 || firstColumn < 0) {
            return Optional.empty();
        }
        return Optional.of(new Address(firstColumn, firstRow));
    }

    /**
     * Gets either the minimum or maximum row or column number, considering only cells with data.
     *
     * @param row         If true, the min or max row is returned, otherwise the column
     * @param min         If true, the min value of the row or column is defined, otherwise the max value
     * @param ignoreEmpty If true, empty cell values are ignored, otherwise considered without checking the content
     * @return Min or max number, or -1 if not defined
     */
    private int getBoundaryDataNumber(boolean row, boolean min, boolean ignoreEmpty) {
        if (cells.isEmpty()) {
            return -1;
        }

        IntStream values = cells.values().stream()
                .filter(cell -> !ignoreEmpty
                        || (cell.getValue() != null
                        && !cell.getValue().toString().isEmpty()))
                .mapToInt(cell -> row
                        ? cell.getRowNumber()
                        : cell.getColumnNumber());

        return min
                ? values.min().orElse(-1)
                : values.max().orElse(-1);
    }

    /**
     * Gets either the minimum or maximum row or column number, considering all available data.
     *
     * @param row If true, the min or max row is returned, otherwise the column
     * @param min If true, the min value of the row or column is defined, otherwise the max value
     * @return Min or max number, or -1 if not defined
     */
    private int getBoundaryNumber(boolean row, boolean min) {
        int cellBoundary = getBoundaryDataNumber(row, min, false);

        if (row) {
            int heightBoundary = -1;
            if (!rowHeights.isEmpty()) {
                heightBoundary = min
                        ? rowHeights.keySet().stream().mapToInt(Integer::intValue).min().orElse(-1)
                        : rowHeights.keySet().stream().mapToInt(Integer::intValue).max().orElse(-1);
            }

            int hiddenBoundary = -1;
            if (!hiddenRows.isEmpty()) {
                hiddenBoundary = min
                        ? hiddenRows.keySet().stream().mapToInt(Integer::intValue).min().orElse(-1)
                        : hiddenRows.keySet().stream().mapToInt(Integer::intValue).max().orElse(-1);
            }

            return min
                    ? getMinRow(cellBoundary, heightBoundary, hiddenBoundary)
                    : getMaxRow(cellBoundary, heightBoundary, hiddenBoundary);
        }

        int columnDefBoundary = -1;
        if (!columns.isEmpty()) {
            columnDefBoundary = min
                    ? columns.keySet().stream().mapToInt(Integer::intValue).min().orElse(-1)
                    : columns.keySet().stream().mapToInt(Integer::intValue).max().orElse(-1);
        }

        if (min) {
            return cellBoundary >= 0 && cellBoundary < columnDefBoundary
                    ? cellBoundary
                    : columnDefBoundary;
        }

        return cellBoundary >= 0 && cellBoundary > columnDefBoundary
                ? cellBoundary
                : columnDefBoundary;
    }

    /**
     * Gets the maximum row coordinate either from cell data, height definitions or hidden rows
     *
     * @param cellBoundary   Row number of max cell data
     * @param heightBoundary Row number of max defined row height
     * @param hiddenBoundary Row number of max defined hidden row
     * @return Max row number or -1 if nothing valid defined
     */
    private static int getMaxRow(int cellBoundary, int heightBoundary, int hiddenBoundary) {
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
     * Gets the minimum row coordinate either from cell data, height definitions or hidden rows
     *
     * @param cellBoundary   Row number of min cell data
     * @param heightBoundary Row number of min defined row height
     * @param hiddenBoundary Row number of min defined hidden row
     * @return Min row number or -1 if nothing valid defined
     */
    private static int getMinRow(int cellBoundary, int heightBoundary, int hiddenBoundary) {
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

// Insert-Search-Replace

    /**
     * Inserts 'count' rows below the specified 'rowNumber'. Existing cells are moved down by the number of new rows.
     * The inserted, new rows inherits the style of the original cell at the defined row number. The inserted cells are
     * empty. The values can be set later
     *
     * <p>Remarks: Formulas / references are not adjusted</p>
     *
     * @param rowNumber       Row number below which the new row(s) will be inserted.
     * @param numberOfNewRows Number of rows to insert.
     */
    public void insertRow(int rowNumber, int numberOfNewRows) {
        // All cells below the first row must receive a new address (row + count);
        List<Cell> upperRow = this.getRow(rowNumber);

        // Identify all cells below the insertion point to adjust their addresses
        List<Cell> cellsToChange = cells.values().stream().filter(c -> c.getCellAddress2().row() > rowNumber).toList();

        // Make a copy of the cells to be moved and then delete the original cells;
        List<Cell> newCells = new ArrayList<>();
        for (Cell cell : cellsToChange) {
            int row = cell.getCellAddress2().row();
            int col = cell.getCellAddress2().column();
            Address newAddress = new Address(col, row + numberOfNewRows);
            Cell newCell = new Cell(cell.getValue(), cell.getDataType(), newAddress);
            if (cell.getCellStyle() != null) {
                newCell.setStyle(cell.getCellStyle());
            }
            newCells.add(newCell);
            cells.remove(new CellKey(col, row));
        }

        // Fill the gap with new cells, using the same style as the first row.
        for (Cell cell : upperRow) {
            for (int i = 0; i < numberOfNewRows; i++) {
                Address newAddress = new Address(cell.getCellAddress2().column(), cell.getCellAddress2().row() + 1 + i);
                Cell newCell = new Cell(null, Cell.CellType.EMPTY, newAddress);
                if (cell.getCellStyle() != null) {
                    newCell.setStyle(cell.getCellStyle());
                }
                cells.put(new CellKey(newAddress.column(), newAddress.row()), newCell);
            }
        }

        // Re-add the displaced cells with their new addresses.
        for (Cell newCell : newCells) {
            cells.put(new CellKey(newCell.getColumnNumber(), newCell.getRowNumber()), newCell);
        }
    }

    /**
     * Inserts 'count' columns right of the specified 'columnNumber'. Existing cells are moved to the right by the
     * number of new columns. The inserted, new columns inherits the style of the original cell at the defined column
     * number. The inserted cells are empty. The values can be set later
     *
     * <p>Remarks: Formulas are not adjusted</p>
     *
     * @param columnNumber       Column number right which the new column(s) will be inserted.
     * @param numberOfNewColumns Number of columns to insert.
     */
    public void insertColumn(int columnNumber, int numberOfNewColumns) {
        List<Cell> leftColumn = this.getColumn(columnNumber);
        List<Cell> cellsToChange = cells.values().stream().filter(c -> c.getCellAddress2().column() > columnNumber)
                .toList();

        List<Cell> newCells = new ArrayList<>();
        for (Cell cell : cellsToChange) {
            int row = cell.getCellAddress2().row();
            int col = cell.getCellAddress2().column();
            Address newAddress = new Address(col + numberOfNewColumns, row);
            Cell newCell = new Cell(cell.getValue(), cell.getDataType(), newAddress);
            if (cell.getCellStyle() != null) {
                newCell.setStyle(cell.getCellStyle());
            }
            newCells.add(newCell);
            cells.remove(new CellKey(col, row));
        }

        // Fill the gap with new cells, using the same style as the left column.
        for (Cell cell : leftColumn) {
            for (int i = 0; i < numberOfNewColumns; i++) {
                Address newAddress = new Address(cell.getCellAddress2().column() + 1 + i, cell.getCellAddress2().row());
                Cell newCell = new Cell(null, Cell.CellType.EMPTY, newAddress);
                if (cell.getCellStyle() != null) {
                    newCell.setStyle(cell.getCellStyle());
                }
                cells.put(new CellKey(newAddress.column(), newAddress.row()), newCell);
            }
        }

        // Re-add the displaced cells with their new addresses.
        for (Cell newCell : newCells) {
            cells.put(new CellKey(newCell.getColumnNumber(), newCell.getRowNumber()), newCell);
        }
    }

    /**
     * Searches for the first occurrence of the value.
     *
     * @param searchValue The value to search for.
     * @return The first cell containing the searched value or empty if the value was not found
     */
    public Optional<Cell> firstCellByValue(Object searchValue) {
        return cells.values().stream().filter(c -> Objects.equals(c.getValue(), searchValue)).findFirst();
    }

    /**
     * Searches for the first cell matching the predicate.
     *
     * @param predicate predicate used to test cells
     * @return the first matching cell, or empty if no cell was found
     */
    public Optional<Cell> firstOrDefaultCell(Predicate<Cell> predicate) {
        return cells.values().stream()
                .filter(c -> c != null && (c.getValue() == null || predicate.test(c)))
                .findFirst();
    }

    /**
     * Searches for cells that contain the specified value and returns a list of these cells.
     *
     * @param searchValue The value to search for.
     * @return A list of cells that contain the specified value.
     */
    public List<Cell> cellsByValue(Object searchValue) {
        return cells.values().stream().filter(c -> Objects.equals(c.getValue(), searchValue)).toList();
    }

    /**
     * Replaces all occurrences of 'oldValue' with 'newValue' and returns the number of replacements.
     *
     * @param oldValue Old value
     * @param newValue New value that should replace the old one
     * @return Count of replaced Cell values
     */
    public int replaceCellValue(Object oldValue, Object newValue) {
        int count = 0;
        List<Cell> foundCells = this.cellsByValue(oldValue);
        for (var cell : foundCells) {
            cell.setValue(newValue);
            count++;
        }
        return count;
    }

// common_methods

    /**
     * Method to add allowed actions if the worksheet is protected. If one or more values are added, UseSheetProtection
     * will be set to true
     *
     * <p>Remarks: If {@link SheetProtectionValue#SELECT_LOCKED_CELLS} is added,
     * {@link SheetProtectionValue#SELECT_UNLOCKED_CELLS} is added automatically</p>
     *
     * @param typeOfProtection Allowed action on the worksheet or cells
     */
    public void addAllowedActionOnSheetProtection(SheetProtectionValue typeOfProtection) {
        if (!sheetProtectionValues.contains(typeOfProtection)) {
            if (typeOfProtection == SheetProtectionValue.SELECT_LOCKED_CELLS &&
                    !sheetProtectionValues.contains(SheetProtectionValue.SELECT_UNLOCKED_CELLS)) {
                sheetProtectionValues.add(SheetProtectionValue.SELECT_UNLOCKED_CELLS);
            }
            sheetProtectionValues.add(typeOfProtection);
            sheetProtection = true;
        }
    }

    /**
     * Sets the defined column as hidden
     *
     * @param columnNumber Column number to hide on the worksheet
     * @throws RangeException Throws a RangeException if the passed column number is out of range
     */
    public void addHiddenColumn(int columnNumber) {
        setColumnHiddenState(columnNumber, true);
    }

    /**
     * Sets the defined column as hidden
     *
     * @param columnAddress Column address to hide on the worksheet
     * @throws RangeException Throws a RangeException if the passed column address is out of range
     */
    public void addHiddenColumn(String columnAddress) {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnHiddenState(columnNumber, true);
    }

    /// <summary>
    /// Sets the defined row as hidden
    /// </summary>
    /// <param name="rowNumber">Row number to hide on the worksheet</param>
    /// <exception cref="RangeException">Throws a RangeException if the passed row number is out of range</exception>
    public void addHiddenRow(int rowNumber) {
        setRowHiddenState(rowNumber, true);
    }

    /// <summary>
    /// Clears the active style of the worksheet. All later added calls will contain no style unless another active
    /// style is set
    /// </summary>
    public void clearActiveStyle() {
        useActiveStyle = false;
        activeStyle = null;
    }

    /**
     * Gets the cell of the specified address
     *
     * @param address Address of the cell
     * @return Cell object
     * @throws WorksheetException Trows a WorksheetException if the cell was not found on the cell table of this
     *                            worksheet
     */
    public Cell getCell(Address address) {
        CellKey key = new CellKey(address.column(), address.row());
        if (!cells.containsKey(key)) {
            throw new WorksheetException(
                    "The cell with the address " + address.getAddress() + " does not exist in this worksheet");
        }
        return cells.get(key);
    }

    /**
     * Gets the cell of the specified column and row number (zero-based)
     *
     * @param columnNumber Column number of the cell
     * @param rowNumber    Row number of the cell
     * @return Cell object
     * @throws WorksheetException Trows a WorksheetException if the cell was not found on the cell table of this
     *                            worksheet
     */
    public Cell getCell(int columnNumber, int rowNumber) {
        return getCell(new Address(columnNumber, rowNumber));
    }

    /**
     * Gets whether the specified address exists in the worksheet. Existing means that a value was stored at the
     * address
     *
     * @param address Address to check
     * @return {@code true} if the cell exists, otherwise {@code false}.
     */
    public boolean hasCell(Address address) {
        return cells.containsKey(new CellKey(address.column(), address.row()));
    }

    /**
     * Gets whether the specified address exists in the worksheet. Existing means that a value was stored at the
     * address
     *
     * @param columnNumber Column number of the cell to check (zero-based)
     * @param rowNumber    Row number of the cell to check (zero-based)
     * @return {@code true} if the cell exists, otherwise {@code false}.
     * @throws RangeException A RangeException is thrown if the column or row number is invalid
     */
    public boolean hasCell(int columnNumber, int rowNumber) {
        return hasCell(new Address(columnNumber, rowNumber));
    }

    /**
     * Resets the defined column, if existing. The corresponding instance will be removed from {@link #getColumns()}.
     *
     * <p>Remarks: If the column is inside an autoFilter-Range, the column cannot be entirely removed from
     * {@link #getColumns()}. The hidden state will be set to false and width to default, in this case.</p>
     *
     * @param columnNumber Column number to reset (zero-based)
     */
    public void resetColumn(int columnNumber) {
        if (columns.containsKey(columnNumber) &&
                !columns.get(columnNumber).hasAutoFilter()) // AutoFilters cannot have gaps
        {
            columns.remove(columnNumber);
        } else if (columns.containsKey(columnNumber)) {
            Column column = columns.get(columnNumber);
            column.setHidden(false);
            column.setWidth(DEFAULT_WORKSHEET_COLUMN_WIDTH);
        }
    }

    /**
     * Gets a row as list of cell objects
     *
     * @param rowNumber Row number (zero-based)
     * @return List of cell objects. If the row doesn't exist, an empty list is returned
     */
    public List<Cell> getRow(int rowNumber) {
        List<Cell> list = new ArrayList<>();
        for (Cell cell : cells.values()) {
            if (cell.getRowNumber() == rowNumber) {
                list.add(cell);
            }
        }
        list.sort(Comparator.comparingInt(Cell::getColumnNumber));
        return List.copyOf(list);
    }

    /**
     * Gets a column as list of cell objects
     *
     * @param columnAddress Column address
     * @return List of cell objects. If the column doesn't exist, an empty list is returned
     * @throws RangeException A range exception is thrown if the address is not valid
     */
    public List<Cell> getColumn(String columnAddress) {
        int column = Cell.resolveColumn(columnAddress);
        return getColumn(column);
    }

    /**
     * Gets a column as list of cell objects
     *
     * @param columnNumber Column number (zero-based)
     * @return List of cell objects. If the column doesn't exist, an empty list is returned
     */
    public List<Cell> getColumn(int columnNumber) {
        List<Cell> list = new ArrayList<>();
        for (Cell cell : cells.values()) {
            if (cell.getColumnNumber() == columnNumber) {
                list.add(cell);
            }
        }
        list.sort(Comparator.comparingInt(Cell::getRowNumber));
        return List.copyOf(list);
    }

    /**
     * Moves the current position to the next column
     */
    public void goToNextColumn() {
        currentColumnNumber++;
        currentRowNumber = 0;
        Cell.validateColumnNumber(currentColumnNumber);
    }

    /**
     * Moves the current position to the next column with the number of cells to move, moving to the first row number
     * (0)
     *
     * <p>Remarks: The value can also be negative. However, resulting column numbers below 0 or above 16383 will cause
     * an exception</p>
     *
     * @param numberOfColumns Number of columns to move
     */
    public void goToNextColumn(int numberOfColumns) {
        goToNextColumn(numberOfColumns, false);
    }

    /**
     * Moves the current position to the next column with the number of cells to move
     *
     * <p>Remarks: The value can also be negative. However, resulting column numbers below 0 or above 16383 will cause
     * an exception</p>
     *
     * @param numberOfColumns Number of columns to move
     * @param keepRowPosition If true, the row position is preserved, otherwise set to 0
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
        currentRowNumber++;
        currentColumnNumber = 0;
        Cell.validateRowNumber(currentRowNumber);
    }

    /**
     * Moves the current position to the next row with the number of cells to move (use for a new line), moving to the
     * first column number (0)
     *
     * <p>Remarks: The value can also be negative. However, resulting row numbers below 0 or above 1048575 will cause
     * an exception</p>
     *
     * @param numberOfRows Number of rows to move
     */
    public void goToNextRow(int numberOfRows) {
        goToNextRow(numberOfRows, false);
    }

    /**
     * Moves the current position to the next row with the number of cells to move (use for a new line)
     *
     * <p>Remarks: The value can also be negative. However, resulting row numbers below 0 or above 1048575 will cause
     * an exception</p>
     *
     * @param numberOfRows       Number of rows to move
     * @param keepColumnPosition If true, the column position is preserved, otherwise set to 0
     */
    public void goToNextRow(int numberOfRows, boolean keepColumnPosition) {
        currentRowNumber += numberOfRows;
        if (!keepColumnPosition) {
            currentColumnNumber = 0;
        }
        Cell.validateRowNumber(currentRowNumber);
    }

    /**
     * Merges the defined cell range
     *
     * @param cellRange Range to merge
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     * @throws RangeException Throws a RangeException if the passed cell range is out of range
     */
    public String mergeCells(Range cellRange) {
        return mergeCells(cellRange.startAddress(), cellRange.endAddress());
    }

    /**
     * Merges the defined cell range
     *
     * @param cellRange Range to merge (e.g. 'A1:B12')
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     * @throws RangeException  Throws a RangeException if the passed cell range is out of range
     * @throws FormatException Throws a FormatException if the passed cell range is malformed
     */
    public String MergeCells(String cellRange) {
        Range range = Cell.resolveCellRange(cellRange);
        return mergeCells(range.startAddress(), range.endAddress());
    }

    /**
     * Merges the defined cell range
     *
     * @param startAddress Start address of the merged cell range
     * @param endAddress   End address of the merged cell range
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     * @throws RangeException Throws a RangeException if one of the passed cell addresses is out of range or if one or
     *                        more cell addresses are already occupied in another merge range
     */
    public String mergeCells(Address startAddress, Address endAddress) {
        String key = startAddress + ":" + endAddress;
        Range value = new Range(startAddress, endAddress);
        List<Address> result = value.resolveEnclosedAddresses();
        for (Map.Entry<String, Range> item : mergedCells.entrySet()) {
            if (item.getValue().resolveEnclosedAddresses().stream().anyMatch(result::contains)) {
                throw new RangeException(
                        "The passed range: " + value + " contains cells that are already in the defined merge range: " +
                                item.getKey());
            }
        }
        mergedCells.put(key, value);
        return key;
    }

    /**
     * Method to recalculate the auto filter (columns) of this worksheet. This is an internal method. There is no need
     * to use it
     */
    void recalculateAutoFilter() {
        if (autoFilterRange == null) {
            return;
        }
        int start = autoFilterRange.startAddress().column();
        int end = autoFilterRange.endAddress().column();
        int endRow = 0;
        for (Cell item : cellValues) {
            if (item.getColumnNumber() < start || item.getColumnNumber() > end) {
                continue;
            }
            if (item.getRowNumber() > endRow) {
                endRow = item.getRowNumber();
            }
        }
        Column c;
        for (int i = start; i <= end; i++) {
            if (!columns.containsKey(i)) {
                c = new Column(i);
                c.setAutoFilter(true);
                columns.put(i, c);
            } else {
                columns.get(i).setAutoFilter(true);
            }
        }
        autoFilterRange = new Range(start, 0, end, endRow);
    }

    /**
     * Method to recalculate the collection of columns of this worksheet. This is an internal method. There is no need
     * to use it
     */
    void recalculateColumns() {
        List<Integer> columnsToDelete = new ArrayList<>();
        for (Map.Entry<Integer, Column> col : columns.entrySet()) {
            if (!col.getValue().hasAutoFilter() && !col.getValue().isHidden() &&
                    Comparators.compareDimensions(col.getValue().getWidth(), DEFAULT_WORKSHEET_COLUMN_WIDTH) == 0 &&
                    col.getValue().getDefaultColumnStyle() == null) {
                columnsToDelete.add(col.getKey());
            }
        }
        for (int index : columnsToDelete) {
            columns.remove(index);
        }
    }

    /**
     * Method to resolve all merged cells of the worksheet. Only the value of the very first cell of the locked cells
     * range will be visible. The other values are still present (set to EMPTY) but will not be stored in the
     * worksheet.<br /> This is an internal method. There is no need to use it
     *
     * @throws StyleException Throws a StyleException if one of the styles of the merged cells cannot be referenced or
     *                        is null
     */
    void resolveMergedCells() {
        Style mergeStyle = BasicStyles.getMergeCellStyle();
        Cell cell;
        for (Map.Entry<String, Range> range : mergedCells.entrySet()) {
            int pos = 0;
            List<Address> addresses = Cell.getCellRange(range.getValue().startAddress(), range.getValue().endAddress());
            for (Address address : addresses) {
                CellKey key = new CellKey(address.column(), address.row());
                if (!cells.containsKey(key)) {
                    cell = new Cell();
                    cell.setDataType(Cell.CellType.EMPTY);
                    cell.setRowNumber(address.row());
                    cell.setColumnNumber(address.column());
                    addCell(cell, cell.getColumnNumber(), cell.getRowNumber());
                }
                if (pos != 0) {
                    cell = cells.get(key);
                    cell.setDataType(Cell.CellType.EMPTY);
                    if (cell.getCellStyle() == null) {
                        cell.setStyle(mergeStyle);
                    } else {
                        Style mixedMergeStyle = cell.getCellStyle();
                        // TODO: There should be a better possibility to identify particular style elements that
                        //  deviates
                        mixedMergeStyle.getCurrentCellXf()
                                .setForceApplyAlignment(mergeStyle.getCurrentCellXf().isForceApplyAlignment());
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
        autoFilterRange = null;
    }

    /**
     * Sets a previously defined, hidden column as visible again
     *
     * @param columnNumber Column number to make visible again
     * @throws RangeException Throws a RangeException if the passed column number is out of range
     */
    public void removeHiddenColumn(int columnNumber) {
        setColumnHiddenState(columnNumber, false);
    }

    /**
     * Sets a previously defined, hidden column as visible again
     *
     * @param columnAddress Column address to make visible again
     * @throws RangeException Throws a RangeException if the column address out of range
     */
    public void removeHiddenColumn(String columnAddress) {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnHiddenState(columnNumber, false);
    }

    /**
     * Sets a previously defined, hidden row as visible again
     *
     * @param rowNumber Row number to hide on the worksheet
     * @throws RangeException Throws a RangeException if the passed row number is out of range
     */
    public void removeHiddenRow(int rowNumber) {
        setRowHiddenState(rowNumber, false);
    }

    /**
     * Removes the defined merged cell range
     *
     * @param range Cell range to remove the merging
     * @throws RangeException Throws a RangeException if the passed cell range was not merged earlier
     */
    public void removeMergedCells(String range) {
        range = ParserUtils.toUpper(range);
        if (range == null || !mergedCells.containsKey(range)) {
            throw new RangeException("The cell range " + range + " was not found in the list of merged cell ranges");
        }

        List<Address> addresses = Cell.getCellRange(range);
        for (Address address : addresses) {
            CellKey key = new CellKey(address.column(), address.row());
            if (cells.containsKey(key)) {
                Cell cell = cells.get(key);
                if (BasicStyles.getMergeCellStyle().equals(cell.getCellStyle())) {
                    cell.removeStyle();
                }
                cell.resolveCellType(); // resets the type
            }
        }
        mergedCells.remove(range);
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
     * Sets the active style of the worksheet. This style will be assigned to all later added cells
     *
     * @param style Style to set as active style
     */
    public void setActiveStyle(Style style) {
        if (style == null) {
            useActiveStyle = false;
        } else {
            useActiveStyle = true;
        }
        activeStyle = style;
    }

    /**
     * Sets the column auto filter within the defined column range
     *
     * @param startColumn Column number with the first appearance of an auto filter drop down
     * @param endColumn   Column number with the last appearance of an auto filter drop down
     * @throws RangeException Throws a RangeException if the start or end address out of range
     */
    public void setAutoFilter(int startColumn, int endColumn) {
        String start = Cell.resolveCellAddress(startColumn, 0);
        String end = Cell.resolveCellAddress(endColumn, 0);
        if (endColumn < startColumn) {
            setAutoFilter(end + ":" + start);
        } else {
            setAutoFilter(start + ":" + end);
        }
    }

    /**
     * Sets the column auto filter within the defined column range
     *
     * @param range Range to apply auto filter on. The range could be 'A1:C10' for instance. The end row will be
     *              recalculated automatically when saving the file
     * @throws RangeException  Throws a RangeException if the passed range out of range
     * @throws FormatException Throws a FormatException if the passed range is malformed
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
     * @throws RangeException Throws a RangeException if the column number out of range
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
                Comparators.compareDimensions(columns.get(columnNumber).getWidth(), DEFAULT_WORKSHEET_COLUMN_WIDTH) ==
                        0 && !columns.get(columnNumber).hasAutoFilter()) {
            columns.remove(columnNumber);
        }
    }

    /**
     * Sets the width of the passed column address
     *
     * @param columnAddress Column address (A - XFD)
     * @param width         Width from 0 to 255.0
     * @throws RangeException Throws a RangeException:<br />a) If the passed column address is out of range<br />b) if
     *                        the column width is out of range (0 - 255.0)
     */
    public void setColumnWidth(String columnAddress, float width) {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnWidth(columnNumber, width);
    }

    /**
     * Sets the width of the passed column number (zero-based)
     *
     * @param columnNumber Column number (zero-based, from 0 to 16383)
     * @param width        Width from 0 to 255.0
     * @throws RangeException Throws a RangeException:<br />a) If the passed column number is out of range<br />b) if
     *                        the column width is out of range (0 - 255.0)
     */
    public void setColumnWidth(int columnNumber, float width) {
        Cell.validateColumnNumber(columnNumber);
        if (width < MIN_COLUMN_WIDTH || width > MAX_COLUMN_WIDTH) {
            throw new RangeException(
                    "The column width (" + width + ") is out of range. Range is from " + MIN_COLUMN_WIDTH + " to " +
                            MAX_COLUMN_WIDTH + " (chars).");
        }
        if (columns.containsKey(columnNumber)) {
            columns.get(columnNumber).setWidth(width);
        } else {
            Column c = new Column(columnNumber);
            c.setWidth(width);
            columns.put(columnNumber, c);
        }
    }

    /**
     * Sets the default column style of the passed column address
     *
     * @param columnAddress Column address (A - XFD)
     * @param style         Style to set as default. If null, the style is cleared
     * @return Assigned style or null if cleared
     * @throws RangeException Throws a RangeException:<br />a) If the passed column address is out of range<br />b) if
     *                        the column width is out of range (0 - 255.0)
     */
    public Style setColumnDefaultStyle(String columnAddress, Style style) {
        int columnNumber = Cell.resolveColumn(columnAddress);
        return setColumnDefaultStyle(columnNumber, style);
    }

    /**
     * Sets the default column style of the passed column number (zero-based)
     *
     * @param columnNumber Column number (zero-based, from 0 to 16383)
     * @param style        Style to set as default. If null, the style is cleared
     * @return Assigned style or null if cleared
     * @throws RangeException Throws a RangeException:<br />a) If the passed column number is out of range<br />b) if
     *                        the column width is out of range (0 - 255.0)
     */
    public Style setColumnDefaultStyle(int columnNumber, Style style) {
        Cell.validateColumnNumber(columnNumber);
        if (this.columns.containsKey(columnNumber)) {
            return columns.get(columnNumber).setDefaultColumnStyle(style);
        } else {
            Column c = new Column(columnNumber);
            Style returnStyle = c.setDefaultColumnStyle(style);
            this.columns.put(columnNumber, c);
            return returnStyle;
        }
    }

    /**
     * Set the current cell address
     *
     * @param columnNumber Column number (zero based)
     * @param rowNumber    Row number (zero based)
     * @throws RangeException Throws a RangeException if one of the passed cell addresses is out of range
     */
    public void setCurrentCellAddress(int columnNumber, int rowNumber) {
        setCurrentColumnNumber(columnNumber);
        setCurrentRowNumber(rowNumber);
    }

    /**
     * Set the current cell address
     *
     * @param address Cell address in the format A1 - XFD1048576
     * @throws RangeException  Throws a RangeException if the passed cell address is out of range
     * @throws FormatException Throws a FormatException if the passed cell address is malformed
     */
    public void setCurrentCellAddress(String address) {
        Address addr = Cell.resolveCellCoordinate(address);
        setCurrentCellAddress(addr.column(), addr.row());
    }

    /**
     * Sets the current column number (zero based)
     *
     * @param columnNumber Column number (zero based)
     * @throws RangeException Throws a RangeException if the number is out of the valid range. Range is from 0 to 16383
     *                        (16384 columns)
     */
    public void setCurrentColumnNumber(int columnNumber) {
        Cell.validateColumnNumber(columnNumber);
        currentColumnNumber = columnNumber;
    }

    /**
     * Sets the current row number (zero based)
     *
     * @param rowNumber Row number (zero based)
     * @throws RangeException Throws a RangeException if the number is out of the valid range. Range is from 0 to
     *                        1048575 (1048576 rows)
     */
    public void setCurrentRowNumber(int rowNumber) {
        Cell.validateRowNumber(rowNumber);
        currentRowNumber = rowNumber;
    }

    /**
     * Adds a range to the selected cells on this worksheet
     *
     * @param range Cell range to add
     */
    public void addSelectedCells(Range range) {
        selectedCells = DataUtils.mergeRange(selectedCells, range);
    }

    /**
     * Adds a range to the selected cells on this worksheet
     *
     * @param startAddress Start address of the range
     * @param endAddress   End address of the range
     */
    public void addSelectedCells(Address startAddress, Address endAddress) {
        addSelectedCells(new Range(startAddress, endAddress));
    }

    /**
     * Adds a range or cell address to the selected cells on this worksheet
     *
     * @param rangeOrAddress Cell range or address to add
     */
    public void addSelectedCells(String rangeOrAddress) {
        Optional<Range> resolved = parseRange(rangeOrAddress);
        resolved.ifPresent(this::addSelectedCells);
    }

    /**
     * Adds a single cell address to the selected cells on this worksheet
     *
     * @param address Cell address to add
     */
    public void addSelectedCells(Address address) {
        addSelectedCells(new Range(address, address));
    }

    /**
     * Removes all cell selections of this worksheet
     */
    public void clearSelectedCells() {
        selectedCells.clear();
    }

    /**
     * Removes the given range from the selected cell ranges of this worksheet, if existing. If the passed range is
     * overlapping the ranges of the selected cells, only the intersecting addresses will be removed
     *
     * @param range Range to remove
     */
    public void removeSelectedCells(Range range) {
        selectedCells = DataUtils.subtractRange(selectedCells, range);
    }

    /**
     * Removes the given range or cell address from the selected cell ranges of this worksheet, if existing
     *
     * @param rangeOrAddress Range or cell address to remove
     */
    public void removeSelectedCells(String rangeOrAddress) {
        Optional<Range> resolved = parseRange(rangeOrAddress);
        resolved.ifPresent(this::removeSelectedCells);
    }

    /**
     * Removes the given address from the selected cell ranges of this worksheet, if existing
     *
     * @param address Address of the range to remove
     */
    public void removeSelectedCells(Address address) {
        removeSelectedCells(new Range(address, address));
    }

    /// <summary>
    /// Removes the given range from the selected cell ranges of this worksheet, if existing
    /// </summary>
    /// <param name="startAddress">Start address of the range to remove</param>
    /// <param name="endAddress">End address of the range to remove</param>
    public void removeSelectedCells(Address startAddress, Address endAddress) {
        removeSelectedCells(new Range(startAddress, endAddress));
    }

    /**
     * Sets or removes the password for worksheet protection. If set, UseSheetProtection will be also set to true
     *
     * @param password Password (UTF-8) to protect the worksheet. If the password is null or empty, no password will be
     *                 used
     */
    public void setSheetProtectionPassword(String password) {
        if (ParserUtils.isNullOrEmpty(password)) {
            sheetProtectionPassword.unsetPassword();
            sheetProtection = false;
        } else {
            sheetProtectionPassword.setPassword(password);
            sheetProtection = true;
        }
    }

    /**
     * Sets the height of the passed row number (zero-based)
     *
     * @param rowNumber Row number (zero-based, 0 to 1048575)
     * @param height    Height from 0 to 409.5
     * @throws RangeException Throws a RangeException:<br />a) If the passed row number is out of range<br />b) if the
     *                        row height is out of range (0 - 409.5)
     */
    public void setRowHeight(int rowNumber, float height) {
        Cell.validateRowNumber(rowNumber);
        if (height < MIN_ROW_HEIGHT || height > MAX_ROW_HEIGHT) {
            throw new RangeException(
                    "The row height (" + height + ") is out of range. Range is from " + MIN_ROW_HEIGHT + " to " +
                            MAX_ROW_HEIGHT + " (equals 546px).");
        }
        rowHeights.put(rowNumber, height);
    }

    /**
     * Sets the defined row as hidden or visible
     *
     * @param rowNumber Row number to make visible again
     * @param state     If true, the row will be hidden, otherwise visible
     * @throws RangeException Throws a RangeException if the passed row number was out of range
     */
    private void setRowHiddenState(int rowNumber, boolean state) {
        Cell.validateRowNumber(rowNumber);
        if (hiddenRows.containsKey(rowNumber)) {
            if (state) {
                hiddenRows.put(rowNumber, true);
            } else {
                hiddenRows.remove(rowNumber);
            }
        } else if (state) {
            hiddenRows.put(rowNumber, true);
        }
    }

    /**
     * Sets the name of the worksheet with the option of name sanitation
     *
     * @param name     Name of the worksheet
     * @param sanitize If true, the filename will be sanitized automatically according to the specifications of Excel
     * @throws WorksheetException Thrown if no workbook is referenced. This information is necessary to determine
     *                            whether the name already exists
     */
    public void setSheetName(String name, boolean sanitize) {
        if (sanitize) {
            sheetName = ""; // Empty name (temporary) to prevent conflicts during sanitizing
            sheetName = sanitizeWorksheetName(name, workbookReference);
        } else {
            setSheetName(name);
        }
    }

    /**
     * Sets the horizontal split of the worksheet into two panes. The measurement in characters cannot be used to freeze
     * panes
     *
     * @param topPaneHeight Height (similar to row height) from top of the worksheet to the split line in characters
     * @param topLeftCell   Top Left cell address of the bottom right pane (if applicable). Only the row component is
     *                      important in a horizontal split
     */
    public void setHorizontalSplit(float topPaneHeight, Address topLeftCell) {
        setHorizontalSplit(topPaneHeight, topLeftCell, null);
    }

    /**
     * Sets the horizontal split and its active pane.
     *
     * @param topPaneHeight Height from the top of the worksheet to the split line in characters
     * @param topLeftCell   Top left cell address of the bottom right pane
     * @param activePane    Active pane, or {@code null} if no active pane is defined
     */
    public void setHorizontalSplit(float topPaneHeight, Address topLeftCell, WorksheetPane activePane) {
        setSplit(null, topPaneHeight, topLeftCell, activePane);
    }

    /**
     * Sets the horizontal split of the worksheet into two panes. The measurement in rows can be used to split and
     * freeze panes
     *
     * @param numberOfRowsFromTop Number of rows from top of the worksheet to the split line. The particular row heights
     *                            are considered
     * @param freeze              If true, all panes are frozen, otherwise remains movable
     * @param topLeftCell         Top Left cell address of the bottom right pane (if applicable). Only the row component
     *                            is important in a horizontal split
     * @throws WorksheetException WorksheetException Thrown if the row number of the top left cell is smaller the split
     *                            panes number of rows from top, if freeze is applied
     */
    public void setHorizontalSplit(int numberOfRowsFromTop, boolean freeze, Address topLeftCell) {
        setHorizontalSplit(numberOfRowsFromTop, freeze, topLeftCell, null);
    }

    /**
     * Sets the horizontal split and its active pane using a number of rows.
     *
     * @param numberOfRowsFromTop Number of rows from the top of the worksheet to the split line
     * @param freeze              If true, all panes are frozen; otherwise they remain movable
     * @param topLeftCell         Top left cell address of the bottom right pane
     * @param activePane          Active pane, or {@code null} if no active pane is defined
     * @throws WorksheetException If the top-left cell is above the split when freezing is applied
     */
    public void setHorizontalSplit(
            int numberOfRowsFromTop, boolean freeze, Address topLeftCell, WorksheetPane activePane) {
        setSplit(null, numberOfRowsFromTop, freeze, topLeftCell, activePane);
    }

    /**
     * Sets the vertical split of the worksheet into two panes. The measurement in characters cannot be used to freeze
     * panes
     *
     * @param leftPaneWidth Width (similar to column width) from left of the worksheet to the split line in characters
     * @param topLeftCell   Top Left cell address of the bottom right pane (if applicable). Only the column component is
     *                      important in a vertical split
     */
    public void setVerticalSplit(float leftPaneWidth, Address topLeftCell) {
        setVerticalSplit(leftPaneWidth, topLeftCell, null);
    }

    /**
     * Sets the vertical split and its active pane.
     *
     * @param leftPaneWidth Width from the left of the worksheet to the split line in characters
     * @param topLeftCell   Top left cell address of the bottom right pane
     * @param activePane    Active pane, or {@code null} if no active pane is defined
     */
    public void setVerticalSplit(float leftPaneWidth, Address topLeftCell, WorksheetPane activePane) {
        setSplit(leftPaneWidth, null, topLeftCell, activePane);
    }

    /**
     * Sets the vertical split of the worksheet into two panes. The measurement in columns can be used to split and
     * freeze panes
     *
     * @param numberOfColumnsFromLeft Number of columns from left of the worksheet to the split line. The particular
     *                                column widths are considered
     * @param freeze                  If true, all panes are frozen, otherwise remains movable
     * @param topLeftCell             Top Left cell address of the bottom right pane (if applicable). Only the column
     *                                component is important in a vertical split
     * @throws WorksheetException WorksheetException Thrown if the column number of the top left cell is smaller the
     *                            split panes number of columns from left, if freeze is applied
     */
    public void setVerticalSplit(int numberOfColumnsFromLeft, boolean freeze, Address topLeftCell) {
        setVerticalSplit(numberOfColumnsFromLeft, freeze, topLeftCell, null);
    }

    /**
     * Sets the vertical split and its active pane using a number of columns.
     *
     * @param numberOfColumnsFromLeft Number of columns from the left of the worksheet to the split line
     * @param freeze                  If true, all panes are frozen; otherwise they remain movable
     * @param topLeftCell             Top left cell address of the bottom right pane
     * @param activePane              Active pane, or {@code null} if no active pane is defined
     * @throws WorksheetException If the top-left cell is left of the split when freezing is applied
     */
    public void setVerticalSplit(
            int numberOfColumnsFromLeft, boolean freeze, Address topLeftCell, WorksheetPane activePane) {
        setSplit(numberOfColumnsFromLeft, null, freeze, topLeftCell, activePane);
    }

    /**
     * Sets the horizontal and vertical split of the worksheet into four panes using rows and columns.
     *
     * @param numberOfColumnsFromLeft Number of columns from the left, or {@code null} for no vertical split
     * @param numberOfRowsFromTop     Number of rows from the top, or {@code null} for no horizontal split
     * @param freeze                  If true, all panes are frozen; otherwise they remain movable
     * @param topLeftCell             Top left cell address of the bottom right pane
     * @throws WorksheetException If the top-left cell precedes the split when freezing is applied
     */
    public void setSplit(
            Integer numberOfColumnsFromLeft, Integer numberOfRowsFromTop, boolean freeze, Address topLeftCell
    ) {
        setSplit(numberOfColumnsFromLeft, numberOfRowsFromTop, freeze, topLeftCell, null);
    }

    /**
     * Sets the horizontal and vertical split of the worksheet into four panes using rows and columns.
     *
     * @param numberOfColumnsFromLeft Number of columns from the left, or {@code null} for no vertical split
     * @param numberOfRowsFromTop     Number of rows from the top, or {@code null} for no horizontal split
     * @param freeze                  If true, all panes are frozen; otherwise they remain movable
     * @param topLeftCell             Top left cell address of the bottom right pane
     * @param activePane              Active pane, or {@code null} if no active pane is defined
     * @throws WorksheetException If the top-left cell precedes the split when freezing is applied
     */
    public void setSplit(
            Integer numberOfColumnsFromLeft, Integer numberOfRowsFromTop, boolean freeze,
            Address topLeftCell, WorksheetPane activePane
    ) {
        Objects.requireNonNull(topLeftCell, "topLeftCell");
        if (freeze) {
            if (numberOfColumnsFromLeft != null && topLeftCell.column() < numberOfColumnsFromLeft) {
                throw new WorksheetException("The column number " + topLeftCell.column() +
                        " is not valid for a frozen, vertical split with the split pane column number " +
                        numberOfColumnsFromLeft);
            }
            if (numberOfRowsFromTop != null && topLeftCell.row() < numberOfRowsFromTop) {
                throw new WorksheetException("The row number " + topLeftCell.row() +
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
     * Sets the horizontal and vertical split of the worksheet into four panes. The measurement in characters cannot be
     * used to freeze panes
     *
     * @param leftPaneWidth Width (similar to column width) from left of the worksheet to the split line in
     *                      characters.<br /> The parameter is optional. If left null, the method acts identical to
     *                      {@link #setHorizontalSplit(float, Address, WorksheetPane)}
     * @param topPaneHeight Height (similar to row height) from top of the worksheet to the split line in characters.<br
     *                      /> The parameter is optional. If left null, the method acts identical to
     *                      {@link #setVerticalSplit(float, Address, WorksheetPane)}
     * @param topLeftCell   Top Left cell address of the bottom right pane (if applicable)
     */
    public void setSplit(Float leftPaneWidth, Float topPaneHeight, Address topLeftCell) {
        setSplit(leftPaneWidth, topPaneHeight, topLeftCell, null);
    }

    /**
     * Sets the horizontal and vertical split of the worksheet into four movable panes using character measurements.
     *
     * @param leftPaneWidth Width from the left, or {@code null} for no vertical split
     * @param topPaneHeight Height from the top, or {@code null} for no horizontal split
     * @param topLeftCell   Top left cell address of the bottom right pane
     * @param activePane    Active pane, or {@code null} if no active pane is defined
     */
    public void setSplit(
            Float leftPaneWidth, Float topPaneHeight, Address topLeftCell, WorksheetPane activePane
    ) {
        Objects.requireNonNull(topLeftCell, "topLeftCell");
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
     * Creates a (dereferenced) deep copy of this worksheet
     *
     * <p>Remarks: Not considered in the copy are the internal ID, the worksheet name and the workbook reference.
     * Since styles are managed in a shared repository, no dereferencing is applied (Styles are not deep-copied). Use
     * {@link Workbook#copyWorksheetTo(Worksheet, String, Workbook, boolean)} or
     * {@link Workbook#copyWorksheetIntoThis(Worksheet, String, boolean)} to add a copy of worksheet to a workbook.
     * These methods will set the internal ID, name and workbook reference.
     * </p>
     *
     * @return Copy of this worksheet
     */
    public Worksheet copy() {
        Worksheet copy = new Worksheet();
        for (Cell cell : this.cells.values()) {
            copy.addCell(cell.copy(), cell.getColumnNumber(), cell.getRowNumber());
        }
        copy.activePane = this.activePane;
        copy.activeStyle = this.activeStyle;
        if (this.autoFilterRange != null) {
            copy.autoFilterRange = this.autoFilterRange.copy();
        }
        for (Map.Entry<Integer, Column> column : this.getColumns().entrySet()) {
            copy.columns.put(column.getKey(), column.getValue().copy());
        }
        copy.currentCellDirection = this.currentCellDirection;
        copy.currentColumnNumber = this.currentColumnNumber;
        copy.currentRowNumber = this.currentRowNumber;
        copy.defaultColumnWidth = this.defaultColumnWidth;
        copy.defaultRowHeight = this.defaultRowHeight;
        copy.freezeSplitPanes = this.freezeSplitPanes;
        copy.hidden = this.hidden;
        copy.hiddenRows.putAll(this.hiddenRows);
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
        copy.rowHeights.putAll(this.rowHeights);
        for (Range range : selectedCells) {
            copy.addSelectedCells(range);
        }
        copy.sheetProtectionPassword.copyFrom(this.sheetProtectionPassword);
        copy.sheetProtectionValues.addAll(this.sheetProtectionValues);
        copy.useActiveStyle = this.useActiveStyle;
        copy.sheetProtection = this.sheetProtection;
        copy.showGridLines = this.showGridLines;
        copy.showRowColumnHeaders = this.showRowColumnHeaders;
        copy.showRuler = this.showRuler;
        copy.viewType = this.viewType;
        copy.zoomFactors.clear();
        for (Map.Entry<SheetViewType, Integer> zoomFactor : this.zoomFactors.entrySet()) {
            copy.setZoomFactor(zoomFactor.getKey(), zoomFactor.getValue());
        }
        return copy;
    }

    /**
     * Sets a zoom factor for a given {@link SheetViewType}. If {@link #AUTO_ZOOM_FACTOR}, the zoom factor is set to
     * automatic
     *
     * <p>Remarks: This factor is not the currently set factor. use the setter {@link #setZoomFactor(int)}  to set the
     * factor for the current {@link #getViewType()}</p>
     *
     * @param sheetViewType Sheet view type to apply the zoom factor on
     * @param zoomFactor    Zoom factor in percent
     * @throws WorksheetException Throws a WorksheetException if the zoom factor is not {@link #AUTO_ZOOM_FACTOR} or
     *                            below {@link #MIN_ZOOM_FACTOR} or above {@link #MAX_ZOOM_FACTOR}
     */
    public void setZoomFactor(SheetViewType sheetViewType, int zoomFactor) {
        if (zoomFactor != AUTO_ZOOM_FACTOR && (zoomFactor < MIN_ZOOM_FACTOR || zoomFactor > MAX_ZOOM_FACTOR)) {
            throw new WorksheetException(
                    "The zoom factor " + zoomFactor + " is not valid. Valid are values between " + MIN_ZOOM_FACTOR +
                            " and " + MAX_ZOOM_FACTOR + ", or " + AUTO_ZOOM_FACTOR + " (automatic)");
        }
        this.zoomFactors.put(sheetViewType, zoomFactor);
    }

// static methods

    /**
     * Sanitizes a worksheet name
     *
     * @param input    Name to sanitize
     * @param workbook Workbook reference
     * @return Name of the sanitized worksheet
     * @throws WorksheetException Thrown if the workbook reference is null, since all worksheets have to be considered
     *                            during sanitation
     */
    public static String sanitizeWorksheetName(String input, Workbook workbook) {
        if (ParserUtils.isNullOrEmpty(input)) {
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
     * Parses a string to a range. If the string is a single address, the range consists of this as start and end
     * address
     *
     * @param rangeOrAddress Range or address expression
     * @return Range or null if the
     */
    private static Optional<Range> parseRange(String rangeOrAddress) {
        if (ParserUtils.isNullOrEmpty(rangeOrAddress)) {
            return Optional.empty();
        }
        Range range;
        if (rangeOrAddress.contains(":")) {
            range = Cell.resolveCellRange(rangeOrAddress);
        } else {
            Address address = Cell.resolveCellCoordinate(rangeOrAddress);
            range = new Range(address, address);
        }
        return Optional.of(range);
    }

    /**
     * Determines the next unused worksheet name in the passed workbook
     *
     * <p>Remarks: The 'rare' case where 10^31 Worksheets exists (leads to a crash) is deliberately not handled,
     * since such a number of sheets would consume at least one quintillion bytes of RAM... what is vastly out of the 64
     * bit range
     * </p>
     *
     * @param name     Original name to start the check
     * @param workbook Workbook to look for existing worksheets
     * @return Not yet used worksheet name
     * @throws WorksheetException Thrown if the workbook reference is null, since all worksheets have to be considered
     *                            during sanitation
     */
    private static String getUnusedWorksheetName(String name, Workbook workbook) {
        if (workbook == null) {
            throw new WorksheetException("The workbook reference is null");
        }
        if (!worksheetExists(name, workbook)) {
            return name;
        }
        Pattern pattern = Pattern.compile("^(.*?)(\\d{1,31})$");
        Matcher match = pattern.matcher(name);
        String prefix = name;
        int number = 1;
        if (match.matches()) {
            prefix = match.group(1);
            number = ParserUtils.tryParseInt(match.group(2)).orElse(0);
            // Match C# int.TryParse: an unrepresentable suffix starts at zero.
        }
        while (true) {
            String numberString = ParserUtils.toString(number);
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
            if (name.equals(workbook.getWorksheets().get(i).getSheetName())) {
                return true;
            }
        }
        return false;
    }

    /**
     * Converts a SheetProtectionValue enum to its string representation
     */
    static String getSheetProtectionName(SheetProtectionValue protection) {
        return switch (protection) {
            case SheetProtectionValue.OBJECTS -> "objects";
            case SheetProtectionValue.SCENARIOS -> "scenarios";
            case SheetProtectionValue.FORMAT_CELLS -> "formatCells";
            case SheetProtectionValue.FORMAT_COLUMNS -> "formatColumns";
            case SheetProtectionValue.FORMAT_ROWS -> "formatRows";
            case SheetProtectionValue.INSERT_COLUMNS -> "insertColumns";
            case SheetProtectionValue.INSERT_ROWS -> "insertRows";
            case SheetProtectionValue.INSERT_HYPERLINKS -> "insertHyperlinks";
            case SheetProtectionValue.DELETE_COLUMNS -> "deleteColumns";
            case SheetProtectionValue.DELETE_ROWS -> "deleteRows";
            case SheetProtectionValue.SELECT_LOCKED_CELLS -> "selectLockedCells";
            case SheetProtectionValue.SORT -> "sort";
            case SheetProtectionValue.AUTO_FILTER -> "autoFilter";
            case SheetProtectionValue.PIVOT_TABLES -> "pivotTables";
            case SheetProtectionValue.SELECT_UNLOCKED_CELLS -> "selectUnlockedCells";
        };
    }

    /**
     * Converts a string representation of a worksheet pane to its enum value
     *
     * @param pane String value
     * @return Enum value or null if not matching
     */
    static Optional<WorksheetPane> getWorksheetPaneEnum(String pane) {
        WorksheetPane output = switch (pane) {
            case "topLeft" -> WorksheetPane.TOP_LEFT;
            case "topRight" -> WorksheetPane.TOP_RIGHT;
            case "bottomLeft" -> WorksheetPane.BOTTOM_LEFT;
            case "bottomRight" -> WorksheetPane.BOTTOM_RIGHT;
            default -> null;
        };
        if (output == null) {
            return Optional.empty();
        }
        return Optional.of(output);
    }

    /**
     * Converts a string representation of a sheet view type to its enum value
     *
     * @param viewType String value
     * @return Enum value
     */
    static SheetViewType getSheetViewTypeEnum(String viewType) {
        SheetViewType output = SheetViewType.NORMAL;
        output = switch (viewType) {
            case "pageBreakPreview" -> SheetViewType.PAGE_BREAK_PREVIEW;
            case "pageLayout" -> SheetViewType.PAGE_LAYOUT;
            default -> output;
        };
        return output;
    }
}
