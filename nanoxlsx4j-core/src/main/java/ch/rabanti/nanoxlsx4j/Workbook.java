/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.IntStream;

import ch.rabanti.nanoxlsx4j.colors.Color;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.internal.AuxiliaryData;
import ch.rabanti.nanoxlsx4j.internal.FeatureSet;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Password;
import ch.rabanti.nanoxlsx4j.themes.Theme;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;
import ch.rabanti.nanoxlsx4j.utils.Validators;

public class Workbook {

    // TODO check whether PloginLoadr.initialize() is necessary here (see C# implementation)

// privateFields

    private String filename;
    private List<Worksheet> worksheets;
    private Worksheet currentWorksheet;
    private Metadata workbookMetadata;
    private Password workbookProtectionPassword;
    private boolean lockWindowsIfProtected;
    private boolean lockStructureIfProtected;
    private int selectedWorksheet;
    private Shortener shortener;
    private final List<Color> mruColors = new ArrayList<>();
    private final List<DefinedName> definedNames = new ArrayList<>();
    private AuxiliaryData auxiliaryData;
    private boolean useWorkbookProtection;
    private boolean hidden;
    private Theme workbookTheme = Theme.getDefaultTheme();
    private FeatureSet features = new FeatureSet();

    // TODO define where this is in Java (no internal as in C#)
    boolean importInProgress; // Used by NanoXLSX.Reader

// Getters&Setters

    /**
     * Optional auxiliary data object. This object is used to store additional information about the workbook. The data
     * is not stored in the file but can be used by plug-ins
     *
     * @return Auxiliary data instance
     */
    AuxiliaryData getAuxiliaryData() {
        return auxiliaryData;
    }

    /**
     * Gets the shortener object for the current worksheet
     *
     * @return Shortener instance
     */
    public Shortener getShortener() {
        return shortener;
    }

    /**
     * Gets the current worksheet
     *
     * @return Current worksheet
     */
    public Worksheet getCurrentWorksheet() {
        return currentWorksheet;
    }

    /**
     * Gets the filename of the workbook.
     *
     * @return Filename of the workbook
     */
    public String getFilename() {
        return filename;
    }

    /**
     * Sets the filename of the workbook.
     *
     * <p>Remarks: Note that the file name is not sanitized. If a filename is set that is not compliant to the file
     * system, saving of the workbook may fail</p>
     *
     * @param filename Filename of the workbook
     */
    public void setFilename(String filename) {
        this.filename = filename;
    }

    /**
     * Gets whether the structure are locked if workbook is protected. See also {@link #setWorkbookProtection}
     *
     * @return True if windows are locked on protection, otherwise false
     */
    public boolean isLockStructureIfProtected() {
        return lockStructureIfProtected;
    }

    /**
     * Gets whether the windows are locked if workbook is protected. See also {@link #setWorkbookProtection}
     *
     * @return True if windows are locked on protection, otherwise false
     */
    public boolean isLockWindowsIfProtected() {
        return lockWindowsIfProtected;
    }

    /**
     * Gets the metadata object of the workbook
     *
     * @return Workbook metadata
     */
    public Metadata getWorkbookMetadata() {
        return workbookMetadata;
    }

    /**
     * Sets teh Metadata object of the workbook
     *
     * @param workbookMetadata Workbook metadata
     */
    public void setWorkbookMetadata(Metadata workbookMetadata) {
        this.workbookMetadata = workbookMetadata;
    }

    /**
     * Gets the selected worksheet. The selected worksheet is not the current worksheet while design time but the
     * selected sheet in the output file
     *
     * @return Index of the selected worksheet
     */
    public int getSelectedWorksheet() {
        return selectedWorksheet;
    }

    /**
     * Gets whether the workbook is protected
     *
     * @return True if workbook protection is used, otherwise false
     */
    public boolean isUseWorkbookProtection() {
        return useWorkbookProtection;
    }

    /**
     * Sets whether the workbook is protected
     *
     * @param useWorkbookProtection True if workbook protection is used, otherwise false
     */
    public void setUseWorkbookProtection(boolean useWorkbookProtection) {
        this.useWorkbookProtection = useWorkbookProtection;
    }

    /**
     * Password instance of the protected workbook. If a password was set, the pain text representation and the hash can
     * be read from the instance
     *
     * <p>Remarks: The password of this property is stored in plain text at runtime but not stored to a workbook. The
     * plain text password cannot be recovered when loading a workbook. The hash is retrieved and can be reused, if no
     * changes are made in the area of workbook protection
     * ({@link #setWorkbookProtection(boolean, boolean, boolean, String)})
     * </p>
     */
    public Password getWorkbookProtectionPassword() {
        return workbookProtectionPassword;
    }

    /**
     * Gets the list of worksheets in the workbook
     */
    public List<Worksheet> getWorksheets() {
        return worksheets;
    }

    /**
     * Gets whether the whole workbook is hidden
     *
     * <p>Remarks: A hidden workbook can only be made visible, using another, already visible Excel window</p>
     *
     * @return True if the workbook is hidden, otherwise false
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Sets whether the whole workbook is hidden
     *
     * <p>Remarks: A hidden workbook can only be made visible, using another, already visible Excel window</p>
     *
     * @param hidden True if the workbook is hidden, otherwise false
     */
    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }

    /**
     * Gets the theme of the workbook. The default is defined by {@link Theme#getDefaultTheme()}. However, the theme can
     * be nullified
     *
     * @return Workbook theme
     */
    public Theme getWorkbookTheme() {
        return workbookTheme;
    }

    /**
     * Sets the theme of the workbook. The default is defined by {@link Theme#getDefaultTheme()}. However, the theme can
     * be nullified
     *
     * @param workbookTheme Workbook theme
     */
    public void setWorkbookTheme(Theme workbookTheme) {
        this.workbookTheme = workbookTheme;
    }

    /**
     * Gets the feature set of the workbook
     *
     * @return Feature set
     */
    public FeatureSet getFeatures() {
        return features;
    }

// Constructors

    /**
     * Default constructor. No initial worksheet is created. Use {@link #addWorksheet(String)} (or overloads) to add
     * one
     */
    public Workbook() {
        init();
    }

    /**
     * Constructor with additional parameter to create a default worksheet. This constructor can be used to define a
     * workbook that is saved as stream
     *
     * @param createWorkSheet If true, a default worksheet with the name 'Sheet1' will be crated and set as current
     *                        worksheet
     */
    public Workbook(boolean createWorkSheet) {
        init();
        if (createWorkSheet) {
            addWorksheet("Sheet1");
        }
    }

    /**
     * Constructor with additional parameter to create a default worksheet with the specified name. This constructor can
     * be used to define a workbook that is saved as stream
     *
     * @param sheetName Filename of the workbook.  The name will be sanitized automatically according to the
     *                  specifications of Excel
     */
    public Workbook(String sheetName) {
        init();
        addWorksheet(sheetName, true);
    }

    /**
     * Constructor with filename ant the name of the first worksheet
     *
     * @param filename  Filename of the workbook.  The name will be sanitized automatically according to the
     *                  specifications of Excel
     * @param sheetName Name of the first worksheet. The name will be sanitized automatically according to the
     *                  specifications of Excel
     * @throws WorksheetException Throws a WorksheetException if the name of the worksheet already exists
     * @throws FormatException    Thrown if the name contains illegal characters or is out of range (length between 1 an
     *                            31 characters)
     */
    public Workbook(String filename, String sheetName) {
        init();
        this.filename = filename;
        addWorksheet(sheetName, true);
    }

    /**
     * Constructor with filename ant the name of the first worksheet
     *
     * @param filename          Filename of the workbook
     * @param sheetName         Name of the first worksheet
     * @param sanitizeSheetName If true, the name of the worksheet will be sanitized automatically according to the
     *                          specifications of Excel
     * @throws WorksheetException Throws a WorksheetException if the name of the worksheet already exists
     * @throws FormatException    Thrown if the name contains illegal characters or is out of range (length between 1 an
     *                            31 characters)
     */
    public Workbook(String filename, String sheetName, boolean sanitizeSheetName) {
        init();
        this.filename = filename;
        if (sanitizeSheetName) {
            addWorksheet(Worksheet.sanitizeWorksheetName(sheetName, this));
        } else {
            addWorksheet(sheetName);
        }
    }

// methods

    // definedNames

    /**
     * Adds a defined name, pointing to a single cell, to the workbook
     *
     * @param name        Unique name of defined name
     * @param worksheet   Worksheet that contains the target cell (cannot be null)
     * @param cellAddress Cell address as string (cannot be null)
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or cell address is null
     * @throws FormatException    Thrown if the name of the defined name or the cell address is invalid
     */
    public DefinedName addDefinedNameCell(String name, Worksheet worksheet, String cellAddress) {
        return addDefinedNameCell(name, worksheet, cellAddress, null, null);
    }

    /**
     * Adds a defined name, pointing to a single cell, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param cellAddress    Cell address as string (cannot be null)
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or cell address is null
     * @throws FormatException    Thrown if the name of the defined name or the cell address is invalid
     */
    public DefinedName addDefinedNameCell(
            String name, Worksheet worksheet, String cellAddress, Worksheet localWorksheet) {
        return addDefinedNameCell(name, worksheet, cellAddress, localWorksheet, null);
    }

    /**
     * Adds a defined name, pointing to a single cell, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param cellAddress    Cell address as string (cannot be null)
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @param comment        Optional comment of the defined name
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or cell address is null
     * @throws FormatException    Thrown if the name of the defined name or the cell address is invalid
     */
    public DefinedName addDefinedNameCell(
            String name, Worksheet worksheet, String cellAddress, Worksheet localWorksheet, String comment) {
        Address address = new Address(cellAddress);
        return addDefinedNameCell(name, worksheet, address, localWorksheet, comment);
    }

    /**
     * Adds a defined name, pointing to a single cell, to the workbook
     *
     * @param name      Unique name of defined name
     * @param worksheet Worksheet that contains the target cell (cannot be null)
     * @param column    Column number (zero-based) of the target cell
     * @param row       Row number (zero-based) of the target cell
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or cell address is null
     * @throws FormatException    Thrown if the name of the defined name or the cell address is invalid
     */
    public DefinedName addDefinedNameCell(String name, Worksheet worksheet, int column, int row) {
        return addDefinedNameCell(name, worksheet, column, row, null, null);
    }

    /**
     * Adds a defined name, pointing to a single cell, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param column         Column number (zero-based) of the target cell
     * @param row            Row number (zero-based) of the target cell
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or cell address is null
     * @throws FormatException    Thrown if the name of the defined name or the cell address is invalid
     */
    public DefinedName addDefinedNameCell(
            String name, Worksheet worksheet, int column, int row, Worksheet localWorksheet) {
        return addDefinedNameCell(name, worksheet, column, row, localWorksheet, null);
    }

    /**
     * Adds a defined name, pointing to a single cell, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param column         Column number (zero-based) of the target cell
     * @param row            Row number (zero-based) of the target cell
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @param comment        Optional comment of the defined name
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or cell address is null
     * @throws FormatException    Thrown if the name of the defined name or the cell address is invalid
     */
    public DefinedName addDefinedNameCell(
            String name, Worksheet worksheet, int column, int row, Worksheet localWorksheet, String comment) {
        Address address = new Address(column, row);
        return addDefinedNameCell(name, worksheet, address, localWorksheet, comment);
    }

    /**
     * Adds a defined name, pointing to a single cell, to the workbook
     *
     * @param name        Unique name of defined name
     * @param worksheet   Worksheet that contains the target cell (cannot be null)
     * @param cellAddress Address object of the target cell
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or cell address is null
     * @throws FormatException    Thrown if the name of the defined name or the cell address is invalid
     */
    public DefinedName addDefinedNameCell(String name, Worksheet worksheet, Address cellAddress) {
        return addDefinedNameCell(name, worksheet, cellAddress, null, null);
    }

    /**
     * Adds a defined name, pointing to a single cell, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param cellAddress    Address object of the target cell
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or cell address is null
     * @throws FormatException    Thrown if the name of the defined name or the cell address is invalid
     */
    public DefinedName addDefinedNameCell(
            String name, Worksheet worksheet, Address cellAddress, Worksheet localWorksheet) {
        return addDefinedNameCell(name, worksheet, cellAddress, localWorksheet, null);
    }

    /**
     * Adds a defined name, pointing to a single cell, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param cellAddress    Address object of the target cell
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @param comment        Optional comment of the defined name
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or cell address is null
     * @throws FormatException    Thrown if the name of the defined name or the cell address is invalid
     */
    public DefinedName addDefinedNameCell(
            String name, Worksheet worksheet, Address cellAddress, Worksheet localWorksheet, String comment) {
        if (worksheet == null) {
            throw new WorksheetException("A defined name to a cell must have a worksheet");
        }
        if (cellAddress == null) {
            throw new WorksheetException("The cell address pointing to a defined name cannot be null");
        }
        return addDefinedName(name, DefinedName.NameType.CELL, cellAddress, worksheet, localWorksheet, comment);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * <p>Remarks: The range string must contain an explicit start and end address, for example {@code A1:B2}. To
     * define a name for one cell, use {@link #addDefinedNameCell(String, Worksheet, String, Worksheet, String)}.</p>
     *
     * @param name         Unique name of defined name
     * @param worksheet    Worksheet that contains the target cell (cannot be null)
     * @param rangeAddress Address of the target range as string
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(String name, Worksheet worksheet, String rangeAddress) {
        return addDefinedNameRange(name, worksheet, rangeAddress, null, null);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * <p>Remarks: The range string must contain an explicit start and end address, for example {@code A1:B2}. To
     * define a name for one cell, use {@link #addDefinedNameCell(String, Worksheet, String, Worksheet, String)}.</p>
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param rangeAddress   Address of the target range as string
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(
            String name, Worksheet worksheet, String rangeAddress, Worksheet localWorksheet) {
        return addDefinedNameRange(name, worksheet, rangeAddress, localWorksheet, null);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * <p>Remarks: The range string must contain an explicit start and end address, for example {@code A1:B2}. To
     * define a name for one cell, use {@link #addDefinedNameCell(String, Worksheet, String, Worksheet, String)}.</p>
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param rangeAddress   Address of the target range as string
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @param comment        Optional comment of the defined name
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(
            String name, Worksheet worksheet, String rangeAddress, Worksheet localWorksheet, String comment) {
        Range range = new Range(rangeAddress);
        return addDefinedNameRange(name, worksheet, range, localWorksheet, comment);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * @param name         Unique name of defined name
     * @param worksheet    Worksheet that contains the target cell (cannot be null)
     * @param startAddress Start address object of the target range
     * @param endAddress   End address object of the target range
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(String name, Worksheet worksheet, Address startAddress, Address endAddress) {
        return addDefinedNameRange(name, worksheet, startAddress, endAddress, null, null);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param startAddress   Start address object of the target range
     * @param endAddress     End address object of the target range
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(
            String name, Worksheet worksheet, Address startAddress, Address endAddress, Worksheet localWorksheet) {
        return addDefinedNameRange(name, worksheet, startAddress, endAddress, localWorksheet, null);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param startAddress   Start address object of the target range
     * @param endAddress     End address object of the target range
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @param comment        Optional comment of the defined name
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(
            String name, Worksheet worksheet, Address startAddress, Address endAddress,
            Worksheet localWorksheet, String comment
    ) {
        Range range = new Range(startAddress, endAddress);
        return addDefinedNameRange(name, worksheet, range, localWorksheet, comment);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * @param name        Unique name of defined name
     * @param worksheet   Worksheet that contains the target cell (cannot be null)
     * @param startColumn Start column number (zero-based) of the target range
     * @param startRow    Start row number (zero-based) of the target range
     * @param endColumn   End column number (zero-based) of the target range
     * @param endRow      End row number (zero-based) of the target range
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(
            String name, Worksheet worksheet, int startColumn, int startRow, int endColumn, int endRow) {
        return addDefinedNameRange(name, worksheet, startColumn, startRow, endColumn, endRow, null, null);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param startColumn    Start column number (zero-based) of the target range
     * @param startRow       Start row number (zero-based) of the target range
     * @param endColumn      End column number (zero-based) of the target range
     * @param endRow         End row number (zero-based) of the target range
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(
            String name, Worksheet worksheet, int startColumn, int startRow,
            int endColumn, int endRow, Worksheet localWorksheet
    ) {
        return addDefinedNameRange(name, worksheet, startColumn, startRow, endColumn, endRow, localWorksheet, null);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param startColumn    Start column number (zero-based) of the target range
     * @param startRow       Start row number (zero-based) of the target range
     * @param endColumn      End column number (zero-based) of the target range
     * @param endRow         End row number (zero-based) of the target range
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @param comment        Optional comment of the defined name
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(
            String name, Worksheet worksheet, int startColumn, int startRow,
            int endColumn, int endRow, Worksheet localWorksheet, String comment
    ) {
        Range range = new Range(startColumn, startRow, endColumn, endRow);
        return addDefinedNameRange(name, worksheet, range, localWorksheet, comment);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * @param name         Unique name of defined name
     * @param worksheet    Worksheet that contains the target cell (cannot be null)
     * @param rangeAddress Range object of the target range
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(String name, Worksheet worksheet, Range rangeAddress) {
        return addDefinedNameRange(name, worksheet, rangeAddress, null, null);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param rangeAddress   Range object of the target range
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(
            String name, Worksheet worksheet, Range rangeAddress, Worksheet localWorksheet) {
        return addDefinedNameRange(name, worksheet, rangeAddress, localWorksheet, null);
    }

    /**
     * Adds a defined name, pointing to a cell range, to the workbook
     *
     * @param name           Unique name of defined name
     * @param worksheet      Worksheet that contains the target cell (cannot be null)
     * @param rangeAddress   Range object of the target range
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @param comment        Optional comment of the defined name
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the worksheet or range address is null
     * @throws FormatException    Thrown if the name of the defined name or the range address is invalid
     */
    public DefinedName addDefinedNameRange(
            String name, Worksheet worksheet, Range rangeAddress, Worksheet localWorksheet, String comment) {
        if (worksheet == null) {
            throw new WorksheetException("A defined name to a cell must have a worksheet");
        }
        return addDefinedName(name, DefinedName.NameType.RANGE, rangeAddress, worksheet, localWorksheet, comment);
    }

    /**
     * Adds a defined name, pointing to a constant value, to the workbook
     *
     * <p>Remarks: A constant should be one of the compatible types: String, int, double, float, long, short,
     * BigDecimal, byte, Date, Duration, boolean. Other types will be treated as string, using the default toString()
     * method.</p>
     *
     * @param name  Unique name of defined name
     * @param value Constant value
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the constant value is null
     * @throws FormatException    Thrown if the name of the defined name is invalid
     */
    public DefinedName addDefinedNameConstant(String name, Object value) {
        return addDefinedNameConstant(name, value, null, null);
    }

    /**
     * Adds a defined name, pointing to a constant value, to the workbook
     *
     * <p>Remarks: A constant should be one of the compatible types: String, int, double, float, long, short,
     * BigDecimal, byte, Date, Duration, boolean. Other types will be treated as string, using the default toString()
     * method.</p>
     *
     * @param name           Unique name of defined name
     * @param value          Constant value
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the constant value is null
     * @throws FormatException    Thrown if the name of the defined name is invalid
     */
    public DefinedName addDefinedNameConstant(String name, Object value, Worksheet localWorksheet) {
        return addDefinedNameConstant(name, value, localWorksheet, null);
    }

    /**
     * Adds a defined name, pointing to a constant value, to the workbook
     *
     * <p>Remarks: A constant should be one of the compatible types: String, int, double, float, long, short,
     * BigDecimal, byte, Date, Duration, boolean. Other types will be treated as string, using the default toString()
     * method.</p>
     *
     * @param name           Unique name of defined name
     * @param value          Constant value
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @param comment        Optional comment of the defined name
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the constant value is null
     * @throws FormatException    Thrown if the name of the defined name is invalid
     */
    public DefinedName addDefinedNameConstant(String name, Object value, Worksheet localWorksheet, String comment) {
        if (value == null) {
            throw new WorksheetException("A constant value pointing to a defined name cannot be nul");
        }
        return addDefinedName(name, DefinedName.NameType.CONSTANT, value, null, localWorksheet, comment);
    }

    /**
     * Adds a defined name, pointing to a formula expression, to the workbook. Do not add a leading equal sign (=) to
     * the formula. For simple cell references (e.g. "A1") use
     * {@link #addDefinedNameCell(String, Worksheet, Address, Worksheet, String)} or one of the overloaded methods. For
     * cell range references (e.g. "A1:C3") use
     * {@link #addDefinedNameRange(String, Worksheet, Address, Address, Worksheet, String)} or one of the overloaded
     * methods.
     *
     * <p>Remarks: The formula is not evaluated. Also do not add references to external workbooks (will cause an
     * exception on save), as longs no NanoXLSX extension is loaded that can handle external links.</p>
     *
     * @param name    Unique name of defined name
     * @param formula Formula value
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the formula value is null or empty
     * @throws FormatException    Thrown if the name of the defined name is invalid
     */
    public DefinedName addDefinedNameFormula(String name, String formula) {
        return addDefinedNameFormula(name, formula, null, null);
    }

    /**
     * Adds a defined name, pointing to a formula expression, to the workbook. Do not add a leading equal sign (=) to
     * the formula. For simple cell references (e.g. "A1") use
     * {@link #addDefinedNameCell(String, Worksheet, Address, Worksheet, String)} or one of the overloaded methods. For
     * cell range references (e.g. "A1:C3") use
     * {@link #addDefinedNameRange(String, Worksheet, Address, Address, Worksheet, String)} or one of the overloaded
     * methods.
     *
     * <p>Remarks: The formula is not evaluated. Also do not add references to external workbooks (will cause an
     * exception on save), as longs no NanoXLSX extension is loaded that can handle external links.</p>
     *
     * @param name           Unique name of defined name
     * @param formula        Formula value
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the formula value is null or empty
     * @throws FormatException    Thrown if the name of the defined name is invalid
     */
    public DefinedName addDefinedNameFormula(String name, String formula, Worksheet localWorksheet) {
        return addDefinedNameFormula(name, formula, localWorksheet, null);
    }

    /**
     * Adds a defined name, pointing to a formula expression, to the workbook. Do not add a leading equal sign (=) to
     * the formula. For simple cell references (e.g. "A1") use
     * {@link #addDefinedNameCell(String, Worksheet, Address, Worksheet, String)} or one of the overloaded methods. For
     * cell range references (e.g. "A1:C3") use
     * {@link #addDefinedNameRange(String, Worksheet, Address, Address, Worksheet, String)} or one of the overloaded
     * methods.
     *
     * <p>Remarks: The formula is not evaluated. Also do not add references to external workbooks (will cause an
     * exception on save), as longs no NanoXLSX extension is loaded that can handle external links.</p>
     *
     * @param name           Unique name of defined name
     * @param formula        Formula value
     * @param localWorksheet If not null, the defined name will only be available on the given worksheet. The
     *                       combination of 'name' and 'worksheet' must be unique
     * @param comment        Optional comment of the defined name
     * @return Returns the added defined name object
     * @throws WorksheetException Thrown if the formula value is null or empty
     * @throws FormatException    Thrown if the name of the defined name is invalid
     */
    public DefinedName addDefinedNameFormula(String name, String formula, Worksheet localWorksheet, String comment) {
        if (ParserUtils.isNullOrWhiteSpace(formula)) {
            throw new WorksheetException("A formula value pointing to a defined name cannot be null or empty");
        }
        // Note: Added strings like '[0]' will out-of-the-box throw an exception on saving a workbook
        return addDefinedName(name, DefinedName.NameType.FORMULA, formula, null, localWorksheet, comment);
    }

    /**
     * Internal method to add a defined name
     *
     * @param name            Name of the defined name.
     * @param type            Type of the reference.
     * @param value           Reference text (cell, range, formula, or constant).
     * @param targetWorksheet Target worksheet in case of a cell or range. Pass null for constant or formula.
     * @param localWorksheet  Optional worksheet that scopes the defined name. Pass null for workbook scope.
     * @param comment         Optional comment.
     * @return Returns the added defined name.
     */
    DefinedName addDefinedName(
            String name, DefinedName.NameType type, Object value, Worksheet targetWorksheet, Worksheet localWorksheet,
            String comment
    ) {
        // Validation is implemented in DefinedName class
        DefinedName definedName = new DefinedName(
                this, type, name, value, targetWorksheet, localWorksheet, null, Optional.empty());
        addDefinedName(definedName);
        return definedName;
    }

    /**
     * Internal method to add a defined name object without validation
     *
     * @param definedName Defined name object
     */
    void addDefinedName(DefinedName definedName) {
        definedNames.add(definedName);
    }

    /**
     * Removes the defined name with the supplied name.
     *
     * <p>Remarks: <b>Important:</b> Removing defined names may remove references in worksheets. If a workbook is saved
     * in this state, it can lead to broken formula references, indicated by "#NAME?"</p>
     *
     * @param name Name of the defined name to remove.
     * @return True if a matching defined name was removed, false if no match was found.
     */
    public boolean removeDefinedName(String name) {
        return removeDefinedName(name);
    }

    /**
     * Removes the defined name with the supplied name and scope.
     *
     * <p>Remarks: <b>Important:</b> Removing defined names may remove references in worksheets. If a workbook is saved
     * in this state, it can lead to broken formula references, indicated by "#NAME?"</p>
     *
     * @param name       Name of the defined name to remove.
     * @param localSheet Worksheet scope, or null for workbook scope.
     * @return True if a matching defined name was removed, false if no match was found.
     */
    public boolean removeDefinedName(String name, Worksheet localSheet) {
        return removeDefinedName(name, true, localSheet);
    }

    /**
     * Removes the defined name with the supplied name and scope with optional invalidation of formula references.
     *
     * <p>Remarks: <b>Important:</b> Removing defined names may remove references in worksheets. If a workbook is saved
     * in this state, it can lead to broken formula references, indicated by "#NAME?"</p>
     *
     * @param name       Name of the defined name to remove.
     * @param invalidate If true, all references in formula cells will be removed, otherwise left untouched
     * @param localSheet Worksheet scope, or null for workbook scope.
     * @return True if a matching defined name was removed, false if no match was found.
     */
    boolean removeDefinedName(String name, boolean invalidate, Worksheet localSheet) {
        int index = findDefinedNameIndex(name, localSheet);
        if (index < 0) {
            return false;
        }
        DefinedName definedName = definedNames.get(index);
        if (invalidate) {
            invalidateDefinedNameReferences(definedName);
        }
        definedName.getFeatures().remove(features);  // Decrease counter of defined name features
        definedNames.remove(index);
        return true;
    }

    /**
     * Removes all defined names of the workbook and removes their references from formula cells
     *
     * <p>Remarks: <b>Important:</b> Removing defined names may remove references in worksheets. If a workbook is saved
     * in this state, it can lead to broken formula references, indicated by "#NAME?"</p>
     */
    public void clearDefinedNames() {
        for (int i = definedNames.size() - 1; i >= 0; i--) {
            DefinedName definedName = definedNames.get(i);
            invalidateDefinedNameReferences(definedName);
            definedName.getFeatures().remove(features);  // Decrease counter of defined name features
            definedNames.remove(i);
        }
    }

    /**
     * Gets the defined name with the supplied name and scope.
     *
     * @param name       Name of the defined name.
     * @param localSheet Worksheet scope, or null for workbook scope.
     * @return The matching {@link DefinedName} instance, or null if no match was found.
     */
    public DefinedName getDefinedName(String name, Worksheet localSheet) {
        int index = findDefinedNameIndex(name, localSheet);
        return index < 0 ? null : definedNames.get(index);
    }

    /**
     * Gets a read-only view of all defined names in this workbook in insertion order.
     *
     * @return Read-only list of {@link DefinedName} instances.
     */
    public List<DefinedName> getDefinedNames() {
        return List.copyOf(definedNames);
    }

    /**
     * Locates the index of a defined name by name and scope.
     *
     * @param name       Name to find.
     * @param localSheet Worksheet scope, or null for workbook scope.
     * @return Index in the internal list, or -1 if not found.
     */
    int findDefinedNameIndex(String name, Worksheet localSheet) {
        for (int i = 0; i < definedNames.size(); i++) {
            DefinedName candidate = definedNames.get(i);
            if (ParserUtils.equalsIgnoreCase(candidate.getName(), name) && candidate.getLocalSheet() == localSheet) {
                return i;
            }
        }
        return -1;
    }

// other methods

    /**
     * Adds a color value (HEX; 6-digit RGB or 8-digit ARGB) to the MRU list
     *
     * @param color RGB code in hex format (either 6 characters, e.g. FF00AC or 8 characters with leading alpha value).
     *              Alpha will be set to full opacity (FF) in case of 6 characters
     */
    public void addMruColor(String color) {
        Validators.validateGenericColor(color);
        mruColors.add(Color.createRgb(color));
    }

    /**
     * Adds a generic color value. This can be an RGB/ARGB color, Auto, Theme, Indexed or System color
     *
     * @param color Color instance
     */
    public void addMruColor(Color color) {
        mruColors.add(color);
    }

    /**
     * Gets the MRU color list
     *
     * @return Immutable list of color instances
     */
    public List<Color> getMruColors() {
        return List.copyOf(mruColors);
    }

    /**
     * Clears the MRU color list
     */
    public void clearMruColors() {
        mruColors.clear();
    }

    /**
     * Adding a new Worksheet. The new worksheet will be defined as current worksheet
     *
     * @param name Name of the new worksheet
     * @throws WorksheetException Throws a WorksheetException if the name of the worksheet already exists
     * @throws FormatException    Thrown if the name contains illegal characters or is out of range (length between 1 an
     *                            31 characters)
     */
    public void addWorksheet(String name) {
        for (Worksheet item : worksheets) {
            if (item.getSheetName().equals(name)) {
                throw new WorksheetException("The worksheet with the name '" + name + "' already exists.");
            }
        }
        int number = getNextWorksheetId();
        Worksheet newWs = new Worksheet(name, number, this);
        currentWorksheet = newWs;
        worksheets.add(newWs);
        newWs.getFeatures().add(features);
        shortener.setCurrentWorksheetInternal(currentWorksheet);
    }

    /**
     * Adding a new Worksheet with a sanitizing option. The new worksheet will be defined as current worksheet
     *
     * @param name              Name of the new worksheet
     * @param sanitizeSheetName If true, the name of the worksheet will be sanitized automatically according to the
     *                          specifications of Excel
     * @throws WorksheetException WorksheetException is thrown if the name of the worksheet already exists and
     *                            sanitizeSheetName is false
     * @throws FormatException    Thrown if the worksheet name contains illegal characters or is out of range (length
     *                            between 1 an 31) and sanitizeSheetName is false
     */
    public void addWorksheet(String name, boolean sanitizeSheetName) {
        if (sanitizeSheetName) {
            String sanitized = Worksheet.sanitizeWorksheetName(name, this);
            addWorksheet(sanitized);
        } else {
            addWorksheet(name);
        }
    }

    /**
     * Adding a new Worksheet. The new worksheet will be defined as current worksheet
     *
     * @param worksheet Prepared worksheet object
     * @throws WorksheetException WorksheetException is thrown if the name of the worksheet already exists
     * @throws FormatException    FormatException is thrown if the worksheet name contains illegal characters or is out
     *                            of range (length between 1 an 31)
     */
    public void addWorksheet(Worksheet worksheet) {
        addWorksheet(worksheet, false);
    }

    /**
     * Adding a new Worksheet. The new worksheet will be defined as current worksheet
     *
     * @param worksheet         Prepared worksheet object
     * @param sanitizeSheetName If true, the name of the worksheet will be sanitized automatically according to the
     *                          specifications of Excel
     * @throws WorksheetException Thrown if the name of the worksheet already exists, when sanitation is false
     * @throws FormatException    Thrown if the worksheet name contains illegal characters or is out of range (length
     *                            between 1 an 31) and sanitation is false
     */
    public void addWorksheet(Worksheet worksheet, boolean sanitizeSheetName) {
        if (sanitizeSheetName) {
            String name = Worksheet.sanitizeWorksheetName(worksheet.getSheetName(), this);
            worksheet.setSheetName(name);
        } else {
            if (ParserUtils.isNullOrEmpty(worksheet.getSheetName())) {
                throw new WorksheetException("The name of the passed worksheet is null or empty.");
            }
            for (int i = 0; i < worksheets.size(); i++) {
                if (worksheets.get(i).getSheetName() != null &&
                        worksheets.get(i).getSheetName().equals(worksheet.getSheetName())) {
                    throw new WorksheetException(
                            "The worksheet with the name '" + worksheet.getSheetName() + "' already exists.");
                }
            }
        }
        worksheet.setSheetId(getNextWorksheetId());
        currentWorksheet = worksheet;
        worksheets.add(worksheet);
        worksheet.setWorkbookReference(this);
        worksheet.getFeatures().add(features);
    }

    /**
     * Removes the defined worksheet based on its name. If the worksheet is the current or selected worksheet, the
     * current and / or the selected worksheet will be set to the last worksheet of the workbook. If the last worksheet
     * is removed, the selected worksheet will be set to 0 and the current worksheet to null.
     *
     * @param name Name of the worksheet
     * @throws WorksheetException Throws a WorksheetException if the name of the worksheet is unknown
     */
    public void removeWorksheet(String name) {
        Optional<Worksheet> worksheetToRemove = worksheets.stream()
                .filter(w -> w.getSheetName() != null && w.getSheetName().equals(name)).findFirst();
        if (worksheetToRemove.isEmpty()) {
            throw new WorksheetException("The worksheet with the name '" + name + "' does not exist.");
        }
        int index = worksheets.indexOf(worksheetToRemove.get());
        boolean resetCurrentWorksheet = worksheetToRemove.get() == currentWorksheet;
        removeWorksheet(index, resetCurrentWorksheet);
    }

    /**
     * Removes the defined worksheet based on its index. If the worksheet is the current or selected worksheet, the
     * current and / or the selected worksheet will be set to the last worksheet of the workbook. If the last worksheet
     * is removed, the selected worksheet will be set to 0 and the current worksheet to null.
     *
     * @param index Index within the worksheets list
     * @throws WorksheetException Throws a WorksheetException if the index is out of range
     */

    public void removeWorksheet(int index) {
        if (index < 0 || index >= worksheets.size()) {
            throw new WorksheetException("The worksheet index " + index + " is out of range");
        }
        boolean resetCurrentWorksheet = worksheets.get(index) == currentWorksheet;
        removeWorksheet(index, resetCurrentWorksheet);
    }

    /**
     * Method to resolve all merged cells in all worksheets. Only the value of the very first cell of the locked cells
     * range will be visible. The other values are still present (set to EMPTY) but will not be stored in the
     * worksheet.<br /> This is an internal method. There is no need to use it
     *
     * @throws StyleException Throws a StyleException if one of the styles of the merged cells cannot be referenced or
     *                        is null
     */
    void resolveMergedCells() {
        for (Worksheet worksheet : worksheets) {
            worksheet.resolveMergedCells();
        }
    }

    /**
     * Sets the current worksheet
     *
     * @param name Name of the worksheet
     * @return Returns the current worksheet
     * @throws WorksheetException Thrown if the name of the worksheet is unknown
     */
    public Worksheet setCurrentWorksheet(String name) {
        currentWorksheet = getWorksheet(name);
        shortener.setCurrentWorksheetInternal(currentWorksheet);
        return currentWorksheet;
    }

    /// <summary>
    /// Sets the current worksheet
    /// </summary>
    /// <param name="worksheetIndex">Zero-based worksheet index</param>
    /// <returns>Returns the current worksheet</returns>
    /// <exception cref="WorksheetException">Thrown if the name of the worksheet is unknown</exception>
    public Worksheet setCurrentWorksheet(int worksheetIndex) {
        currentWorksheet = getWorksheet(worksheetIndex);
        shortener.setCurrentWorksheetInternal(currentWorksheet);
        return currentWorksheet;
    }

    /**
     * Sets the current worksheet
     *
     * @param worksheet Worksheet object (must be in the collection of worksheets)
     * @throws WorksheetException Thrown if the worksheet was not found in the worksheet collection
     */
    public void setCurrentWorksheet(Worksheet worksheet) {
        int index = worksheets.indexOf(worksheet);
        if (index < 0) {
            throw new WorksheetException("The passed worksheet object is not in the worksheet collection.");
        }
        currentWorksheet = worksheets.get(index);
        shortener.setCurrentWorksheetInternal(worksheet);
    }

    /**
     * Sets the selected worksheet in the output workbook
     *
     * @param name Name of the worksheet
     * @throws WorksheetException Thrown if the name of the worksheet is unknown
     */
    public void setSelectedWorksheet(String name) {
        int index = IntStream.range(0, worksheets.size())
                .filter(i -> Objects.equals(worksheets.get(i).getSheetName(), name)).findFirst().orElse(-1);
        if (index < 0) {
            throw new WorksheetException("No worksheet with the name '" + name + "' was found in this workbook.");
        }
        selectedWorksheet = index;
    }

    /**
     * Sets the selected worksheet in the output workbook
     *
     * <p>Remarks: This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead
     * for this</p>
     *
     * @param worksheetIndex Zero-based worksheet index
     * @throws RangeException     Throws a RangeException if the index of the worksheet is out of range
     * @throws WorksheetException Thrown if the worksheet to be set selected is hidden
     */
    public void setSelectedWorksheet(int worksheetIndex) {
        if (worksheetIndex < 0 || worksheetIndex > worksheets.size() - 1) {
            throw new RangeException("The worksheet index " + worksheetIndex + " is out of range");
        }
        selectedWorksheet = worksheetIndex;
        validateWorksheets();
    }

    /**
     * Sets the selected worksheet in the output workbook
     *
     * <p>Remarks: This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead
     * for this</p>
     *
     * @param worksheet Worksheet object (must be in the collection of worksheets)
     * @throws WorksheetException Thrown if the worksheet was not found in the worksheet collection or if it is hidden
     */
    public void setSelectedWorksheet(Worksheet worksheet) {
        selectedWorksheet = worksheets.indexOf(worksheet);
        if (selectedWorksheet < 0) {
            throw new WorksheetException("The passed worksheet object is not in the worksheet collection.");
        }
        validateWorksheets();
    }

    /**
     * Gets a worksheet from this workbook by name
     *
     * @param name Name of the worksheet
     * @return Worksheet with the passed name
     * @throws WorksheetException Thrown if the worksheet was not found in the worksheet collection
     */
    public Worksheet getWorksheet(String name) {
        int index = IntStream.range(0, worksheets.size())
                .filter(i -> Objects.equals(worksheets.get(i).getSheetName(), name)).findFirst().orElse(-1);
        if (index < 0) {
            throw new WorksheetException("No worksheet with the name '" + name + "' was found in this workbook.");
        }
        return worksheets.get(index);
    }

    /**
     * Gets a worksheet from this workbook by index
     *
     * @param index Index of the worksheet
     * @return Worksheet with the passed index
     * @throws WorksheetException Thrown if the worksheet was not found in the worksheet collection
     */
    public Worksheet getWorksheet(int index) {
        if (index < 0 || index > worksheets.size() - 1) {
            throw new RangeException("The worksheet index " + index + " is out of range");
        }
        return worksheets.get(index);
    }

    /**
     * Sets or removes the workbook protection. If protectWindows and protectStructure are both false, the workbook will
     * not be protected
     *
     * @param state            If true, the workbook will be protected, otherwise not
     * @param protectWindows   If true, the windows will be locked if the workbook is protected
     * @param protectStructure If true, the structure will be locked if the workbook is protected
     * @param password         Optional password. If null or empty, no password will be set in case of protection
     */
    public void setWorkbookProtection(
            boolean state, boolean protectWindows, boolean protectStructure, String password) {
        lockWindowsIfProtected = protectWindows;
        lockStructureIfProtected = protectStructure;
        workbookProtectionPassword.setPassword(password);
        if (!protectWindows && !protectStructure) {
            useWorkbookProtection = false;
        } else {
            useWorkbookProtection = state;
        }
    }

    /**
     * Copies a worksheet of the current workbook by its name
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept</p>
     *
     * @param sourceWorksheetName Name of the worksheet to copy, originated in this workbook
     * @param newWorksheetName    Name of the new worksheet (copy)
     * @return Copied worksheet
     */
    public Worksheet copyWorksheetIntoThis(String sourceWorksheetName, String newWorksheetName) {
        return copyWorksheetIntoThis(sourceWorksheetName, newWorksheetName, true);
    }

    /**
     * Copies a worksheet of the current workbook by its name
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept</p>
     *
     * @param sourceWorksheetName Name of the worksheet to copy, originated in this workbook
     * @param newWorksheetName    Name of the new worksheet (copy)
     * @param sanitizeSheetName   If true, the new name will be automatically sanitized if a name collision occurs
     * @return Copied worksheet
     */
    public Worksheet copyWorksheetIntoThis(
            String sourceWorksheetName, String newWorksheetName, boolean sanitizeSheetName) {
        Worksheet sourceWorksheet = getWorksheet(sourceWorksheetName);
        return copyWorksheetTo(sourceWorksheet, newWorksheetName, this, sanitizeSheetName);
    }

    /**
     * Copies a worksheet of the current workbook by its index
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept</p>
     *
     * @param sourceWorksheetIndex Index of the worksheet to copy, originated in this workbook
     * @param newWorksheetName     Name of the new worksheet (copy)
     * @return Copied worksheet
     */
    public Worksheet copyWorksheetIntoThis(int sourceWorksheetIndex, String newWorksheetName) {
        return copyWorksheetIntoThis(sourceWorksheetIndex, newWorksheetName, true);
    }

    /**
     * Copies a worksheet of the current workbook by its index
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept</p>
     *
     * @param sourceWorksheetIndex Index of the worksheet to copy, originated in this workbook
     * @param newWorksheetName     Name of the new worksheet (copy)
     * @param sanitizeSheetName    If true, the new name will be automatically sanitized if a name collision occurs
     * @return Copied worksheet
     */
    public Worksheet copyWorksheetIntoThis(
            int sourceWorksheetIndex, String newWorksheetName, boolean sanitizeSheetName) {
        Worksheet sourceWorksheet = getWorksheet(sourceWorksheetIndex);
        return copyWorksheetTo(sourceWorksheet, newWorksheetName, this, sanitizeSheetName);
    }

    /**
     * Copies a worksheet of any workbook into the current workbook
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept. The source worksheet can
     * originate from any workbook</p>
     *
     * @param sourceWorksheet  Worksheet to copy
     * @param newWorksheetName Name of the new worksheet (copy)
     * @return Copied worksheet
     */
    public Worksheet copyWorksheetIntoThis(Worksheet sourceWorksheet, String newWorksheetName) {
        return copyWorksheetIntoThis(sourceWorksheet, newWorksheetName, true);
    }

    /**
     * Copies a worksheet of any workbook into the current workbook
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept. The source worksheet can
     * originate from any workbook</p>
     *
     * @param sourceWorksheet   Worksheet to copy
     * @param newWorksheetName  Name of the new worksheet (copy)
     * @param sanitizeSheetName If true, the new name will be automatically sanitized if a name collision occurs
     * @return Copied worksheet
     */
    public Worksheet copyWorksheetIntoThis(
            Worksheet sourceWorksheet, String newWorksheetName, boolean sanitizeSheetName) {
        return copyWorksheetTo(sourceWorksheet, newWorksheetName, this, sanitizeSheetName);
    }

    /**
     * Copies a worksheet of the current workbook by its name into another workbook
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept</p>
     *
     * @param sourceWorksheetName Name of the worksheet to copy, originated in this workbook
     * @param newWorksheetName    Name of the new worksheet (copy)
     * @param targetWorkbook      Workbook to copy the worksheet into
     * @return Copied worksheet
     */
    public Worksheet copyWorksheetTo(String sourceWorksheetName, String newWorksheetName, Workbook targetWorkbook) {
        return copyWorksheetTo(sourceWorksheetName, newWorksheetName, targetWorkbook, true);
    }

    /**
     * Copies a worksheet of the current workbook by its name into another workbook
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept</p>
     *
     * @param sourceWorksheetName Name of the worksheet to copy, originated in this workbook
     * @param newWorksheetName    Name of the new worksheet (copy)
     * @param targetWorkbook      Workbook to copy the worksheet into
     * @param sanitizeSheetName   If true, the new name will be automatically sanitized if a name collision occurs
     * @return Copied worksheet
     */
    public Worksheet copyWorksheetTo(
            String sourceWorksheetName, String newWorksheetName, Workbook targetWorkbook, boolean sanitizeSheetName) {
        Worksheet sourceWorksheet = getWorksheet(sourceWorksheetName);
        return copyWorksheetTo(sourceWorksheet, newWorksheetName, targetWorkbook, sanitizeSheetName);
    }

    /**
     * Copies a worksheet of the current workbook by its index into another workbook
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept</p>
     *
     * @param sourceWorksheetIndex Index of the worksheet to copy, originated in this workbook
     * @param newWorksheetName     Name of the new worksheet (copy)
     * @param targetWorkbook       Workbook to copy the worksheet into
     * @return Copied worksheet
     */
    public Worksheet copyWorksheetTo(int sourceWorksheetIndex, String newWorksheetName, Workbook targetWorkbook) {
        return copyWorksheetTo(sourceWorksheetIndex, newWorksheetName, targetWorkbook, true);
    }

    /**
     * Copies a worksheet of the current workbook by its index into another workbook
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept</p>
     *
     * @param sourceWorksheetIndex Index of the worksheet to copy, originated in this workbook
     * @param newWorksheetName     Name of the new worksheet (copy)
     * @param targetWorkbook       Workbook to copy the worksheet into
     * @param sanitizeSheetName    If true, the new name will be automatically sanitized if a name collision occurs
     * @return Copied worksheet
     */
    public Worksheet copyWorksheetTo(
            int sourceWorksheetIndex, String newWorksheetName, Workbook targetWorkbook, boolean sanitizeSheetName) {
        Worksheet sourceWorksheet = getWorksheet(sourceWorksheetIndex);
        return copyWorksheetTo(sourceWorksheet, newWorksheetName, targetWorkbook, sanitizeSheetName);
    }

    /**
     * Copies a worksheet of any workbook into the another workbook
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept</p>
     *
     * @param sourceWorksheet  Worksheet to copy
     * @param newWorksheetName Name of the new worksheet (copy)
     * @param targetWorkbook   Workbook to copy the worksheet into
     * @return Copied worksheet
     */
    public static Worksheet copyWorksheetTo(
            Worksheet sourceWorksheet, String newWorksheetName, Workbook targetWorkbook) {
        return copyWorksheetTo(sourceWorksheet, newWorksheetName, targetWorkbook, true);
    }

    /**
     * Copies a worksheet of any workbook into the another workbook
     *
     * <p>Remarks: The copy is not set as current worksheet. The existing one is kept</p>
     *
     * @param sourceWorksheet   Worksheet to copy
     * @param newWorksheetName  Name of the new worksheet (copy)
     * @param targetWorkbook    Workbook to copy the worksheet into
     * @param sanitizeSheetName If true, the new name will be automatically sanitized if a name collision occurs
     * @return Copied worksheet
     */
    public static Worksheet copyWorksheetTo(
            Worksheet sourceWorksheet, String newWorksheetName, Workbook targetWorkbook, boolean sanitizeSheetName) {
        if (targetWorkbook == null) {
            throw new WorksheetException("The target workbook cannot be null");
        }
        if (sourceWorksheet == null) {
            throw new WorksheetException("The source worksheet cannot be null");
        }
        Worksheet copy = sourceWorksheet.copy();
        copy.setSheetName(newWorksheetName);
        Worksheet currentWorksheet = targetWorkbook.getCurrentWorksheet();
        targetWorkbook.addWorksheet(copy, sanitizeSheetName);
        targetWorkbook.setCurrentWorksheet(currentWorksheet);
        return copy;
    }

    /**
     * Validates the worksheets regarding several conditions that must be met:<br /> - At least one worksheet must be
     * defined<br /> - A hidden worksheet cannot be the selected one<br /> - At least one worksheet must be visible<br
     * /> If one of the conditions is not met, an exception is thrown
     *
     * <p>Remarks: If an import is in progress, these rules are disabled to avoid conflicts by the order of loaded
     * worksheets</p>
     */
    void validateWorksheets() {
        if (importInProgress) {
            // No validation during import
            return;
        }
        int worksheetCount = worksheets.size();
        if (worksheetCount == 0) {
            throw new WorksheetException("The workbook must contain at least one worksheet");
        }
        for (int i = 0; i < worksheetCount; i++) {
            if (worksheets.get(i).isHidden()) {
                if (i == selectedWorksheet) {
                    throw new WorksheetException("The worksheet with the index " + selectedWorksheet +
                            " cannot be set as selected, since it is set hidden");
                }
            }
        }
    }

    /// <summary>
    /// Removes all references of invalid defined names in Worksheets
    /// </summary>
    /// <param name="definedName">Defined name object to invalidate</param>
    private void invalidateDefinedNameReferences(DefinedName definedName) {
        for (Worksheet worksheet : worksheets) {
            if (!worksheet.getFeatures().containsDefinedNameReferences()) {
                continue;
            }
            for (Map.Entry<String, Cell> cell : worksheet.getCells().entrySet()) {
                if (cell.getValue().getDataType() == Cell.CellType.FORMULA) {
                    FormulaData formula = cell.getValue().getFormula();
                    if (formula != null && formula.getDefinedNameReference() == definedName) {
                        formula.setDefinedNameReference(null);
                    }
                }
            }
        }
    }

    /**
     * Removes the worksheet at the defined index and relocates current and selected worksheet references
     *
     * @param index                 Index within the worksheets list
     * @param resetCurrentWorksheet If true, the current worksheet will be relocated to the last worksheet in the list
     */
    private void removeWorksheet(int index, boolean resetCurrentWorksheet) {
        worksheets.get(index).getFeatures().remove(features); // Remove cascading features
        worksheets.remove(index);
        if (!worksheets.isEmpty()) {
            for (int i = 0; i < worksheets.size(); i++) {
                worksheets.get(i).setSheetId(i + 1);
            }
            if (resetCurrentWorksheet) {
                currentWorksheet = worksheets.getLast();
            }
            if (selectedWorksheet == index || selectedWorksheet > worksheets.size() - 1) {
                selectedWorksheet = worksheets.size() - 1;
            }
        } else {
            currentWorksheet = null;
            selectedWorksheet = 0;
        }
        validateWorksheets();
    }

    /**
     * Gets the next free worksheet ID
     *
     * @return Worksheet ID
     */
    private int getNextWorksheetId() {
        if (worksheets.isEmpty()) {
            return 1;
        }
        return worksheets.stream().map(Worksheet::getSheetId).max(Integer::compare).get() + 1;
    }

    /// <summary>
    /// Init method called in the constructors
    /// </summary>
    private void init() {
        this.worksheets = new ArrayList<>();
        this.workbookMetadata = new Metadata();
        this.shortener = new Shortener(this);
        this.workbookProtectionPassword = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        this.auxiliaryData = new AuxiliaryData();
    }
}
