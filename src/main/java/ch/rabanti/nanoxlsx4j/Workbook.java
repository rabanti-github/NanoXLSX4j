/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.*;
import ch.rabanti.nanoxlsx4j.lowLevel.XlsxReader;
import ch.rabanti.nanoxlsx4j.lowLevel.XlsxWriter;
import ch.rabanti.nanoxlsx4j.styles.*;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/**
 * Class representing a workbook
 *
 * @author Raphael Stoeckli
 */
public class Workbook {

    // ### P R I V A T E  F I E L D S ###
    private Worksheet currentWorksheet;
    private String filename;
    private boolean lockStructureIfProtected;
    private boolean lockWindowsIfProtected;
    private int selectedWorksheet;
    private boolean useWorkbookProtection;
    private Metadata workbookMetadata;
    private String workbookProtectionPassword;
    private List<Worksheet> worksheets;
    private boolean hidden;
    private boolean importInProgress = false;

    /**
     * Shortener omits getter and setter to simplify the access (Can throw a WorksheetException if not defined)
     */
    public Shortener WS;

// ### G E T T E R S  &  S E T T E R S ###

    /**
     * Gets the current worksheet
     *
     * @return Current worksheet reference
     */
    public Worksheet getCurrentWorksheet() {
        return currentWorksheet;
    }


    /**
     * Gets the filename of the workbook
     *
     * @return Filename of the workbook
     * @apiNote Note that the file name is not sanitized. If a filename is set that is not compliant to the file system, saving of the workbook may fail
     */
    public String getFilename() {
        return filename;
    }

    /**
     * Sets the filename of the workbook
     *
     * @param filename Filename of the workbook
     */
    public void setFilename(String filename) {
        this.filename = filename;
    }

    /**
     * Gets the selected worksheet. The selected worksheet is not the current worksheet while design time but the selected sheet in the output file
     *
     * @return Zero-based worksheet index
     */
    public int getSelectedWorksheet() {
        return selectedWorksheet;
    }

    /**
     * Gets the meta data object of the workbook
     *
     * @return Meta data object
     */
    public Metadata getWorkbookMetadata() {
        return workbookMetadata;
    }

    /**
     * Sets the meta data object of the workbook
     *
     * @param workbookMetadata Meta data object
     */
    public void setWorkbookMetadata(Metadata workbookMetadata) {
        this.workbookMetadata = workbookMetadata;
    }

    /**
     * Sets whether the workbook is protected
     *
     * @param useWorkbookProtection If true, the workbook is protected otherwise not
     */
    public void setWorkbookProtection(boolean useWorkbookProtection) {
        this.useWorkbookProtection = useWorkbookProtection;
    }

    /**
     * Gets the password used for workbook protection
     *
     * @return Password (UTF-8)
     * @apiNote The password of this filed is stored in plan text. Encryption is performed when the workbook is saved
     */
    public String getWorkbookProtectionPassword() {
        return workbookProtectionPassword;
    }

    /**
     * Gets the list of worksheets in the workbook
     *
     * @return List of worksheet objects
     */
    public List<Worksheet> getWorksheets() {
        return worksheets;
    }

    /**
     * Gets whether the structure are locked if workbook is protected
     *
     * @return True if the structure is locked when the workbook is protected
     */
    public boolean isStructureLockedIfProtected() {
        return lockStructureIfProtected;
    }

    /**
     * Gets whether the windows are locked if workbook is protected
     *
     * @return True if the windows are locked when the workbook is protected
     */
    public boolean isWindowsLockedIfProtected() {
        return lockWindowsIfProtected;
    }

    /**
     * Gets whether the workbook is protected
     *
     * @return If true, the workbook is protected otherwise not
     */
    public boolean isWorkbookProtectionUsed() {
        return useWorkbookProtection;
    }

    /**
     * Gets whether the workbook is hidden
     *
     * @return If true hidden, otherwise visible
     * @implNote A hidden workbook can only be made visible, using another, already visible Excel window
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Sets whether the workbook is hidden
     *
     * @param hidden If true hidden, otherwise visible
     * @implNote A hidden workbook can only be made visible, using another, already visible Excel window
     */
    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Default constructor. No initial worksheet is created. Use {@link #addWorksheet(String)} (or overloads) to add one
     */
    public Workbook() {
        init();
    }

    /**
     * Constructor with additional parameter to create a default worksheet. This constructor can be used to define a workbook that is saved as stream
     *
     * @param createWorksheet If true, a default worksheet with the name 'Sheet1' will be created and set as current worksheet
     */
    public Workbook(boolean createWorksheet) {
        init();
        if (createWorksheet == true) {
            addWorksheet("Sheet1");
        }
    }

    /**
     * Constructor with additional parameter to create a default worksheet with the specified name. This constructor can be used to define a workbook that is saved as stream
     *
     * @param sheetName Name of the first worksheet. The name will be sanitized automatically according to the specifications of Excel
     * @throws FormatException Thrown if the worksheet name contains illegal characters
     */
    public Workbook(String sheetName) {
        init();
        addWorksheet(sheetName, true);
    }

    /**
     * Constructor with filename ant the name of the first worksheet. This constructor can be used to define a workbook that is saved as stream
     *
     * @param filename  Filename of the workbook
     * @param sheetName Name of the first worksheet. The name will be sanitized automatically according to the specifications of Excel
     * @throws FormatException Thrown if the worksheet name contains illegal characters
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
     * @param sanitizeSheetName If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel
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

// ### M E T H O D S ###


    /**
     * Adds a style to the style manager
     *
     * @param style Style to add
     * @return The managed style of the style manager
     * @deprecated This method has no direct impact on the generated file and is deprecated
     */
    public Style addStyle(Style style) {
        return StyleRepository.getInstance().addStyle(style);
    }

    /**
     * Adds a style component to a style
     *
     * @param baseStyle    Style to append a component
     * @param newComponent Component to add to the baseStyle
     * @return The managed style of the style manager
     * @deprecated This method has no direct impact on the generated file and is deprecated.
     */
    public Style addStyleComponent(Style baseStyle, AbstractStyle newComponent) {

        if (newComponent instanceof Border) {
            baseStyle.setBorder((Border) newComponent);
        } else if (newComponent instanceof CellXf) {
            baseStyle.setCellXf((CellXf) newComponent);
        } else if (newComponent instanceof Fill) {
            baseStyle.setFill((Fill) newComponent);
        } else if (newComponent instanceof Font) {
            baseStyle.setFont((Font) newComponent);
        } else if (newComponent instanceof NumberFormat) {
            baseStyle.setNumberFormat((NumberFormat) newComponent);
        }
        return StyleRepository.getInstance().addStyle(baseStyle);
    }

    /**
     * Adding a new Worksheet. The new worksheet will be defined as current worksheet
     *
     * @param name Name of the new worksheet
     * @throws WorksheetException Thrown if the name of the worksheet already exists
     * @throws FormatException    Thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31)
     */
    public void addWorksheet(String name) {
        for (int i = 0; i < this.worksheets.size(); i++) {
            if (this.worksheets.get(i).getSheetName().equals(name)) {
                throw new WorksheetException("The worksheet with the name '" + name + "' already exists.");
            }
        }
        int number = getNextWorksheetId();
        Worksheet newWs = new Worksheet(name, number, this);
        this.currentWorksheet = newWs;
        this.worksheets.add(newWs);
        this.WS.setCurrentWorksheetInternal(this.currentWorksheet);
    }

    /**
     * Adding a new Worksheet with a sanitizing option. The new worksheet will be defined as current worksheet
     *
     * @param name              Name of the new worksheet
     * @param sanitizeSheetName If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel
     * @throws WorksheetException Thrown if the name of the worksheet already exists and sanitizeSheetName is false
     * @throws FormatException    Thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31) and sanitizeSheetName is false
     */
    public void addWorksheet(String name, boolean sanitizeSheetName) {
        if (sanitizeSheetName == true) {
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
     * @throws WorksheetException Thrown if the name of the worksheet already exists
     * @throws FormatException    Thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31
     */
    public void addWorksheet(Worksheet worksheet) {
        addWorksheet(worksheet, false);
    }

    /**
     * Adding a new Worksheet. The new worksheet will be defined as current worksheet
     *
     * @param worksheet Prepared worksheet object
     * @throws WorksheetException Thrown if the name of the worksheet already exists, when sanitation is false
     * @throws FormatException    Thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31) and sanitation is false
     */
    public void addWorksheet(Worksheet worksheet, boolean sanitizeSheetName) {
        if (sanitizeSheetName) {
            String name = Worksheet.sanitizeWorksheetName(worksheet.getSheetName(), this);
            worksheet.setSheetName(name);
        } else {
            if (worksheet.getSheetName() == null || worksheet.getSheetName().isEmpty()) {
                throw new WorksheetException("The name of the passed worksheet is null or empty.");
            }
            for (int i = 0; i < this.worksheets.size(); i++) {
                if (this.worksheets.get(i).getSheetName().equals(worksheet.getSheetName())) {
                    throw new WorksheetException("The worksheet with the name '" + worksheet.getSheetName() + "' already exists.");
                }
            }
        }
        int number = getNextWorksheetId();
        worksheet.setSheetID(number);
        this.currentWorksheet = worksheet;
        this.worksheets.add(worksheet);
        worksheet.setWorkbookReference(this);
    }

    /**
     * Removes the passed style from the style sheet
     *
     * @param style Style to remove
     * @apiNote Note: This method is available due to compatibility reasons.
     * Added styles are actually not removed by it since unused styles are disposed automatically
     * @deprecated This method has no direct impact on the generated file and is deprecated.
     */
    public void removeStyle(Style style) {
        removeStyle(style, false);
    }

    /**
     * Removes the defined style from the style sheet of the workbook
     *
     * @param styleName Name of the style to be removed
     * @apiNote Note: This method is available due to compatibility reasons.
     * Added styles are actually not removed by it since unused styles are disposed automatically
     * @deprecated This method has no direct impact on the generated file and is deprecated.
     */
    public void removeStyle(String styleName) {
        removeStyle(styleName, false);
    }

    /**
     * Removes the defined style from the style manager of the workbook
     *
     * @param style        Style to remove
     * @param onlyIfUnused If true, the style will only be removed if not used in any cell
     * @apiNote Note: This method is available due to compatibility reasons.
     * Added styles are actually not removed by it since unused styles are disposed automatically
     * @deprecated This method has no direct impact on the generated file and is deprecated.
     */
    public void removeStyle(Style style, boolean onlyIfUnused) {
        if (style == null) {
            throw new StyleException("The style to remove is not defined");
        }
        removeStyle(style.getName(), onlyIfUnused);
    }

    /**
     * Removes the defined style from the style manager of the workbook
     *
     * @param styleName    Name of the style to remove
     * @param onlyIfUnused If true, the style will only be removed if not used in any cell
     * @apiNote Note: This method is available due to compatibility reasons.
     * Added styles are actually not removed by it since unused styles are disposed automatically
     * @deprecated This method has no direct impact on the generated file and is deprecated.
     */
    public void removeStyle(String styleName, boolean onlyIfUnused) {
        if (Helper.isNullOrEmpty(styleName)) {
            throw new StyleException("The style to remove is not defined (no name specified)");
        }
        // noOp / deprecated
    }

    /**
     * Removes the defined worksheet based on its name. If the worksheet is the current or selected worksheet, the current and / or the selected worksheet will be set to the last worksheet of the workbook.
     * If the last worksheet is removed, the selected worksheet will be set to 0 and the current worksheet to null.
     *
     * @param name Name of the worksheet
     * @throws WorksheetException thrown if the name of the worksheet is unknown
     */
    public void removeWorksheet(String name) {
        Optional<Worksheet> worksheetToRemove = this.worksheets.stream().filter(w -> w.getSheetName().equals(name)).findFirst();
        if (worksheetToRemove.isEmpty()) {
            throw new WorksheetException("The worksheet with the name '" + name + "' does not exist.");
        }
        int index = this.worksheets.indexOf(worksheetToRemove.get());
        boolean resetCurrentWorksheet = worksheetToRemove.get() == this.currentWorksheet;
        removeWorksheet(index, resetCurrentWorksheet);
    }

    /**
     * Removes the defined worksheet based on its index. If the worksheet is the current or selected worksheet, the current and / or the selected worksheet will be set to the last worksheet of the workbook.
     * If the last worksheet is removed, the selected worksheet will be set to 0 and the current worksheet to null.
     *
     * @param index Index within the worksheets list
     * @throws WorksheetException thrown if the index is out of range
     */
    public void removeWorksheet(int index) {
        if (index < 0 || index >= worksheets.size()) {
            throw new WorksheetException("The worksheet index " + index + " is out of range");
        }
        boolean resetCurrentWorksheet = worksheets.get(index) == currentWorksheet;
        removeWorksheet(index, resetCurrentWorksheet);
    }

    /**
     * Method to resolve all merged cells in all worksheets. Only the value of the very first cell of the merged cells range will be visible. The other values are still present (set to EMPTY) but will not be stored in the worksheet.<br/>
     * This is an internal method. There is no need to use it.
     *
     * @throws StyleException Thrown if an unreferenced style was in the style sheet
     * @throws RangeException Thrown if the cell range was not found
     */
    public void resolveMergedCells() {
        for (Worksheet worksheet : worksheets) {
            worksheet.resolveMergedCells();
        }
    }

    /**
     * Saves the workbook
     *
     * @throws IOException Throws IOException in case of an error
     */
    public void save() throws IOException {
        XlsxWriter l = new XlsxWriter(this);
        l.save();
    }

    /**
     * Saves the workbook with the defined name
     *
     * @param filename Filename of the saved workbook
     * @throws IOException Thrown in case of an error
     */
    public void saveAs(String filename) throws IOException {
        String backup = this.filename;
        this.filename = filename;
        XlsxWriter l = new XlsxWriter(this);
        l.save();
        this.filename = backup;
    }

    /**
     * Save the workbook to an output stream
     *
     * @param stream Output Stream
     * @throws IOException Thrown in case of an error
     */
    public void saveAsStream(OutputStream stream) throws IOException {
        XlsxWriter l = new XlsxWriter(this);
        l.saveAsStream(stream);
    }

    /**
     * Sets the current worksheet
     *
     * @param name Name of the worksheet
     * @return Returns the current worksheet
     * @throws WorksheetException Thrown if the name of the worksheet is unknown
     */
    public Worksheet setCurrentWorksheet(String name) {
        Optional<Worksheet> worksheet = worksheets.stream().filter(w -> w.getSheetName().equals(name)).findFirst();
        if (worksheet.isEmpty()) {
            throw new WorksheetException("The worksheet with the name '" + name + "' does not exist.");
        } else {
            this.currentWorksheet = worksheet.get();
            this.WS.setCurrentWorksheetInternal(worksheet.get());
        }
        return this.currentWorksheet;
    }

    /**
     * Sets the current worksheet
     *
     * @param worksheetIndex Zero-based worksheet index
     * @return Returns the current worksheet
     * @throws WorksheetException Thrown if the name of the worksheet is unknown
     */
    public Worksheet setCurrentWorksheet(int worksheetIndex) {
        if (worksheetIndex < 0 || worksheetIndex > worksheets.size() - 1) {
            throw new RangeException("The worksheet index " + worksheetIndex + " is out of range");
        }
        currentWorksheet = worksheets.get(worksheetIndex);
        this.WS.setCurrentWorksheetInternal(currentWorksheet);
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
        this.WS.setCurrentWorksheetInternal(worksheet);
    }

    /**
     * Sets the selected worksheet in the output workbook
     *
     * @param name Name of the worksheet
     * @throws WorksheetException Throws a WorksheetException if the name of the worksheet is unknown or if it is hidden
     * @throws RangeException     Throws a RangeException if the index of the worksheet is out of range
     */
    public void setSelectedWorksheet(String name) {
        Optional<Worksheet> worksheet = worksheets.stream().filter(w -> w.getSheetName().equals(name)).findFirst();
        if (worksheet.isEmpty()) {
            throw new WorksheetException("The worksheet with the name '" + name + "' does not exist.");
        }
        this.selectedWorksheet = worksheets.indexOf(worksheet.get());
        validateWorksheets();
    }

    /**
     * Sets the selected worksheet in the output workbook<br>Note: This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this
     *
     * @param worksheetIndex Zero-based worksheet index
     * @throws RangeException     Throws a RangeException if the index of the worksheet is out of range
     * @throws WorksheetException Throws a WorksheetException if the worksheet is hidden
     */
    public void setSelectedWorksheet(int worksheetIndex) {
        if (worksheetIndex < 0 || worksheetIndex > this.worksheets.size() - 1) {
            throw new RangeException("The worksheet index " + worksheetIndex + " is out of range");
        }
        this.selectedWorksheet = worksheetIndex;
        validateWorksheets();
    }

    /**
     * Sets the selected worksheet in the output workbook<br>Note: This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this
     *
     * @param worksheet Worksheet object (must be in the collection of worksheets)
     * @throws WorksheetException Throws a WorksheetException if the worksheet was not found in the worksheet collection or if it is hidden
     */
    public void setSelectedWorksheet(Worksheet worksheet) {
        boolean check = false;
        for (int i = 0; i < this.worksheets.size(); i++) {
            if (this.worksheets.get(i).equals(worksheet)) {
                this.selectedWorksheet = i;
                check = true;
                break;
            }
        }
        if (!check) {
            throw new WorksheetException("The passed worksheet object is not in the worksheet collection.");
        }
        validateWorksheets();
    }

    /**
     * Sets or removes the workbook protection. If protectWindows and protectStructure are both false, the workbook will not be protected
     *
     * @param state            If true, the workbook will be protected, otherwise not
     * @param protectWindows   If true, the windows will be locked if the workbook is protected
     * @param protectStructure If true, the structure will be locked if the workbook is protected
     * @param password         Optional password. If null or empty, no password will be set in case of protection
     */
    public void setWorkbookProtection(boolean state, boolean protectWindows, boolean protectStructure, String password) {
        this.lockWindowsIfProtected = protectWindows;
        this.lockStructureIfProtected = protectStructure;
        this.workbookProtectionPassword = password;
        if (!protectWindows && !protectStructure) {
            this.useWorkbookProtection = false;
        } else {
            this.useWorkbookProtection = state;
        }
    }

    /**
     * Removes the worksheet at the defined index and relocates current and selected worksheet references
     *
     * @param index                 Index within the worksheets list
     * @param resetCurrentWorksheet If true, the current worksheet will be relocated to the last worksheet in the list
     */
    private void removeWorksheet(int index, boolean resetCurrentWorksheet) {
        this.worksheets.remove(index);
        if (!this.worksheets.isEmpty()) {
            for (int i = 0; i < worksheets.size(); i++) {
                this.worksheets.get(i).setSheetID(i + 1);
            }
            if (resetCurrentWorksheet) {
                currentWorksheet = worksheets.get(worksheets.size() - 1);
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
     * Validates the worksheets regarding several conditions that must be met:<br/>
     * - At least one worksheet must be defined<br/>
     * - A hidden worksheet cannot be the selected one<br/>
     * - At least one worksheet must be visible<br/>
     * If one of the conditions is not met, an exception is thrown
     * @apiNote If an import is in progress, these rules are disabled to avoid conflicts by the order of loaded worksheets
     */
    public void validateWorksheets() {
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
                    throw new WorksheetException("The worksheet with the index " + selectedWorksheet + " cannot be set as selected, since it is set hidden");
                }
            }
        }
    }

    /**
     * Gets the next free worksheet ID
     *
     * @return Worksheet ID
     */
    private int getNextWorksheetId() {
        if (this.worksheets.isEmpty()) {
            return 1;
        }
        return this.worksheets.stream().max((w1, w2) -> Integer.compare(w1.getSheetID(), w2.getSheetID())).get().getSheetID() + 1;
    }

    /**
     * Init method called in the constructors
     */
    private void init() {
        this.worksheets = new ArrayList<>();
        this.workbookMetadata = new Metadata();
        this.WS = new Shortener(this);
    }

    /** --------------- NANO - PART --------------- (PICO part is above) */

    /**
     * Loads a workbook from a file
     *
     * @param filename Filename of the workbook
     * @return Workbook object
     * @throws IOException Throws IOException in case of an error
     */
    public static Workbook load(String filename) throws IOException, java.io.IOException {
        return load(filename, null);
    }

    /**
     * Loads a workbook from a file with import options
     *
     * @param filename      Filename of the workbook
     * @param importOptions Import options to override the data types of columns or cells. These options can be used to cope with wrong interpreted data, caused by irregular styles
     * @return Workbook object
     * @throws IOException Throws IOException in case of an error
     */
    public static Workbook load(String filename, ImportOptions importOptions) throws IOException, java.io.IOException {
        XlsxReader r = new XlsxReader(filename, importOptions);
        r.read();
        return r.getWorkbook();
    }

    /**
     * Loads a workbook from an input stream
     *
     * @param stream Stream containing the workbook
     * @return Workbook object
     * @throws IOException Throws IOException in case of an error
     */
    public static Workbook load(InputStream stream) throws IOException, java.io.IOException {
        return load(stream, null);
    }

    /**
     * Loads a workbook from an input stream with import options
     *
     * @param stream        Stream containing the workbook
     * @param importOptions Import options to override the data types of columns or cells. These options can be used to cope with wrong interpreted data, caused by irregular styles
     * @return Workbook object
     * @throws IOException Throws IOException in case of an error
     */
    public static Workbook load(InputStream stream, ImportOptions importOptions) throws IOException, java.io.IOException {
        XlsxReader r = new XlsxReader(stream, importOptions);
        r.read();
        return r.getWorkbook();
    }

    /**
     *  Sets the import state. If an import is in progress, no validity checks on are performed to avoid conflicts by incomplete data (e.g. hidden worksheets)
     * @param state True if an import is in progress, otherwise false
     */
    public void setImportState(boolean state) {
        this.importInProgress = state;
    }



}
