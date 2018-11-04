/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.*;
import ch.rabanti.nanoxlsx4j.lowLevel.XlsXWriter;
import ch.rabanti.nanoxlsx4j.lowLevel.XlsxReader;
import ch.rabanti.nanoxlsx4j.styles.*;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Class representing a workbook
 * @author Raphael Stoeckli
 */
public class Workbook {
    
// ### P R I V A T E  F I E L D S ###    
    private Worksheet currentWorksheet;
    private String filename;
    private boolean lockStructureIfProtected;
    private boolean lockWindowsIfProtected;
    private int selectedWorksheet;
    private StyleManager styleManager;
    private boolean useWorkbookProtection;
    private Metadata workbookMetadata;
    private String workbookProtectionPassword;
    private List<Worksheet> worksheets;
    
    /**
     * Shortener omits getter and setter to simplify the access (Can throw a WorksheetException if not defined)
     */
    public Shortener WS;
    
// ### G E T T E R S  &  S E T T E R S ###    
    /**
     * Gets the current worksheet
     * @return Current worksheet reference
     */
    public Worksheet getCurrentWorksheet() {
        return currentWorksheet;
    }
    
    
    /**
     * Gets the filename of the workbook
     * @return Filename of the workbook
     */
    public String getFilename() {
        return filename;
    }

    /**
     * Sets the filename of the workbook
     * @param filename Filename of the workbook
     */
    public void setFilename(String filename) {
        this.filename = filename;
    }
    /**
     * Gets the selected worksheet. The selected worksheet is not the current worksheet while design time but the selected sheet in the output file
     * @return Zero-based worksheet index
     */
    public int getSelectedWorksheet() {
        return selectedWorksheet;
    }
    /**
     * Sets the selected worksheet in the output workbook<br>Note: This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this
     * @param worksheet Worksheet object (must be in the collection of worksheets)
     * @throws WorksheetException Throws a WorksheetException if the worksheet was not found in the worksheet collection
     */
    public void setSelectedWorksheet(Worksheet worksheet)
    {
        boolean check = false;
        for (int i = 0; i < this.worksheets.size(); i++)
        {
            if (this.worksheets.get(i).equals(worksheet))
            {
                this.selectedWorksheet = i;
                check = true;
                break;
            }
        }
        if (check == false)
        {
            throw new WorksheetException("MissingReferenceException","The passed worksheet object is not in the worksheet collection.");
        }        
    }
    /**
     * Sets the selected worksheet in the output workbook<br>Note: This method does not set the current worksheet while design time. Use SetCurrentWorksheet instead for this
     * @param worksheetIndex Zero-based worksheet index
     * @throws RangeException Throws a RangeException if the index of the worksheet is out of range
     */
    public void setSelectedWorksheet(int worksheetIndex)
    {
        if (worksheetIndex < 0 || worksheetIndex > this.worksheets.size() - 1)
        {
            throw new RangeException("OutOfRangeException","The worksheet index " + Integer.toString(worksheetIndex) + " is out of range");
        }
        this.selectedWorksheet = worksheetIndex;
    }
    
    /**
     * Gets the meta data object of the workbook
     * @return Meta data object
     */
    public Metadata getWorkbookMetadata() {
        return workbookMetadata;        
    }

    /**
     * Sets the meta data object of the workbook
     * @param workbookMetadata Meta data object
     */
    public void setWorkbookMetadata(Metadata workbookMetadata) {
        this.workbookMetadata = workbookMetadata;
    }    

    /**
     * Sets whether the workbook is protected
     * @param useWorkbookProtection If true, the workbook is protected otherwise not
     */
    public void setWorkbookProtection(boolean useWorkbookProtection) {
        this.useWorkbookProtection = useWorkbookProtection;
    }

    /**
     * Gets the password used for workbook protection
     * @return Password (UTF-8)
     */
    public String getWorkbookProtectionPassword() {
        return workbookProtectionPassword;
    }
    /**
     * Gets the list of worksheets in the workbook
     * @return List of worksheet objects
     */
    public List<Worksheet> getWorksheets() {
        return worksheets;
    }

    /**
     * Gets whether the structure are locked if workbook is protected
     * @return True if the structure is locked when the workbook is protected
     */
    public boolean isStructureLockedIfProtected() {
        return lockStructureIfProtected;
    }
    /**
     * Gets whether the windows are locked if workbook is protected
     * @return True if the windows are locked when the workbook is protected
     */
    public boolean isWindowsLockedIfProtected() {
        return lockWindowsIfProtected;
    }    
    /**
     * Gets whether the workbook is protected
     * @return If true, the workbook is protected otherwise not
     */
    public boolean isWorkbookProtectionUsed() {
        return useWorkbookProtection;
    }    

    /**
     * Gets the style manager of this workbook
     * @return Style manager object
     */
    public StyleManager getStyleManager() {
        return styleManager;
    }
    
    
// ### C O N S T R U C T O R S ###
    /**
     * Constructor with additional parameter to create a default worksheet. This constructor can be used to define a workbook that is saved as stream
     * @param createWorksheet If true, a default worksheet with the name 'Sheet1' will be created and set as current worksheet
     */
    public Workbook(boolean createWorksheet)
    { 
            init();
            if (createWorksheet == true)
            {
                addWorksheet("Sheet1");
            }       
    }

    /**
     * Constructor with additional parameter to create a default worksheet with the specified name. This constructor can be used to define a workbook that is saved as stream
     * @param sheetName Name of the first worksheet. The name will be sanitized automatically according to the specifications of Excel
     * @throws FormatException Thrown if the worksheet name contains illegal characters
     */
    public Workbook(String sheetName)
    {
        init();
        addWorksheet(sheetName, true);
    }

    /**
     * Constructor with filename ant the name of the first worksheet. This constructor can be used to define a workbook that is saved as stream
     * @param filename Filename of the workbook
     * @param sheetName Name of the first worksheet. The name will be sanitized automatically according to the specifications of Excel
     * @throws FormatException Thrown if the worksheet name contains illegal characters
     */
    public Workbook(String filename, String sheetName)
    {
        init();   
        this.filename = filename;
        addWorksheet(sheetName, true);
    }

    /**
     * Constructor with filename ant the name of the first worksheet
     * @param filename Filename of the workbook
     * @param sheetName Name of the first worksheet
     * @param sanitizeSheetName If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel
     */
    public Workbook(String filename, String sheetName, boolean sanitizeSheetName)
    {
        init();   
        this.filename = filename;
        addWorksheet(Worksheet.sanitizeWorksheetName(sheetName, this));
    }    
    
// ### M E T H O D S ###


    /**
     * Adds a style to the style manager
     * @param style Style to add
     * @return The managed style of the style manager
     */
    public Style addStyle(Style style)
    {
       return this.styleManager.addStyle(style);
    }
    
    /**
     * Adds a style component to a style
     * @param baseStyle Style to append a component
     * @param newComponent Component to add to the baseStyle 
     * @return The managed style of the style manager
     */
    public Style addStyleComponent(Style baseStyle, AbstractStyle newComponent)
    {
        
        if (newComponent instanceof Border)
        {
            baseStyle.setBorder((Border)newComponent);
        }
        else if (newComponent instanceof CellXf)
        {
            baseStyle.setCellXf((CellXf)newComponent);
        }
        else if (newComponent instanceof Fill)
        {
            baseStyle.setFill((Fill)newComponent);
        }
        else if (newComponent instanceof Font)
        {
            baseStyle.setFont((Font)newComponent);
        }
        else if (newComponent instanceof NumberFormat)
        {
            baseStyle.setNumberFormat((NumberFormat)newComponent);
        }
        return this.styleManager.addStyle(baseStyle);
    }
    
    /**
     * Adding a new Worksheet. The new worksheet will be defined as current worksheet
     * @param name Name of the new worksheet
     * @throws WorksheetException Thrown if the name of the worksheet already exists
     * @throws FormatException Thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31)
     */
    public void addWorksheet(String name)
    {
        for (int i = 0; i < this.worksheets.size(); i++)
        {
            if (this.worksheets.get(i).getSheetName().equals(name))
            {
                throw new WorksheetException("WorksheetNameAlreadyExistsException","The worksheet with the name '" + name + "' already exists.");
            }
        }
        int number = this.worksheets.size() + 1;
        Worksheet newWs = new Worksheet(name, number, this);
        this.currentWorksheet = newWs;
        this.worksheets.add(newWs);
        this.WS.setCurrentWorksheet(this.currentWorksheet);
    }
    
    /**
     * Adding a new Worksheet with a sanitizing option. The new worksheet will be defined as current worksheet
     * @param name Name of the new worksheet
     * @param sanitizeSheetName If true, the name of the worksheet will be sanitized automatically according to the specifications of Excel
     * @throws WorksheetException Thrown if the name of the worksheet already exists and sanitizeSheetName is false
     * @throws FormatException Thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31) and sanitizeSheetName is false
     */
    public void addWorksheet(String name, boolean sanitizeSheetName)
    {
        if (sanitizeSheetName == true)
        {
            String sanitized = Worksheet.sanitizeWorksheetName(name, this);
            addWorksheet(sanitized);
        }
        else
        {
            addWorksheet(name);
        }
    }
    
    /**
     * Adding a new Worksheet. The new worksheet will be defined as current worksheet
     * @param worksheet Prepared worksheet object
     * @throws WorksheetException Thrown if the name of the worksheet already exists
     * @throws FormatException Thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31
     */
    public void addWorksheet(Worksheet worksheet)
    {
        for (int i = 0; i < this.worksheets.size(); i++)
        {
            if (this.worksheets.get(i).getSheetName().equals(worksheet.getSheetName()))
            {
                throw new WorksheetException("WorksheetNameAlreadyExistsException","The worksheet with the name '" + worksheet.getSheetName() + "' already exists.");
            }
        }
        int number = this.worksheets.size() + 1;
        worksheet.setSheetID(number);
        worksheet.setWorkbookReference(this);
        this.currentWorksheet = worksheet;
        this.worksheets.add(worksheet);
    }    
    
    /**
     * Init method called in the constructors
     */
    private void init()
    {
        this.worksheets = new ArrayList<>();
        this.styleManager = new StyleManager();
        this.styleManager.addStyle(new Style("default", 0, true));
        Style borderStyle = new Style("default_border_style", 1, true);
        borderStyle.setBorder(BasicStyles.DottedFill_0_125().getBorder());
        borderStyle.setFill(BasicStyles.DottedFill_0_125().getFill());
        this.styleManager.addStyle(borderStyle);
        this.workbookMetadata = new Metadata();
        this.WS = new Shortener();
    }
        
    
    /**
    * Removes the passed style from the style sheet
    * @param style Style to remove
    * @throws StyleException Thrown if the style is not defined in the style sheet
    */
    public void removeStyle(Style style)
    {
        removeStyle(style, false);
    }
    
    /**
    * Removes the defined style from the style sheet of the workbook
    * @param styleName Name of the style to be removed
    * @throws StyleException Thrown if the style is not defined in the style sheet
    */
    public void removeStyle(String styleName)
    {
        removeStyle(styleName, false);
    }    

    /**
    * Removes the defined style from the style manager of the workbook
    * @param style Style to remove
    * @param onlyIfUnused If true, the style will only be removed if not used in any cell
    * @throws StyleException Thrown if the style is not defined in the style sheet
    */
    public void removeStyle(Style style, boolean onlyIfUnused)
    {
        if (style == null)
        {
            throw new StyleException("MissingReferenceException","The style to remove is not defined");
        }
        removeStyle(style.getName(), onlyIfUnused);
    }
    
    /**
     * Removes the defined style from the style manager of the workbook
     * @param styleName Name of the style to remove
     * @param onlyIfUnused If true, the style will only be removed if not used in any cell
     */
    public void removeStyle(String styleName, boolean onlyIfUnused)
    {
        if (Helper.isNullOrEmpty(styleName))
        {
            throw new StyleException("MissingReferenceException","The style to remove is not defined (no name specified)");
        }
        
        if (onlyIfUnused == true)
        {
                boolean styleInUse = false;
                Cell cell;
                for(int i = 0; i < this.worksheets.size(); i++)
                {
                    for(Object item: this.worksheets.get(i).getCells().entrySet())
                    {
                        cell = (Cell)item;
                        if (cell.getCellStyle() == null) { continue; }
                        if (cell.getCellStyle().getName().equals(styleName))
                        {
                            styleInUse = true;
                            break;
                        }
                    }
                    if (styleInUse == true)
                    {
                        break;
                    }
                }
                if (styleInUse == false)
                {
                    this.styleManager.removeStyle(styleName);
                }
        }
        else
        {
            this.styleManager.removeStyle(styleName);
        }
    }           
    
    /**
     * Removes the defined worksheet
     * @param name Name of the worksheet
     * @throws WorksheetException Thrown if the name of the worksheet is unknown
     */
    public void removeWorksheet(String name)
    {
        boolean exists = false;
        boolean resetCurrent = false;
        int index = 0;
        for (int i = 0; i < this.worksheets.size(); i++)
        {
            if (this.worksheets.get(index).getSheetName().equals(name))
            {
                index = i;
                exists = true;
                break;
            }
        }
        if (exists == false)
        {
            throw new WorksheetException("MissingReferenceException","The worksheet with the name '" + name + "' does not exist.");
        }
        if (this.worksheets.get(index).getSheetName().equals(this.currentWorksheet.getSheetName()) )
        {
            resetCurrent = true;
        }
        this.worksheets.remove(index);
        if (this.worksheets.size() > 0)
        {
            for (int i = 0; i < this.worksheets.size(); i++)
            {
                this.worksheets.get(index).setSheetID(i + 1);
                if (resetCurrent == true && i == 0)
                {
                    this.currentWorksheet = this.worksheets.get(i);
                }
            }
        }
        else
        {
            this.currentWorksheet = null;
        }
        if (this.selectedWorksheet > this.worksheets.size() - 1)
        {
            this.selectedWorksheet = this.worksheets.size() - 1;
        }
    }
    
    /**
     * Method to resolve all merged cells in all worksheets. Only the value of the very first cell of the locked cells range will be visible. The other values are still present (set to EMPTY) but will not be stored in the worksheet.
     * @throws StyleException Thrown if an unreferenced style was in the style sheet
     * @throws RangeException Thrown if the cell range was not found
     */
    public void resolveMergedCells()
    {
        Style mergeStyle = BasicStyles.MergeCellStyle();
        int pos;
        List<Address> addresses;
        Cell cell;
        Worksheet sheet;
        Address address;
        Iterator<Map.Entry<String, Range>> itr;
        Map.Entry<String, Range> range;
        for (int i = 0; i < this.worksheets.size(); i++)
        {
            sheet = this.worksheets.get(i);
            itr = sheet.getMergedCells().entrySet().iterator();
            while (itr.hasNext())
            {
                range = itr.next();
                pos = 0;
                addresses = Cell.getCellRange(range.getValue().StartAddress, range.getValue().EndAddress);
                for (int j = 0; j < addresses.size(); j++)
                {
                    address = addresses.get(j);
                    if (sheet.getCells().containsKey(address.toString()) == false)
                    {
                        cell = new Cell();
                        cell.setDataType(Cell.CellType.EMPTY);
                        cell.setRowNumber(address.Row);
                        cell.setColumnNumber(address.Column);
                        sheet.addCell(cell, cell.getColumnNumber(), cell.getRowNumber());
                    }
                    else
                    {
                        cell = sheet.getCells().get(address.toString());
                    }
                    if (pos != 0)
                    {
                        cell.setDataType(Cell.CellType.EMPTY);
                        cell.setStyle(mergeStyle);
                    }
                    pos++;
                }
            }
        }
    }

    /**
    * Saves the workbook
    * @throws IOException Throws IOException in case of an error
    */
    public void save() throws IOException
    {
        XlsXWriter l = new XlsXWriter(this);
        l.save();
    }
    
    /**
    * Saves the workbook with the defined name
    * @param filename Filename of the saved workbook
    * @throws IOException Thrown in case of an error
    */
    public void saveAs(String filename) throws IOException
    {
        String backup = this.filename;
        this.filename = filename;
        XlsXWriter l = new XlsXWriter(this);
        l.save();
        this.filename = backup;
    }
    
    /**
     * Save the workbook to a output stream
     * @param stream Output Stream
     * @throws IOException Thrown in case of an error
     */
    public void saveAsStream(OutputStream stream) throws IOException
    {
        XlsXWriter l = new XlsXWriter(this);
        l.saveAsStream(stream);
    }
    
    /**
     * Sets the current worksheet
     * @param name Name of the worksheet
     * @return Returns the current worksheet
     * @throws WorksheetException Thrown if the name of the worksheet is unknown
     */
    public Worksheet setCurrentWorksheet(String name)
    {
        boolean exists = false;
        for(int i = 0; i < this.worksheets.size(); i++)
        {
            if (this.worksheets.get(i).getSheetName().equals(name))
            {
                this.currentWorksheet = this.worksheets.get(i);
                exists = true;
                break;
            }
        }
        if (exists == false)
        {
            throw new WorksheetException("MissingReferenceException","The worksheet with the name '" + name + "' does not exist.");
        }
        this.WS.setCurrentWorksheet(this.currentWorksheet);
        return this.currentWorksheet;
    }
    /**
     * Sets or removes the workbook protection. If protectWindows and protectStructure are both false, the workbook will not be protected
     * @param state If true, the workbook will be protected, otherwise not
     * @param protectWindows If true, the windows will be locked if the workbook is protected
     * @param protectStructure If true, the structure will be locked if the workbook is protected
     * @param password Optional password. If null or empty, no password will be set in case of protection
     */
    public void setWorkbookProtection(boolean state, boolean protectWindows, boolean protectStructure, String password)
    {
        this.lockWindowsIfProtected = protectWindows;
        this.lockStructureIfProtected = protectStructure;
        this.workbookProtectionPassword = password;
        if (protectWindows == false && protectStructure == false)
        {
            this.useWorkbookProtection = false;
        }
        else
        {
            this.useWorkbookProtection = state;
        }
    }

    /** --------------- NANO - PART --------------- (PICO part is above) */

    /**
     * Loads a workbook from a file
     * @param filename Filename of the workbook
     * @return Workbook object
     * @throws IOException Throws IOException in case of an error
     */
    public static Workbook load(String filename) throws IOException
    {
        XlsxReader r = new XlsxReader(filename);
        r.read();
        return r.getWorkbook();
    }

    /**
     * Loads a workbook from a input stream
     * @param stream Stream containing the workbook
     * @return Workbook object
     * @throws IOException Throws IOException in case of an error
     */
    public static Workbook load(InputStream stream) throws IOException
    {
        XlsxReader r = new XlsxReader(stream);
        r.read();
        return r.getWorkbook();
    }
    
    
        
}
