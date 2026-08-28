package ch.rabanti.nanoxlsx4j;

import java.util.ArrayList;
import java.util.List;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.internal.AuxiliaryData;
import ch.rabanti.nanoxlsx4j.internal.FeatureSet;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Password;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

public class Workbook {

    private List<Worksheet> worksheets;
    private FeatureSet features = new FeatureSet();
    private Metadata workbookMetadata;
    private final List<DefinedName> definedNames = new ArrayList<>();
    private Password workbookProtectionPassword;
    private Shortener shortener;
    private AuxiliaryData auxiliaryData;
    private Worksheet currentWorksheet;

    public List<Worksheet> getWorksheets() {
        return worksheets;
    }

    /**
     * Gets the feature set of the workbook
     *
     * @return Feature set
     */
    public FeatureSet getFeatures() {
        return features;
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
     * Adding a new Worksheet. The new worksheet will be defined as current worksheet
     *
     * @param name Name of the new worksheet
     * @throws WorksheetException Throws a WorksheetException if the name of the worksheet already exists
     * @throws FormatException    Thrown if the name contains illegal characters or is out of range (length between 1 an
     *                            31 characters)
     */
    public void AddWorksheet(String name) {
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
            AddWorksheet(sanitized);
        } else {
            AddWorksheet(name);
        }
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
