package ch.rabanti.nanoxlsx4j;

import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.internal.FeatureSet;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

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
//privateField
    private int sheetId;
    private String sheetName;
    private FeatureSet features = new FeatureSet();

//getters&setters

    /**
     * Gets the internal ID of the worksheet
     *
     * @return Sheet ID
     */
    public int getSheetId() {
        return sheetId;
    }

    public String getSheetName() {
        return sheetName;
    }

    FeatureSet getFeatures() {
        return this.features;
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

    public Worksheet(String name, int id, Workbook workbook) {
        // TODO implement
    }

//methods

    public void addCellFormula(String formula, int columnNumber, int rowNumber) {
        // TODO implement
    }

    public void addNextCellFormula(String formula) {
        // TODO implement
    }

//staticMethods

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
        // Regex regex = new Regex(@"^(.*?)(\d{1,31})$");
        Matcher match = pattern.matcher(name);
        String prefix = name;
        int number = 1;
        if (match.groupCount() > 1) {
            prefix = match.group(1);
            Optional<Integer> parsedNumber = ParserUtils.tryParseInt(match.group(2));
            if (parsedNumber.isPresent()) {
                number = parsedNumber.get();
            }
            // if this failed, the start number is 0 (parsed number was >max. int32)
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


}
