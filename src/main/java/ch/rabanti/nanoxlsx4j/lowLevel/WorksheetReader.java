/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.ImportOptions;

import java.io.InputStream;
import java.time.LocalTime;
import java.util.*;

import static ch.rabanti.nanoxlsx4j.Cell.CellType.DATE;
import static ch.rabanti.nanoxlsx4j.Cell.CellType.TIME;

/**
 * Class representing a reader for worksheets of XLSX files
 *
 * @author Raphael Stoeckli
 */
public class WorksheetReader {

    private int worksheetNumber;
    private final String name;
    private final Map<String, Cell> data;
    private final SharedStringsReader sharedStrings;
    private final Map<String, String> styleAssignment = new HashMap<>();
    private final ImportOptions importOptions;
    private List<String> dateStyles;
    private List<String> timeStyles;

    /**
     * Gets the assignment of resolved styles to cell addresses
     *
     * @return Maps of cell address-style number tuples
     */
    public Map<String, String> getStyleAssignment() {
        return styleAssignment;
    }

    /**
     * Gets the name of the worksheet
     *
     * @return Name of the worksheet
     */
    public String getName() {
        return name;
    }

    /**
     * Gets the number of the worksheet
     *
     * @return Number of the worksheet
     */
    public int getWorksheetNumber() {
        return worksheetNumber;
    }

    /**
     * Sets the number of the worksheet
     *
     * @param worksheetNumber Worksheet number
     */
    public void setWorksheetNumber(int worksheetNumber) {
        this.worksheetNumber = worksheetNumber;
    }

    /**
     * Gets the data of the worksheet as Hashmap of cell address-cell object tuples
     *
     * @return Hashmap of cell address-cell object tuples
     */
    public Map<String, Cell> getData() {
        return data;
    }

    /**
     * Constructor with parameters
     *
     * @param sharedStrings        SharedStringsReader object
     * @param name                 Worksheet name
     * @param number               Worksheet number
     * @param styleReaderContainer Resolved styles, used to determine dates or times
     */
    public WorksheetReader(SharedStringsReader sharedStrings, String name, int number, StyleReaderContainer styleReaderContainer) {
        data = new HashMap<>();
        this.name = name;
        this.worksheetNumber = number;
        this.sharedStrings = sharedStrings;
        this.importOptions = null;
        processStyles(styleReaderContainer);
    }

    /**
     * Constructor with parameters and import options
     *
     * @param sharedStrings        SharedStringsReader object
     * @param name                 Worksheet name
     * @param number               Worksheet number
     * @param styleReaderContainer Resolved styles, used to determine dates or times
     */
    public WorksheetReader(SharedStringsReader sharedStrings, String name, int number, StyleReaderContainer styleReaderContainer, ImportOptions options) {
        data = new HashMap<>();
        this.name = name;
        this.worksheetNumber = number;
        this.sharedStrings = sharedStrings;
        this.importOptions = options;
        processStyles(styleReaderContainer);
    }

    /**
     * Determine which of the resolved styles are either to define a time or a date. Stores also the styles into a map
     *
     * @param styleReaderContainer Resolved styles from the style reader
     */
    private void processStyles(StyleReaderContainer styleReaderContainer) {
        this.dateStyles = new ArrayList<>();
        this.timeStyles = new ArrayList<>();
        for (int i = 0; i < styleReaderContainer.getStyleCount(); i++) {
            StyleReaderContainer.StyleResult result = styleReaderContainer.evaluateDateTimeStyle(i, true);
            if (result.isDateStyle()) {
                this.dateStyles.add(Integer.toString(i));
            }
            if (result.isTimeStyle()) {
                this.timeStyles.add(Integer.toString(i));
            }
        }
    }

    /**
     * Gets whether the specified column exists in the data
     *
     * @param columnAddress Column address as string
     * @return Column address as string
     */
    public boolean hasColumn(String columnAddress) {
        if (Helper.isNullOrEmpty(columnAddress)) {
            return false;
        }
        int columnNumber = Cell.resolveColumn(columnAddress);
        for (Map.Entry<String, Cell> cell : data.entrySet()) {
            if (cell.getValue().getColumnNumber() == columnNumber) {
                return true;
            }
        }
        return false;
    }

    /**
     * Gets whether the passed row (represented as list of cell objects) contains the specified column numbers
     *
     * @param cells         List of cell objects to check
     * @param columnNumbers Array of column numbers
     * @return True if all column numbers were found, otherwise false
     */
    public boolean rowHasColumns(List<Cell> cells, int[] columnNumbers) {
        if (columnNumbers == null || cells == null) {
            return false;
        }
        int len = columnNumbers.length;
        int len2 = cells.size();
        int j;
        boolean match;
        if (len < 1 || len2 < 1) {
            return false;
        }
        for (int columnNumber : columnNumbers) {
            match = false;
            for (j = 0; j < len2; j++) {
                if (cells.get(j).getColumnNumber() == columnNumber) {
                    match = true;
                    break;
                }
            }
            if (!match) {
                return false;
            }
        }
        return true;
    }

    /**
     * Gets the number of rows
     *
     * @return Number of rows
     */
    public int getRowCount() {
        int count = -1;
        for (Map.Entry<String, Cell> cell : data.entrySet()) {
            if (cell.getValue().getRowNumber() > count) {
                count = cell.getValue().getRowNumber();
            }
        }
        return count + 1;
    }

    /**
     * Gets a row as list of cell objects
     *
     * @param rowNumber Row number
     * @return List of cell objects
     */
    public List<Cell> getRow(int rowNumber) {
        List<Cell> list = new ArrayList<>();
        for (Map.Entry<String, Cell> cell : data.entrySet()) {
            if (cell.getValue().getRowNumber() == rowNumber) {
                list.add(cell.getValue());
            }
        }
        list.sort(Comparator.comparingInt(Cell::getColumnNumber));
        return list;
    }

    /**
     * Reads the XML file form the passed stream and processes the worksheet data
     *
     * @param stream Stream of the XML file
     */
    public void read(InputStream stream) throws java.io.IOException {
        data.clear();
        try {
            XmlDocument xr = new XmlDocument();
            xr.load(stream);
            XmlDocument.XmlNodeList rows = xr.getDocumentElement().getElementsByTagName("row", true);
            for (int i = 0; i < rows.size(); i++) {
                XmlDocument.XmlNode row = rows.get(i);
                if (row.hasChildNodes()) {
                    for (XmlDocument.XmlNode rowChild : row.getChildNodes()) {
                        readCell(rowChild);
                    }
                }
            }
        } catch (Exception ex) {
            if (stream != null) {
                stream.close();
            }
        }
    }

    /**
     * Reads one cell in a worksheet
     *
     * @param rowChild Current child row as XmlNode
     */
    private void readCell(XmlDocument.XmlNode rowChild) {
        String type = "s";
        String styleNumber = "";
        String address = "A1";
        String value = "";
        String formula = null;
        if (rowChild.getName().equalsIgnoreCase("c")) {
            address = rowChild.getAttribute("r"); // Mandatory
            type = rowChild.getAttribute("t"); // can be null if not existing
            styleNumber = rowChild.getAttribute("s"); // can be null
            if (rowChild.hasChildNodes()) {
                for (XmlDocument.XmlNode valueNode : rowChild.getChildNodes()) {
                    if (valueNode.getName().equalsIgnoreCase("v")) {
                        value = valueNode.getInnerText();
                    }
                    if (valueNode.getName().equalsIgnoreCase("f")) {
                        formula = valueNode.getInnerText();
                    }
                }
            }
        }
        resolveCellData(address, type, value, styleNumber, formula);
    }

    /**
     * Resolves the data of a read cell, transforms it into a cell object and adds it to the data
     *
     * @param addressString Address of the cell
     * @param type          Expected data type
     * @param value         Raw value as string
     * @param styleNumber   Style number as string (can be null)
     * @param formula       Formula as string (can be null; data type determines whether value or formula is used)
     */
    private void resolveCellData(String addressString, String type, String value, String styleNumber, String formula) {
        Address address = new Address(addressString);
        String key = addressString.toUpperCase();
        this.styleAssignment.put(key, styleNumber);
        if (importOptions == null) {
            this.data.put(key, autoResolveCellData(address, type, value, styleNumber, formula));
        } else {
            this.data.put(key, resolveCellDataConditionally(address, type, value, styleNumber, formula));
        }
    }

    /**
     * Resolves the data of a read cell with conditions of import options, transforms it into a cell object
     *
     * @param address     Address of the cell
     * @param type        Expected data type
     * @param value       Raw value as string
     * @param styleNumber Style number as string (can be null)
     * @param formula     Formula as string (can be null; data type determines whether value or formula is used)
     * @return The resolved Cell
     */
    private Cell resolveCellDataConditionally(Address address, String type, String value, String styleNumber, String formula) {
        if (address.Row < importOptions.getEnforcingStartRowNumber()) {
            return autoResolveCellData(address, type, value, styleNumber, formula); // Skip enforcing
        }
        if (importOptions.getEnforcedColumnTypes().containsKey(address.Column)) {
            ImportOptions.ColumnType importType = importOptions.getEnforcedColumnTypes().get(address.Column);
            switch (importType) {
                case bool:
                    return getBooleanValue(value, address);
                case date:
                    if (importOptions.isEnforceDateTimesAsNumbers()) {
                        return getNumericValue(value, address);
                    } else {
                        return getDateTimeValue(value, address, DATE);
                    }
                case time:
                    if (importOptions.isEnforceDateTimesAsNumbers()) {
                        return getNumericValue(value, address);
                    } else {
                        return getDateTimeValue(value, address, TIME);
                    }
                case numeric:
                    return getNumericValue(value, address);
                case string:
                    return getStringValue(value, address);
                default:
                    return autoResolveCellData(address, type, value, styleNumber, formula);
            }
        } else {
            return autoResolveCellData(address, type, value, styleNumber, formula);
        }
    }

    /**
     * Resolves the data of a read cell automatically, transforms it into a cell object
     *
     * @param address     Address of the cell
     * @param type        Expected data type
     * @param value       Raw value as string
     * @param styleNumber Style number as string (can be null)
     * @param formula     Formula as string (can be null; data type determines whether value or formula is used)
     * @return The resolved Cell
     */
    private Cell autoResolveCellData(Address address, String type, String value, String styleNumber, String formula) {
        if (type.equals("s")) // string (declared)
        {
            return getStringValue(value, address);
        } else if (type.equals("b")) // boolean
        {
            return getBooleanValue(value, address);
        } else if (dateStyles.contains(styleNumber))  // date (priority)
        {
            return getDateTimeValue(value, address, DATE);
        } else if (timeStyles.contains(styleNumber)) // time
        {
            return getDateTimeValue(value, address, TIME);
        } else if (Helper.isNullOrEmpty(type) || type.equals("n")) // try numeric if not parsed as date or time, before numeric
        {
            return getNumericValue(value, address);
        } else if (formula != null) // formula before string
        {
            return new Cell(formula, Cell.CellType.FORMULA, address);
        } else // fall back to sting
        {
            return getStringValue(value, address);
        }
    }

    /**
     * Parses the numeric value of a raw cell. The order of possible number types are: int, float double. If nothing applies, a string is returned
     *
     * @param raw Raw value as string
     * @return Cell of the type int, float, double or string as fall-back type<
     */
    private static Cell getNumericValue(String raw, Address address) {
        try {
            int i = Integer.parseInt(raw);
            return new Cell(i, Cell.CellType.NUMBER, address);
        } catch (Exception ignored) {
        }
        try {
            long l = Long.parseLong(raw);
            return new Cell(l, Cell.CellType.NUMBER, address);
        } catch (Exception ignored) {
        }
        try {
            float f = Float.parseFloat(raw);
            return new Cell(f, Cell.CellType.NUMBER, address);
        } catch (Exception ignored) {
        }
        try {
            double d = Double.parseDouble(raw);
            return new Cell(d, Cell.CellType.NUMBER, address);
        } catch (Exception dummy) {
            return new Cell(raw, Cell.CellType.STRING, address);
        }
    }



    /**
     * Parses the string value of a raw cell. May take the value from the shared string table, if available
     *
     * @param raw     Raw value as string
     * @param address Address of the cell
     * @return Cell of the type string
     */
    private Cell getStringValue(String raw, Address address) {
        try {
            int stringId = Integer.parseInt(raw);
            String resolvedString = sharedStrings.getString(stringId);
            if (resolvedString == null) {
                return new Cell(raw, Cell.CellType.STRING, address);
            } else {
                return new Cell(resolvedString, Cell.CellType.STRING, address);
            }
        } catch (Exception ex) {
            return new Cell(raw, Cell.CellType.STRING, address);
        }
    }


    /**
     * Parses the boolean value of a raw cell
     *
     * @param raw     Raw value as string
     * @param address Address of the cell
     * @return Cell of the type boolean or the defined fall-back type
     */
    private static Cell getBooleanValue(String raw, Address address) {
        if (raw.equals("0")) {
            return new Cell(false, Cell.CellType.BOOL, address);
        } else if (raw.equals("1")) {
            return new Cell(true, Cell.CellType.BOOL, address);
        } else {
            try {
                boolean bool = Boolean.parseBoolean(raw);
                return new Cell(bool, Cell.CellType.BOOL, address);
            } catch (Exception ex) {
                return new Cell(raw, Cell.CellType.STRING, address);
            }
        }
    }

    /**
     * Parses the date (Date) or time (LocalTime) value of a raw cell. If the value is numeric, but out of range of a OAdate, a numeric value will be returned instead. If invalid, the string representation will be returned.
     *
     * @param raw     Raw value as string
     * @param address Address of the cell
     * @param type    Type of the value to be converted: Valid values are DATE and TIME
     * @return CellResolverTuple with information about the validity and resolved data
     * @throws IllegalArgumentException Thrown if an unsupported type was passed
     */
    private static Cell getDateTimeValue(String raw, Address address, Cell.CellType type) {
        try {
            double d = Double.parseDouble(raw);
            if (d < XlsxWriter.MIN_OADATE_VALUE || d > XlsxWriter.MAX_OADATE_VALUE) {
                return new Cell(d, Cell.CellType.NUMBER, address); // Invalid OAdate == plain number
            } else {
                switch (type) {
                    case DATE:
                        Date date = Helper.getDateFromOA(d);
                        return new Cell(date, DATE, address);
                    case TIME:
                        LocalTime time = LocalTime.ofSecondOfDay((long) (d * 86400));
                        return new Cell(time, TIME, address);
                    default:
                        throw new IllegalArgumentException("The defined type is not supported to be uses as date or time");
                }
            }
        } catch (Exception e) {
            return new Cell(raw, Cell.CellType.STRING, address);
        }
    }

}