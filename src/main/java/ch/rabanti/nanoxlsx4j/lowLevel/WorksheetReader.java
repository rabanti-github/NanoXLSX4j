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

import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoField;
import java.time.temporal.TemporalAccessor;
import java.util.*;

import static ch.rabanti.nanoxlsx4j.Cell.CellType.DATE;
import static ch.rabanti.nanoxlsx4j.Cell.CellType.TIME;

/**
 * Class representing a reader for worksheets of XLSX files
 *
 * @author Raphael Stoeckli
 */
public class WorksheetReader {

    private static final double ZERO_THRESHOLD = 0.000001d;
    private static Calendar CALENDAR = Calendar.getInstance();

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
     * Reads the XML file form the passed stream and processes the worksheet data
     *
     * @param stream Stream of the XML file
     * @throws IOException thrown if the document could not be read
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
            throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
        } finally {
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
        Cell tempCell = autoResolveCellData(address, type, value, styleNumber, formula);
        switch (importOptions.getGlobalEnforcingType()) {
            case AllNumbersToDouble:
                if (tempCell.getDataType().equals(Cell.CellType.NUMBER)) {
                    Number number = (Number) tempCell.getValue();
                    return new Cell(number.doubleValue(), tempCell.getDataType(), address);
                } else if (tempCell.getDataType().equals(Cell.CellType.BOOL)) {
                    double tempDouble = ((boolean) tempCell.getValue()) ? 1d : 0d;
                    return new Cell(tempDouble, Cell.CellType.NUMBER, address);
                } else if (tempCell.getDataType().equals(DATE) || tempCell.getDataType().equals(TIME)) {
                    return new Cell(Double.valueOf(value), Cell.CellType.NUMBER, address);
                } else if (tempCell.getDataType().equals(Cell.CellType.STRING)) {
                    Number number = tryParseDecimal(tempCell.getValue().toString());
                    if (number != null) {
                        return new Cell(number.doubleValue(), Cell.CellType.NUMBER, address);
                    }
                }
                return tempCell;
            case AllNumbersToInt:
                if (tempCell.getDataType().equals(Cell.CellType.NUMBER)) {
                    Number number = (Number) tempCell.getValue();
                    return new Cell((int) Math.round(number.doubleValue()), tempCell.getDataType(), address);
                } else if (tempCell.getDataType().equals(Cell.CellType.BOOL)) {
                    int tempInt = ((boolean) tempCell.getValue()) ? 1 : 0;
                    return new Cell(tempInt, Cell.CellType.NUMBER, address);
                } else if (tempCell.getDataType().equals(DATE) || tempCell.getDataType().equals(TIME)) {
                    return new Cell((int) Math.round(Double.parseDouble(value)), Cell.CellType.NUMBER, address);
                } else if (tempCell.getDataType().equals(Cell.CellType.STRING)) {
                    Number number = tryParseDecimal(tempCell.getValue().toString());
                    if (number != null) {
                        return new Cell(number.intValue(), Cell.CellType.NUMBER, address);
                    }
                }
                return tempCell;
            case EverythingToString:
                return getEnforcedStingValue(address, type, value, styleNumber, formula, importOptions);
        }
        if (Helper.isNullOrEmpty(value) && Helper.isNullOrEmpty(formula)) {
            if (importOptions.isEnforceEmptyValuesAsString()) {
                return new Cell("", Cell.CellType.STRING, address);
            } else {
                return new Cell(null, Cell.CellType.EMPTY, address);
            }
        }
        if (importOptions.getEnforcedColumnTypes().containsKey(address.Column)) {

            ImportOptions.ColumnType importType = importOptions.getEnforcedColumnTypes().get(address.Column);
            if (type != null && type.equals("s")) {
                // Resolve shared string first
                value = resolveSharedString(value);
            }
            switch (importType) {
                case Bool:
                    tempCell = getBooleanValue(value, address);
                    if (tempCell == null) {
                        return autoResolveCellData(address, type, value, styleNumber, formula);
                    }
                    return tempCell;
                case Date:
                    if (importOptions.isEnforceDateTimesAsNumbers()) {
                        if (!Helper.isNullOrEmpty(formula)) {
                            return tempCell;
                        }
                        return getNumericValue(value, address, styleNumber);
                    } else {
                        return getDateTimeValue(value, address, DATE, importOptions, type);
                    }
                case Time:
                    if (importOptions.isEnforceDateTimesAsNumbers()) {
                        if (!Helper.isNullOrEmpty(formula)) {
                            return tempCell;
                        }
                        return getNumericValue(value, address, styleNumber);
                    } else {
                        return getDateTimeValue(value, address, TIME, importOptions, type);
                    }
                case Numeric:
                    return getNumericValue(value, address);
                case Double:
                    return getDoubleValue(value, address);
                case String:
                    return getStringValue(value, address, type, styleNumber, importOptions);
            }
        }
        return autoResolveCellData(address, type, value, styleNumber, formula);
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
        if (type != null && type.equals("s")) // string (declared)
        {
            return getStringValue(value, address, type, null, null);
        } else if (type != null && type.equals("b")) // boolean
        {
            Cell tempCell = getBooleanValue(value, address);
            if (tempCell == null) {
                return autoResolveCellData(address, null, value, styleNumber, formula);
            }
            return tempCell;
        } else if (dateStyles.contains(styleNumber))  // date (priority)
        {
            return getDateTimeValue(value, address, DATE, importOptions);
        } else if (timeStyles.contains(styleNumber)) // time
        {
            return getDateTimeValue(value, address, TIME, importOptions);
        } else if (Helper.isNullOrEmpty(type) || type.equals("n")) // try numeric if not parsed as date or time, before numeric
        {
            if (Helper.isNullOrEmpty(value)) {
                return new Cell(null, Cell.CellType.EMPTY, address);
            }
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
     * Handles the value of a raw cell as string. An appropriate formatting is applied to Date and LocalTime values. Null values are left on type EMPTY
     *
     * @param address     Address of the cell
     * @param type        Expected data type
     * @param value       Raw value as string
     * @param styleNumber Style number as string (can be null)
     * @param formula     Formula as string (can be null; data type determines whether value or formula is used)
     * @param options     Options instance to determine appropriate formatting information
     * @return Cell of the type string
     */
    private Cell getEnforcedStingValue(Address address, String type, String value, String styleNumber, String formula, ImportOptions options) {
        Cell parsed = autoResolveCellData(address, type, value, styleNumber, formula);
        if (parsed.getDataType().equals(Cell.CellType.EMPTY)) {
            return parsed;
        } else if (parsed.getDataType().equals(Cell.CellType.DATE)) {
            return getStringValue(options.getDateFormatter().format((Date) parsed.getValue()), address);
        } else if (parsed.getDataType().equals(Cell.CellType.TIME)) {
            return getStringValue(options.getLocalTimeFormatter().format((LocalTime) parsed.getValue()), address);
        } else {
            return getStringValue(parsed.getValue().toString(), address);
        }
    }

    /**
     * Parses the numeric value of a raw cell. The order of possible number types are: int, float double. If nothing applies, a string is returned
     *
     * @param raw     Raw value as string
     * @param address Address of cell
     * @return Cell of the type int, float, double or string as fall-back type<
     */
    private Cell getNumericValue(String raw, Address address) {
        return getNumericValue(raw, address, null);
    }

    /**
     * Parses the numeric value of a raw cell. The order of possible number types are: int, float double. If nothing applies, a string is returned
     *
     * @param raw         Raw value as string
     * @param address     Address of cell
     * @param styleNumber Parameter to determine whether a double is enforced in case of a date or time style
     * @return Cell of the type int, float, double or string as fall-back type<
     */
    private Cell getNumericValue(String raw, Address address, String styleNumber) {
        if (dateStyles.contains(styleNumber) || timeStyles.contains(styleNumber)) {
            return getDoubleValue(raw, address);
        }
        Integer i = tryParseInt(raw);
        if (i != null) {
            return new Cell(i, Cell.CellType.NUMBER, address);
        }
        Long l = tryParseLong(raw);
        if (l != null) {
            return new Cell(l, Cell.CellType.NUMBER, address);
        }

        Number n = tryParseDecimal(raw);
        if (n instanceof Float) {
            return new Cell(n.floatValue(), Cell.CellType.NUMBER, address);
        }
        return getDoubleValue(raw, address);
    }

    /**
     * Parses a raw value as double
     *
     * @param raw     Raw value as string
     * @param address Address of the cell
     * @return Cell of the type double or string as fall-back type
     */
    private static Cell getDoubleValue(String raw, Address address) {
        try {
            double d = Double.parseDouble(raw);
            return new Cell(d, Cell.CellType.NUMBER, address);
        } catch (Exception ex) {
            return new Cell(raw, Cell.CellType.STRING, address);
        }
    }

    /**
     * Tries to parse a string to a decimal (either float or double, if out of range for float)
     *
     * @param value Raw string value
     * @return Decimal with either a float or double value
     */
    private static Number tryParseDecimal(String value) {
        try {
            double d = Double.parseDouble(value);
            float f = Float.parseFloat(value);
            if (Float.isFinite(f) && (f != 0.0 || (float) d == d)) {
                return f;
            }
            return d;
        } catch (Exception ignore) {
            return null;
        }
    }

    /**
     * Tries to parse a string to an int
     *
     * @param value Raw input string
     * @return Parsed int or null if not a valid integer32
     */
    private static Integer tryParseInt(String value) {
        try {
            return Integer.parseInt(value);
        } catch (Exception ignore) {
            return null;
        }
    }

    /**
     * Tries to parse a string to an long
     *
     * @param value Raw input string
     * @return Parsed long or null if not a valid integer64
     */
    private static Long tryParseLong(String value) {
        try {
            return Long.parseLong(value);
        } catch (Exception ignore) {
            return null;
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
        return new Cell(resolveSharedString(raw), Cell.CellType.STRING, address);
    }

    /**
     * Parses the string value of a raw cell. May take the value from the shared string table, if available
     *
     * @param raw         Raw value as string
     * @param address     Address of the cell
     * @param type        Type to check whether the raw value is already a resolved string
     * @param styleNumber Optional style number that may indicate a date or time value
     * @param options     Optional import options to determine the date and time formatting information
     * @return Cell of the type string
     */
    private Cell getStringValue(String raw, Address address, String type, String styleNumber, ImportOptions options) {
        if (type != null && type.equals("s")) {
            return new Cell(resolveSharedString(raw), Cell.CellType.STRING, address);
        } else if (type != null && type.equals("b")) {
            Cell tempCell = getBooleanValue(raw, address);
            if (tempCell != null) {
                return new Cell(tempCell.getValue().toString(), Cell.CellType.STRING, address);
            }
        } else if (styleNumber != null) {
            Cell tempCell = null;
            if (dateStyles.contains(styleNumber))  // date (priority)
            {
                tempCell = getDateTimeValue(raw, address, Cell.CellType.DATE, options);
            } else if (timeStyles.contains(styleNumber)) // time
            {
                tempCell = getDateTimeValue(raw, address, Cell.CellType.TIME, options);
            }
            if (tempCell != null && tempCell.getDataType().equals(DATE)) {
                DateFormat format = options.getDateFormatter();
                return new Cell(format.format((Date) tempCell.getValue()), Cell.CellType.STRING, address);
            } else if (tempCell != null && tempCell.getDataType().equals(TIME)) {
                DateTimeFormatter format = options.getLocalTimeFormatter();
                return new Cell(format.format((LocalTime) tempCell.getValue()), Cell.CellType.STRING, address);
            }
        }
        return new Cell(raw, Cell.CellType.STRING, address);
    }

    /**
     * Tries to resolve a shared string from its ID
     *
     * @param raw Raw value that can be either an ID of a shared string or an actual string value
     * @return Resolved string or the raw value if no shared string could be determined
     */
    private String resolveSharedString(String raw) {
        try {
            int stringId = Integer.parseInt(raw);
            String resolvedString = sharedStrings.getString(stringId);
            if (resolvedString == null) {
                return raw;
            } else {
                return resolvedString;
            }
        } catch (Exception ex) {
            return raw;
        }
    }

    /**
     * Parses the boolean value of a raw cell
     *
     * @param raw     Raw value as string
     * @param address Address of the cell
     * @return Cell of the type boolean or null if not able to parse
     */
    private static Cell getBooleanValue(String raw, Address address) {
        if (raw == null || raw.isBlank()) {
            return new Cell(raw, Cell.CellType.STRING, address);
        }
        String str = raw.toLowerCase();
        if (str.equals("1") || str.equals("true")) {
            return new Cell(true, Cell.CellType.BOOL, address);
        } else if (str.equals("0") || str.equals("false")) {
            return new Cell(false, Cell.CellType.BOOL, address);
        } else {
            try {
                double d = Double.parseDouble(raw);
                if (d >= -ZERO_THRESHOLD && d <= ZERO_THRESHOLD) {
                    return new Cell(false, Cell.CellType.BOOL, address);
                } else if (d - 1 >= -ZERO_THRESHOLD && d - 1 <= ZERO_THRESHOLD) {
                    return new Cell(true, Cell.CellType.BOOL, address);
                }
            } catch (Exception ex) {
                return null;
            }
        }
        return null;
    }

    /**
     * Parses the date (Date) or time (LocalTime) value of a raw cell. If the value is numeric, but out of range of a OAdate, a numeric value will be returned instead. If invalid, the string representation will be returned.
     *
     * @param raw     Raw value as string
     * @param address Address of the cell
     * @param type    Type of the value to be converted: Valid values are DATE and TIME
     * @param options Import options instance to take the Date or LocalTime parser from
     * @return CellResolverTuple with information about the validity and resolved data
     * @throws IllegalArgumentException Thrown if an unsupported type was passed
     */
    private static Cell getDateTimeValue(String raw, Address address, Cell.CellType type, ImportOptions options) {
        return getDateTimeValue(raw, address, type, options, null);
    }

    /**
     * Parses the date (Date) or time (LocalTime) value of a raw cell. If the value is numeric, but out of range of a OAdate, a numeric value will be returned instead. If invalid, the string representation will be returned.
     *
     * @param raw       Raw value as string
     * @param address   Address of the cell
     * @param valueType Type of the value to be converted: Valid values are DATE and TIME
     * @param type      Parameter to check whether the raw value should be tried to be parsed as date or time
     * @param options   Import options instance to take the Date or LocalTime parser from
     * @return CellResolverTuple with information about the validity and resolved data
     * @throws IllegalArgumentException Thrown if an unsupported type was passed
     */
    private static Cell getDateTimeValue(String raw, Address address, Cell.CellType valueType, ImportOptions options, String type) {
        if (type != null && type.equals("b")) {
            return getBooleanValue(raw, address);
        }
        try {
            if (type != null && type.equals("s")) {
                Date tempDate = tryParseDate(raw, options);
                if (tempDate != null && valueType == Cell.CellType.DATE) {
                    return getTemporalCell(tempDate, address, options);
                } else if (tempDate != null && valueType == Cell.CellType.TIME) {
                    CALENDAR.setTime(tempDate);
                    return getTemporalCell(LocalTime.of(CALENDAR.get(Calendar.HOUR_OF_DAY), CALENDAR.get(Calendar.MINUTE), CALENDAR.get(Calendar.SECOND)), address, options);
                }
                LocalTime tempTime = tryParseTime(raw, options);
                if (tempTime != null && valueType == Cell.CellType.TIME) {
                    return getTemporalCell(tempTime, address, options);
                }
            }
            double d = Double.parseDouble(raw);
            if (d < XlsxWriter.MIN_OADATE_VALUE || d > XlsxWriter.MAX_OADATE_VALUE || (options != null && options.isEnforceDateTimesAsNumbers())) {
                return new Cell(d, Cell.CellType.NUMBER, address); // Invalid OAdate / enforced number == plain number
            } else {
                Date date = Helper.getDateFromOA(d);
                switch (valueType) {
                    case DATE:
                        if (date.getTime() >= Helper.FIRST_ALLOWED_EXCEL_DATE.getTime()) {
                            return getTemporalCell(date, address, options);
                        } else {
                            // Prevent to import 00.01.1900, since it will lead to trouble when exporting / writing
                            return new Cell(d, Cell.CellType.NUMBER, address);
                        }
                    case TIME:
                        CALENDAR.setTime(date);
                        return getTemporalCell(LocalTime.of(CALENDAR.get(Calendar.HOUR_OF_DAY), CALENDAR.get(Calendar.MINUTE), CALENDAR.get(Calendar.SECOND)), address, options);
                    default:
                        throw new IllegalArgumentException("The defined type is not supported to be uses as date or time");
                }
            }
        } catch (Exception e) {
            return new Cell(raw, Cell.CellType.STRING, address);
        }
    }

    /**
     * Gets a cell either as Date, LocalTime or as double if enforced by import options
     *
     * @param dateTimeValue Value of the cell
     * @param address       Address of the cell
     * @param options       Import options
     * @return Casted cell
     */
    private static Cell getTemporalCell(Object dateTimeValue, Address address, ImportOptions options) {
        if (options != null && options.isEnforceDateTimesAsNumbers()) {
            if (dateTimeValue instanceof Date) {
                return new Cell(Helper.getOADate((Date) dateTimeValue), Cell.CellType.NUMBER, address);
            } else {
                return new Cell(Helper.getOATime((LocalTime) dateTimeValue), Cell.CellType.NUMBER, address);
            }
        }
        if (dateTimeValue instanceof Date) {
            return new Cell((Date) dateTimeValue, Cell.CellType.DATE, address);
        } else {
            return new Cell((LocalTime) dateTimeValue, Cell.CellType.TIME, address);
        }
    }

    /**
     * Tris to parse a Date instance from a string
     *
     * @param raw     String to parse
     * @param options Import options to take the parsing patterns of
     * @return LocalTime instance or null if not possible to parse
     */
    private static Date tryParseDate(String raw, ImportOptions options) {
        try {
            if (options == null || options.getDateFormatter() == null) {
                // no generic parsing available
                return null;
            }
            Date date = options.getDateFormatter().parse(raw);
            long d = date.getTime();
            if (d >= Helper.FIRST_ALLOWED_EXCEL_DATE.getTime() && d <= Helper.LAST_ALLOWED_EXCEL_DATE.getTime()) {
                return date;
            }
            return null;
        } catch (Exception ex) {
            return null;
        }
    }

    /**
     * Tris to parse a LocalTime instance from a string
     *
     * @param raw     String to parse
     * @param options Import options to take the parsing patterns of
     * @return LocalTime instance or null if not possible to parse
     */
    private static LocalTime tryParseTime(String raw, ImportOptions options) {
        try {
            if (options == null || options.getLocalTimeFormatter() == null) {
                // no generic parsing available
                return null;
            }
            TemporalAccessor time = options.getLocalTimeFormatter().parse(raw);
            return LocalTime.of(time.get(ChronoField.HOUR_OF_DAY), time.get(ChronoField.MINUTE_OF_HOUR), time.get(ChronoField.SECOND_OF_MINUTE));
        } catch (Exception ex) {
            return null;
        }
    }


}