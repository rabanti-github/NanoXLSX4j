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
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.styles.Style;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoField;
import java.time.temporal.ChronoUnit;
import java.time.temporal.TemporalAccessor;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
    private static final DecimalFormat DECIMAL_FORMAT = new DecimalFormat("#.#########");

    private final Map<String, Cell> data;
    private final SharedStringsReader sharedStrings;
    private final Map<String, String> styleAssignment = new HashMap<>();
    private final ImportOptions importOptions;
    private List<String> dateStyles;
    private List<String> timeStyles;
    private Map<String, Style> resolvedStyles;
    private Range autoFilterRange = null;

    /**
     * Gets the data of the worksheet as Hashmap of cell address-cell object tuples
     *
     * @return Hashmap of cell address-cell object tuples
     */
    public Map<String, Cell> getData() {
        return data;
    }

    /**
     * Gets the assignment of resolved styles to cell addresses
     *
     * @return Map of cell address-style number tuples
     */
    public Map<String, String> getStyleAssignment() {
        return styleAssignment;
    }

    /**
     * gets the auto filter range
     *
     * @return Auto filter range if defined, otherwise null
     */
    public Range getAutoFilterRange() {
        return autoFilterRange;
    }

    /**
     * Constructor with parameters and import options
     *
     * @param sharedStrings        SharedStringsReader object
     * @param styleReaderContainer Resolved styles, used to determine dates or times
     */
    public WorksheetReader(SharedStringsReader sharedStrings, StyleReaderContainer styleReaderContainer, ImportOptions options) {
        this.data = new HashMap<>();
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
        this.resolvedStyles = new HashMap<>();
        for (int i = 0; i < styleReaderContainer.getStyleCount(); i++) {
            String index = Integer.toString(i);
            StyleReaderContainer.StyleResult result = styleReaderContainer.evaluateDateTimeStyle(i, true);
            if (result.isDateStyle()) {
                this.dateStyles.add(index);
            }
            if (result.isTimeStyle()) {
                this.timeStyles.add(index);
            }
            resolvedStyles.put(index, result.getResult());
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
            XmlDocument.XmlNodeList autoFilterRanges = xr.getDocumentElement().getElementsByTagName("autoFilter", true);
            if (autoFilterRanges != null && autoFilterRanges.size() > 0) {
                String auoFilterRef = autoFilterRanges.get(0).getAttribute("ref");
                if (auoFilterRef != null) {
                    this.autoFilterRange = new Range(auoFilterRef);
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
                        value = valueNode.getInnerText();
                    }
                }
            }
        }
        String key = address.toUpperCase();
        styleAssignment.put(key, styleNumber);
        data.put(key, resolveCellData(value, type, styleNumber, address));
    }

    private Cell resolveCellData(String raw, String type, String styleNumber, String address) {
        Cell.CellType importedType = Cell.CellType.DEFAULT;
        Object rawValue;
        if (checkType(type, "b")) {
            rawValue = tryParseBool(raw);
            if (rawValue != null) {
                importedType = Cell.CellType.BOOL;
            } else {
                rawValue = getNumericValue(raw);
                if (rawValue != null) {
                    importedType = Cell.CellType.NUMBER;
                }
            }
        } else if (checkType(type, "s")) {
            importedType = Cell.CellType.STRING;
            rawValue = resolveSharedString(raw);
        } else if (checkType(type, "str")) {
            importedType = Cell.CellType.FORMULA;
            rawValue = raw;
        } else if (dateStyles.contains(styleNumber) && (checkType(type, null) || checkType(type, "") || checkType(type, "n"))) {
            Result<Object, Cell.CellType> result = getDateTimeValue(raw, DATE);
            rawValue = result.result1;
            importedType = result.result2;
        } else if (timeStyles.contains(styleNumber) && (checkType(type, null) || checkType(type, "") || checkType(type, "n"))) {
            Result<Object, Cell.CellType> result = getDateTimeValue(raw, TIME);
            rawValue = result.result1;
            importedType = result.result2;
        } else {
            importedType = Cell.CellType.NUMBER;
            rawValue = getNumericValue(raw);
        }
        if (rawValue == null && raw.equals("")) {
            importedType = Cell.CellType.EMPTY;
            rawValue = null;
        } else if (rawValue == null && raw.length() > 0) {
            importedType = Cell.CellType.STRING;
            rawValue = raw;
        }
        Address cellAddress = new Address(address);
        if (importOptions != null) {
            if (importOptions.getEnforcedColumnTypes().size() > 0) {
                rawValue = getEnforcedColumnValue(rawValue, importedType, cellAddress);
            }
            rawValue = getGloballyEnforcedValue(rawValue, cellAddress);
            rawValue = getGloballyEnforcedFlagValues(rawValue, cellAddress);
            importedType = resolveType(rawValue, importedType);
            if (importedType == Cell.CellType.DATE && rawValue instanceof Date && ((Date) rawValue).getTime() < Helper.FIRST_ALLOWED_EXCEL_DATE.getTime()) {
                // Fix conversion from time to date, where time has no days
                rawValue = addTemporalUnits((Date) rawValue, 1, 0, 0, 0);
            }
        }
        return createCell(rawValue, importedType, cellAddress, styleNumber);
    }

    private boolean checkType(String type, String expectation) {
        if (type == null && expectation != null) {
            return false;
        } else if (type == null) {
            return true;
        }
        return expectation.equals(type);
    }

    private Cell.CellType resolveType(Object value, Cell.CellType defaultType) {
        if (defaultType == Cell.CellType.FORMULA) {
            return defaultType;
        }
        if (value == null) {
            return Cell.CellType.EMPTY;
        }
        Class<?> cls = value.getClass();
        if (BigDecimal.class.equals(cls) || Long.class.equals(cls) || Short.class.equals(cls) || Float.class.equals(cls) || Double.class.equals(cls) || Byte.class.equals(cls) || Integer.class.equals(cls)) {
            return Cell.CellType.NUMBER;
        } else if (Date.class.equals(cls)) {
            return DATE;
        } else if (Duration.class.equals(cls)) {
            return TIME;
        } else if (Boolean.class.equals(cls)) {
            return Cell.CellType.BOOL;
        } else {
            return Cell.CellType.STRING;
        }
    }

    private Object getGloballyEnforcedFlagValues(Object data, Address address) {
        if (address.Row < importOptions.getEnforcingStartRowNumber()) {
            return data;
        }
        if (importOptions.isEnforceDateTimesAsNumbers()) {
            if (data instanceof Date) {
                data = Helper.getOADate((Date) data, true);
            } else if (data instanceof Duration) {
                data = Helper.getOATime((Duration) data);
            }
        }
        if (importOptions.isEnforceEmptyValuesAsString()) {
            if (data == null) {
                return "";
            }
        }
        return data;
    }

    private Object getGloballyEnforcedValue(Object data, Address address) {
        if (address.Row < importOptions.getEnforcingStartRowNumber()) {
            return data;
        }
        if (importOptions.getGlobalEnforcingType().equals(ImportOptions.GlobalType.AllNumbersToDouble)) {
            Object tempDouble = convertToDouble(data);
            if (tempDouble != null) {
                return tempDouble;
            }
        }
        if (importOptions.getGlobalEnforcingType().equals(ImportOptions.GlobalType.AllNumbersToBigDecimal)) {
            Object tempBigDecimal = convertToBigDecimal(data);
            if (tempBigDecimal != null) {
                return tempBigDecimal;
            }
        } else if (importOptions.getGlobalEnforcingType().equals(ImportOptions.GlobalType.AllNumbersToInt)) {
            Object tempInt = convertToInt(data);
            if (tempInt != null) {
                return tempInt;
            }
        } else if (importOptions.getGlobalEnforcingType().equals(ImportOptions.GlobalType.EverythingToString)) {
            return convertToString(data);
        }
        return data;
    }

    private Object getEnforcedColumnValue(Object data, Cell.CellType importedTyp, Address address) {
        if (address.Row < importOptions.getEnforcingStartRowNumber()) {
            return data;
        }
        if (!importOptions.getEnforcedColumnTypes().containsKey(address.Column)) {
            return data;
        }
        if (importedTyp == Cell.CellType.FORMULA) {
            return data;
        }
        switch (importOptions.getEnforcedColumnTypes().get(address.Column)) {
            case Numeric:
                return getNumericValue(data, importedTyp);
            case BigDecimal:
                return convertToBigDecimal(data);
            case Double:
                return convertToDouble(data);
            case Date:
                return convertToDate(data);
            case Time:
                return convertToTime(data);
            case Bool:
                return convertToBool(data);
            default:
                return convertToString(data);
        }
    }

    private Object convertToBool(Object data) {
        if (data == null) {
            return null;
        }
        Class<?> cls = data.getClass();
        if (Boolean.class.equals(cls)) {
            return data;
        } else if (Long.class.equals(cls) || Short.class.equals(cls) || Float.class.equals(cls) || Double.class.equals(cls) || Byte.class.equals(cls) || BigDecimal.class.equals(cls) || Integer.class.equals(cls)) {
            Object tempObject = convertToDouble(data);
            if (tempObject instanceof Double) {
                double tempDouble = (double) tempObject;
                if (compareDouble(tempDouble, 0d)) {
                    return false;
                } else if (compareDouble(tempDouble, 1d)) {
                    return true;
                }
            }
        } else if (String.class.equals(cls)) {
            String tempString = (String) data;
            Boolean tempBool = tryParseBool(tempString);
            if (tempBool != null) {
                return tempBool;
            }
        }
        return data;
    }

    private Boolean tryParseBool(String raw) {
        if (Helper.isNullOrEmpty(raw)) {
            return null;
        }
        Object nValue = getNumericValue(raw);
        if (nValue != null) {
            Number n = (Number) nValue;
            if (n.intValue() == 1) {
                return true;
            } else if (n.intValue() == 0) {
                return false;
            } else {
                return null;
            }
        }
        if (raw.equalsIgnoreCase("true")) {
            return true;
        } else if (raw.equalsIgnoreCase("false")) {
            return false;
        }
        return null;
    }

    private Object convertToDouble(Object data) {
        Object value = convertToBigDecimal(data);
        if (value instanceof BigDecimal) {
            double tempDouble = ((BigDecimal) value).doubleValue();
            if (Double.isFinite(tempDouble)) {
                return tempDouble;
            }
        }
        return value;
    }

    private Object convertToBigDecimal(Object data) {
        if (data == null) {
            return null;
        }
        Class<?> cls = data.getClass();
        if (BigDecimal.class.equals(cls)) {
            return data;
        } else if (Long.class.equals(cls) || Short.class.equals(cls) || Float.class.equals(cls) || Double.class.equals(cls) || Byte.class.equals(cls) || Integer.class.equals(cls)) {
            Number number = (Number) data;
            return BigDecimal.valueOf(number.doubleValue());
        } else if (Boolean.class.equals(cls)) {
            if (Boolean.TRUE.equals(data)) {
                return BigDecimal.ONE;
            } else {
                return BigDecimal.ZERO;
            }
        } else if (Date.class.equals(cls)) {
            return BigDecimal.valueOf(Helper.getOADate((Date) data));
        } else if (Duration.class.equals(cls)) {
            return BigDecimal.valueOf(Helper.getOATime((Duration) data));
        } else if (String.class.equals(cls)) {
            String tempString = (String) data;
            BigDecimal dValue = tryParseBigDecimal(tempString);
            if (dValue != null) {
                return dValue;
            }
            Date tempDate = tryParseDate(tempString, importOptions);
            if (tempDate != null) {
                return BigDecimal.valueOf(Helper.getOADate(tempDate));
            }
            Duration tempTime = tryParseTime(tempString, importOptions);
            if (tempTime != null) {
                return BigDecimal.valueOf(Helper.getOATime(tempTime));
            }
        }
        return data;
    }

    private Object convertToInt(Object data) {
        if (data == null) {
            return null;
        }
        double tempDouble;
        Class<?> cls = data.getClass();
        if (Date.class.equals(cls)) {
            tempDouble = Helper.getOADate((Date) data, true);
            return convertDoubleToInt(tempDouble);
        } else if (Duration.class.equals(cls)) {
            tempDouble = Helper.getOATime((Duration) data);
            return convertDoubleToInt(tempDouble);
        } else if (Float.class.equals(cls) || BigDecimal.class.equals(cls) || Double.class.equals(cls)) {
            Object tempInt = tryConvertDoubleToInt(data);
            if (tempInt != null) {
                return tempInt;
            }
        } else if (Boolean.class.equals(cls)) {
            return (boolean) data ? 1 : 0;
        } else if (String.class.equals(cls)) {
            Integer tempInt2 = tryParseInt((String) data);
            if (tempInt2 != null) {
                return tempInt2;
            }
        }
        return null;
    }

    private Object convertToDate(Object data) {
        if (data == null) {
            return null;
        }
        Class<?> cls = data.getClass();
        if (Date.class.equals(cls)) {
            return data;
        } else if (Duration.class.equals(cls)) {
            Date root = Helper.FIRST_ALLOWED_EXCEL_DATE;
            TimeComponent t = new TimeComponent(((Duration) data).get(ChronoUnit.SECONDS));
            root = addTemporalUnits(root, -1, t.getHours(), t.getMinutes(), t.getSeconds());
            return root;
        } else if (Double.class.equals(cls) || BigDecimal.class.equals(cls) || Long.class.equals(cls) || Short.class.equals(cls) || Float.class.equals(cls) || Byte.class.equals(cls) || Integer.class.equals(cls)) {
            return convertDateFromDouble(data);
        } else if (String.class.equals(cls)) {
            Date date2 = tryParseDate((String) data, importOptions.getDateFormatter());
            if (date2 != null) {
                return date2;
            }
            return convertDateFromDouble(data);
        }
        return data;
    }

    private static Date tryParseDate(String raw, SimpleDateFormat formatter) {
        try {
            Date date;
            date = formatter.parse(raw);
            if (date.getTime() >= Helper.FIRST_ALLOWED_EXCEL_DATE.getTime() && date.getTime() <= Helper.LAST_ALLOWED_EXCEL_DATE.getTime()) {
                return date;
            }
        } catch (Exception ex) {
        }
        return null;
    }

    private Object convertToTime(Object data) {
        if (data == null) {
            return null;
        }
        Class<?> cls = data.getClass();
        if (Date.class.equals(cls)) {
            return convertTimeFromDouble(data);
        } else if (Duration.class.equals(cls)) {
            return data;
        } else if (Double.class.equals(cls) || BigDecimal.class.equals(cls) || Long.class.equals(cls) || Short.class.equals(cls) || Float.class.equals(cls) || Byte.class.equals(cls) || Integer.class.equals(cls)) {
            return convertTimeFromDouble(data);
        } else if (String.class.equals(cls)) {
            Duration time = tryParseTime((String) data, importOptions.getTimeFormatter());
            if (time != null) {
                return time;
            }
            return convertTimeFromDouble(data);
        }
        return data;
    }


    private static Duration tryParseTime(String raw, DateTimeFormatter formatter) {
        try {
            Duration time;
            time = Helper.parseTime(raw, formatter);
            double days = time.get(ChronoUnit.SECONDS) / 86400d;
            if (days >= 0d && days < Helper.MAX_OADATE_VALUE) {
                return time;
            } else {
                return null;
            }
        } catch (Exception ex) {
            return null;
        }
    }

    private Result<Object, Cell.CellType> getDateTimeValue(String raw, Cell.CellType valueType) {
        Double dValue = tryParseDouble(raw);
        if (dValue == null) {
            return new Result<>(raw, Cell.CellType.STRING);
        }
        if ((valueType == Cell.CellType.DATE && (dValue < Helper.MIN_OADATE_VALUE || dValue > Helper.MAX_OADATE_VALUE)) || (valueType == Cell.CellType.TIME && (dValue < 0.0 || dValue > Helper.MAX_OADATE_VALUE))) {
            // fallback to number (cannot be anything else)
            return new Result<>(getNumericValue(raw), Cell.CellType.NUMBER);
        }
        Date tempDate = Helper.getDateFromOA(dValue);
        if (dValue < 1.0) {
            tempDate = addTemporalUnits(tempDate, 1, 0, 0, 0); // Modify wrong 1st date when < 1
        }
        if (valueType == Cell.CellType.DATE) {
            return new Result<>(tempDate, DATE);
        } else {
            CALENDAR.setTime(tempDate);
            return new Result<>(Helper.createDuration(dValue.intValue(), CALENDAR.get(Calendar.HOUR_OF_DAY), CALENDAR.get(Calendar.MINUTE), CALENDAR.get(Calendar.SECOND)), TIME);
        }
    }

    private Object convertDateFromDouble(Object data) {
        Object oaDate = convertToDouble(data);
        if (oaDate instanceof Double && (Double) oaDate < Helper.MAX_OADATE_VALUE) {
            Date date = Helper.getDateFromOA((Double) oaDate);
            if (date.getTime() >= Helper.FIRST_ALLOWED_EXCEL_DATE.getTime() && date.getTime() <= Helper.LAST_ALLOWED_EXCEL_DATE.getTime()) {
                return date;
            }
        }
        return data;
    }

    private Object convertTimeFromDouble(Object data) {
        Object oaDate = convertToDouble(data);
        if (oaDate instanceof Double) {
            double d = (Double) oaDate;
            if (d >= Helper.MIN_OADATE_VALUE && d <= Helper.MAX_OADATE_VALUE) {
                Date date = Helper.getDateFromOA(d);
                CALENDAR.setTime(date);
                return Helper.createDuration((int) d, CALENDAR.get(Calendar.HOUR_OF_DAY), CALENDAR.get(Calendar.MINUTE), CALENDAR.get(Calendar.SECOND));
            }
        }
        return data;
    }

    private Object tryConvertDoubleToInt(Object data) {
        Number number = (Number) data;
        double dValue = number.doubleValue();
        if (dValue > Integer.MIN_VALUE && dValue < Integer.MAX_VALUE) {
            return (int) Math.round(number.doubleValue());
        }
        return null;
    }

    public Object convertDoubleToInt(Object data) {
        Number number = (Number) data;
        return (int) Math.round(number.doubleValue());
    }

    private String convertToString(Object data) {
        if (data == null) {
            return null;
        }
        Class<?> cls = data.getClass();
        if (Integer.class.equals(cls)) {
            return ((Integer) data).toString();
        } else if (Long.class.equals(cls)) {
            return ((Long) data).toString();
        } else if (Float.class.equals(cls)) {
            return ((Float) data).toString();
        } else if (Double.class.equals(cls)) {
            return ((Double) data).toString();
        } else if (BigDecimal.class.equals(cls)) {
            return ((BigDecimal) data).toString();
        } else if (Boolean.class.equals(cls)) {
            return ((Boolean) data).toString();
        } else if (Date.class.equals(cls)) {
            return importOptions.getDateFormatter().format((Date) data);
        } else if (Duration.class.equals(cls)) {
            TimeComponent t = new TimeComponent((int) ((Duration) data).toSeconds());
            LocalTime tempTime = LocalTime.of(t.getHours(), t.getMinutes(), t.getSeconds(), t.getDays());
            return importOptions.getTimeFormatter().format(tempTime);
        }
        return data.toString();
    }

    private Object getNumericValue(Object raw, Cell.CellType importedType) {
        if (raw == null) {
            return null;
        }
        Object tempObject;
        switch (importedType) {
            case STRING:
                String tempString = raw.toString();
                tempObject = getNumericValue(tempString);
                if (tempObject != null) {
                    return tempObject;
                }
                Date tempDate = tryParseDate(tempString, importOptions);
                if (tempDate != null) {
                    return Helper.getOADate(tempDate);
                }
                Duration tempTime = tryParseTime(tempString, importOptions);
                if (tempTime != null) {
                    return Helper.getOATime(tempTime);
                }
                tempObject = convertToBool(raw);
                if (tempObject instanceof Boolean) {
                    return (boolean) tempObject ? 1 : 0;
                }
                break;
            case NUMBER:
                return raw;
            case DATE:
                return Helper.getOADate((Date) raw);
            case TIME:
                return Helper.getOATime((Duration) raw);
            case BOOL:
                if ((boolean) raw) {
                    return 1;
                }
                return 0;
        }
        return raw;
    }

    private Object getNumericValue(String raw) {
        // integer section
        Integer iValue = tryParseInt(raw);
        if (iValue != null) {
            return iValue;
        }
        Long lValue = tryParseLong(raw);
        if (lValue != null) {
            return lValue;
        }
        // float section
        Number dcValue = tryParseDecimal(raw);
        if (dcValue != null && dcValue instanceof Float) {
            return dcValue.floatValue();
        } else if (dcValue != null && dcValue instanceof Double) {
            return dcValue.doubleValue();
        }
        return null;
    }

    private static class Result<R1, R2> {
        public final R1 result1;
        public final R2 result2;

        public Result(R1 result1, R2 result2) {
            this.result1 = result1;
            this.result2 = result2;
        }
    }

    private static Date addTemporalUnits(Date root, int days, int hours, int minutes, int seconds) {
        CALENDAR.setTime(root);
        if (days != 0) {
            CALENDAR.add(Calendar.DATE, days);
        }
        if (hours != 0) {
            CALENDAR.add(Calendar.HOUR, hours);
        }
        if (minutes != 0) {
            CALENDAR.add(Calendar.MINUTE, minutes);
        }
        if (seconds != 0) {
            CALENDAR.add(Calendar.SECOND, seconds);
        }
        return CALENDAR.getTime();
    }

    private static Double tryParseDouble(String raw) {
        try {
            return Double.parseDouble(raw);
        } catch (Exception ex) {
            return null;
        }
    }

    private static BigDecimal tryParseBigDecimal(String raw) {
        try {
            return new BigDecimal(raw);
        } catch (Exception ex) {
            return null;
        }
    }

    private static boolean compareDouble(double d1, double d2) {
        final double epsilon = 0.000001d;
        return Math.abs(d1 - d2) < epsilon;
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
            String[] dString = DECIMAL_FORMAT.format(d).split("\\.");
            int numberOfDigits = 0;
            if (dString.length == 2) {
                numberOfDigits = dString[1].length();
            }
            float f = Float.parseFloat(value);
            if (Float.isFinite(f) && numberOfDigits < 7 && (f != 0.0 && d != 0.0)) {
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
     * Tris to parse a Date instance from a string
     *
     * @param raw     String to parse
     * @param options Import options to take the parsing patterns of
     * @return LocalTime instance or null if not possible to parse
     */
    private static Date tryParseDate(String raw, ImportOptions options) {
        try {
            if (options != null && options.getDateFormatter() != null) {
                Date date = options.getDateFormatter().parse(raw);
                long d = date.getTime();
                if (d >= Helper.FIRST_ALLOWED_EXCEL_DATE.getTime() && d <= Helper.LAST_ALLOWED_EXCEL_DATE.getTime()) {
                    return date;
                }
            }
        } catch (Exception ex) {
            // Ignore
        }
        return null;
    }


    /**
     * Creates a generic cell with style information
     *
     * @param value       value of the cell
     * @param type        Cell type
     * @param address     Cell address
     * @param styleNumber Style number of the cell
     * @return Resolved cell
     */
    private Cell createCell(Object value, Cell.CellType type, Address address, String styleNumber) {
        Cell cell = new Cell(value, type, address);
        if (styleNumber != null && resolvedStyles.containsKey(styleNumber)) {
            cell.setStyle(resolvedStyles.get(styleNumber));
        }
        return cell;
    }

    /**
     * Tris to parse a Duration instance from a string
     *
     * @param raw     String to parse
     * @param options Import options to take the parsing patterns of
     * @return Duration instance or null if not possible to parse
     */
    private static Duration tryParseTime(String raw, ImportOptions options) {
        try {
            TemporalAccessor time = options.getTimeFormatter().parse(raw);
            int days = time.get(ChronoField.NANO_OF_SECOND);
            int hours = time.get(ChronoField.HOUR_OF_DAY);
            int minutes = time.get(ChronoField.MINUTE_OF_HOUR);
            int seconds = time.get(ChronoField.SECOND_OF_MINUTE);
            Duration duration = Duration.ofSeconds(hours * 3600 + minutes * 60 + seconds, 0);
            return duration.plusDays(days);
        } catch (Exception ex) {
            return null;
        }
    }

    /**
     * Class to represent the components of a time with an optional number of days
     */
    private static class TimeComponent {
        private int hours;
        private int minutes;
        private int seconds;
        private int days;

        public int getHours() {
            return hours;
        }

        public int getMinutes() {
            return minutes;
        }

        public int getSeconds() {
            return seconds;
        }

        public int getDays() {
            return days;
        }

        public TimeComponent(long totalSeconds) {
            calculateComponents((int) totalSeconds);
        }

        private void calculateComponents(int totalSeconds) {
            this.days = totalSeconds / 86400;
            this.hours = (totalSeconds - (this.days * 86400)) / 3600;
            this.minutes = (totalSeconds - (this.days * 86400) - (this.hours * 3600)) / 60;
            this.seconds = (totalSeconds - (this.days * 86400) - (this.hours * 3600) - (this.minutes * 60));
        }

    }


}