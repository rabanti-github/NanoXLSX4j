/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.exception.IOException;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamReader;
import java.io.InputStream;
import java.util.*;

/**
 * Class representing a reader for worksheets of XLSX files
 * @author Raphael Stoeckli
 */
public class WorksheetReader
{

    private int worksheetNumber;
    private final String name;
    private final Map<String, Cell> data;
    private final SharedStringsReader sharedStrings;


    /**
     * Gets the name of the worksheet
     * @return Name of the worksheet
     */
    public String getName() {
        return name;
    }

    /**
     * Gets the number of the worksheet
     * @return Number of the worksheet
     */
    public int getWorksheetNumber() {
        return worksheetNumber;
    }

    /**
     * Sets the number of the worksheet
     * @param worksheetNumber Worksheet number
     */
    public void setWorksheetNumber(int worksheetNumber) {
        this.worksheetNumber = worksheetNumber;
    }

    /**
     * Gets the data of the worksheet as Hashmap of cell address-cell object tuples
     * @return Hashmap of cell address-cell object tuples
     */
    public Map<String, Cell> getData() {
        return data;
    }

    /**
     * Constructor with parameters
     * @param sharedStrings SharedStringsReader object
     * @param name Worksheet name
     * @param number Worksheet number
     */
    public WorksheetReader(SharedStringsReader sharedStrings, String name, int number)
    {
        data = new HashMap<>();
        this.name = name;
        this.worksheetNumber = number;
        this.sharedStrings = sharedStrings;
    }

    /**
     * Gets whether the specified column exists in the data
     * @param columnAddress Column address as string
     * @return True, if the column exists
     */
    public boolean hasColumn(String columnAddress)
    {
        if (Helper.isNullOrEmpty(columnAddress)) { return false; }
        int columnNumber = Cell.resolveColumn(columnAddress);
        for (Map.Entry<String, Cell> cell : data.entrySet())
        {
            if (cell.getValue().getColumnNumber() == columnNumber)
            {
                return true;
            }
        }
        return false;
    }

    /**
     * Gets whether the passed row (represented as list of cell objects) contains the specified column numbers
     * @param cells List of cell objects to check
     * @param columnNumbers Array of column numbers
     * @return True if all column numbers were found, otherwise false
     */
    public boolean rowHasColumns(List<Cell> cells, int[] columnNumbers)
    {
        if (columnNumbers == null || cells == null) { return false; }
        int len = columnNumbers.length;
        int len2 = cells.size();
        int j;
        boolean match;
        if (len < 1 || len2 < 1){ return false; }
        for (int i = 0; i < len; i++)
        {
            match = false;
            for (j = 0; j < len2; j++)
            {
                if (cells.get(j).getColumnNumber() == columnNumbers[i])
                {
                    match = true;
                    break;
                }
            }
            if (match == false) { return false; }
        }
        return true;
    }

    /**
     * Gets the number of rows
     * @return Number of rows
     */
    public int getRowCount()
    {
        int count = -1;
        for (Map.Entry<String, Cell> cell : data.entrySet())
        {
            if (cell.getValue().getRowNumber() > count)
            {
                count = cell.getValue().getRowNumber();
            }
        }
        return count + 1;
    }

    /**
     * Gets a row as list of cell objects
     * @param rowNumber Row number
     * @return List of cell objects
     */
    public List<Cell> getRow(int rowNumber)
    {
        List<Cell> list = new ArrayList<>();
        for (Map.Entry<String, Cell> cell : data.entrySet())
        {
            if (cell.getValue().getRowNumber() == rowNumber)
            {
                list.add(cell.getValue());
            }
        }
        list.sort((c1, c2) -> Integer.compare(c1.getColumnNumber(), c2.getColumnNumber())); // Lambda sort (java 8 and higher)
        return list;
    }

    /**
     * Reads the xlsx file form the passed stream and processes the worksheet data
     * @param stream Stream of the xlsx file
     * @throws IOException Throws IOException in case of an error
     */
    public void read(InputStream stream) throws IOException {
        data.clear();
        boolean isSheetData = false, isCell = false, isCellValue = false, isFormula = false;
        String value = "", formula = null;
        String type = "s";
        String style = "";
        String address = "A1";
        String name;
        XMLStreamReader xr;
        XMLInputFactory factory = XMLInputFactory.newFactory();
        int nodeType;
        try {
            xr = factory.createXMLStreamReader(stream);
            while (xr.hasNext()) {
                nodeType = xr.next();
                if (nodeType == XMLStreamReader.START_ELEMENT) {
                    name = xr.getName().getLocalPart().toLowerCase();
                    if (name.equals("sheetdata")) {
                        isSheetData = true;
                    } else if (name.equals("c") && isSheetData == true) {
                        address = xr.getAttributeValue(null, "r"); // mandatory
                        type = xr.getAttributeValue(null, "t"); // can be null if not existing
                        style = xr.getAttributeValue(null, "s"); // can be null; if "1" then date
                        isCell = true;
                    } else if (name.equals("f") && isCell == true) {
                        isFormula = true;
                    } else if (name.equals("v") && isCell == true) {
                        isCellValue = true;
                    }
                } else if (nodeType == XMLStreamReader.CHARACTERS) {
                    if (isFormula == true) {
                        formula = xr.getText();
                    } else if (isCellValue == true) {
                        value = xr.getText();
                    }
                } else if (nodeType == XMLStreamReader.END_ELEMENT) {
                    name = xr.getName().getLocalPart().toLowerCase();
                    if (name.equals("c") && isCell == true && isCellValue == true) {
                        isCell = false;
                        isCellValue = false;
                        isFormula = false;
                        resolveCellData(address, type, value, style, formula);
                        formula = null;
                        value = "";
                    } else if (name.equals("sheetdata") && isSheetData == true) {
                        isSheetData = false;
                    }
                }
            }
        } catch (Exception ex) {
            throw new IOException("XMLStreamException", "The XML entry could not be read from the input stream. Please see the inner exception:", ex);
        }
        finally
        {
            if (stream != null)
            {
                try {
                    stream.close();
                } catch (java.io.IOException ex) {
                    throw new IOException("XMLStreamException", "The XML entry stream could not be closed. Please see the inner exception:", ex);
                }
            }
        }
    }

    /**
     * Resolves the data of a read cell, transforms it into a cell object and adds it to the data
     * @param address Address of the cell
     * @param type Expected data type
     * @param value Raw value as string
     * @param style Style definition as string (can be null)
     * @param formula Formula as string (can be null; data type determines whether value or formula is used)
     */
    private void resolveCellData(String address, String type, String value, String style, String formula)
    {
        address = address.toUpperCase();
        double d;
        boolean b;
        int i;
        String s;
        Cell cell;
        CellResolverTuple tuple;
        if (style != null && style.equals("1")) // Date must come before numeric
        {
            tuple = getNumericValue(value);
            if (tuple.isValid() == true)
            {
                cell = new Cell(tuple.getData(), Cell.CellType.DATE, address);
            }
            else
            {
                cell = new Cell(value, Cell.CellType.STRING, address);
            }
        }
        else if (type == null) // try numeric
        {
            tuple = getNumericValue(value);
            if (tuple.isValid() == true)
            {
                cell = new Cell(tuple.getData(), Cell.CellType.NUMBER, address);
            }
            else
            {
                cell = new Cell(value, Cell.CellType.STRING, address);
            }
        }
        else if (type.equals("b"))
        {
            tuple = getBooleanValue(value);
            if (tuple.isValid() == true)
            {
                cell = new Cell(tuple.getData(), Cell.CellType.BOOL, address);
            }
            else
            {
                cell = new Cell(value, Cell.CellType.STRING, address);
            }
        }
        else if (formula != null)
        {
            cell = new Cell(formula, Cell.CellType.FORMULA, address);
        }
        else if (type.equals("s"))
        {
            tuple = getIntValue(value);
            if (tuple.isValid() == false)
            {
                cell = new Cell(value, Cell.CellType.STRING, address);
            }
            else
            {
                s = sharedStrings.getString((int)tuple.getData());
                if (s != null)
                {
                    cell = new Cell(s, Cell.CellType.STRING, address);
                }
                else
                {
                    cell = new Cell(value, Cell.CellType.STRING, address);
                }
            }
        }
        else
        {
            cell = new Cell(value, Cell.CellType.STRING, address);
        }
        data.put(address, cell);
    }

    /**
     * Parses the numeric (double) value of a raw cell
     * @param raw Raw value as string
     * @return CellResolverTuple with information about the validity and resolved data
     */
    private static CellResolverTuple getNumericValue(String raw)
    {
        CellResolverTuple t;
        try
        {
            double d = Double.parseDouble(raw);
            t = new CellResolverTuple(true, d, Double.class);
        }
        catch (Exception e)
        {
            t = new CellResolverTuple(false, 0, Double.class);
        }
        return t;
    }

    /**
     * Parses the integer value of a raw cell
     * @param raw Raw value as string
     * @return CellResolverTuple with information about the validity and resolved data
     */
    private static CellResolverTuple getIntValue(String raw)
    {
        CellResolverTuple t;
        try
        {
            int i = Integer.parseInt(raw);
            t = new CellResolverTuple(true, i, Integer.class);
        }
        catch (Exception e)
        {
            t = new CellResolverTuple(false, 0, Integer.class);
        }
        return t;
    }

    /**
     * Parses the boolean value of a raw cell
     * @param raw Raw value as string
     * @return CellResolverTuple with information about the validity and resolved data
     */
    private static CellResolverTuple getBooleanValue(String raw)
    {
        boolean value;
        boolean state;
        if (raw.equals("0"))
        {
            value = false;
            state = true;
        }
        else if (raw.equals("1"))
        {
            value = true;
            state = true;
        }
        else
        {
            try
            {
                value = Boolean.parseBoolean(raw);
                state = true;
            }
            catch (Exception e)
            {
                value = false;
                state = false;
            }
        }
        return new CellResolverTuple(state, value, Boolean.class);
    }

    /**
     * Helper class representing a tuple of cell data and is state (valid or invalid). And additional type is also available
     */
    static final class CellResolverTuple
    {
        private final boolean isValid;
        private final Object data;
        private final Class<?> type;

        /**
         * Gets whether the cell is valid
         * @return True if valid, otherwise false
         */
        public boolean isValid() {
            return isValid;
        }

        /**
         * Gets the data as object
         * @return Generic object
         */
        public Object getData() {
            return data;
        }

        /**
         * Gets the type of the cell
         * @return Data type
         */
        public Class<?> getType() {
            return type;
        }

        /**
         * Default constructor with parameters
         * @param isValid If true, the resolved cell contains valid data
         * @param data Data object
         * @param type Type of the cell
         */
        public CellResolverTuple(boolean isValid, Object data, Class<?> type)
        {
            this.data = data;
            this.isValid = isValid;
            this.type = type;
        }

    }

}