/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2019
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.*;
import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.styles.*;
import org.w3c.dom.Document;
import org.xml.sax.InputSource;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.io.StringReader;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;


/**
 * Class for low level handling (XML, formatting, preparing of packing)<br>This class is only for internal use. Use the high level API (e.g. class Workbook) to manipulate data and create Excel files.
 * @author Raphael Stoeckli
 */
public class XlsXWriter {
     
// ### P R I V A T E  F I E L D S ###    
    private final SortedMap sharedStrings;
    private int sharedStringsTotalCount;
    private final Workbook workbook;
    private boolean interceptDocuments;
    private HashMap<String, Document> interceptedDocuments;
    
// ### G E T T E R S   &   S E T T E R S ###

    /**
     * Gets whether XML documents are intercepted during creation
     * @return If true, documents will be intercepted and stored into interceptedDocuments 
     */
    public boolean getDocumentInterception() {
        return interceptDocuments;
    }

    /**
     * Set whether XML documents are intercepted during creation
     * @param interceptDocuments If true, documents will be intercepted and stored into interceptedDocuments
     */
    public void setDocumentInterception(boolean interceptDocuments)
    {
        this.interceptDocuments = interceptDocuments;
        if (interceptDocuments == true && this.interceptedDocuments == null)
        {
            this.interceptedDocuments = new HashMap<>();
        }
        else if (interceptDocuments == false)
        {
            this.interceptedDocuments = null;
        }
    }

    /**
     * Gets the intercepted documents if interceptDocuments is set to true
     * @return HashMap with a String as key and a XML document as value  
     */
    public HashMap<String, Document> getInterceptedDocuments() {
        return interceptedDocuments;
    }
    
    
  
// ### C O N S T R U C T O R S ###
    /**
     * Constructor with defined workbook object
     * @param workbook Workbook to process
     */
    public XlsXWriter(Workbook workbook)
    {
       this.workbook = workbook;
       this.sharedStrings = new SortedMap();
       this.sharedStringsTotalCount = 0;
    }

// ### M E T H O D S ###    
    /**
     * Method to append a simple XML tag with an enclosed value to the passed StringBuilder
     * @param sb StringBuilder to append
     * @param value Value of the XML element
     * @param tagName Tag name of the XML element
     * @param nameSpace Optional XML name space. Can be empty or null
     */
    private void appendXMLtag(StringBuilder sb, String value, String tagName, String nameSpace)
    {
        if (Helper.isNullOrEmpty(value)) { return; }
        if (sb == null || Helper.isNullOrEmpty(tagName)) { return; }
        boolean hasNoNs = Helper.isNullOrEmpty(nameSpace);
        sb.append('<');
        if (hasNoNs == false)
        {
            sb.append(nameSpace);
            sb.append(':');
        }
        sb.append(tagName).append(">");
        sb.append(escapeXMLChars(value));
        sb.append("</");
        if (hasNoNs == false)
        {
            sb.append(nameSpace);
            sb.append(':');
        }
        sb.append(tagName);
        sb.append(">");
    }
   
    
    /**
     * Method to create the app-properties (part of meta data) as XML document
     * @return Formatted XML document
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    private Document createAppPropertiesDocument() throws IOException
    {
        StringBuilder sb = new StringBuilder();
        sb.append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");
        sb.append(createAppString());
        sb.append("</Properties>");
        return createXMLDocument(sb.toString(), "APPPROPERTIES");
    }    
    /**
     * Method to create the XML string for the app-properties document
     * @return String with formatted XML data
     */
    private String createAppString()
    {
        if (this.workbook.getWorkbookMetadata() == null) { return ""; }
        Metadata md = this.workbook.getWorkbookMetadata();
        StringBuilder sb = new StringBuilder();
        appendXMLtag(sb, "0", "TotalTime", null);
        appendXMLtag(sb, md.getApplication(), "Application", null);
        appendXMLtag(sb, "0", "DocSecurity", null);
        appendXMLtag(sb, "false", "ScaleCrop", null);
        appendXMLtag(sb, md.getManager(), "Manager", null);
        appendXMLtag(sb, md.getCompany(), "Company", null);
        appendXMLtag(sb, "false", "LinksUpToDate", null);
        appendXMLtag(sb, "false", "SharedDoc", null);
        appendXMLtag(sb, md.getHyperlinkBase(), "HyperlinkBase", null);
        appendXMLtag(sb, "false", "HyperlinksChanged", null);
        appendXMLtag(sb, md.getApplicationVersion(), "AppVersion", null);
        return sb.toString();
    }
    /**
     * Method to create the columns as XML string. This is used to define the width of columns
     * @param worksheet Worksheet to process
     * @return String with formatted XML data
     */
    private String createColsString(Worksheet worksheet)
    {
        if (worksheet.getColumns().size() > 0)
        {
            String col;
            String hidden = "";
            StringBuilder sb = new StringBuilder();
            
            for (Map.Entry<Integer, Column> column : worksheet.getColumns().entrySet())
            {
                if (column.getValue().getWidth() == worksheet.getDefaultColumnWidth() && column.getValue().isHidden() == false) { continue; }
                if (worksheet.getColumns().containsKey(column.getKey()))
                {
                    if (worksheet.getColumns().get(column.getKey()).isHidden() == true)
                    {
                        hidden = " hidden=\"1\"";
                    }
                }
                col = Integer.toString(column.getKey() + 1); // Add 1 for Address
                sb.append("<col customWidth=\"1\" width=\"").append(column.getValue().getWidth()).append("\" max=\"").append(col).append("\" min=\"").append(col).append("\"").append(hidden).append("/>");
            }
            String value = sb.toString();
            if (value.length() > 0)
            {
                return value;
            }
            else
            {
                return "";
            }
        }
        else
        {
            return "";
        }    
    }
    
    /**
     * Method to create the core-properties (part of meta data) as XML document
     * @return Formatted XML document
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    private Document createCorePropertiesDocument() throws IOException
    {
        StringBuilder sb = new StringBuilder();
        sb.append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
        sb.append(createCorePropertiesString());
        sb.append("</cp:coreProperties>");
        return createXMLDocument(sb.toString(),"COREPROPERTIES");
    }    
    /**
     * Method to create the XML string for the core-properties document
     * @return String with formatted XML data
     */
    private String createCorePropertiesString()
    {
        if (this.workbook.getWorkbookMetadata() == null) { return ""; }
        Metadata md = this.workbook.getWorkbookMetadata();
        StringBuilder sb = new StringBuilder();
        appendXMLtag(sb, md.getTitle(), "title", "dc");
        appendXMLtag(sb, md.getSubject(), "subject", "dc");
        appendXMLtag(sb, md.getCreator(), "creator", "dc");
        appendXMLtag(sb, md.getCreator(), "lastModifiedBy", "cp");
        appendXMLtag(sb, md.getKeywords(), "keywords", "cp");
        appendXMLtag(sb, md.getDescription(), "description", "dc");
        
        Calendar cal = new GregorianCalendar();
        DateFormat  df = new SimpleDateFormat("yyyy-MM-dd'T'hh:mm:ss'Z'");
        df.setCalendar(cal);
        Date now = cal.getTime();
        String time = df.format(now);
        
        sb.append("<dcterms:created xsi:type=\"dcterms:W3CDTF\">").append(time).append("</dcterms:created>");
        sb.append("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">").append(time).append("</dcterms:modified>");
        
        appendXMLtag(sb, md.getCategory(), "category", "cp");
        appendXMLtag(sb, md.getContentStatus(), "contentStatus", "cp");
        
        return sb.toString();
    }    
    /**
     * Method to create the merged cells string of the passed worksheet
     * @param sheet Worksheet to process
     * @return Formatted string with merged cell ranges
     */
    private String createMergedCellsString(Worksheet sheet)
    {
        if (sheet.getMergedCells().size() < 1)
        {
            return "";
        }
        Iterator<Map.Entry<String, Range>> itr;
        Map.Entry<String, Range> range;
        StringBuilder sb = new StringBuilder();
        sb.append("<mergeCells count=\"").append(sheet.getMergedCells().size()).append("\">");
        itr = sheet.getMergedCells().entrySet().iterator();
        while (itr.hasNext())
        {
            range = itr.next();
            sb.append("<mergeCell ref=\"").append(range.getValue().toString()).append("\"/>");
        }
        sb.append("</mergeCells>");
        return sb.toString();
    }
    /**
     * Method to create the XML string for the color-MRU part of the style sheet document (recent colors)
     * @return String with formatted XML data
     */
    private String createMruColorsString()
    {
        Font[] fonts = this.workbook.getStyleManager().getFonts();
        Fill[] fills = this.workbook.getStyleManager().getFills();
        StringBuilder sb = new StringBuilder();
        List<String> tempColors = new ArrayList<>();
        for (Font font : fonts) {
            if (Helper.isNullOrEmpty(font.getColorValue()) == true) { continue; }
            if (font.getColorValue().equals(Fill.DEFAULTCOLOR)) { continue; }
            if (tempColors.contains(font.getColorValue()) == false) { tempColors.add(font.getColorValue()); }
        }
        for (Fill fill : fills) {
            if (Helper.isNullOrEmpty(fill.getBackgroundColor()) == false)
            {
                if (fill.getBackgroundColor().equals(Fill.DEFAULTCOLOR) == false)
                {
                    if (tempColors.contains(fill.getBackgroundColor()) == false) { tempColors.add(fill.getBackgroundColor()); }
                }
            }
            if (Helper.isNullOrEmpty(fill.getForegroundColor()) == false)
            {
                if (fill.getForegroundColor().equals(Fill.DEFAULTCOLOR) == false)
                {
                    if (tempColors.contains(fill.getForegroundColor()) == false) { tempColors.add(fill.getForegroundColor()); }
                }
            }
        }
        if (tempColors.size() > 0)
        {
            sb.append("<mruColors>");
            for(int i = 0; i < tempColors.size(); i++)
            {
                sb.append("<color rgb=\"").append(tempColors.get(i)).append("\"/>");
            }
            sb.append("</mruColors>");
            return sb.toString();
        }
        else
        {
            return "";
        }
    }
    /**
     * Method to create a row string
     * @param columnFields List of cells
     * @param worksheet Worksheet to process
     * @return Formatted row string
     */
    private String createRowString(List<Cell> columnFields, Worksheet worksheet)
    {
        int rowNumber = columnFields.get(0).getRowNumber();
        String height = "";
        String hidden = "";
        if (worksheet.getRowHeights().containsKey(rowNumber))
        {
            if (worksheet.getRowHeights().get(rowNumber) != worksheet.getDefaultRowHeight())
            {
                height = " x14ac:dyDescent=\"0.25\" customHeight=\"1\" ht=\"" + worksheet.getRowHeights().get(rowNumber) + "\"";
            }
        }
        if (worksheet.getHiddenRows().containsKey(rowNumber))
        {
            if (worksheet.getHiddenRows().get(rowNumber) == true)
            {
                hidden = " hidden=\"1\"";
            }
        }
        int colNum = columnFields.size();
        StringBuilder sb = new StringBuilder(43 * colNum); // A row string size is according to statistics (random value) 43 times the column number
        //StringBuilder sb = new StringBuilder();
        if (colNum > 0)
        {
            sb.append("<row r=\"");
            sb.append((rowNumber + 1));
            sb.append("\"").append(height).append(hidden).append(">");
        }
        else
        {
            sb.append("<row").append(height).append(">");
        }
        String typeAttribute;
        String sValue, tValue;
        String value = "";
        boolean bVal;
        
        Date dVal;
        int col = 0;
        Cell item;
        for (int i = 0; i < colNum; i++)
        {
            item = columnFields.get(i);
            tValue = " ";
            if (item.getCellStyle() != null)
            {
                sValue = " s=\"" + item.getCellStyle().getInternalID() + "\" ";
            }
            else
            {
                sValue = "";
            }
            item.resolveCellType(); // Recalculate the type (for handling DEFAULT)
            if (item.getDataType() == Cell.CellType.BOOL)
            {
                typeAttribute = "b";
                tValue = " t=\"" + typeAttribute + "\" ";
                bVal = (boolean)item.getValue();
                if (bVal == true) { value = "1"; }
                else { value = "0"; }

            }
            // Number casting
            else if (item.getDataType() == Cell.CellType.NUMBER)
            {
                typeAttribute = "n";
                tValue = " t=\"" + typeAttribute + "\" ";
                Object o = item.getValue();
                if (o instanceof Byte)            { value = Byte.toString((byte)item.getValue()); }
                else if (o instanceof BigDecimal) { value = item.getValue().toString(); }
                else if (o instanceof Double)     { value = Double.toString((double)item.getValue()); }
                else if (o instanceof Float)      { value = Float.toString((float)item.getValue());}
                else if (o instanceof Integer)    { value = Integer.toString((int)item.getValue()); }
                else if (o instanceof Long)       { value = Long.toString((long)item.getValue()); }
                else if (o instanceof Short)      { value = Short.toString((short)item.getValue()); }
            }
            // Date parsing
            else if (item.getDataType() == Cell.CellType.DATE)
            {
                typeAttribute = "d";
                dVal = (Date)item.getValue();
                value = Double.toString(Helper.getOADate(dVal));
            }
            // String parsing
            else
            {
                if (item.getValue() == null)
                {
                    typeAttribute = "str";
                    value = "";
                }
                else // handle shared Strings
                {
                    //  value = item.getValue().toString();
                    if (item.getDataType() == Cell.CellType.FORMULA)
                    {
                        typeAttribute = "str";
                        value = item.getValue().toString();
                    }
                    else
                    {
                        typeAttribute = "s";
                        value = item.getValue().toString();
                        if (this.sharedStrings.containsKey(value) == false)
                        {
                            this.sharedStrings.add(value, Integer.toString(sharedStrings.size()));
                        }
                        value = this.sharedStrings.get(value);
                        this.sharedStringsTotalCount++;
                    }
                }
                tValue = " t=\"" + typeAttribute + "\" ";
            }
            if (item.getDataType() != Cell.CellType.EMPTY)
            {
                sb.append("<c").append(tValue).append("r=\"").append(item.getCellAddress()).append("\"").append(sValue).append(">");
                if (item.getDataType() == Cell.CellType.FORMULA)
                {
                    sb.append("<f>").append(XlsXWriter.escapeXMLChars(item.getValue().toString())).append("</f>");
                }
                else
                {
                    sb.append("<v>").append(XlsXWriter.escapeXMLChars(value)).append("</v>");
                }
                sb.append("</c>");
            }
            else // Empty cell
            {
                sb.append("<c").append(tValue).append("r=\"").append(item.getCellAddress()).append("\"").append(sValue).append("/>");
            }
            col++;
        }
        sb.append("</row>");
        return sb.toString();
    }
    
     /**
     * Method to create shared strings as XML document
     * @return Formatted XML document
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    private Document createSharedStringsDocument() throws IOException
    {
            StringBuilder sb = new StringBuilder();
            sb.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"");
            sb.append(this.sharedStringsTotalCount);
            sb.append("\" uniqueCount=\"");
            sb.append(this.sharedStrings.size());
            sb.append("\">");
            ArrayList<String> keys = this.sharedStrings.getKeys();
            //for (Map.Entry<String, String> str : sharedStrings.entrySet())
            for(int i = 0; i < keys.size(); i++)
            {  
                sb.append("<si><t>");
                //sb.append(escapeXMLChars(str.getKey()));
                sb.append(escapeXMLChars(keys.get(i)));
                sb.append("</t></si>");
            }
            sb.append("</sst>");
            return createXMLDocument(sb.toString(), "SHAREDSTRINGS");
    }
    /**
     * Method to create the protection string of the passed worksheet
     * @param sheet Worksheet to process
     * @return Formatted string with protection statement of the worksheet
     */
    private String createSheetProtectionString(Worksheet sheet)
    {
        if (sheet.isUseSheetProtection() == false)
        {
            return "";
        }
        HashMap<Worksheet.SheetProtectionValue, Integer> actualLockingValues = new HashMap<>();
        if (sheet.getSheetProtectionValues().isEmpty())
        {
            actualLockingValues.put(Worksheet.SheetProtectionValue.selectLockedCells, 1);
            actualLockingValues.put(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
        }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.objects) == false)
        {
            actualLockingValues.put(Worksheet.SheetProtectionValue.objects, 1);
        }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.scenarios) == false)
        {
            actualLockingValues.put(Worksheet.SheetProtectionValue.scenarios, 1);
        }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.selectLockedCells) == false )
        {
            if (actualLockingValues.containsKey(Worksheet.SheetProtectionValue.selectLockedCells) == false)
            {
                actualLockingValues.put(Worksheet.SheetProtectionValue.selectLockedCells, 1);
            }
        }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.selectUnlockedCells) == false || sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.selectLockedCells) == false)
        {
            if (actualLockingValues.containsKey(Worksheet.SheetProtectionValue.selectUnlockedCells) == false)
            {
                actualLockingValues.put(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
            }
        }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.formatCells)) { actualLockingValues.put(Worksheet.SheetProtectionValue.formatCells, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.formatColumns)) { actualLockingValues.put(Worksheet.SheetProtectionValue.formatColumns, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.formatRows)) { actualLockingValues.put(Worksheet.SheetProtectionValue.formatRows, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.insertColumns)) { actualLockingValues.put(Worksheet.SheetProtectionValue.insertColumns, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.insertRows)) { actualLockingValues.put(Worksheet.SheetProtectionValue.insertRows, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.insertHyperlinks)) { actualLockingValues.put(Worksheet.SheetProtectionValue.insertHyperlinks, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.deleteColumns)) { actualLockingValues.put(Worksheet.SheetProtectionValue.deleteColumns, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.deleteRows)) { actualLockingValues.put(Worksheet.SheetProtectionValue.deleteRows, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.sort)) { actualLockingValues.put(Worksheet.SheetProtectionValue.sort, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.autoFilter)) { actualLockingValues.put(Worksheet.SheetProtectionValue.autoFilter, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.pivotTables)) { actualLockingValues.put(Worksheet.SheetProtectionValue.pivotTables, 0); }
        StringBuilder sb = new StringBuilder();
        sb.append("<sheetProtection");
        String temp;
        Iterator<Map.Entry<Worksheet.SheetProtectionValue, Integer>> itr;
        Map.Entry<Worksheet.SheetProtectionValue, Integer> item;
        itr = actualLockingValues.entrySet().iterator();
        while (itr.hasNext())
        {
            item = itr.next();
            temp = item.getKey().name();// Note! If the enum names differs from the OOXML definitions, this method will cause invalid OOXML entries
            sb.append(" ").append(temp).append("=\"").append(item.getValue()).append("\"");
        }
        if (Helper.isNullOrEmpty(sheet.getSheetProtectionPassword()) == false)
        {
            String hash = generatePasswordHash(sheet.getSheetProtectionPassword());
            sb.append(" password=\"").append(hash).append("\"");
        }
        sb.append(" sheet=\"1\"/>");
        return sb.toString();
    }    
    
    /**
     * Method to create the XML string for the border part of the style sheet document

     * @return String with formatted XML data
     */
    private String createStyleBorderString()
    {
        Border[] borderStyles = this.workbook.getStyleManager().getBorders();
        StringBuilder sb = new StringBuilder();
        for (Border borderStyle : borderStyles) {
            if (borderStyle.isDiagonalDown() == true && borderStyle.isDiagonalUp() == false) { sb.append("<border diagonalDown=\"1\">"); }
            else if (borderStyle.isDiagonalDown() == false && borderStyle.isDiagonalUp() == true) { sb.append("<border diagonalUp=\"1\">"); }
            else if (borderStyle.isDiagonalDown() == true && borderStyle.isDiagonalUp() == true) { sb.append("<border diagonalDown=\"1\" diagonalUp=\"1\">"); }
            else { sb.append("<border>"); }
            if (borderStyle.getLeftStyle() != Border.StyleValue.none)
            {
                sb.append("<left style=\"").append(Border.getStyleName(borderStyle.getLeftStyle())).append("\">");
                if (Helper.isNullOrEmpty(borderStyle.getLeftColor()) == true) { sb.append("<color rgb=\"").append(borderStyle.getLeftColor()).append("\"/>"); }
                else { sb.append("<color auto=\"1\"/>"); }
                sb.append("</left>");
            }
            else
            {
                sb.append("<left/>");
            }
            if (borderStyle.getRightStyle() != Border.StyleValue.none)
            {
                sb.append("<right style=\"").append(Border.getStyleName(borderStyle.getRightStyle())).append("\">");
                if (Helper.isNullOrEmpty(borderStyle.getRightColor()) == true) { sb.append("<color rgb=\"").append(borderStyle.getRightColor()).append("\"/>"); }
                else { sb.append("<color auto=\"1\"/>"); }
                sb.append("</right>");
            }
            else
            {
                sb.append("<right/>");
            }
            if (borderStyle.getTopStyle() != Border.StyleValue.none)
            {
                sb.append("<top style=\"").append(Border.getStyleName(borderStyle.getTopStyle())).append("\">");
                if (Helper.isNullOrEmpty(borderStyle.getTopColor()) == true) { sb.append("<color rgb=\"").append(borderStyle.getTopColor()).append("\"/>"); }
                else { sb.append("<color auto=\"1\"/>"); }
                sb.append("</top>");
            }
            else
            {
                sb.append("<top/>");
            }
            if (borderStyle.getBottomStyle() != Border.StyleValue.none)
            {
                sb.append("<bottom style=\"").append(Border.getStyleName(borderStyle.getBottomStyle())).append("\">");
                if (Helper.isNullOrEmpty(borderStyle.getBottomColor()) == true) { sb.append("<color rgb=\"").append(borderStyle.getBottomColor()).append("\"/>"); }
                else { sb.append("<color auto=\"1\"/>"); }
                sb.append("</bottom>");
            }
            else
            {
                sb.append("<bottom/>");
            }
            if (borderStyle.getDiagonalStyle() != Border.StyleValue.none)
            {
                sb.append("<diagonal style=\"").append(Border.getStyleName(borderStyle.getDiagonalStyle())).append("\">");
                if (Helper.isNullOrEmpty(borderStyle.getDiagonalColor()) == true) { sb.append("<color rgb=\"").append(borderStyle.getDiagonalColor()).append("\"/>"); }
                else { sb.append("<color auto=\"1\"/>"); }
                sb.append("</diagonal>");
            }
            else
            {
                sb.append("<diagonal/>");
            }
            sb.append("</border>");
        }
        return sb.toString();
    }    
    
    /**
     * Method to create the XML string for the fill part of the style sheet document
     * @return String with formatted XML data
     */
    private String createStyleFillString()
    {
        Fill[] fillStyles = this.workbook.getStyleManager().getFills();
        StringBuilder sb = new StringBuilder();
        for (Fill fillStyle : fillStyles) {
            sb.append("<fill>");
            sb.append("<patternFill patternType=\"").append(Fill.getPatternName(fillStyle.getPatternFill())).append("\"");
            if (fillStyle.getPatternFill() == Fill.PatternValue.solid)
            {
                sb.append(">");
                sb.append("<fgColor rgb=\"").append(fillStyle.getForegroundColor()).append("\"/>");
                sb.append("<bgColor indexed=\"");
                sb.append(fillStyle.getIndexedColor());
                sb.append("\"/>");
                sb.append("</patternFill>");
            }
            else if (fillStyle.getPatternFill() == Fill.PatternValue.mediumGray || fillStyle.getPatternFill() == Fill.PatternValue.lightGray || fillStyle.getPatternFill() == Fill.PatternValue.gray0625 || fillStyle.getPatternFill() == Fill.PatternValue.darkGray)
            {
                sb.append(">");
                sb.append("<fgColor rgb=\"").append(fillStyle.getForegroundColor()).append("\"/>");
                if (Helper.isNullOrEmpty(fillStyle.getBackgroundColor()) == false)
                {
                    sb.append("<bgColor rgb=\"").append(fillStyle.getBackgroundColor()).append("\"/>");
                }
                sb.append("</patternFill>");
            }
            else
            {
                sb.append("/>");
            }
            sb.append("</fill>");
        }
        return sb.toString();
    }  
    /**
     * Method to create the XML string for the font part of the style sheet document
     * @return String with formatted XML data
     */
    private String createStyleFontString()
    {
        Font[] fontStyles = this.workbook.getStyleManager().getFonts();
        StringBuilder sb = new StringBuilder();
        for (Font fontStyle : fontStyles) {
            sb.append("<font>");
            if (fontStyle.isBold() == true) { sb.append("<b/>"); }
            if (fontStyle.isItalic() == true) { sb.append("<i/>"); }
            if (fontStyle.isUnderline() == true) { sb.append("<u/>"); }
            if (fontStyle.isDoubleUnderline() == true) { sb.append("<u val=\"double\"/>"); }
            if (fontStyle.isStrike() == true) { sb.append("<strike/>"); }
            if (fontStyle.getVerticalAlign() == Font.VerticalAlignValue.subscript) { sb.append("<vertAlign val=\"subscript\"/>"); }
            else if (fontStyle.getVerticalAlign() == Font.VerticalAlignValue.superscript) { sb.append("<vertAlign val=\"superscript\"/>"); }
            sb.append("<sz val=\"");
            sb.append(fontStyle.getSize());
            sb.append("\"/>");
            if (Helper.isNullOrEmpty(fontStyle.getColorValue()))
            {
                sb.append("<color theme=\"");
                sb.append(fontStyle.getColorTheme());
                sb.append("\"/>");
            }
            else
            {
                sb.append("<color rgb=\"").append(fontStyle.getColorValue()).append("\"/>");
            }
            sb.append("<name val=\"").append(fontStyle.getName()).append("\"/>");
            sb.append("<family val=\"").append(fontStyle.getFamily()).append("\"/>");
            if (fontStyle.getScheme() != Font.SchemeValue.none)
            {
                if (fontStyle.getScheme() == Font.SchemeValue.major)
                { sb.append("<scheme val=\"major\"/>"); }
                else if (fontStyle.getScheme() == Font.SchemeValue.minor)
                { sb.append("<scheme val=\"minor\"/>"); }
            }
            if (Helper.isNullOrEmpty(fontStyle.getCharset()) == false)
            {
                sb.append("<charset val=\"").append(fontStyle.getCharset()).append("\"/>");
            }
            sb.append("</font>");
        }
        return sb.toString();
    }    
    /**
     * Method to create the XML string for the number format part of the style sheet document
     * @return String with formatted XML data
     */
    private String createStyleNumberFormatString()
    {
        NumberFormat[] numberFormatStyles = this.workbook.getStyleManager().getNumberFormats();
        StringBuilder sb = new StringBuilder();
        for (NumberFormat numberFormatStyle : numberFormatStyles) {
            if (numberFormatStyle.isCustomFormat() == true)
            {
                sb.append("<numFmt formatCode=\"").append(numberFormatStyle.getCustomFormatCode()).append("\" numFmtId=\"");
                sb.append(numberFormatStyle.getCustomFormatID());
                sb.append("\"/>");
            }
        }
        return sb.toString();
    }
    /**
     * Method to create a style sheet as XML document
     * @return Formatted XML document
     * @throws StyleException Thrown if a style was not referenced in the style sheet
     * @throws RangeException Thrown if a referenced cell was out of range
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    private Document createStyleSheetDocument() throws IOException
    {
        String bordersString = createStyleBorderString();
        String fillsString = createStyleFillString();
        String fontsString = createStyleFontString();
        String numberFormatsString = createStyleNumberFormatString();
        int numFormatCount = getNumberFormatStringCounter();
        String xfsStings = createStyleXfsString();
        String mruColorString = createMruColorsString();
        StringBuilder sb = new StringBuilder();
        
        sb.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
        if (numFormatCount > 0)
        {
            sb.append("<numFmts count=\"");
            sb.append(numFormatCount);
            sb.append("\">");
            sb.append(numberFormatsString).append("</numFmts>");
        }
        sb.append("<fonts x14ac:knownFonts=\"1\" count=\"");    
        sb.append(this.workbook.getStyleManager().getFontStyleNumber());
        sb.append("\">");
        sb.append(fontsString).append("</fonts>");
        sb.append("<fills count=\"");
        sb.append(this.workbook.getStyleManager().getFillStyleNumber());
        sb.append("\">");
        sb.append(fillsString).append("</fills>");
        sb.append("<borders count=\"");
        sb.append(this.workbook.getStyleManager().getBorderStyleNumber());
        sb.append("\">");
        sb.append(bordersString).append("</borders>");
        sb.append("<cellXfs count=\"");
        sb.append(this.workbook.getStyleManager().getStyleNumber());
        sb.append("\">");
        sb.append(xfsStings).append("</cellXfs>");
        if (this.workbook.getWorkbookMetadata() != null)
        {
            if (Helper.isNullOrEmpty(mruColorString) == false && this.workbook.getWorkbookMetadata().isUseColorMRU() == true)
            {
                sb.append("<colors>");
                sb.append(mruColorString);
                sb.append("</colors>");
            }
        }
        sb.append("</styleSheet>");
        return createXMLDocument(sb.toString(), "STYLESHEET");
    }
    
    /**
     * Method to create the XML string for the XF part of the style sheet document
     * @return String with formatted XML data
     * @throws RangeException Thrown if a referenced cell was out of range
     */
    private String createStyleXfsString()
    {
        Style[] styles = this.workbook.getStyleManager().getStyles();
        StringBuilder sb = new StringBuilder();
        StringBuilder sb2;
        String alignmentString, protectionString;
        int formatNumber, textRotation;
        for (Style style : styles) {
            textRotation = style.getCellXf().calculateInternalRotation();
            alignmentString = "";
            protectionString = "";
            if (style.getCellXf().getHorizontalAlign() != CellXf.HorizontalAlignValue.none || style.getCellXf().getVerticalAlign() != CellXf.VerticalAlignValue.none || style.getCellXf().getAlignment() != CellXf.TextBreakValue.none || textRotation != 0)
            {
                sb2 = new StringBuilder();
                sb2.append("<alignment");
                if (style.getCellXf().getHorizontalAlign() != CellXf.HorizontalAlignValue.none)
                {
                    sb2.append(" horizontal=\"");
                    if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.center) { sb2.append("center"); }
                    else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.right) { sb2.append("right"); }
                    else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.centerContinuous) { sb2.append("centerContinuous"); }
                    else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.distributed) { sb2.append("distributed"); }
                    else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.fill) { sb2.append("fill"); }
                    else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.general) { sb2.append("general"); }
                    else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.justify) { sb2.append("justify"); }
                    else { sb2.append("left"); }
                    sb2.append("\"");
                }
                if (style.getCellXf().getVerticalAlign() != CellXf.VerticalAlignValue.none)
                {
                    sb2.append(" vertical=\"");
                    if (style.getCellXf().getVerticalAlign() == CellXf.VerticalAlignValue.center) { sb2.append("center"); }
                    else if (style.getCellXf().getVerticalAlign() == CellXf.VerticalAlignValue.distributed) { sb2.append("distributed"); }
                    else if (style.getCellXf().getVerticalAlign() == CellXf.VerticalAlignValue.justify) { sb2.append("justify"); }
                    else if (style.getCellXf().getVerticalAlign() == CellXf.VerticalAlignValue.top) { sb2.append("top"); }
                    else { sb2.append("bottom"); }
                    sb2.append("\"");
                }

                if (style.getCellXf().getAlignment() != CellXf.TextBreakValue.none)
                {
                    if (style.getCellXf().getAlignment() == CellXf.TextBreakValue.shrinkToFit) { sb2.append(" shrinkToFit=\"1"); }
                    else { sb2.append(" wrapText=\"1"); }
                    sb2.append("\"");
                }
                if (textRotation != 0)
                {
                    sb2.append(" textRotation=\"");
                    sb2.append(textRotation);
                    sb2.append("\"");
                }
                sb2.append("/>"); // </xf>
                alignmentString = sb2.toString();
            }
            if (style.getCellXf().isHidden() == true || style.getCellXf().isLocked() == true)
            {
                if (style.getCellXf().isHidden() == true && style.getCellXf().isLocked() == true)
                {
                    protectionString = "<protection locked=\"1\" hidden=\"1\"/>";
                }
                else if (style.getCellXf().isHidden() == true && style.getCellXf().isLocked() == false)
                {
                    protectionString = "<protection hidden=\"1\" locked=\"0\"/>";
                }
                else
                {
                    protectionString = "<protection hidden=\"0\" locked=\"1\"/>";
                }
            }
            sb.append("<xf numFmtId=\"");
            if (style.getNumberFormat().isCustomFormat() == true)
            {
                sb.append(style.getNumberFormat().getCustomFormatID());
            }
            else
            {
                formatNumber = style.getNumberFormat().getNumber().getValue();
                sb.append(formatNumber);
            }
            sb.append("\" borderId=\"");
            sb.append(style.getBorder().getInternalID());
            sb.append("\" fillId=\"");
            sb.append(style.getFill().getInternalID());
            sb.append("\" fontId=\"");
            sb.append(style.getFont().getInternalID());
            if (style.getFont().isDefaultFont() == false)
            {
                sb.append("\" applyFont=\"1");
            }
            if (style.getFill().getPatternFill() != Fill.PatternValue.none)
            {
                sb.append("\" applyFill=\"1");
            }
            if (style.getBorder().isEmpty() == false)
            {
                sb.append("\" applyBorder=\"1");
            }
            if (alignmentString.isEmpty() == false || style.getCellXf().isForceApplyAlignment() == true)
            {
                sb.append("\" applyAlignment=\"1");
            }
            if (protectionString.isEmpty() == false)
            {
                sb.append("\" applyProtection=\"1");
            }            
            if (style.getNumberFormat().getNumber() != NumberFormat.FormatNumber.none)
            {
                sb.append("\" applyNumberFormat=\"1\"");
            }
            else
            {
                sb.append("\""); 
            }
            if (alignmentString.isEmpty() == false  || protectionString.isEmpty() == false)
            {
                sb.append(">");
                sb.append(alignmentString);
                sb.append(protectionString);
                sb.append("</xf>");
            }
            else
            {
                sb.append("/>");
            }
        }
        return sb.toString();
    }    
    /**
     * Method to create a workbook as XML document
     * @return Formatted XML document
     * @throws RangeException Thrown if a referenced cell was out of range
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    private Document createWorkbookDocument() throws IOException
    {
        if (this.workbook.getWorksheets().isEmpty())
        {
            throw new RangeException("UnknownRangeException", "The workbook can not be created because no worksheet was defined.");
        }
        StringBuilder sb = new StringBuilder();
        sb.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        if (this.workbook.getSelectedWorksheet() > 0)
        {
            sb.append("<bookViews><workbookView activeTab=\"");
            sb.append(this.workbook.getSelectedWorksheet());
            sb.append("\"/></bookViews>");
        }
        if (this.workbook.isWorkbookProtectionUsed() == true)
        {
            sb.append("<workbookProtection");
            if (this.workbook.isWindowsLockedIfProtected() == true)
            {
                sb.append(" lockWindows=\"1\"");
            }
            if (this.workbook.isStructureLockedIfProtected() == true)
            {
                sb.append(" lockStructure=\"1\"");
            }
            if (Helper.isNullOrEmpty(this.workbook.getWorkbookProtectionPassword()) == false)
            {
                sb.append("workbookPassword=\"");
                sb.append(generatePasswordHash(this.workbook.getWorkbookProtectionPassword()));
                sb.append("\"");
            }
            sb.append("/>");
        }
        sb.append("<sheets>");
        int id;
        for (int i = 0; i < this.workbook.getWorksheets().size(); i++)
        {
            id = this.workbook.getWorksheets().get(i).getSheetID();
            sb.append("<sheet r:id=\"rId");
            sb.append(id);
            sb.append("\" sheetId=\"");
            sb.append(id);
            sb.append("\" name=\"").append(XlsXWriter.escapeXMLAttributeChars(this.workbook.getWorksheets().get(i).getSheetName())).append("\"/>");
        }
        sb.append("</sheets>");
        sb.append("</workbook>");
        return createXMLDocument(sb.toString(), "WORKBOOK");
    }
    /**
     * Method to create a worksheet part as XML document
     * @param worksheet worksheet object to process
     * @return Formatted XML document
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    private Document createWorksheetPart(Worksheet worksheet) throws IOException
    {
        worksheet.recalculateAutoFilter();
        worksheet.recalculateColumns();
        List<List<Cell>> celldata = getSortedSheetData(worksheet);
        StringBuilder sb = new StringBuilder();
        String line;
        sb.append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
        
        if (worksheet.getSelectedCells() != null)
        {
            sb.append("<sheetViews><sheetView workbookViewId=\"0\"");
            if (this.workbook.getSelectedWorksheet() == worksheet.getSheetID() - 1)
            {
                sb.append(" tabSelected=\"1\"");
            }
            sb.append("><selection sqref=\"");
            sb.append(worksheet.getSelectedCells().toString());
            sb.append("\" activeCell=\"");
            sb.append(worksheet.getSelectedCells().StartAddress.toString());
            sb.append("\"/></sheetView></sheetViews>");
        }

        sb.append("<sheetFormatPr x14ac:dyDescent=\"0.25\" defaultRowHeight=\"");
        sb.append(worksheet.getDefaultRowHeight());
        sb.append("\" baseColWidth=\"");
        sb.append(worksheet.getDefaultColumnWidth());
        sb.append("\"/>");
        String colWidths = createColsString(worksheet);
        if (Helper.isNullOrEmpty(colWidths) == false)
        {
            sb.append("<cols>");
            sb.append(colWidths);
            sb.append("</cols>");
        }
        sb.append("<sheetData>");
        for(int i = 0; i < celldata.size(); i++)
        {
            line = createRowString(celldata.get(i), worksheet);
            sb.append(line);
        }
        sb.append("</sheetData>");
        
        sb.append(createMergedCellsString(worksheet));
        sb.append(createSheetProtectionString(worksheet));
        if (worksheet.getAutoFilterRange() != null)
        {
            sb.append("<autoFilter ref=\"").append(worksheet.getAutoFilterRange().toString()).append("\"/>");
        }
        sb.append("</worksheet>");
        //testing.Performance.SaveLoggedValues("LineLength.xlsx");
        return createXMLDocument(sb.toString(), "WORKSHEET: " + worksheet.getSheetName());
    }
    
    /**
     * Creates a XML document from a string
     * @param rawInput String to process
     * @param title Title for interception / debugging purpose
     * @return Formatted XML document
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    public Document createXMLDocument(String rawInput, String title) throws IOException
    {
        try
        {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();   
        DocumentBuilder docBuilder = factory.newDocumentBuilder();
        InputSource input = new InputSource(new StringReader( rawInput ));
        input.setEncoding("UTF-8");
        Document doc = docBuilder.parse( input );
        doc.setXmlVersion("1.0");
        doc.setXmlStandalone(true);
        if (this.interceptDocuments == true)
        {
            this.interceptedDocuments.put(title, doc);
            System.out.println("DEBUG: Document '" + title + "' was intercepted");
        }
        return doc;
        }
        catch(Exception e)
        {
            throw new IOException("XmlDocumentException","There was an error while creating the XML document. Please see the inner exception.", e);
        }
    }    
    
    /**
     * Gets the number of custom number formats
     * @return Number of custom number formats to apply in the style document
     */
    private int getNumberFormatStringCounter()
    {
        NumberFormat[] numberFormatStyles = this.workbook.getStyleManager().getNumberFormats();
        int counter = 0;
        for (NumberFormat numberFormatStyle : numberFormatStyles) {
            if (numberFormatStyle.isCustomFormat() == true)
            {
                counter++;
            }            
        }
        return counter;
    }    
      
    /**
     * Method to sort the cells of a worksheet as preparation for the XML document
     * @param sheet Worksheet to process 
     * @return Two dimensional array of Cell object
     */
    private List<List<Cell>> getSortedSheetData(Worksheet sheet)
     {
         List<Cell> temp = new ArrayList<>();
        Map.Entry<String, Cell> entry;
        Iterator<Map.Entry<String, Cell>> itr = sheet.getCells().entrySet().iterator();
        while (itr.hasNext())
        {  
            entry = itr.next();
            temp.add(entry.getValue());
        }
         Collections.sort(temp);             
         List<Cell> line = new ArrayList<>();
         List<List<Cell>> output = new ArrayList<>();
         if (temp.size() > 0)
         {
             int rowNumber = temp.get(0).getRowNumber();
             for (int i = 0; i < temp.size(); i++)
             {
                 if (temp.get(i).getRowNumber() != rowNumber)
                 {
                     output.add(line);
                     line = new ArrayList<>();
                     rowNumber = temp.get(i).getRowNumber();
                 }
                 line.add(temp.get(i));
             }
             if (line.size() > 0)
             {
                 output.add(line);
             }
         }
         return output;
     }

    
    /**
     * Method to save the workbook
     * @throws IOException Thrown in case of an error
     */
    public void save() throws IOException
    {
        try
        {
            FileOutputStream dest = new FileOutputStream(this.workbook.getFilename());
            saveAsStream(dest);
        }
        catch (Exception e)
        {
            throw new IOException("SaveException","There was an error while creating the workbook document during saving to a file. Please see the inner exception:" + e.getMessage(), e);
        }   
    }
    
    public void saveAsStream(OutputStream stream) throws IOException
    {
          try
        {
            this.workbook.resolveMergedCells();
            Document doc;
            Document app = createAppPropertiesDocument();
            Document core = createCorePropertiesDocument();
            Document styles = createStyleSheetDocument();
            Document book = createWorkbookDocument();
            String file;
            Worksheet sheet;
            Packer p = new Packer(this);
            Packer.Relationship rel = p.createRelationship("_rels/.rels");
            rel.addRelationshipEntry("/xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
            rel.addRelationshipEntry("/docProps/core.xml", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
            rel.addRelationshipEntry("/docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
            rel = p.createRelationship("xl/_rels/workbook.xml.rels");
            for(int i = 0; i < this.workbook.getWorksheets().size(); i++)
            {
                sheet = this.workbook.getWorksheets().get(i);
                doc = createWorksheetPart(sheet);
                file = "sheet" + sheet.getSheetID() + ".xml";
                rel.addRelationshipEntry("/xl/worksheets/" + file, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                p.addPart("xl/worksheets/" + file, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", doc);
            }
            rel.addRelationshipEntry("/xl/styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            rel.addRelationshipEntry("/xl/sharedStrings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
            p.addPart("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml", core);
            p.addPart("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml", app);
            p.addPart("xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", createSharedStringsDocument());
            p.addPart("xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", book, false);
            p.addPart("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", styles);
            p.pack(stream);
        }
        catch (Exception e)
        {
            throw new IOException("SaveException","There was an error while creating the workbook document during writing to a stream. Please see the inner exception:" + e.getMessage(), e);
        }      
    }    
    
// ### S T A T I C   M E T H O D S ###        
    /**
     * Method to convert an XML document to a byte array
     * @param document Document to process
     * @return array of bytes (UTF-8)
     * @throws IOException Thrown if the document could not be converted to a byte array
     */
    public static byte[] createBytesFromDocument(Document document) throws IOException
    {
        try
        {
            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            ByteArrayOutputStream bs = new ByteArrayOutputStream();
            Result output = new StreamResult(bs);
            Source input = new DOMSource(document);
            
            transformer.transform(input, output);
            bs.flush();
            byte[] bytes = bs.toByteArray();
            bs.close();
            return bytes;
        }
        catch(Exception e)
        {
            throw new IOException("ByteSteamException","There was an error while creating the byte array. Please see the inner exception.", e);
        }
    }    

    /**
     * Method to escape XML characters in an XML attribute
     * @param input Input string to process
     * @return Escaped string
     */
    private static String escapeXMLAttributeChars(String input)
    {
        input = escapeXMLChars(input); // Sanitize string from illegal characters beside quotes
        input = input.replace("\"", "&quot;");
        return input;
    }
    /**
     * Method to escape XML characters between two XML tags<br>Note: The XML specs allow characters up to the character value of 0x10FFFF. However, the Java char range is only up to 0xFFFF. PicoXLSX4j will neglect all values above this level in the sanitizing check. Illegal characters like 0x1 will be replaced with a white space (0x20)
     * @param input Input string to process
     * @return Escaped string
     */
    private static String escapeXMLChars(String input)
    {
        int len = input.length();
        List<Integer> illegalCharacters = new ArrayList<>(len);
        List<Integer> characterTypes = new ArrayList<>(len);
        int i;
        char c;
        for (i = 0; i < len; i++)
        {
            c = input.charAt(i);
            if  ((c < 0x9) || (c > 0xA && c < 0xD) || (c > 0xD && c < 0x20) || (c > 0xD7FF && c < 0xE000) || (c > 0xFFFD))
            {
                illegalCharacters.add(i);
                characterTypes.add(0);
                continue;
            } // Note: XML specs allow characters up to 0x10FFFF. However, the Java char range is only up to 0xFFFF; Higher values are neglected here
            if (c == 0x3C) // <
            {
                illegalCharacters.add(i);
                characterTypes.add(1);
            }
            else if (c == 0x3E) // >
            {
                illegalCharacters.add(i);
                characterTypes.add(2);
            }
            else if (c == 0x26) // &
            {
                illegalCharacters.add(i);
                characterTypes.add(3);
            }
        }
        if (illegalCharacters.isEmpty())
        {
            return input;
        }  
        StringBuilder sb = new StringBuilder(len);
        int lastIndex = 0;
        len = illegalCharacters.size();
        int j, type;
        for (i = 0; i < len; i++)
        {
            j = illegalCharacters.get(i);
            type = characterTypes.get(i);
            sb.append(input, lastIndex, j);
            if (type == 0)
            {
                sb.append(' '); // Whitespace as fall back on illegal character
            }
            else if (type == 1) // replace <
            {
                sb.append("&lt;");
            }
            else if (type == 2) // replace >
            {
                sb.append("&gt;");
            }
            else if (type == 3) // replace &
            {
                sb.append("&amp;");
            }
            lastIndex = j + 1;
        }
        sb.append(input.substring(lastIndex));
        return sb.toString();
    }
    
    /**
     * Method to generate an Excel internal password hash to protect workbooks or worksheets<br>
     * This method is derived from the c++ implementation by Kohei Yoshida (<a href="http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/">http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/</a>)<br>
     * <b>WARNING!</b> Do not use this method to encrypt 'real' passwords or data outside from PicoXLSX4j. This is only a minor security feature. Use a proper cryptography method instead.
     * @param password Password as plain text
     * @return Encoded password
     */
    private static String generatePasswordHash(String password)
    {
        if (Helper.isNullOrEmpty(password)) { return ""; }
        int passwordLength = password.length();
        int passwordHash = 0;
        char character;
        for(int i = passwordLength; i > 0; i--)
        {
            character = password.charAt(i - 1);
            passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
            passwordHash ^= character;
        }
        passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
        passwordHash ^= (0x8000 | ('N' << 8) | 'K');
        passwordHash ^= passwordLength;
        return Integer.toHexString(passwordHash).toUpperCase();
    }    
    
}
