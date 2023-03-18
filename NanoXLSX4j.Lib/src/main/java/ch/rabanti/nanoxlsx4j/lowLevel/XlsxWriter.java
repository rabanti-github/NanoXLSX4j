/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.io.StringReader;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.xml.sax.InputSource;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Column;
import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.Metadata;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.styles.Border;
import ch.rabanti.nanoxlsx4j.styles.CellXf;
import ch.rabanti.nanoxlsx4j.styles.Fill;
import ch.rabanti.nanoxlsx4j.styles.Font;
import ch.rabanti.nanoxlsx4j.styles.NumberFormat;
import ch.rabanti.nanoxlsx4j.styles.Style;
import ch.rabanti.nanoxlsx4j.styles.StyleManager;

/**
 * Class for low level handling (XML, formatting, preparing of packing)<br>
 * This class is only for internal use. Use the high level API (e.g. class
 * Workbook) to manipulate data and create Excel files.
 *
 * @author Raphael Stoeckli
 */
public class XlsxWriter {

	// ### C O N S T A N T S ###

	/**
	 * Minimum valid OAdate value (1900-01-01)
	 */
	public static final double MIN_OADATE_VALUE = 0d;
	/**
	 * Maximum valid OAdate value (9999-12-31)
	 */
	public static final double MAX_OADATE_VALUE = 2958465.999988426d;

	// ### P R I V A T E F I E L D S ###
	private final SortedMap sharedStrings;
	private int sharedStringsTotalCount;
	private final Workbook workbook;
	private StyleManager styles;

	// ### C O N S T R U C T O R S ###

	/**
	 * Constructor with defined workbook object
	 *
	 * @param workbook
	 *            Workbook to process
	 */
	public XlsxWriter(Workbook workbook) {
		this.workbook = workbook;
		this.sharedStrings = new SortedMap();
		this.sharedStringsTotalCount = 0;
	}

	// ### M E T H O D S ###

	/**
	 * Method to append a simple XML tag with an enclosed value to the passed
	 * StringBuilder
	 *
	 * @param sb
	 *            StringBuilder to append
	 * @param value
	 *            Value of the XML element
	 * @param tagName
	 *            Tag name of the XML element
	 * @param nameSpace
	 *            Optional XML name space. Can be empty or null
	 */
	private void appendXmlTag(StringBuilder sb, String value, String tagName, String nameSpace) {
		if (Helper.isNullOrEmpty(value)) {
			return;
		}
		boolean hasNoNs = Helper.isNullOrEmpty(nameSpace);
		sb.append('<');
		if (!hasNoNs) {
			sb.append(nameSpace);
			sb.append(':');
		}
		sb.append(tagName).append(">");
		sb.append(escapeXMLChars(value));
		sb.append("</");
		if (!hasNoNs) {
			sb.append(nameSpace);
			sb.append(':');
		}
		sb.append(tagName);
		sb.append(">");
	}

	/**
	 * Method to create the app-properties (part of meta data) as XML document
	 *
	 * @return Formatted XML document
	 * @throws IOException
	 *             Thrown in case of an error while creating the XML document
	 */
	private Document createAppPropertiesDocument() throws IOException {
        String sb = "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                createAppString() +
                "</Properties>";
		return createXMLDocument(sb);
	}

	/**
	 * Method to create the XML string for the app-properties document
	 *
	 * @return String with formatted XML data
	 */
	private String createAppString() {
		if (this.workbook.getWorkbookMetadata() == null) {
			return "";
		}
		Metadata md = this.workbook.getWorkbookMetadata();
		StringBuilder sb = new StringBuilder();
		appendXmlTag(sb, "0", "TotalTime", null);
		appendXmlTag(sb, md.getApplication(), "Application", null);
		appendXmlTag(sb, "0", "DocSecurity", null);
		appendXmlTag(sb, "false", "ScaleCrop", null);
		appendXmlTag(sb, md.getManager(), "Manager", null);
		appendXmlTag(sb, md.getCompany(), "Company", null);
		appendXmlTag(sb, "false", "LinksUpToDate", null);
		appendXmlTag(sb, "false", "SharedDoc", null);
		appendXmlTag(sb, md.getHyperlinkBase(), "HyperlinkBase", null);
		appendXmlTag(sb, "false", "HyperlinksChanged", null);
		appendXmlTag(sb, md.getApplicationVersion(), "AppVersion", null);
		return sb.toString();
	}

	/**
	 * Method to create the columns as XML string. This is used to define the width
	 * of columns
	 *
	 * @param worksheet
	 *            Worksheet to process
	 * @return String with formatted XML data
	 */
	private String createColsString(Worksheet worksheet) {
		if (worksheet.getColumns().size() > 0) {
			String col;
			String hidden = "";
			StringBuilder sb = new StringBuilder();

			for (Map.Entry<Integer, Column> column : worksheet.getColumns().entrySet()) {
				if (column.getValue().getWidth() == worksheet.getDefaultColumnWidth() && !column.getValue().isHidden()) {
					continue;
				}
				if (worksheet.getColumns().containsKey(column.getKey())) {
					if (worksheet.getColumns().get(column.getKey()).isHidden()) {
						hidden = " hidden=\"1\"";
					}
				}
				col = Integer.toString(column.getKey() + 1); // Add 1 for Address
				float width = Helper.getInternalColumnWidth(column.getValue().getWidth());
				sb.append("<col customWidth=\"1\" width=\"").append(width).append("\" max=\"").append(col).append("\" min=\"").append(col).append("\"")
						.append(hidden).append("/>");
			}
			String value = sb.toString();
			if (value.length() > 0) {
				return value;
			}
			else {
				return "";
			}
		}
		else {
			return "";
		}
	}

	/**
	 * Method to create the core-properties (part of meta data) as XML document
	 *
	 * @return Formatted XML document
	 * @throws IOException
	 *             Thrown in case of an error while creating the XML document
	 */
	private Document createCorePropertiesDocument() throws IOException {
        String sb = "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" +
                createCorePropertiesString() +
                "</cp:coreProperties>";
		return createXMLDocument(sb);
	}

	/**
	 * Method to create the XML string for the core-properties document
	 *
	 * @return String with formatted XML data
	 */
	private String createCorePropertiesString() {
		if (this.workbook.getWorkbookMetadata() == null) {
			return "";
		}
		Metadata md = this.workbook.getWorkbookMetadata();
		StringBuilder sb = new StringBuilder();
		appendXmlTag(sb, md.getTitle(), "title", "dc");
		appendXmlTag(sb, md.getSubject(), "subject", "dc");
		appendXmlTag(sb, md.getCreator(), "creator", "dc");
		appendXmlTag(sb, md.getCreator(), "lastModifiedBy", "cp");
		appendXmlTag(sb, md.getKeywords(), "keywords", "cp");
		appendXmlTag(sb, md.getDescription(), "description", "dc");

		Calendar cal = new GregorianCalendar();
		DateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'hh:mm:ss'Z'");
		df.setCalendar(cal);
		Date now = cal.getTime();
		String time = df.format(now);

		sb.append("<dcterms:created xsi:type=\"dcterms:W3CDTF\">").append(time).append("</dcterms:created>");
		sb.append("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">").append(time).append("</dcterms:modified>");

		appendXmlTag(sb, md.getCategory(), "category", "cp");
		appendXmlTag(sb, md.getContentStatus(), "contentStatus", "cp");

		return sb.toString();
	}

	/**
	 * Method to create the merged cells string of the passed worksheet
	 *
	 * @param sheet
	 *            Worksheet to process
	 * @return Formatted string with merged cell ranges
	 */
	private String createMergedCellsString(Worksheet sheet) {
		if (sheet.getMergedCells().size() < 1) {
			return "";
		}
		Iterator<Map.Entry<String, Range>> itr;
		Map.Entry<String, Range> range;
		StringBuilder sb = new StringBuilder();
		sb.append("<mergeCells count=\"").append(sheet.getMergedCells().size()).append("\">");
		itr = sheet.getMergedCells().entrySet().iterator();
		while (itr.hasNext()) {
			range = itr.next();
			sb.append("<mergeCell ref=\"").append(range.getValue().toString()).append("\"/>");
		}
		sb.append("</mergeCells>");
		return sb.toString();
	}

	/**
	 * Method to create the XML string for the color-MRU part of the style sheet
	 * document (recent colors)
	 *
	 * @return String with formatted XML data
	 */
	private String createMruColorsString() {
		StringBuilder sb = new StringBuilder();
		List<String> tempColors = new ArrayList<>();
		for (String item : this.workbook.getMruColors()) {
			if (item.equals(Fill.DEFAULT_COLOR)) {
				continue;
			}
			if (!tempColors.contains(item)) {
				tempColors.add(item);
			}
		}
		if (tempColors.size() > 0) {
			sb.append("<mruColors>");
			for (int i = 0; i < tempColors.size(); i++) {
				sb.append("<color rgb=\"").append(tempColors.get(i)).append("\"/>");
			}
			sb.append("</mruColors>");
			return sb.toString();
		}
		else {
			return "";
		}
	}

	/**
	 * Method to create a row string
	 *
	 * @param dynamicRow
	 *            Dynamic row with List of cells, heights and hidden states
	 * @param worksheet
	 *            Worksheet to process
	 * @return Formatted row string
	 */
	private String createRowString(DynamicRow dynamicRow, Worksheet worksheet) {
		int rowNumber = dynamicRow.getRowNumber();
		String height = "";
		String hidden = "";
		if (worksheet.getRowHeights().containsKey(rowNumber)) {
			if (worksheet.getRowHeights().get(rowNumber) != worksheet.getDefaultRowHeight()) {
				height = " x14ac:dyDescent=\"0.25\" customHeight=\"1\" ht=\"" + Helper.getInternalRowHeight(worksheet.getRowHeights().get(rowNumber)) + "\"";
			}
		}
		if (worksheet.getHiddenRows().containsKey(rowNumber)) {
			if (worksheet.getHiddenRows().get(rowNumber)) {
				hidden = " hidden=\"1\"";
			}
		}
		StringBuilder sb = new StringBuilder();
		sb.append("<row r=\"").append((rowNumber + 1)).append("\"").append(height).append(hidden).append(">");
		String typeAttribute;
		String styleDef = "";
		String typeDef = "";
		String value = "";
		String tValue = "";
		boolean boolValue;

		int col = 0;
		for (Cell item : dynamicRow.getCellDefinitions()) {
			typeDef = " ";
			if (item.getCellStyle() != null) {
				styleDef = " s=\"" + item.getCellStyle().getInternalID() + "\" ";
			}
			else {
				styleDef = "";
			}
			item.resolveCellType(); // Recalculate the type (for handling DEFAULT)
			if (item.getDataType().equals(Cell.CellType.BOOL)) {
				typeAttribute = "b";
				typeDef = " t=\"" + typeAttribute + "\" ";
				boolValue = (boolean) item.getValue();
				if (boolValue) {
					value = "1";
				}
				else {
					value = "0";
				}

			}
			// Number casting
			else if (item.getDataType() == Cell.CellType.NUMBER) {
				typeAttribute = "n";
				tValue = " t=\"" + typeAttribute + "\" ";
				Object o = item.getValue();
				if (o instanceof Byte) {
					value = Byte.toString((byte) item.getValue());
				}
				else if (o instanceof BigDecimal) {
					value = item.getValue().toString();
				}
				else if (o instanceof Double) {
					value = Double.toString((double) item.getValue());
				}
				else if (o instanceof Float) {
					value = Float.toString((float) item.getValue());
				}
				else if (o instanceof Integer) {
					value = Integer.toString((int) item.getValue());
				}
				else if (o instanceof Long) {
					value = Long.toString((long) item.getValue());
				}
				else if (o instanceof Short) {
					value = Short.toString((short) item.getValue());
				}
			}
			// Date parsing
			else if (item.getDataType().equals(Cell.CellType.DATE)) {
				typeAttribute = "d";
				Date date = (Date) item.getValue();
				value = Helper.getOADateString(date);
			}
			// Time parsing
			else if (item.getDataType().equals(Cell.CellType.TIME)) {
				typeAttribute = "d";
				// TODO: 'd' is probably an outdated attribute (to be checked for dates and
				// times)
				Duration time = (Duration) item.getValue();
				value = Helper.getOATimeString(time);
			}
			else {
				if (item.getValue() == null) {
					typeAttribute = null;
					value = null;
				}
				else // Handle sharedStrings
				{
					if (item.getDataType().equals(Cell.CellType.FORMULA)) {
						typeAttribute = "str";
						value = item.getValue().toString();
					}
					else {
						typeAttribute = "s";
						value = sharedStrings.add(item.getValue().toString(), Integer.toString(sharedStrings.size()));
						sharedStringsTotalCount++;
					}
				}
				typeDef = " t=\"" + typeAttribute + "\" ";
			}
			if (!item.getDataType().equals(Cell.CellType.EMPTY)) {
				sb.append("<c").append(typeDef).append("r=\"").append(item.getCellAddress()).append("\"").append(styleDef).append(">");
				if (item.getDataType().equals(Cell.CellType.FORMULA)) {
					sb.append("<f>").append(XlsxWriter.escapeXMLChars(item.getValue().toString())).append("</f>");
				}
				else {
					sb.append("<v>").append(XlsxWriter.escapeXMLChars(value)).append("</v>");
				}
				sb.append("</c>");
			}
			else if (value == null || item.getDataType().equals(Cell.CellType.EMPTY)) // Empty cell
			{
				sb.append("<c r=\"").append(item.getCellAddress()).append("\"").append(styleDef).append("/>");
			}
			col++;
		}
		sb.append("</row>");
		return sb.toString();
	}

	/**
	 * Method to create shared strings as XML document
	 *
	 * @return Formatted XML document
	 * @throws IOException
	 *             Thrown in case of an error while creating the XML document
	 */
	private Document createSharedStringsDocument() throws IOException {
		StringBuilder sb = new StringBuilder();
		sb.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"");
		sb.append(this.sharedStringsTotalCount);
		sb.append("\" uniqueCount=\"");
		sb.append(this.sharedStrings.size());
		sb.append("\">");
		List<String> keys = this.sharedStrings.getKeys();
		for (int i = 0; i < keys.size(); i++) {
			appendSharedString(sb, keys.get(i));
		}
		sb.append("</sst>");
		return createXMLDocument(sb.toString());
	}

	private void appendSharedString(StringBuilder sb, String value) {
		int len = value.length();
		value = escapeXMLChars(value);
		sb.append("<si>");
		if (len == 0) {
			sb.append("<t></t>");
		}
		else {
			if (Character.isWhitespace(value.charAt(0)) || Character.isWhitespace(value.charAt(len - 1))) {
				sb.append("<t xml:space=\"preserve\">");
			}
			else {
				sb.append("<t>");
			}
			sb.append(normalizeNewLines(value)).append("</t>");
		}
		sb.append("</si>");
	}

	/**
	 * Method to normalize all newlines to CR+LF
	 *
	 * @param value
	 *            Input value
	 * @return Normalized value
	 */
	private String normalizeNewLines(String value) {
		if (value == null || (value.indexOf('\n') == -1 && value.indexOf('\r') == -1)) {
			return value;
		}
		String normalized = value.replace("\r\n", "\n").replace("\r", "\n");
		return normalized.replace("\n", "\r\n");
	}

	/**
	 * Method to create the protection string of the passed worksheet
	 *
	 * @param sheet
	 *            Worksheet to process
	 * @return Formatted string with protection statement of the worksheet
	 */
	private String createSheetProtectionString(Worksheet sheet) {
		if (!sheet.isUseSheetProtection()) {
			return "";
		}
		HashMap<Worksheet.SheetProtectionValue, Integer> actualLockingValues = new HashMap<>();
		if (sheet.getSheetProtectionValues().isEmpty()) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.selectLockedCells, 1);
			actualLockingValues.put(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
		}
		if (!sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.objects)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.objects, 1);
		}
		if (!sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.scenarios)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.scenarios, 1);
		}
		if (!sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.selectLockedCells)) {
			if (!actualLockingValues.containsKey(Worksheet.SheetProtectionValue.selectLockedCells)) {
				actualLockingValues.put(Worksheet.SheetProtectionValue.selectLockedCells, 1);
			}
		}
		if (!sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.selectUnlockedCells) ||
				!sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.selectLockedCells)) {
			if (!actualLockingValues.containsKey(Worksheet.SheetProtectionValue.selectUnlockedCells)) {
				actualLockingValues.put(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
			}
		}
		if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.formatCells)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.formatCells, 0);
		}
		if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.formatColumns)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.formatColumns, 0);
		}
		if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.formatRows)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.formatRows, 0);
		}
		if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.insertColumns)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.insertColumns, 0);
		}
		if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.insertRows)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.insertRows, 0);
		}
		if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.insertHyperlinks)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.insertHyperlinks, 0);
		}
		if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.deleteColumns)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.deleteColumns, 0);
		}
		if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.deleteRows)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.deleteRows, 0);
		}
		if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.sort)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.sort, 0);
		}
		if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.autoFilter)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.autoFilter, 0);
		}
		if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.pivotTables)) {
			actualLockingValues.put(Worksheet.SheetProtectionValue.pivotTables, 0);
		}
		StringBuilder sb = new StringBuilder();
		sb.append("<sheetProtection");
		String temp;
		Iterator<Map.Entry<Worksheet.SheetProtectionValue, Integer>> itr;
		Map.Entry<Worksheet.SheetProtectionValue, Integer> item;
		itr = actualLockingValues.entrySet().iterator();
		while (itr.hasNext()) {
			item = itr.next();
			temp = item.getKey().name();// Note! If the enum names differs from the OOXML definitions, this method will
										// cause invalid OOXML entries
			sb.append(" ").append(temp).append("=\"").append(item.getValue()).append("\"");
		}
		if (!Helper.isNullOrEmpty(sheet.getSheetProtectionPasswordHash())) {
			sb.append(" password=\"").append(sheet.getSheetProtectionPasswordHash()).append("\"");
		}
		sb.append(" sheet=\"1\"/>");
		return sb.toString();
	}

	/**
	 * Method to create the XML string for the border part of the style sheet
	 * document
	 *
	 * @return String with formatted XML data
	 */
	private String createStyleBorderString() {
		Border[] borderStyles = this.styles.getBorders();
		StringBuilder sb = new StringBuilder();
		for (Border borderStyle : borderStyles) {
			if (borderStyle.isDiagonalDown() && !borderStyle.isDiagonalUp()) {
				sb.append("<border diagonalDown=\"1\">");
			}
			else if (!borderStyle.isDiagonalDown() && borderStyle.isDiagonalUp()) {
				sb.append("<border diagonalUp=\"1\">");
			}
			else if (borderStyle.isDiagonalDown() && borderStyle.isDiagonalUp()) {
				sb.append("<border diagonalDown=\"1\" diagonalUp=\"1\">");
			}
			else {
				sb.append("<border>");
			}
			if (borderStyle.getLeftStyle() != Border.StyleValue.none) {
				sb.append("<left style=\"").append(Border.getStyleName(borderStyle.getLeftStyle())).append("\">");
				if (!Helper.isNullOrEmpty(borderStyle.getLeftColor())) {
					sb.append("<color rgb=\"").append(borderStyle.getLeftColor()).append("\"/>");
				}
				else {
					sb.append("<color auto=\"1\"/>");
				}
				sb.append("</left>");
			}
			else {
				sb.append("<left/>");
			}
			if (borderStyle.getRightStyle() != Border.StyleValue.none) {
				sb.append("<right style=\"").append(Border.getStyleName(borderStyle.getRightStyle())).append("\">");
				if (!Helper.isNullOrEmpty(borderStyle.getRightColor())) {
					sb.append("<color rgb=\"").append(borderStyle.getRightColor()).append("\"/>");
				}
				else {
					sb.append("<color auto=\"1\"/>");
				}
				sb.append("</right>");
			}
			else {
				sb.append("<right/>");
			}
			if (borderStyle.getTopStyle() != Border.StyleValue.none) {
				sb.append("<top style=\"").append(Border.getStyleName(borderStyle.getTopStyle())).append("\">");
				if (!Helper.isNullOrEmpty(borderStyle.getTopColor())) {
					sb.append("<color rgb=\"").append(borderStyle.getTopColor()).append("\"/>");
				}
				else {
					sb.append("<color auto=\"1\"/>");
				}
				sb.append("</top>");
			}
			else {
				sb.append("<top/>");
			}
			if (borderStyle.getBottomStyle() != Border.StyleValue.none) {
				sb.append("<bottom style=\"").append(Border.getStyleName(borderStyle.getBottomStyle())).append("\">");
				if (!Helper.isNullOrEmpty(borderStyle.getBottomColor())) {
					sb.append("<color rgb=\"").append(borderStyle.getBottomColor()).append("\"/>");
				}
				else {
					sb.append("<color auto=\"1\"/>");
				}
				sb.append("</bottom>");
			}
			else {
				sb.append("<bottom/>");
			}
			if (borderStyle.getDiagonalStyle() != Border.StyleValue.none) {
				sb.append("<diagonal style=\"").append(Border.getStyleName(borderStyle.getDiagonalStyle())).append("\">");
				if (!Helper.isNullOrEmpty(borderStyle.getDiagonalColor())) {
					sb.append("<color rgb=\"").append(borderStyle.getDiagonalColor()).append("\"/>");
				}
				else {
					sb.append("<color auto=\"1\"/>");
				}
				sb.append("</diagonal>");
			}
			else {
				sb.append("<diagonal/>");
			}
			sb.append("</border>");
		}
		return sb.toString();
	}

	/**
	 * Method to create the XML string for the fill part of the style sheet document
	 *
	 * @return String with formatted XML data
	 */
	private String createStyleFillString() {
		Fill[] fillStyles = this.styles.getFills();
		StringBuilder sb = new StringBuilder();
		for (Fill fillStyle : fillStyles) {
			sb.append("<fill>");
			sb.append("<patternFill patternType=\"").append(Fill.getPatternName(fillStyle.getPatternFill())).append("\"");
			if (fillStyle.getPatternFill() == Fill.PatternValue.solid) {
				sb.append(">");
				sb.append("<fgColor rgb=\"").append(fillStyle.getForegroundColor()).append("\"/>");
				sb.append("<bgColor indexed=\"");
				sb.append(fillStyle.getIndexedColor());
				sb.append("\"/>");
				sb.append("</patternFill>");
			}
			else if (fillStyle.getPatternFill() == Fill.PatternValue.mediumGray ||
					fillStyle.getPatternFill() == Fill.PatternValue.lightGray ||
					fillStyle.getPatternFill() == Fill.PatternValue.gray0625 ||
					fillStyle.getPatternFill() == Fill.PatternValue.darkGray) {
				sb.append(">");
				sb.append("<fgColor rgb=\"").append(fillStyle.getForegroundColor()).append("\"/>");
				if (!Helper.isNullOrEmpty(fillStyle.getBackgroundColor())) {
					sb.append("<bgColor rgb=\"").append(fillStyle.getBackgroundColor()).append("\"/>");
				}
				sb.append("</patternFill>");
			}
			else {
				sb.append("/>");
			}
			sb.append("</fill>");
		}
		return sb.toString();
	}

	/**
	 * Method to create the XML string for the font part of the style sheet document
	 *
	 * @return String with formatted XML data
	 */
	private String createStyleFontString() {
		Font[] fontStyles = this.styles.getFonts();
		StringBuilder sb = new StringBuilder();
		for (Font fontStyle : fontStyles) {
			sb.append("<font>");
			if (fontStyle.isBold()) {
				sb.append("<b/>");
			}
			if (fontStyle.isItalic()) {
				sb.append("<i/>");
			}
			if (fontStyle.isStrike()) {
				sb.append("<strike/>");
			}
			if (fontStyle.getUnderline() == null || fontStyle.getUnderline() != Font.UnderlineValue.none) {
				if (fontStyle.getUnderline() == Font.UnderlineValue.u_double) {
					sb.append("<u val=\"double\"/>");
				}
				else if (fontStyle.getUnderline() == Font.UnderlineValue.singleAccounting) {
					sb.append("<u val=\"singleAccounting\"/>");
				}
				else if (fontStyle.getUnderline() == Font.UnderlineValue.doubleAccounting) {
					sb.append("<u val=\"doubleAccounting\"/>");
				}
				else {
					sb.append("<u/>");
				}
			}
			if (fontStyle.getVerticalAlign() == Font.VerticalAlignValue.subscript) {
				sb.append("<vertAlign val=\"subscript\"/>");
			}
			else if (fontStyle.getVerticalAlign() == Font.VerticalAlignValue.superscript) {
				sb.append("<vertAlign val=\"superscript\"/>");
			}
			sb.append("<sz val=\"");
			sb.append(fontStyle.getSize());
			sb.append("\"/>");
			if (Helper.isNullOrEmpty(fontStyle.getColorValue())) {
				sb.append("<color theme=\"");
				sb.append(fontStyle.getColorTheme());
				sb.append("\"/>");
			}
			else {
				sb.append("<color rgb=\"").append(fontStyle.getColorValue()).append("\"/>");
			}
			sb.append("<name val=\"").append(fontStyle.getName()).append("\"/>");
			sb.append("<family val=\"").append(fontStyle.getFamily()).append("\"/>");
			if (fontStyle.getScheme() != Font.SchemeValue.none) {
				if (fontStyle.getScheme() == Font.SchemeValue.major) {
					sb.append("<scheme val=\"major\"/>");
				}
				else if (fontStyle.getScheme() == Font.SchemeValue.minor) {
					sb.append("<scheme val=\"minor\"/>");
				}
			}
			if (!Helper.isNullOrEmpty(fontStyle.getCharset())) {
				sb.append("<charset val=\"").append(fontStyle.getCharset()).append("\"/>");
			}
			sb.append("</font>");
		}
		return sb.toString();
	}

	/**
	 * Method to create the XML string for the number format part of the style sheet
	 * document
	 *
	 * @return String with formatted XML data
	 */
	private String createStyleNumberFormatString() {
		NumberFormat[] numberFormatStyles = this.styles.getNumberFormats();
		StringBuilder sb = new StringBuilder();
		for (NumberFormat numberFormatStyle : numberFormatStyles) {
			if (numberFormatStyle.isCustomFormat()) {

				if (Helper.isNullOrEmpty(numberFormatStyle.getCustomFormatCode())) {
					throw new FormatException(
							"The number format style component with the ID " + numberFormatStyle.getCustomFormatID() + " cannot be null or empty");
				}
				sb.append("<numFmt formatCode=\"").append(escapeXMLAttributeChars(numberFormatStyle.getCustomFormatCode())).append("\" numFmtId=\"")
						.append(numberFormatStyle.getCustomFormatID()).append("\"/>");
			}
		}
		return sb.toString();
	}

	/**
	 * Method to create a style sheet as XML document
	 *
	 * @return Formatted XML document
	 * @throws StyleException
	 *             Thrown if a style was not referenced in the style sheet
	 * @throws RangeException
	 *             Thrown if a referenced cell was out of range
	 * @throws IOException
	 *             Thrown in case of an error while creating the XML document
	 */
	private Document createStyleSheetDocument() throws IOException {
		String bordersString = createStyleBorderString();
		String fillsString = createStyleFillString();
		String fontsString = createStyleFontString();
		String numberFormatsString = createStyleNumberFormatString();
		int numFormatCount = getNumberFormatStringCounter();
		String xfsStings = createStyleXfsString();
		String mruColorString = createMruColorsString();
		StringBuilder sb = new StringBuilder();

		sb.append(
				"<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
		if (numFormatCount > 0) {
			sb.append("<numFmts count=\"");
			sb.append(numFormatCount);
			sb.append("\">");
			sb.append(numberFormatsString).append("</numFmts>");
		}
		sb.append("<fonts x14ac:knownFonts=\"1\" count=\"");
		sb.append(this.styles.getFontStyleNumber());
		sb.append("\">");
		sb.append(fontsString).append("</fonts>");
		sb.append("<fills count=\"");
		sb.append(this.styles.getFillStyleNumber());
		sb.append("\">");
		sb.append(fillsString).append("</fills>");
		sb.append("<borders count=\"");
		sb.append(this.styles.getBorderStyleNumber());
		sb.append("\">");
		sb.append(bordersString).append("</borders>");
		sb.append("<cellXfs count=\"");
		sb.append(this.styles.getStyleNumber());
		sb.append("\">");
		sb.append(xfsStings).append("</cellXfs>");
		if (this.workbook.getWorkbookMetadata() != null) {
			if (!Helper.isNullOrEmpty(mruColorString)) {
				sb.append("<colors>");
				sb.append(mruColorString);
				sb.append("</colors>");
			}
		}
		sb.append("</styleSheet>");
		return createXMLDocument(sb.toString());
	}

	/**
	 * Method to create the XML string for the XF part of the style sheet document
	 *
	 * @return String with formatted XML data
	 * @throws RangeException
	 *             Thrown if a referenced cell was out of range
	 */
	private String createStyleXfsString() {
		Style[] styleItems = this.styles.getStyles();
		StringBuilder sb = new StringBuilder();
		StringBuilder sb2;
		String alignmentString, protectionString;
		int formatNumber, textRotation;
		for (Style style : styleItems) {
			textRotation = style.getCellXf().calculateInternalRotation();
			alignmentString = "";
			protectionString = "";
			if (style.getCellXf().getHorizontalAlign() != CellXf.HorizontalAlignValue.none ||
					style.getCellXf().getVerticalAlign() != CellXf.VerticalAlignValue.none ||
					style.getCellXf().getAlignment() != CellXf.TextBreakValue.none ||
					textRotation != 0) {
				sb2 = new StringBuilder();
				sb2.append("<alignment");
				if (style.getCellXf().getHorizontalAlign() != CellXf.HorizontalAlignValue.none) {
					sb2.append(" horizontal=\"");
					if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.center) {
						sb2.append("center");
					}
					else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.right) {
						sb2.append("right");
					}
					else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.centerContinuous) {
						sb2.append("centerContinuous");
					}
					else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.distributed) {
						sb2.append("distributed");
					}
					else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.fill) {
						sb2.append("fill");
					}
					else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.general) {
						sb2.append("general");
					}
					else if (style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.justify) {
						sb2.append("justify");
					}
					else {
						sb2.append("left");
					}
					sb2.append("\"");
				}
				if (style.getCellXf().getVerticalAlign() != CellXf.VerticalAlignValue.none) {
					sb2.append(" vertical=\"");
					if (style.getCellXf().getVerticalAlign() == CellXf.VerticalAlignValue.center) {
						sb2.append("center");
					}
					else if (style.getCellXf().getVerticalAlign() == CellXf.VerticalAlignValue.distributed) {
						sb2.append("distributed");
					}
					else if (style.getCellXf().getVerticalAlign() == CellXf.VerticalAlignValue.justify) {
						sb2.append("justify");
					}
					else if (style.getCellXf().getVerticalAlign() == CellXf.VerticalAlignValue.top) {
						sb2.append("top");
					}
					else {
						sb2.append("bottom");
					}
					sb2.append("\"");
				}
				if (style.getCellXf().getIndent() > 0 &&
						(style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.left ||
								style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.right ||
								style.getCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.distributed)) {
					sb2.append(" indent=\"");
					sb2.append(style.getCellXf().getIndent());
					sb2.append("\"");
				}

				if (style.getCellXf().getAlignment() != CellXf.TextBreakValue.none) {
					if (style.getCellXf().getAlignment() == CellXf.TextBreakValue.shrinkToFit) {
						sb2.append(" shrinkToFit=\"1");
					}
					else {
						sb2.append(" wrapText=\"1");
					}
					sb2.append("\"");
				}
				if (textRotation != 0) {
					sb2.append(" textRotation=\"");
					sb2.append(textRotation);
					sb2.append("\"");
				}
				sb2.append("/>"); // </xf>
				alignmentString = sb2.toString();
			}
			if (style.getCellXf().isHidden() || style.getCellXf().isLocked()) {
				if (style.getCellXf().isHidden() && style.getCellXf().isLocked()) {
					protectionString = "<protection locked=\"1\" hidden=\"1\"/>";
				}
				else if (style.getCellXf().isHidden() && !style.getCellXf().isLocked()) {
					protectionString = "<protection hidden=\"1\" locked=\"0\"/>";
				}
				else {
					protectionString = "<protection hidden=\"0\" locked=\"1\"/>";
				}
			}
			sb.append("<xf numFmtId=\"");
			if (style.getNumberFormat().isCustomFormat()) {
				sb.append(style.getNumberFormat().getCustomFormatID());
			}
			else {
				formatNumber = style.getNumberFormat().getNumber().getValue();
				sb.append(formatNumber);
			}
			sb.append("\" borderId=\"");
			sb.append(style.getBorder().getInternalID());
			sb.append("\" fillId=\"");
			sb.append(style.getFill().getInternalID());
			sb.append("\" fontId=\"");
			sb.append(style.getFont().getInternalID());
			if (!style.getFont().isDefaultFont()) {
				sb.append("\" applyFont=\"1");
			}
			if (style.getFill().getPatternFill() != Fill.PatternValue.none) {
				sb.append("\" applyFill=\"1");
			}
			if (!style.getBorder().isEmpty()) {
				sb.append("\" applyBorder=\"1");
			}
			if (!alignmentString.isEmpty() || style.getCellXf().isForceApplyAlignment()) {
				sb.append("\" applyAlignment=\"1");
			}
			if (!protectionString.isEmpty()) {
				sb.append("\" applyProtection=\"1");
			}
			if (style.getNumberFormat().getNumber() != NumberFormat.FormatNumber.none) {
				sb.append("\" applyNumberFormat=\"1\"");
			}
			else {
				sb.append("\"");
			}
			if (!alignmentString.isEmpty() || !protectionString.isEmpty()) {
				sb.append(">");
				sb.append(alignmentString);
				sb.append(protectionString);
				sb.append("</xf>");
			}
			else {
				sb.append("/>");
			}
		}
		return sb.toString();
	}

	/**
	 * Method to create a workbook as XML document
	 *
	 * @return Formatted XML document
	 * @throws RangeException
	 *             Thrown if a referenced cell was out of range
	 * @throws IOException
	 *             Thrown in case of an error while creating the XML document
	 */
	private Document createWorkbookDocument() throws IOException {
		StringBuilder sb = new StringBuilder();
		sb.append(
				"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
		if (workbook.getSelectedWorksheet() > 0 || workbook.isHidden()) {
			sb.append("<bookViews><workbookView ");
			if (workbook.isHidden()) {
				sb.append("visibility=\"hidden\"");
			}
			else {
				sb.append("activeTab=\"").append(workbook.getSelectedWorksheet()).append("\"");
			}
			sb.append("/></bookViews>");
		}
		createWorkbookProtectionString(sb);
		sb.append("<sheets>");
		if (!workbook.getWorksheets().isEmpty()) {
			for (Worksheet item : workbook.getWorksheets()) {
				sb.append("<sheet r:id=\"rId").append(item.getSheetID()).append("\" sheetId=\"").append(item.getSheetID()).append("\" name=\"")
						.append(escapeXMLAttributeChars(item.getSheetName())).append("\"");
				if (item.isHidden()) {
					sb.append(" state=\"hidden\"");
				}
				sb.append("/>");
			}
		}
		else {
			// Fallback on empty workbook
			sb.append("<sheet r:id=\"rId1\" sheetId=\"1\" name=\"sheet1\"/>");
		}
		sb.append("</sheets>");
		sb.append("</workbook>");
		return createXMLDocument(sb.toString());
	}

	/**
	 * Method to create the (sub) part of the workbook protection within the
	 * workbook XML document
	 *
	 * @param sb
	 *            reference to the stringbuilder
	 */
	private void createWorkbookProtectionString(StringBuilder sb) {
		if (workbook.isWorkbookProtectionUsed()) {
			sb.append("<workbookProtection");
			if (workbook.isWindowsLockedIfProtected()) {
				sb.append(" lockWindows=\"1\"");
			}
			if (workbook.isStructureLockedIfProtected()) {
				sb.append(" lockStructure=\"1\"");
			}
			if (!Helper.isNullOrEmpty(workbook.getWorkbookProtectionPassword())) {
				sb.append(" workbookPassword=\"");
				sb.append(workbook.getWorkbookProtectionPasswordHash());
				sb.append("\"");
			}
			sb.append("/>");
		}
	}

	/**
	 * Method to create a worksheet part as XML document
	 *
	 * @param worksheet
	 *            worksheet object to process
	 * @return Formatted XML document
	 * @throws IOException
	 *             Thrown in case of an error while creating the XML document
	 */
	private Document createWorksheetPart(Worksheet worksheet) throws IOException {
		worksheet.recalculateAutoFilter();
		worksheet.recalculateColumns();
		StringBuilder sb = new StringBuilder();
		sb.append(
				"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");

		if (worksheet.getSelectedCells() != null ||
				worksheet.getPaneSplitTopHeight() != null ||
				worksheet.getPaneSplitLeftWidth() != null ||
				worksheet.getPaneSplitAddress() != null ||
				worksheet.isHidden()) {
			createSheetViewString(worksheet, sb);
		}

		sb.append("<sheetFormatPr");
		if (!hasPaneSplitting(worksheet)) {
			// TODO: Find the right calculation to compensate baseColWidth when using pane
			// splitting
			sb.append(" defaultColWidth=\"")
					.append(worksheet.getDefaultColumnWidth())
					.append("\"");
		}
		sb.append(" defaultRowHeight=\"")
				.append(worksheet.getDefaultRowHeight())
				.append("\" baseColWidth=\"")
				.append(worksheet.getDefaultColumnWidth())
				.append("\" x14ac:dyDescent=\"0.25\"/>");

		String colDefinitions = createColsString(worksheet);
		if (!Helper.isNullOrEmpty(colDefinitions)) {
			sb.append("<cols>");
			sb.append(colDefinitions);
			sb.append("</cols>");
		}
		sb.append("<sheetData>");
		createRowsString(worksheet, sb);
		sb.append("</sheetData>");

		sb.append(createMergedCellsString(worksheet));
		sb.append(createSheetProtectionString(worksheet));
		if (worksheet.getAutoFilterRange() != null) {
			sb.append("<autoFilter ref=\"").append(worksheet.getAutoFilterRange().toString()).append("\"/>");
		}
		sb.append("</worksheet>");
		return createXMLDocument(sb.toString());
	}

	/**
	 * Checks whether pane splitting is applied in the given worksheet
	 *
	 * @param worksheet
	 *            Checks whether pane splitting is applied in the given worksheet
	 * @return True if applied, otherwise false
	 */
	private boolean hasPaneSplitting(Worksheet worksheet) {
        return worksheet.getPaneSplitLeftWidth() != null || worksheet.getPaneSplitTopHeight() != null || worksheet.getPaneSplitAddress() != null;
    }

	private void createRowsString(Worksheet worksheet, StringBuilder sb) {
		List<DynamicRow> cellData = getSortedSheetData(worksheet);
		String line;
		for (DynamicRow row : cellData) {
			line = createRowString(row, worksheet);
			sb.append(line);
		}
	}

	/**
	 * Method to create the (sub) part of the sheet view (selected cells and panes)
	 * within the worksheet XML document
	 *
	 * @param worksheet
	 *            worksheet object to process
	 * @param sb
	 *            reference to the stringbuilder
	 */
	private void createSheetViewString(Worksheet worksheet, StringBuilder sb) {
		sb.append("<sheetViews><sheetView workbookViewId=\"0\"");
		if (workbook.getSelectedWorksheet() == worksheet.getSheetID() - 1 && !worksheet.isHidden()) {
			sb.append(" tabSelected=\"1\"");
		}
		sb.append(">");
		createPaneString(worksheet, sb);
		if (!worksheet.getSelectedCellRanges().isEmpty()) {
			sb.append("<selection sqref=\"");
			for (int i = 0; i < worksheet.getSelectedCellRanges().size(); i++) {
				sb.append(worksheet.getSelectedCellRanges().get(i).toString());
				if (i < worksheet.getSelectedCellRanges().size() - 1) {
					sb.append(" ");
				}
			}
			sb.append("\" activeCell=\"");
			sb.append(worksheet.getSelectedCellRanges().get(0).StartAddress.toString());
			sb.append("\"/>");
		}
		sb.append("</sheetView></sheetViews>");
	}

	/**
	 * Method to create the (sub) part of the pane (splitting and freezing) within
	 * the worksheet XML document
	 *
	 * @param worksheet
	 *            worksheet">worksheet object to process
	 * @param sb
	 *            reference to the stringbuilder
	 */
	private void createPaneString(Worksheet worksheet, StringBuilder sb) {
		if (worksheet.getPaneSplitLeftWidth() == null && worksheet.getPaneSplitTopHeight() == null && worksheet.getPaneSplitAddress() == null) {
			return;
		}
		sb.append("<pane");
		boolean applyXSplit = false;
		boolean applyYSplit = false;
		if (worksheet.getPaneSplitAddress() != null) {
			boolean freeze = worksheet.getFreezeSplitPanes() != null && worksheet.getFreezeSplitPanes();
			int xSplit = worksheet.getPaneSplitAddress().Column;
			int ySplit = worksheet.getPaneSplitAddress().Row;
			if (xSplit > 0) {
				if (freeze) {
					sb.append(" xSplit=\"").append(xSplit).append("\"");
				}
				else {
					sb.append(" xSplit=\"").append(calculatePaneWidth(worksheet, xSplit)).append("\"");
				}
				applyXSplit = true;
			}
			if (ySplit > 0) {
				if (freeze) {
					sb.append(" ySplit=\"").append(ySplit).append("\"");
				}
				else {
					sb.append(" ySplit=\"").append(calculatePaneHeight(worksheet, ySplit)).append("\"");
				}
				applyYSplit = true;
			}
			if (freeze && applyXSplit && applyYSplit) {
				sb.append(" state=\"frozenSplit\"");
			}
			else if (freeze) {
				sb.append(" state=\"frozen\"");
			}
		}
		else {
			if (worksheet.getPaneSplitLeftWidth() != null) {
				sb.append(" xSplit=\"").append(Helper.getInternalPaneSplitWidth(worksheet.getPaneSplitLeftWidth())).append("\"");
				applyXSplit = true;
			}
			if (worksheet.getPaneSplitTopHeight() != null) {
				sb.append(" ySplit=\"").append(Helper.getInternalPaneSplitHeight(worksheet.getPaneSplitTopHeight())).append("\"");
				applyYSplit = true;
			}
		}
		if ((applyXSplit || applyYSplit) && worksheet.getActivePane() != null) {
			switch (worksheet.getActivePane()) {
				case bottomLeft:
					sb.append(" activePane=\"bottomLeft\"");
					break;
				case bottomRight:
					sb.append(" activePane=\"bottomRight\"");
					break;
				case topLeft:
					sb.append(" activePane=\"topLeft\"");
					break;
				case topRight:
					sb.append(" activePane=\"topRight\"");
					break;
			}
		}
		String topLeftCell = worksheet.getPaneSplitTopLeftCell().getAddress();
		sb.append(" topLeftCell=\"").append(topLeftCell).append("\" ");
		sb.append("/>");
		if (applyXSplit && !applyYSplit) {
			sb.append("<selection pane=\"topRight\" activeCell=\"" + topLeftCell + "\"  sqref=\"" + topLeftCell + "\" />");
		}
		else if (applyYSplit && !applyXSplit) {
			sb.append("<selection pane=\"bottomLeft\" activeCell=\"" + topLeftCell + "\"  sqref=\"" + topLeftCell + "\" />");
		}
		else if (applyYSplit && applyXSplit) {
			sb.append("<selection activeCell=\"" + topLeftCell + "\"  sqref=\"" + topLeftCell + "\" />");
		}
	}

	/**
	 * Method to calculate the pane height, based on the number of rows
	 *
	 * @param worksheet
	 *            worksheet object to get the row definitions from
	 * @param numberOfRows
	 *            Number of rows from the top to the split position
	 * @return Internal height from the top of the worksheet to the pane split
	 *         position
	 */
	private float calculatePaneHeight(Worksheet worksheet, int numberOfRows) {
		float height = 0;
		for (int i = 0; i < numberOfRows; i++) {
			if (worksheet.getRowHeights().containsKey(i)) {
				height += Helper.getInternalRowHeight(worksheet.getRowHeights().get(i));
			}
			else {
				height += Helper.getInternalRowHeight(Worksheet.DEFAULT_ROW_HEIGHT);
			}
		}
		return Helper.getInternalPaneSplitHeight(height);
	}

	/**
	 * Method to calculate the pane width, based on the number of columns
	 *
	 * @param worksheet
	 *            worksheet object to get the column definitions from
	 * @param numberOfColumns
	 *            Number of columns from the left to the split position
	 * @return Internal width from the left of the worksheet to the pane split
	 *         position
	 */
	private float calculatePaneWidth(Worksheet worksheet, int numberOfColumns) {
		float width = 0;
		for (int i = 0; i < numberOfColumns; i++) {
			if (worksheet.getColumns().containsKey(i)) {
				width += Helper.getInternalColumnWidth(worksheet.getColumns().get(i).getWidth());
			}
			else {
				width += Helper.getInternalColumnWidth(Worksheet.DEFAULT_COLUMN_WIDTH);
			}
		}
		// Add padding of 75 per column
		return Helper.getInternalPaneSplitWidth(width) + ((numberOfColumns - 1) * 0f);
	}

	/**
	 * Creates an XML document from a string
	 *
	 * @param rawInput
	 *            String to process
	 * @return Formatted XML document
	 * @throws IOException
	 *             Thrown in case of an error while creating the XML document
	 */
	public Document createXMLDocument(String rawInput) throws IOException {
		try {
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = factory.newDocumentBuilder();
			InputSource input = new InputSource(new StringReader(rawInput));
			input.setEncoding("UTF-8");
			Document doc = docBuilder.parse(input);
			doc.setXmlVersion("1.0");
			doc.setXmlStandalone(true);
			return doc;
		}
		catch (Exception e) {
			throw new IOException("There was an error while creating the XML document. Please see the inner exception.", e);
		}
	}

	/**
	 * Gets the number of custom number formats
	 *
	 * @return Number of custom number formats to apply in the style document
	 */
	private int getNumberFormatStringCounter() {
		NumberFormat[] numberFormatStyles = this.styles.getNumberFormats();
		int counter = 0;
		for (NumberFormat numberFormatStyle : numberFormatStyles) {
			if (numberFormatStyle.isCustomFormat()) {
				counter++;
			}
		}
		return counter;
	}

	/**
	 * Method to sort the cells of a worksheet as preparation for the XML document
	 *
	 * @param sheet
	 *            Worksheet to process
	 * @return Sorted list of dynamic rows that are either defined by cells or row
	 *         widths / hidden states. The list is sorted by row numbers
	 *         (zero-based)
	 */
	private List<DynamicRow> getSortedSheetData(Worksheet sheet) {
		List<Cell> temp = new ArrayList<>();
		for (Map.Entry<String, Cell> item : sheet.getCells().entrySet()) {
			temp.add(item.getValue());
		}
		Collections.sort(temp);
		DynamicRow row = new DynamicRow();
		Map<Integer, DynamicRow> rows = new HashMap<>();
		int rowNumber;
		if (!temp.isEmpty()) {
			rowNumber = temp.get(0).getRowNumber();
			row.setRowNumber(rowNumber);
			for (Cell cell : temp) {
				if (cell.getRowNumber() != rowNumber) {
					rows.put(rowNumber, row);
					row = new DynamicRow();
					row.setRowNumber(cell.getRowNumber());
					rowNumber = cell.getRowNumber();
				}
				row.getCellDefinitions().add(cell);
			}
			if (!row.getCellDefinitions().isEmpty()) {
				rows.put(rowNumber, row);
			}
		}
		for (Map.Entry<Integer, Float> rowHeight : sheet.getRowHeights().entrySet()) {
			if (!rows.containsKey(rowHeight.getKey())) {
				row = new DynamicRow();
				row.setRowNumber(rowHeight.getKey());
				rows.put(rowHeight.getKey(), row);
			}
		}
		for (Map.Entry<Integer, Boolean> hiddenRow : sheet.getHiddenRows().entrySet()) {
			if (!rows.containsKey(hiddenRow.getKey())) {
				row = new DynamicRow();
				row.setRowNumber(hiddenRow.getKey());
				rows.put(hiddenRow.getKey(), row);
			}
		}
		List<DynamicRow> output = new ArrayList<>(rows.values());
		output.sort((r1, r2) -> (Integer.compare(r1.getRowNumber(), r2.getRowNumber()))); // Lambda sort
		return output;
	}

	/**
	 * Method to save the workbook
	 *
	 * @throws IOException
	 *             Thrown in case of an error
	 */
	public void save() throws IOException {
		try {
			FileOutputStream dest = new FileOutputStream(this.workbook.getFilename());
			saveAsStream(dest);
		}
		catch (Exception e) {
			throw new IOException(
					"There was an error while creating the workbook document during saving to a file. Please see the inner exception:" + e.getMessage(), e);
		}
	}

	public void saveAsStream(OutputStream stream) throws IOException {
		try {
			this.workbook.resolveMergedCells();
			this.styles = StyleManager.getManagedStyles(workbook); // After this point, styles must not be changed anymore
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
			if (!workbook.getWorksheets().isEmpty()) {
				for (int i = 0; i < this.workbook.getWorksheets().size(); i++) {
					sheet = this.workbook.getWorksheets().get(i);
					doc = createWorksheetPart(sheet);
					file = "sheet" + sheet.getSheetID() + ".xml";
					rel.addRelationshipEntry("/xl/worksheets/" + file, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
					p.addPart("xl/worksheets/" + file, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", doc);
				}
			}
			else {
				// Fallback on empty workbook
				doc = createWorksheetPart(new Worksheet());
				file = "sheet1.xml";
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
		catch (Exception e) {
			throw new IOException(
					"There was an error while creating the workbook document during writing to a stream. Please see the inner exception:" + e.getMessage(), e);
		}
	}

	// ### S T A T I C M E T H O D S ###

	/**
	 * Method to convert an XML document to a byte array
	 *
	 * @param document
	 *            Document to process
	 * @return array of bytes (UTF-8)
	 * @throws IOException
	 *             Thrown if the document could not be converted to a byte array
	 */
	public static byte[] createBytesFromDocument(Document document) throws IOException {
		try {
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
		catch (Exception e) {
			throw new IOException("There was an error while creating the byte array. Please see the inner exception.", e);
		}
	}

	/**
	 * Method to escape XML characters in an XML attribute
	 *
	 * @param input
	 *            Input string to process
	 * @return Escaped string
	 */
	private static String escapeXMLAttributeChars(String input) {
		input = escapeXMLChars(input); // Sanitize string from illegal characters beside quotes
		input = input.replace("\"", "&quot;");
		return input;
	}

	/**
	 * Method to escape XML characters between two XML tags<br>
	 * Note: The XML specs allow characters up to the character value of 0x10FFFF.
	 * However, the Java char range is only up to 0xFFFF. PicoXLSX4j will neglect
	 * all values above this level in the sanitizing check. Illegal characters like
	 * 0x1 will be replaced with a white space (0x20)
	 *
	 * @param input
	 *            Input string to process
	 * @return Escaped string
	 */
	private static String escapeXMLChars(String input) {
		int len = input.length();
		List<Integer> illegalCharacters = new ArrayList<>(len);
		List<Integer> characterTypes = new ArrayList<>(len);
		int i;
		char c;
		for (i = 0; i < len; i++) {
			c = input.charAt(i);
			if ((c < 0x9) || (c > 0xA && c < 0xD) || (c > 0xD && c < 0x20) || (c > 0xD7FF && c < 0xE000) || (c > 0xFFFD)) {
				illegalCharacters.add(i);
				characterTypes.add(0);
				continue;
			} // Note: XML specs allow characters up to 0x10FFFF. However, the Java char range
				// is only up to 0xFFFF; Higher values are neglected here
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
		if (illegalCharacters.isEmpty()) {
			return input;
		}
		StringBuilder sb = new StringBuilder(len);
		int lastIndex = 0;
		len = illegalCharacters.size();
		int j, type;
		for (i = 0; i < len; i++) {
			j = illegalCharacters.get(i);
			type = characterTypes.get(i);
			sb.append(input, lastIndex, j);
			if (type == 0) {
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

	// ### H E L P E R C L A S S E S ###

	/**
	 * Class representing a row that is either empty or containing cells. Empty rows
	 * can also carry information about height or visibility
	 */
	private static class DynamicRow {

		private final List<Cell> cellDefinitions;
		private int rowNumber;

		/**
		 * Gets the List of cells if not empty
		 *
		 * @return List of cells
		 */
		public List<Cell> getCellDefinitions() {
			return cellDefinitions;
		}

		/**
		 * Gets the row number (zero-based)
		 *
		 * @return Row number
		 */
		public int getRowNumber() {
			return rowNumber;
		}

		/**
		 * Sets the row number (zero-based)
		 *
		 * @param rowNumber
		 *            Row number
		 */
		public void setRowNumber(int rowNumber) {
			this.rowNumber = rowNumber;
		}

		/**
		 * Default constructor. Defines an empty row if no additional operations are
		 * made on the object
		 */
		public DynamicRow() {
			this.cellDefinitions = new ArrayList<>();
		}

	}

}
