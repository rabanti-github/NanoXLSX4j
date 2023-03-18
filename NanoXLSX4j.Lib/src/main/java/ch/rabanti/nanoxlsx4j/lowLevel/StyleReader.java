/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import java.io.InputStream;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import ch.rabanti.nanoxlsx4j.styles.Border;
import ch.rabanti.nanoxlsx4j.styles.CellXf;
import ch.rabanti.nanoxlsx4j.styles.Fill;
import ch.rabanti.nanoxlsx4j.styles.Font;
import ch.rabanti.nanoxlsx4j.styles.NumberFormat;
import ch.rabanti.nanoxlsx4j.styles.Style;

/**
 * Class representing a reader for style definitions of XLSX files
 *
 * @author Raphael Stoeckli
 */
public class StyleReader {

	private final StyleReaderContainer styleReaderContainer;

	/**
	 * Gets the container for raw style components of the reader.
	 *
	 * @return Style reader container with resolved style components
	 */
	public StyleReaderContainer getStyleReaderContainer() {
		return styleReaderContainer;
	}

	/**
	 * Default constructor
	 */
	public StyleReader() {
		this.styleReaderContainer = new StyleReaderContainer();
	}

	/**
	 * Reads the XML file form the passed stream and processes the style information
	 *
	 * @param stream
	 *            Stream of the XML file
	 * @throws IOException
	 *             Throws IOException in case of an error
	 */
	public void read(InputStream stream) throws IOException, java.io.IOException {
		XmlDocument xr = new XmlDocument();
		xr.load(stream);
		for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) {
			if (node.getName().equalsIgnoreCase("numfmts")) { // Handles custom number formats
				this.getNumberFormats(node);
			}
			else if (node.getName().equalsIgnoreCase("borders")) { // handle borders
				this.getBorders(node);
			}
			else if (node.getName().equalsIgnoreCase("fills")) { // handle fills
				this.getFills(node);
			}
			else if (node.getName().equalsIgnoreCase("fonts")) { // handle fonts
				this.getFonts(node);
			}
			else if (node.getName().equalsIgnoreCase("colors")) { // handle MRU colors
				this.getColors(node);
			}
			// TODO: Implement other style components
		}
		for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) { // Redo for composition after all style parts are gathered; standard number
																					// formats
			if (node.getName().equalsIgnoreCase("cellxfs")) {
				this.getCellXfs(node);
			}
		}
	}

	/**
	 * Determines the number formats in an XML node of the style document
	 *
	 * @param node
	 *            Number formats root node
	 */
	private void getNumberFormats(XmlDocument.XmlNode node) {
		for (XmlDocument.XmlNode childNode : node.getChildNodes()) {
			if (childNode.getName().equalsIgnoreCase("numfmt")) {
				NumberFormat numberFormat = new NumberFormat();
				int id = Integer.parseInt(childNode.getAttribute("numFmtId")); // Default/null will (justified) throw an exception
				String code = childNode.getAttribute("formatCode", ""); // code is not un-escaped
				numberFormat.setCustomFormatID(id);
				numberFormat.setNumber(NumberFormat.FormatNumber.custom);
				numberFormat.setInternalID(id);
				numberFormat.setCustomFormatCode(code);
				this.styleReaderContainer.addStyleComponent(numberFormat);
			}
		}
	}

	/**
	 * Determines the borders in an XML node of the style document
	 *
	 * @param node
	 *            Border root node
	 */
	private void getBorders(XmlDocument.XmlNode node) {
		for (XmlDocument.XmlNode border : node.getChildNodes()) {
			Border borderStyle = new Border();
			String diagonalDown = border.getAttribute("diagonalDown");
			String diagonalUp = border.getAttribute("diagonalUp");
			if (diagonalDown != null) {
				int value = ReaderUtils.parseBinaryBoolean(diagonalDown);
				borderStyle.setDiagonalDown(value == 1);
			}
			if (diagonalUp != null) {
				int value = ReaderUtils.parseBinaryBoolean(diagonalUp);
				borderStyle.setDiagonalUp(value == 1);
			}
			XmlDocument.XmlNode innerNode = getChildNode(border, "diagonal");
			if (innerNode != null) {
				borderStyle.setDiagonalStyle(parseBorderStyle(innerNode));
				borderStyle.setDiagonalColor(getColor(innerNode, Border.DEFAULT_BORDER_COLOR));
			}
			innerNode = getChildNode(border, "top");
			if (innerNode != null) {
				borderStyle.setTopStyle(parseBorderStyle(innerNode));
				borderStyle.setTopColor(getColor(innerNode, Border.DEFAULT_BORDER_COLOR));
			}
			innerNode = getChildNode(border, "bottom");
			if (innerNode != null) {
				borderStyle.setBottomStyle(parseBorderStyle(innerNode));
				borderStyle.setBottomColor(getColor(innerNode, Border.DEFAULT_BORDER_COLOR));
			}
			innerNode = getChildNode(border, "left");
			if (innerNode != null) {
				borderStyle.setLeftStyle(parseBorderStyle(innerNode));
				borderStyle.setLeftColor(getColor(innerNode, Border.DEFAULT_BORDER_COLOR));
			}
			innerNode = getChildNode(border, "right");
			if (innerNode != null) {
				borderStyle.setRightStyle(parseBorderStyle(innerNode));
				borderStyle.setRightColor(getColor(innerNode, Border.DEFAULT_BORDER_COLOR));
			}
			borderStyle.setInternalID(styleReaderContainer.getNextBorderId());
			styleReaderContainer.addStyleComponent(borderStyle);
		}
	}

	/**
	 * Tries to parse a border style
	 *
	 * @param innerNode
	 *            Border sub-node
	 * @return Border type or non if parsing was not successful
	 */
	private static Border.StyleValue parseBorderStyle(XmlDocument.XmlNode innerNode) {
		String value = innerNode.getAttribute("style");
		if (value != null) {
			if (value.equalsIgnoreCase("double")) {
				return Border.StyleValue.s_double; // special handling, since double is not a valid enum value
			}
			return getBorderEnumValue(value);
		}
		return Border.StyleValue.none;
	}

	private void getFills(XmlDocument.XmlNode node) {
		for (XmlDocument.XmlNode fill : node.getChildNodes()) {
			Fill fillStyle = new Fill();
			XmlDocument.XmlNode innerNode = getChildNode(fill, "patternFill");
			if (innerNode != null) {
				String pattern = innerNode.getAttribute("patternType");
				fillStyle.setPatternFill(getFillPatternEnumValue(pattern));
				LookupResult attribute = getAttributeOfChild(innerNode, "fgColor", "rgb");
				if (attribute.attributeIsPresent) {
					fillStyle.setForegroundColor(attribute.value);
				}
				XmlDocument.XmlNode backgroundNode = getChildNode(innerNode, "bgColor");
				if (backgroundNode != null) {
					String backgroundArgb = backgroundNode.getAttribute("rgb");
					if (backgroundArgb != null) {
						fillStyle.setBackgroundColor(backgroundArgb);
					}
					String backgroundIndex = backgroundNode.getAttribute("indexed");
					if (backgroundIndex != null) {
						fillStyle.setIndexedColor(Integer.parseInt(backgroundIndex));
					}
				}
			}
			fillStyle.setInternalID(styleReaderContainer.getNextFillId());
			styleReaderContainer.addStyleComponent(fillStyle);
		}
	}

	private void getFonts(XmlDocument.XmlNode node) {
		LookupResult attribute;
		for (XmlDocument.XmlNode font : node.getChildNodes()) {
			Font fontStyle = new Font();
			XmlDocument.XmlNode boldNode = getChildNode(font, "b");
			if (boldNode != null) {
				fontStyle.setBold(true);
			}
			XmlDocument.XmlNode italicNode = getChildNode(font, "i");
			if (italicNode != null) {
				fontStyle.setItalic(true);
			}
			XmlDocument.XmlNode strikeNode = getChildNode(font, "strike");
			if (strikeNode != null) {
				fontStyle.setStrike(true);
			}
			attribute = getAttributeOfChild(font, "u", "val");
			if (attribute.nodeIsPresent) {
				fontStyle.setUnderline(Font.UnderlineValue.u_single); // default
			}
			if (attribute.attributeIsPresent) {
				switch (attribute.value) {
					case "double":
						fontStyle.setUnderline(Font.UnderlineValue.u_double);
						break;
					case "singleAccounting":
						fontStyle.setUnderline(Font.UnderlineValue.singleAccounting);
						break;
					case "doubleAccounting":
						fontStyle.setUnderline(Font.UnderlineValue.doubleAccounting);
						break;
				}
			}

			attribute = getAttributeOfChild(font, "vertAlign", "val");
			if (attribute.attributeIsPresent) {
				fontStyle.setVerticalAlign(getFontVerticalAlignEnumValue(attribute.value));
			}
			attribute = getAttributeOfChild(font, "sz", "val");
			if (attribute.attributeIsPresent) {
				fontStyle.setSize(Float.parseFloat(attribute.value));
			}
			XmlDocument.XmlNode colorNode = getChildNode(font, "color");
			if (colorNode != null) {
				String theme = colorNode.getAttribute("theme");
				if (theme != null) {
					fontStyle.setColorTheme(Integer.parseInt(theme));
				}
				String rgb = colorNode.getAttribute("rgb");
				if (rgb != null) {
					fontStyle.setColorValue(rgb);
				}
			}
			attribute = getAttributeOfChild(font, "name", "val");
			if (attribute.attributeIsPresent) {
				fontStyle.setName(attribute.value);
			}
			attribute = getAttributeOfChild(font, "family", "val");
			if (attribute.attributeIsPresent) {
				fontStyle.setFamily(attribute.value);
			}
			attribute = getAttributeOfChild(font, "scheme", "val");
			if (attribute.attributeIsPresent) {
				switch (attribute.value) {
					case "major":
						fontStyle.setScheme(Font.SchemeValue.major);
						break;
					case "minor":
						fontStyle.setScheme(Font.SchemeValue.minor);
						break;
				}
			}
			attribute = getAttributeOfChild(font, "charset", "val");
			if (attribute.attributeIsPresent) {
				fontStyle.setCharset(attribute.value);
			}

			fontStyle.setInternalID(styleReaderContainer.getNextFontId());
			styleReaderContainer.addStyleComponent(fontStyle);
		}
	}

	/**
	 * Determines the cell XF entries in an XML node of the style document
	 *
	 * @param node
	 *            Cell XF root node
	 */
	private void getCellXfs(XmlDocument.XmlNode node) {

		for (XmlDocument.XmlNode childNode : node.getChildNodes()) {
			if (childNode.getName().equalsIgnoreCase("xf")) {
				CellXf cellXfStyle = new CellXf();
				String attribute = childNode.getAttribute("applyAlignment");
				if (attribute != null) {
					int value = ReaderUtils.parseBinaryBoolean(attribute);
					cellXfStyle.setForceApplyAlignment(value == 1);
				}
				XmlDocument.XmlNode alignmentNode = getChildNode(childNode, "alignment");
				if (alignmentNode != null) {
					attribute = alignmentNode.getAttribute("shrinkToFit");
					if (attribute != null) {
						int value = ReaderUtils.parseBinaryBoolean(attribute);
						if (value == 1) {
							cellXfStyle.setAlignment(CellXf.TextBreakValue.shrinkToFit);
						}
					}
					attribute = alignmentNode.getAttribute("wrapText");
					if (attribute != null) {
						int value = ReaderUtils.parseBinaryBoolean(attribute);
						if (value == 1) {
							cellXfStyle.setAlignment(CellXf.TextBreakValue.wrapText);
						}
					}
					attribute = alignmentNode.getAttribute("horizontal");
					if (attribute != null) {
						cellXfStyle.setHorizontalAlign(getCellXfHorizontalAlignEnumValue(attribute));
					}
					attribute = alignmentNode.getAttribute("vertical");
					if (attribute != null) {
						cellXfStyle.setVerticalAlign(getCellXfVerticalAlignEnumValue(attribute));
					}
					attribute = alignmentNode.getAttribute("indent");
					if (attribute != null) {
						cellXfStyle.setIndent(Integer.parseInt(attribute));
					}
					attribute = alignmentNode.getAttribute("textRotation");
					if (attribute != null) {
						int rotation = Integer.parseInt(attribute);
						cellXfStyle.setTextRotation(rotation > 90 ? 90 - rotation : rotation);
					}
				}
				XmlDocument.XmlNode protectionNode = getChildNode(childNode, "protection");
				if (protectionNode != null) {
					attribute = protectionNode.getAttribute("hidden");
					if (attribute != null) {
						int value = ReaderUtils.parseBinaryBoolean(attribute);
						if (value == 1) {
							cellXfStyle.setHidden(true);
						}
					}
					attribute = protectionNode.getAttribute("locked");
					if (attribute != null) {
						int value = ReaderUtils.parseBinaryBoolean(attribute);
						if (value == 1) {
							cellXfStyle.setLocked(true);
						}
					}
				}

				cellXfStyle.setInternalID(this.styleReaderContainer.getNextCellXFId());
				this.styleReaderContainer.addStyleComponent(cellXfStyle);

				Style style = new Style();
				ReaderUtils.IntParser id;

				id = ReaderUtils.IntParser.tryParse(childNode.getAttribute("numFmtId"));
				NumberFormat format = this.styleReaderContainer.getNumberFormat(id.value);
				if (!id.hasValue || format == null) {
					// Validity is neglected here to prevent unhandled crashes. If invalid, the
					// format will be declared as 'none' Invalid values should not occur at all
					// (malformed Excel files); Undefined values may occur if the file was saved by
					// an Excel version that has implemented yet unknown format numbers (undefined
					// in NanoXLSX)
					NumberFormat.NumberFormatEvaluation evaluation = NumberFormat.tryParseFormatNumber(id.value);
					// Invalid values should not occur at all (malformed Excel files).
					// Undefined values may occur if the file was saved by an Excel version that has
					// implemented yet unknown format numbers (undefined in NanoXLSX)
					format = new NumberFormat();
					format.setNumber(evaluation.getFormatNumber());
					format.setInternalID(this.styleReaderContainer.getNextNumberFormatId());
					this.styleReaderContainer.addStyleComponent(format);
				}
				id = ReaderUtils.IntParser.tryParse(childNode.getAttribute("borderId"));
				Border border = styleReaderContainer.getBorder(id.value);
				if (!id.hasValue || border == null) {
					border = new Border();
					border.setInternalID(styleReaderContainer.getNextBorderId());
				}
				id = ReaderUtils.IntParser.tryParse(childNode.getAttribute("fillId"));
				Fill fill = styleReaderContainer.getFill(id.value);
				if (!id.hasValue || fill == null) {
					fill = new Fill();
					fill.setInternalID(styleReaderContainer.getNextFillId());
				}
				id = ReaderUtils.IntParser.tryParse(childNode.getAttribute("fontId"));
				Font font = styleReaderContainer.getFont(id.value);
				if (!id.hasValue || font == null) {
					font = new Font();
					font.setInternalID(styleReaderContainer.getNextFontId());
				}

				// TODO: Implement other style information
				style.setNumberFormat(format);
				style.setBorder(border);
				style.setFill(fill);
				style.setFont(font);
				style.setCellXf(cellXfStyle);
				style.setInternalID(this.styleReaderContainer.getNextStyleId());

				this.styleReaderContainer.addStyleComponent(style);
			}
		}
	}

	private void getColors(XmlDocument.XmlNode node) {
		for (XmlDocument.XmlNode color : node.getChildNodes()) {
			XmlDocument.XmlNode mruColor = getChildNode(color, "color");
			if (color.getName().equals("mruColors") && mruColor != null) {
				for (XmlDocument.XmlNode value : color.getChildNodes()) {
					String attribute = value.getAttribute("rgb");
					if (attribute != null) {
						styleReaderContainer.addMruColor(attribute);
					}
				}
			}
		}
	}

	/**
	 * Resolves a color value from an XML node, when a rgb attribute exists
	 *
	 * @param node
	 *            Node to check
	 * @param fallback
	 *            Fallback value if the color could not be resolved
	 * @return RGB value as string or the fallback
	 */
	private static String getColor(XmlDocument.XmlNode node, String fallback) {
		XmlDocument.XmlNode childNode = getChildNode(node, "color");
		if (childNode != null) {
			return childNode.getAttribute("rgb");
		}
		return fallback;
	}

	/**
	 * Gets the specified child node
	 *
	 * @param node
	 *            XML node that contains child node
	 * @param name
	 *            Name of the child node
	 * @return Child node or null if not found
	 */
	private static XmlDocument.XmlNode getChildNode(XmlDocument.XmlNode node, String name) {
		if (node != null && node.getChildNodes().size() > 0) {
			for (XmlDocument.XmlNode childNode : node.getChildNodes()) {
				if (childNode.getName().equalsIgnoreCase(name)) {
					return childNode;
				}
			}
		}
		return null;
	}

	/**
	 * Tries to determine a border StyleValue enum entry from a string
	 *
	 * @param styleValue
	 *            String to check
	 * @return Enum value or {@link Border.StyleValue#none} if not found
	 */
	private static Border.StyleValue getBorderEnumValue(String styleValue) {
		try {
			return Border.StyleValue.valueOf(styleValue);
		}
		catch (Exception e) {
			return Border.StyleValue.none;
		}
	}

	/**
	 * Tries to determine a pattern enum entry from a string
	 *
	 * @param styleValue
	 *            String to check
	 * @return Enum value or {@link Fill.PatternValue#none} if not found
	 */
	private static Fill.PatternValue getFillPatternEnumValue(String styleValue) {
		try {
			return Fill.PatternValue.valueOf(styleValue);
		}
		catch (Exception e) {
			return Fill.PatternValue.none;
		}
	}

	/**
	 * Tries to determine a vertical align enum entry from a string
	 *
	 * @param styleValue
	 *            String to check
	 * @return Enum value or {@link Font.VerticalAlignValue#none} if not found
	 */
	private static Font.VerticalAlignValue getFontVerticalAlignEnumValue(String styleValue) {
		try {
			return Font.VerticalAlignValue.valueOf(styleValue);
		}
		catch (Exception e) {
			return Font.VerticalAlignValue.none;
		}
	}

	/**
	 * Tries to determine a horizontal align enum entry of cellXF from a string
	 *
	 * @param styleValue
	 *            String to check
	 * @return Enum value or {@link CellXf.HorizontalAlignValue#none} if not found
	 */
	private static CellXf.HorizontalAlignValue getCellXfHorizontalAlignEnumValue(String styleValue) {
		try {
			return CellXf.HorizontalAlignValue.valueOf(styleValue);
		}
		catch (Exception e) {
			return CellXf.HorizontalAlignValue.none;
		}
	}

	/**
	 * Tries to determine a vertical align enum entry of cellXF from a string
	 *
	 * @param styleValue
	 *            String to check
	 * @return Enum value or {@link CellXf.VerticalAlignValue#none} if not found
	 */
	private static CellXf.VerticalAlignValue getCellXfVerticalAlignEnumValue(String styleValue) {
		try {
			return CellXf.VerticalAlignValue.valueOf(styleValue);
		}
		catch (Exception e) {
			return CellXf.VerticalAlignValue.none;
		}
	}

	/**
	 * Gets the XML attribute from a child node of the passed XML node by its name
	 * and the name of the child node. This method simplifies the process of
	 * gathering one single child node attribute
	 *
	 * @param node
	 *            XML node that contains the child node
	 * @param childNodeName
	 *            Name of the child node
	 * @param attributeName
	 *            Name of the attribute in the child node
	 * @return Result object with the state of the child node and the attribute, as
	 *         well as the value (if available)
	 */
	private static LookupResult getAttributeOfChild(XmlDocument.XmlNode node, String childNodeName, String attributeName) {
		XmlDocument.XmlNode childNode = getChildNode(node, childNodeName);
		if (childNode != null) {
			String value = childNode.getAttribute(attributeName);
			return new LookupResult(value);
		}
		return new LookupResult();
	}

	/**
	 * Plain return class (struct-like) for node and attribute lookups
	 */
	private static class LookupResult {
		public final boolean attributeIsPresent;
		public final boolean nodeIsPresent;
		public final String value;

		public LookupResult() {
			this.attributeIsPresent = false;
			this.nodeIsPresent = false;
			this.value = null;
		}

		public LookupResult(String value) {
			this.attributeIsPresent = value != null;
			this.nodeIsPresent = true;
			this.value = value;
		}
	}

}
