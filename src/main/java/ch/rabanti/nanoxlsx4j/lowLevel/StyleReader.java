/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import ch.rabanti.nanoxlsx4j.styles.Border;
import ch.rabanti.nanoxlsx4j.styles.Fill;
import ch.rabanti.nanoxlsx4j.styles.NumberFormat;
import ch.rabanti.nanoxlsx4j.styles.Style;

import java.io.InputStream;

/**
 * Class representing a reader for style definitions of XLSX files
 *
 * @author Raphael Stoeckli
 */
public class StyleReader {

    private StyleReaderContainer styleReaderContainer;

    /**
     * Gets the container for raw style components of the reader.
     *
     * @return Style reader container with resolved style components
     */
    public StyleReaderContainer getStyleReaderContainer() {
        return styleReaderContainer;
    }

    /**
     * Sets the container for raw style components of the reader.
     *
     * @param styleReaderContainer Style reader container with resolved style components
     */
    public void setStyleReaderContainer(StyleReaderContainer styleReaderContainer) {
        this.styleReaderContainer = styleReaderContainer;
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
     * @param stream Stream of the XML file
     * @throws IOException Throws IOException in case of an error
     */
    public void read(InputStream stream) throws IOException, java.io.IOException {
        XmlDocument xr = new XmlDocument();
        xr.load(stream);
        for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) {
            if (node.getName().equalsIgnoreCase("numfmts")) { // Handles custom number formats
                this.getNumberFormats(node);
            } else if (node.getName().equalsIgnoreCase("borders")) { // handle borders
                this.getBorders(node);
            }
            else if (node.getName().equalsIgnoreCase("fills")) { // handle fills
                this.getFills(node);
            }
            // TODO: Implement other style components
        }
        for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) { // Redo for composition after all style parts are gathered; standard number formats
            if (node.getName().equalsIgnoreCase("cellxfs")) {
                this.getCellXfs(node);
            }
        }
    }

    /**
     * Determines the number formats in an XML node of the style document
     *
     * @param node Number formats root node
     */
    private void getNumberFormats(XmlDocument.XmlNode node) {
        for (XmlDocument.XmlNode childNode : node.getChildNodes()) {
            if (childNode.getName().equalsIgnoreCase("numfmt")) {
                NumberFormat numberFormat = new NumberFormat();
                int id = Integer.parseInt(childNode.getAttribute("numFmtId")); // Default/null will (justified) throw an exception

                String code = childNode.getAttribute("formatCode", "");

                if (id < NumberFormat.CUSTOMFORMAT_START_NUMBER) {
                    NumberFormat.NumberFormatEvaluation enumValue = NumberFormat.tryParseFormatNumber(id);
                    if (enumValue.getRange() == NumberFormat.NumberFormatEvaluation.FormatRange.defined_format) {
                        numberFormat.setNumber(enumValue.getFormatNumber());
                    } else {
                        numberFormat.setCustomFormatID(id);
                        numberFormat.setNumber(NumberFormat.FormatNumber.custom);
                    }
                } else {
                    numberFormat.setCustomFormatID(id);
                    numberFormat.setNumber(NumberFormat.FormatNumber.custom);
                }
                numberFormat.setInternalID(id);
                numberFormat.setCustomFormatCode(code);
                this.styleReaderContainer.addStyleComponent(numberFormat);
            }
        }
    }

    /**
     * Determines the borders in an XML node of the style document
     *
     * @param node Border root node
     */
    private void getBorders(XmlDocument.XmlNode node) {
        for (XmlDocument.XmlNode border : node.getChildNodes()) {
            Border borderStyle = new Border();
            String diagonalDown = border.getAttribute("diagonalDown");
            String diagonalUp = border.getAttribute("diagonalUp");
            if (diagonalDown != null && diagonalDown.equals("1")) {
                borderStyle.setDiagonalDown(true);
            }
            if (diagonalUp != null && diagonalUp.equals("1")) {
                borderStyle.setDiagonalUp(true);
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
     * @param innerNode Border sub-node
     * @return Border type or non if parsing was not successful
     */
    private static Border.StyleValue parseBorderStyle(XmlDocument.XmlNode innerNode)
    {
        String value = innerNode.getAttribute("style");
        if (value != null)
        {
            if (value.equalsIgnoreCase("double"))
            {
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
                XmlDocument.XmlNode foregroundNode = getChildNode(innerNode, "fgColor");
                if (foregroundNode != null){
                    String foregroundRgba = foregroundNode.getAttribute("rgb");
                    if (foregroundRgba != null)
                    {
                        fillStyle.setForegroundColor(foregroundRgba);
                    }
                }
                XmlDocument.XmlNode backgroundNode = getChildNode(innerNode, "bgColor");
                if (backgroundNode != null){
                    String backgroundRgba = backgroundNode.getAttribute("rgb");
                    if (backgroundRgba != null)
                    {
                        fillStyle.setBackgroundColor(backgroundRgba);
                    }
                    String backgroundIndex = backgroundNode.getAttribute("indexed");
                    if (backgroundIndex != null)
                    {
                        fillStyle.setIndexedColor(Integer.parseInt(backgroundIndex));
                    }
                }
            }
            fillStyle.setInternalID(styleReaderContainer.getNextFillId());
            styleReaderContainer.addStyleComponent(fillStyle);
        }
    }

    /**
     * Determines the cell XF entries in a XML node of the style document
     *
     * @param node Cell XF root node
     */
    private void getCellXfs(XmlDocument.XmlNode node) {

        for (XmlDocument.XmlNode childNode : node.getChildNodes()) {
            if (childNode.getName().equalsIgnoreCase("xf")) {
                Style style = new Style();
                int id = Integer.parseInt(childNode.getAttribute("numFmtId"));


                NumberFormat format = this.styleReaderContainer.getNumberFormat(id, true);
                if (format == null) {
                    NumberFormat.NumberFormatEvaluation evaluation = NumberFormat.tryParseFormatNumber(id); // Validity is neglected here to prevent unhandled crashes. If invalid, the format will be declared as 'none'                                          // Invalid values should not occur at all (malformed Excel files); Undefined values may occur if the file was saved by an Excel version that has implemented yet unknown format numbers (undefined in NanoXLSX)
                    // Invalid values should not occur at all (malformed Excel files).
                    // Undefined values may occur if the file was saved by an Excel version that has implemented yet unknown format numbers (undefined in NanoXLSX)
                    format = new NumberFormat();
                    format.setNumber(evaluation.getFormatNumber());
                    format.setInternalID(this.styleReaderContainer.getNextNumberFormatId());
                    this.styleReaderContainer.addStyleComponent(format);
                }
                id = Integer.parseInt(childNode.getAttribute("borderId"));
                Border border = styleReaderContainer.getBorder(id, true);
                if (border == null) {
                    border = new Border();
                    border.setInternalID(styleReaderContainer.getNextBorderId());
                }
                id = Integer.parseInt(childNode.getAttribute("fillId"));
                Fill fill = styleReaderContainer.getFill(id, true);
                if (fill == null) {
                    fill = new Fill();
                    fill.setInternalID(styleReaderContainer.getNextFillId());
                }
                // TODO: Implement other style information
                style.setNumberFormat(format);
                style.setBorder(border);
                style.setFill(fill);
                style.setInternalID(this.styleReaderContainer.getNextStyleId());

                this.styleReaderContainer.addStyleComponent(style);
            }
        }
    }

    /**
     * Resolves a color value from an XML node, when a rgb attribute exists
     *
     * @param node     Node to check
     * @param fallback Fallback value if the color could not be resolved
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
     * @param node XML node that contains child node
     * @param name Name of the child node
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
     * @param styleValue String to check
     * @return Enum value or null if not found
     */
    private static Border.StyleValue getBorderEnumValue(String styleValue) {
        try {
            return Border.StyleValue.valueOf(styleValue);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Tries to determine a pattern value type StyleValue enum entry from a string
     *
     * @param styleValue String to check
     * @return Enum value or null if not found
     */
    private static Fill.PatternValue getFillPatternEnumValue(String styleValue) {
        try {
            return Fill.PatternValue.valueOf(styleValue);
        } catch (Exception e) {
            return null;
        }
    }

}
