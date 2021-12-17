/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import ch.rabanti.nanoxlsx4j.styles.Border;
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
            Border.StyleValue styleType;
            XmlDocument.XmlNode innerNode = getChildNode(border, "diagonal");
            if (innerNode != null) {
                String styleValue = innerNode.getAttribute("style");
                styleType = getEnumValue(styleValue);
                if (styleValue != null) {
                    borderStyle.setDiagonalStyle(styleType);
                }
                borderStyle.setDiagonalColor(getColor(innerNode, Border.DEFAULT_BORDER_COLOR));
            }
            innerNode = getChildNode(border, "top");
            if (innerNode != null) {
                String styleValue = innerNode.getAttribute("style");
                styleType = getEnumValue(styleValue);
                if (styleValue != null) {
                    borderStyle.setTopStyle(styleType);
                }
                borderStyle.setTopColor(getColor(innerNode, Border.DEFAULT_BORDER_COLOR));
            }
            innerNode = getChildNode(border, "bottom");
            if (innerNode != null) {
                String styleValue = innerNode.getAttribute("style");
                styleType = getEnumValue(styleValue);
                if (styleValue != null) {
                    borderStyle.setBottomStyle(styleType);
                }
                borderStyle.setBottomColor(getColor(innerNode, Border.DEFAULT_BORDER_COLOR));
            }
            innerNode = getChildNode(border, "left");
            if (innerNode != null) {
                String styleValue = innerNode.getAttribute("style");
                styleType = getEnumValue(styleValue);
                if (styleValue != null) {
                    borderStyle.setLeftStyle(styleType);
                }
                borderStyle.setLeftColor(getColor(innerNode, Border.DEFAULT_BORDER_COLOR));
            }
            innerNode = getChildNode(border, "right");
            if (innerNode != null) {
                String styleValue = innerNode.getAttribute("style");
                styleType = getEnumValue(styleValue);
                if (styleValue != null) {
                    borderStyle.setRightStyle(styleType);
                }
                borderStyle.setRightColor(getColor(innerNode, Border.DEFAULT_BORDER_COLOR));
            }
            borderStyle.setInternalID(styleReaderContainer.getNextBorderId());
            styleReaderContainer.addStyleComponent(borderStyle);
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
                // TODO: Implement other style information
                style.setNumberFormat(format);
                style.setBorder(border);
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
    private static Border.StyleValue getEnumValue(String styleValue) {
        try {
            return Border.StyleValue.valueOf(styleValue);
        } catch (Exception e) {
            return null;
        }
    }

}
