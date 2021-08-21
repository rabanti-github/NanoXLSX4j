/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;
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
     * Determines the number formats in a XML node of the style document
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
                // TODO: Implement other style information
                style.setNumberFormat(format);
                style.setInternalID(this.styleReaderContainer.getNextStyleId());

                this.styleReaderContainer.addStyleComponent(style);
            }
        }
    }

}
