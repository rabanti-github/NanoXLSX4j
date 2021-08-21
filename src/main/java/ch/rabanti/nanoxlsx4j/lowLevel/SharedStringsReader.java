/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.exceptions.IOException;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Class representing a reader for the shared strings table of XLSX files
 *
 * @author Raphael Stoeckli
 */
public class SharedStringsReader {

    private final List<String> sharedStrings;

    /**
     * Gets the shared strings list
     *
     * @return ArrayList of shared string entries
     */
    public List<String> getSharedStrings() {
        return sharedStrings;
    }

    /**
     * Gets whether the workbook contains shared strings
     *
     * @return True if at least one shared string object exists in the workbook
     */
    public boolean hasElements() {
        return !sharedStrings.isEmpty();
    }

    /**
     * Gets the value of the shared string table by its index
     *
     * @param index Index of the stared string entry
     * @return Determined shared string value. Returns null in case of a invalid index
     */
    public String getString(int index) {
        if (!hasElements() || index > sharedStrings.size() - 1) {
            return null;
        }
        return sharedStrings.get(index);
    }

    /**
     * Default constructor
     */
    public SharedStringsReader() {
        this.sharedStrings = new ArrayList<>();
    }

    /**
     * Reads the XML file form the passed stream and processes the shared strings table
     *
     * @param stream Stream of the XML file
     * @throws IOException Throws IOException in case of an error
     */
    public void read(InputStream stream) throws IOException, java.io.IOException {
        try {
            XmlDocument xr = new XmlDocument();
            xr.load(stream);
            StringBuilder sb = new StringBuilder();
            for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) {
                if (node.getName().equalsIgnoreCase("si")) {
                    sb = new StringBuilder();
                    getTextToken(node, sb);
                    this.sharedStrings.add(sb.toString());
                }
            }
        } catch (Exception ex) {
            throw new IOException("XMLStreamException", "The XML entry could not be read from the input stream. Please see the inner exception:", ex);
        } finally {
            if (stream != null) {
                stream.close();
            }
        }
    }

    /**
     * Function collects text tokens recursively in case of a split by formatting
     *
     * @param node Root node to process
     * @param sb   StringBuilder reference
     */
    private void getTextToken(XmlDocument.XmlNode node, StringBuilder sb) {
        if (node.getName().equalsIgnoreCase("t") && !Helper.isNullOrEmpty(node.getInnerText())) {
            sb.append(node.getInnerText());
        }
        if (node.hasChildNodes()) {
            for (XmlDocument.XmlNode childNode : node.getChildNodes()) {
                getTextToken(childNode, sb);
            }
        }
    }

}