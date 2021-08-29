/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * Class representing a reader to decompile a workbook in an XLSX files
 *
 * @author Raphael Stoeckli
 */
public class WorkbookReader {
    private final Map<Integer, String> worksheetDefinitions;

    /**
     * Hashmap of worksheet definitions. The key is the worksheet number and the value is the worksheet name
     *
     * @return Hashmap of number-name tuples
     */
    public Map<Integer, String> getWorksheetDefinitions() {
        return worksheetDefinitions;
    }

    /**
     * Default constructor
     */
    public WorkbookReader() {
        this.worksheetDefinitions = new HashMap<>();
    }

    /**
     * Reads the XML file form the passed stream and processes the workbook information
     *
     * @param stream Stream of the XML file
     * @throws IOException Throws IOException in case of an error
     */
    public void read(InputStream stream) throws IOException, java.io.IOException {
        try {
            XmlDocument xr = new XmlDocument();
            xr.load(stream);
            for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) {
                getWorkbookInformation(node);
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
     * Finds the workbook information recursively
     *
     * @param node Root node to check
     * @throws IOException Thrown if the workbook information could not be determined
     */
    private void getWorkbookInformation(XmlDocument.XmlNode node) throws IOException {
        if (node.getName().equalsIgnoreCase("sheet")) {
            try {
                String sheetName = node.getAttribute("name", "worksheet1");
                int id = Integer.parseInt(node.getAttribute("sheetId")); // Default will rightly throw an exception
                worksheetDefinitions.put(id, sheetName);
            } catch (Exception e) {
                throw new IOException("The workbook information could not be resolved. Please see the inner exception:", e);
            }
        }
        if (node.hasChildNodes()) {
            for (XmlDocument.XmlNode childNode : node.getChildNodes()) {
                getWorkbookInformation(childNode);
            }
        }
    }

}
