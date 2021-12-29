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
    private final Map<Integer, WorksheetDefinition> worksheetDefinitions;

    /**
     * Hashmap of worksheet definitions. The key is the worksheet number and the value is a WorksheetDefinition object
     * with name, hidden state and other information
     *
     * @return Hashmap of number-name tuples
     */
    public Map<Integer, WorksheetDefinition> getWorksheetDefinitions() {
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
                String state = node.getAttribute("state");
                boolean hidden = false;
                if (state != null && state.equalsIgnoreCase("hidden"))
                {
                    hidden = true;
                }
                WorksheetDefinition definition = new WorksheetDefinition(id, sheetName);
                definition.setHidden(hidden);
                worksheetDefinitions.put(id, definition);
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

// ### S U B - C L A S S E S ###

    /**
     * Class for worksheet Mata-data on import
     */
    public static class WorksheetDefinition{
        private final String worksheetName;
        private final int sheetId;
        private boolean hidden;

        /**
         * Sets the hidden state of the worksheet
         * @param hidden True if hidden
         */
        public void setHidden(boolean hidden) {
            this.hidden = hidden;
        }

        /**
         * Gets the hidden state of the worksheet
         * @return True if hidden
         */
        public boolean isHidden() {
            return hidden;
        }

        /**
         * gets the worksheet name
         * @return Name as string
         */
        public String getWorksheetName() {
            return worksheetName;
        }

        /**
         * Internal worksheet ID
         * @return Intenal id
         */
        public int getSheetId() {
            return sheetId;
        }

        /**
         * Default constructor with parameters
         * @param worksheetName Worksheet name
         * @param sheetId Internal ID
         */
        public WorksheetDefinition(int sheetId, String worksheetName) {
            this.worksheetName = worksheetName;
            this.sheetId = sheetId;
        }
    }

}
