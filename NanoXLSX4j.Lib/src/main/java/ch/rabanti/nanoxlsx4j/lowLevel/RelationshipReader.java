/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Class representing a reader for relationship of XLSX files
 */
public class RelationshipReader {

    private List<RelationShip> relationships;

    /**
     * Gets the list of workbook relationship entries
     * @return
     */
    public List<RelationShip> getRelationships() {
        return relationships;
    }

    /**
     * Default constructor
     */
    public RelationshipReader() {
        this.relationships = new ArrayList<>();
    }

    /**
     * Reads the XML file form the passed stream and processes the relationship declaration
     * table
     *
     * @param stream Stream of the XML file
     * @throws IOException Throws IOException in case of an error
     */
    public void read(InputStream stream) throws IOException, java.io.IOException {
        try {
            XmlDocument xr = new XmlDocument();
            xr.load(stream);
            for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) {
                if (node.getName().equalsIgnoreCase("relationship")) {
                    String id = node.getAttribute("Id");
                    String type = node.getAttribute("Type");
                    String target = node.getAttribute("Target");
                    if (target.startsWith("/")) {
                        target = target.replaceFirst("^/+", "");
                    }
                    if (!target.startsWith("xl/")) {
                        target = "xl/" + target;
                    }
                    relationships.add(new RelationShip(id, type, target));
                }
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
     * Class to represent an OOXML document relationship tuple
     */
    static class RelationShip {
        private final String id;
        private final String type;
        private final String target;

        /**
         * Gets the internal document ID
         *
         * @return ID as string
         */
        public String getId() {
            return id;
        }

        /**
         * Gets the internal document type
         *
         * @return Type as string
         */
        public String getType() {
            return type;
        }

        /**
         * Gets the internal target of the document
         *
         * @return Target as string
         */
        public String getTarget() {
            return target;
        }

        /**
         * Constructor with parameters
         *
         * @param id     Internal ID
         * @param type   Internal type
         * @param target Internal target
         */
        public RelationShip(String id, String type, String target) {
            this.id = id;
            this.type = type;
            this.target = target;
        }
    }

}
