/*
 * NanoXLSX4j is a small Java library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;

import java.io.InputStream;

/**
 * Class representing a reader for the metadata files (docProps) of XLSX files
 */
public class MetaDataReader {

    private String application;

    /**
     * Gets the application that has created an XLSX file. This is an arbitrary text and the default of this library is "NanoXLSX4j"
     * @return Application name
     */
    public String getApplication() {
        return application;
    }

    /**
     * Sets the application that has created an XLSX file
     * @param application Application name
     */
    public void setApplication(String application) {
        this.application = application;
    }

    /**
     * Reads the XML file form the passed stream and processes the AppData section
     * @param stream Stream of the XML file
     */
    public void ReadAppData(InputStream stream) throws IOException, java.io.IOException {
        try {
            XmlDocument xr = new XmlDocument();
            xr.load(stream);
            for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) {
                if (node.getName().equalsIgnoreCase("Application")){
                    this.application = node.getInnerText();
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
     * Reads the XML file form the passed stream and processes the Core section
     * @param stream Stream of the XML file
     */
    public void ReadCoreData(InputStream stream) throws IOException, java.io.IOException {
        try {
            XmlDocument xr = new XmlDocument();
            xr.load(stream);
            for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) {
            // TODO implement
            }
        } catch (Exception ex) {
            throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
        } finally {
            if (stream != null) {
                stream.close();
            }
        }
    }




}
