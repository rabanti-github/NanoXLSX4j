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
    private String applicationVersion;
    private String category;
    private String company;
    private String contentStatus;
    private String creator;

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
     * Gets the version of the application that has created an XLSX file
     * @return Version number as string
     */
    public String getApplicationVersion() {
        return applicationVersion;
    }

    /**
     * Sets the version of the application that has created an XLSX file
     * @param applicationVersion Version number as string
     */
    public void setApplicationVersion(String applicationVersion) {
        this.applicationVersion = applicationVersion;
    }

    /**
     * Gets the document category of an XLSX file
     * @return Document category
     */
    public String getCategory() {
        return category;
    }

    /**
     * Sets the document category of an XLSX file
     * @param category Document category
     */
    public void setCategory(String category) {
        this.category = category;
    }

    /**
     * Gets the responsible company of an XLSX file
     * @return Company name
     */
    public String getCompany() {
        return company;
    }

    /**
     * Sets the responsible company of an XLSX file
     * @param company Company name
     */
    public void setCompany(String company) {
        this.company = company;
    }

    /**
     * Gets the content status of an XLSX file
     * @return Content status
     */
    public String getContentStatus() {
        return contentStatus;
    }

    /**
     * Sets the content status of an XLSX file
     * @param contentStatus Content status
     */
    public void setContentStatus(String contentStatus) {
        this.contentStatus = contentStatus;
    }

    /**
     * Gets the creator of an XLSX file
     * @return Creator name
     */
    public String getCreator() {
        return creator;
    }

    /**
     * Sets the creator of an XLSX file
     * @param creator Creator name
     */
    public void setCreator(String creator) {
        this.creator = creator;
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
                else if (node.getName().equalsIgnoreCase("AppVersion")){
                    this.applicationVersion = node.getInnerText();
                }
                else if (node.getName().equalsIgnoreCase("Company")){
                    this.company = node.getInnerText();
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
                if (node.getName().equalsIgnoreCase("Category")){
                    this.category = node.getInnerText();
                }
                else if (node.getName().equalsIgnoreCase("ContentStatus")){
                    this.contentStatus = node.getInnerText();
                }
                else if (node.getName().equalsIgnoreCase("Creator")){
                    this.creator = node.getInnerText();
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




}
