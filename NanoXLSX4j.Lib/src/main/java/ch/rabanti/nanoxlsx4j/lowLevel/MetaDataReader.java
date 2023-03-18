/*
 * NanoXLSX4j is a small Java library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import java.io.InputStream;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;

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
	private String description;
	private String hyperlinkBase;
	private String keywords;
	private String manager;
	private String subject;
	private String title;

	/**
	 * Gets the application that has created an XLSX file. This is an arbitrary text
	 * and the default of this library is "NanoXLSX4j"
	 *
	 * @return Application name
	 */
	public String getApplication() {
		return application;
	}

	/**
	 * Gets the version of the application that has created an XLSX file
	 *
	 * @return Version number as string
	 */
	public String getApplicationVersion() {
		return applicationVersion;
	}

	/**
	 * Gets the document category of an XLSX file
	 *
	 * @return Document category
	 */
	public String getCategory() {
		return category;
	}

	/**
	 * Gets the responsible company of an XLSX file
	 *
	 * @return Company name
	 */
	public String getCompany() {
		return company;
	}

	/**
	 * Gets the content status of an XLSX file
	 *
	 * @return Content status
	 */
	public String getContentStatus() {
		return contentStatus;
	}

	/**
	 * Gets the creator of an XLSX file
	 *
	 * @return Creator name
	 */
	public String getCreator() {
		return creator;
	}

	/**
	 * Gets the description of an XLSX file
	 *
	 * @return Description text
	 */
	public String getDescription() {
		return description;
	}

	/**
	 * Gets the hyperlink base of an XLSX file
	 *
	 * @return Hyperlink base of the document
	 */
	public String getHyperlinkBase() {
		return hyperlinkBase;
	}

	/**
	 * Gets the keywords of an XLSX file
	 *
	 * @return Keywords
	 */
	public String getKeywords() {
		return keywords;
	}

	/**
	 * Gets the manager (responsible) of an XLSX file
	 *
	 * @return Manager (responsible)
	 */
	public String getManager() {
		return manager;
	}

	/**
	 * Gets the subject of an XLSX file
	 *
	 * @return Subject text
	 */
	public String getSubject() {
		return subject;
	}

	/**
	 * Gets the title of an XLSX file
	 *
	 * @return Tile text
	 */
	public String getTitle() {
		return title;
	}

	/**
	 * Reads the XML file form the passed stream and processes the AppData section
	 *
	 * @param stream
	 *            Stream of the XML file
	 */
	public void ReadAppData(InputStream stream) throws IOException, java.io.IOException {
		try {
			XmlDocument xr = new XmlDocument();
			xr.load(stream);
			for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) {
				if (node.getName().equalsIgnoreCase("Application")) {
					this.application = node.getInnerText();
				}
				else if (node.getName().equalsIgnoreCase("AppVersion")) {
					this.applicationVersion = node.getInnerText();
				}
				else if (node.getName().equalsIgnoreCase("Company")) {
					this.company = node.getInnerText();
				}
				else if (node.getName().equalsIgnoreCase("HyperlinkBase")) {
					this.hyperlinkBase = node.getInnerText();
				}
				else if (node.getName().equalsIgnoreCase("Manager")) {
					this.manager = node.getInnerText();
				}
			}
		}
		catch (Exception ex) {
			throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
		}
		finally {
			if (stream != null) {
				stream.close();
			}
		}
	}

	/**
	 * Reads the XML file form the passed stream and processes the Core section
	 *
	 * @param stream
	 *            Stream of the XML file
	 */
	public void ReadCoreData(InputStream stream) throws IOException, java.io.IOException {
		try {
			XmlDocument xr = new XmlDocument();
			xr.load(stream);
			for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) {
				if (node.getName().equalsIgnoreCase("Category")) {
					this.category = node.getInnerText();
				}
				else if (node.getName().equalsIgnoreCase("ContentStatus")) {
					this.contentStatus = node.getInnerText();
				}
				else if (node.getName().equalsIgnoreCase("Creator")) {
					this.creator = node.getInnerText();
				}
				else if (node.getName().equalsIgnoreCase("Description")) {
					this.description = node.getInnerText();
				}
				else if (node.getName().equalsIgnoreCase("Keywords")) {
					this.keywords = node.getInnerText();
				}
				else if (node.getName().equalsIgnoreCase("Subject")) {
					this.subject = node.getInnerText();
				}
				else if (node.getName().equalsIgnoreCase("Title")) {
					this.title = node.getInnerText();
				}
			}
		}
		catch (Exception ex) {
			throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
		}
		finally {
			if (stream != null) {
				stream.close();
			}
		}
	}

}
