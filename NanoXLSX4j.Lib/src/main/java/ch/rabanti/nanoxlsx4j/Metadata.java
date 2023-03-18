/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

/**
 * Class representing the meta data of a workbook
 *
 * @author Raphael Stoeckli
 */
public class Metadata {

	// ### P R I V A T E F I E L D S ###
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

	// ### G E T T E R S & S E T T E R S ###

	/**
	 * Gets the application which created the workbook. Default is PicoXLSX4j
	 *
	 * @return Application which created the workbook
	 */
	public String getApplication() {
		return application;
	}

	/**
	 * Sets the application which created the workbook. Default is PicoXLSX4j
	 *
	 * @param application
	 *            Application which created the workbook
	 */
	public void setApplication(String application) {
		this.application = application;
	}

	/**
	 * Gets the version of the creation application. Default is the library version
	 * of PicoXLSjX. Use the format xxxxx.yyyyy (e.g. 1.0 or 55.9875) in case of a
	 * custom value.
	 *
	 * @return Version of the creation application
	 */
	public String getApplicationVersion() {
		return applicationVersion;
	}

	/**
	 * Sets the version of the creation application. Default is the library version
	 * of PicoXLSX4j. Use the format xxxxx.yyyyy (e.g. 1.0 or 55.9875) in case of a
	 * custom value.
	 *
	 * @param applicationVersion
	 *            Version of the creation application
	 * @throws FormatException
	 *             Thrown if the passed version results in a higher major or minor
	 *             number of 99999
	 * @apiNote Allowed values are null, empty and fractions from 0.0 to 99999.99999
	 *          (max. number of digits before and after the period is 5)
	 */
	public void setApplicationVersion(String applicationVersion) {
		this.applicationVersion = applicationVersion;
		checkVersion();
	}

	/**
	 * Gets the category of the document. There are no predefined values or
	 * restrictions about the content of this field
	 *
	 * @return Category of the document
	 */
	public String getCategory() {
		return category;
	}

	/**
	 * Sets the category of the document. There are no predefined values or
	 * restrictions about the content of this field
	 *
	 * @param category
	 *            Category of the document
	 */
	public void setCategory(String category) {
		this.category = category;
	}

	/**
	 * Gets the company owning the document. This value is for organizational
	 * purpose. Add more than one manager by using the semicolon (;) between the
	 * words
	 *
	 * @return Company owning the document
	 */
	public String getCompany() {
		return company;
	}

	/**
	 * Sets the company owning the document. This value is for organizational
	 * purpose. Add more than one manager by using the semicolon (;) between the
	 * words
	 *
	 * @param company
	 *            Company owning the document
	 */
	public void setCompany(String company) {
		this.company = company;
	}

	/**
	 * Gets the status of the document. There are no predefined values or
	 * restrictions about the content of this field
	 *
	 * @return Status of the document
	 */
	public String getContentStatus() {
		return contentStatus;
	}

	/**
	 * Sets the status of the document. There are no predefined values or
	 * restrictions about the content of this field
	 *
	 * @param contentStatus
	 *            Status of the document
	 */
	public void setContentStatus(String contentStatus) {
		this.contentStatus = contentStatus;
	}

	/**
	 * Gets the creator of the workbook. Add more than one creator by using the
	 * semicolon (;) between the authors
	 *
	 * @return Creator of the workbook
	 */
	public String getCreator() {
		return creator;
	}

	/**
	 * Sets the creator of the workbook. Add more than one creator by using the
	 * semicolon (;) between the authors
	 *
	 * @param creator
	 *            Creator of the workbook
	 */
	public void setCreator(String creator) {
		this.creator = creator;
	}

	/**
	 * Gets the description of the document or comment about it
	 *
	 * @return Description of the document
	 */
	public String getDescription() {
		return description;
	}

	/**
	 * Sets the description of the document or comment about it
	 *
	 * @param description
	 *            Description of the document
	 */
	public void setDescription(String description) {
		this.description = description;
	}

	/**
	 * Gets the hyper-link base of the document
	 *
	 * @return Hyper-link base of the document
	 */
	public String getHyperlinkBase() {
		return hyperlinkBase;
	}

	/**
	 * Sets the hyper-link base of the document
	 *
	 * @param hyperlinkBase
	 *            Hyper-link base of the document
	 */
	public void setHyperlinkBase(String hyperlinkBase) {
		this.hyperlinkBase = hyperlinkBase;
	}

	/**
	 * Gets the keyword for the workbook. Separate the keywords with semicolons (;)
	 *
	 * @return Keywords for the workbook
	 */
	public String getKeywords() {
		return keywords;
	}

	/**
	 * Sets the keywords for the workbook. Separate the keywords with semicolons (;)
	 *
	 * @param keywords
	 *            Keywords for the workbook
	 */
	public void setKeywords(String keywords) {
		this.keywords = keywords;
	}

	/**
	 * Gets the responsible manager of the document. This value is for
	 * organizational purpose.
	 *
	 * @return Responsible manager of the document
	 */
	public String getManager() {
		return manager;
	}

	/**
	 * Sets the responsible manager of the document. This value is for
	 * organizational purpose.
	 *
	 * @param manager
	 *            Responsible manager of the document
	 */
	public void setManager(String manager) {
		this.manager = manager;
	}

	/**
	 * Gets the subject of the workbook
	 *
	 * @return Subject of the workbook
	 */
	public String getSubject() {
		return subject;
	}

	/**
	 * Sets the subject of the workbook
	 *
	 * @param subject
	 *            Subject of the workbook
	 */
	public void setSubject(String subject) {
		this.subject = subject;
	}

	/**
	 * Gets the title of the workbook
	 *
	 * @return Title of the workbook
	 */
	public String getTitle() {
		return title;
	}

	/**
	 * Sets the title of the workbook
	 *
	 * @param title
	 *            Title of the workbook
	 */
	public void setTitle(String title) {
		this.title = title;
	}

	// ### C O N S T R U C T O R S ###

	/**
	 * Default constructor
	 */
	public Metadata() {
		this.application = Version.APPLICATION_NAME;
		this.applicationVersion = Version.VERSION;

	}

	// ### M E T H O D S ###

	/**
	 * Checks the format of the passed version string. Allowed values are null,
	 * empty and fractions from 0.0 to 99999.99999 (max. number of digits before and
	 * after the period is 5)
	 *
	 * @throws FormatException
	 *             Thrown if the version string is malformed
	 */
	private void checkVersion() {
		if (Helper.isNullOrEmpty(this.applicationVersion)) {
			return;
		}
		String[] split = this.applicationVersion.split("\\.");
		boolean state = true;
		if (split.length != 2) {
			state = false;
		}
		else {
			if (split[1].length() < 1 || split[1].length() > 5) {
				state = false;
			}
			if (split[0].length() < 1 || split[0].length() > 5) {
				state = false;
			}
		}
		if (!state) {
			throw new FormatException("The format of the version in the meta data is wrong (" +
					this.applicationVersion +
					"). Should be in the format and a range from '0.0' to '99999.99999'");
		}
	}

	// ### S T A T I C M E T H O D S ###

	/**
	 * Method to parse a common version (major.minor.revision.build) into the
	 * compatible format (major.minor). The minimum value is 0.0 and the maximum
	 * value is 99999.99999<br>
	 * The minor, revision and build number are joined if possible. If the number is
	 * too long, the additional characters will be removed from the right side down
	 * to five characters (e.g. 785563 will be 78556)
	 *
	 * @param major
	 *            Major number from 0 to 99999
	 * @param minor
	 *            Minor number
	 * @param build
	 *            Build number
	 * @param revision
	 *            Revision number
	 * @return Formatted version number (e.g. 1.0 or 55.987)
	 * @throws FormatException
	 *             Thrown if the major number is too long or one of the numbers is
	 *             negative
	 */
	public static String parseVersion(int major, int minor, int build, int revision) {
		if (major < 0 || minor < 0 || build < 0 || revision < 0) {
			throw new FormatException("The format of the passed version is wrong. No negative number allowed.");
		}
		if (major > 99999) {
			throw new FormatException("The major number may not be bigger than 99999. The passed value is " + major);
		}
		String leftPart = Integer.toString(major);
		String rightPart = Integer.toString(minor) + build + revision;
		rightPart = rightPart.replaceAll("0+$", "");
		if (rightPart.length() == 0) {
			rightPart = "0";
		}
		else {
			if (rightPart.length() > 5) {
				rightPart = rightPart.substring(0, 5);
			}
		}
		return leftPart + "." + rightPart;
	}

}
