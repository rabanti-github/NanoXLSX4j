/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Class representing the metadata of a workbook.
 */
public class Metadata {

    private static final String VERSION_RESOURCE = "/version.properties";
    private static final String VERSION_PROPERTY = "build.version";
    private static final String FALLBACK_VERSION = "0.0";
    private static final Pattern MAVEN_VERSION_PATTERN =
        Pattern.compile("^(\\d+)(?:\\.(\\d+))?(?:\\.(\\d+))?(?:[-+].*)?$");

    /** Default application name, if not otherwise specified. */
    public static final String DEFAULT_APPLICATION_NAME = "NanoXLSX4j";
    /** Default application version, if not otherwise specified. */
    public static final String DEFAULT_APPLICATION_VERSION = loadDefaultApplicationVersion();

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

    /** Creates metadata with the NanoXLSX4j application name and version defaults. */
    public Metadata() {
        this.application = DEFAULT_APPLICATION_NAME;
        this.applicationVersion = DEFAULT_APPLICATION_VERSION;
    }

    /**
     * Gets the application which created the workbook.
     *
     * @return application which created the workbook
     */
    public String getApplication() {
        return application;
    }

    /**
     * Sets the application which created the workbook.
     *
     * @param application application which created the workbook
     */
    public void setApplication(String application) {
        this.application = application;
    }

    /**
     * Gets the version of the application which created the workbook.
     *
     * @return version of the application which created the workbook
     */
    public String getApplicationVersion() {
        return applicationVersion;
    }

    /**
     * Sets the version of the application which created the workbook.
     *
     * @param applicationVersion application version in the format {@code xxxxx.yyyyy}
     * @throws FormatException if the version is malformed
     * @apiNote Allowed values are {@code null}, empty, and values from {@code 0.0} to {@code 99999.99999}, with at
     * most five characters on either side of the period.
     */
    public void setApplicationVersion(String applicationVersion) {
        this.applicationVersion = applicationVersion;
        checkVersion();
    }

    /**
     * Gets the category of the document.
     *
     * @return category of the document
     */
    public String getCategory() {
        return category;
    }

    /**
     * Sets the category of the document.
     *
     * @param category category of the document
     */
    public void setCategory(String category) {
        this.category = category;
    }

    /**
     * Gets the company owning the document.
     *
     * @return company owning the document
     */
    public String getCompany() {
        return company;
    }

    /**
     * Sets the company owning the document.
     *
     * @param company company owning the document
     */
    public void setCompany(String company) {
        this.company = company;
    }

    /**
     * Gets the status of the document.
     *
     * @return status of the document
     */
    public String getContentStatus() {
        return contentStatus;
    }

    /**
     * Sets the status of the document.
     *
     * @param contentStatus status of the document
     */
    public void setContentStatus(String contentStatus) {
        this.contentStatus = contentStatus;
    }

    /**
     * Gets the creator of the workbook.
     *
     * @return creator of the workbook
     */
    public String getCreator() {
        return creator;
    }

    /**
     * Sets the creator of the workbook.
     *
     * @param creator creator of the workbook
     */
    public void setCreator(String creator) {
        this.creator = creator;
    }

    /**
     * Gets the description of the document.
     *
     * @return description of the document
     */
    public String getDescription() {
        return description;
    }

    /**
     * Sets the description of the document.
     *
     * @param description description of the document
     */
    public void setDescription(String description) {
        this.description = description;
    }

    /**
     * Gets the hyperlink base of the document.
     *
     * @return hyperlink base of the document
     */
    public String getHyperlinkBase() {
        return hyperlinkBase;
    }

    /**
     * Sets the hyperlink base of the document.
     *
     * @param hyperlinkBase hyperlink base of the document
     */
    public void setHyperlinkBase(String hyperlinkBase) {
        this.hyperlinkBase = hyperlinkBase;
    }

    /**
     * Gets the keywords of the workbook.
     *
     * @return keywords of the workbook
     */
    public String getKeywords() {
        return keywords;
    }

    /**
     * Sets the keywords of the workbook.
     *
     * @param keywords keywords of the workbook
     */
    public void setKeywords(String keywords) {
        this.keywords = keywords;
    }

    /**
     * Gets the responsible manager of the document.
     *
     * @return responsible manager of the document
     */
    public String getManager() {
        return manager;
    }

    /**
     * Sets the responsible manager of the document.
     *
     * @param manager responsible manager of the document
     */
    public void setManager(String manager) {
        this.manager = manager;
    }

    /**
     * Gets the subject of the workbook.
     *
     * @return subject of the workbook
     */
    public String getSubject() {
        return subject;
    }

    /**
     * Sets the subject of the workbook.
     *
     * @param subject subject of the workbook
     */
    public void setSubject(String subject) {
        this.subject = subject;
    }

    /**
     * Gets the title of the workbook.
     *
     * @return title of the workbook
     */
    public String getTitle() {
        return title;
    }

    /**
     * Sets the title of the workbook.
     *
     * @param title title of the workbook
     */
    public void setTitle(String title) {
        this.title = title;
    }

    /**
     * Converts a four-part version to the metadata-compatible format.
     *
     * @param major major number from 0 to 99999
     * @param minor minor number
     * @param build build number
     * @param revision revision number
     * @return formatted version number, such as {@code 1.0} or {@code 55.987}
     * @throws FormatException if the major number exceeds 99999 or any number is negative
     */
    public static String parseVersion(int major, int minor, int build, int revision) {
        if (major < 0 || minor < 0 || build < 0 || revision < 0) {
            throw new FormatException("The format of the passed version is wrong. No negative number allowed.");
        }
        if (major > 99999) {
            throw new FormatException("The major number may not be bigger than 99999. The passed value is " + major);
        }

        String rightPart = Integer.toString(minor) + build + revision;
        rightPart = rightPart.replaceFirst("0+$", "");
        if (rightPart.isEmpty()) {
            rightPart = "0";
        } else if (rightPart.length() > 5) {
            rightPart = rightPart.substring(0, 5);
        }
        return major + "." + rightPart;
    }

    private void checkVersion() {
        if (applicationVersion == null || applicationVersion.isEmpty()) {
            return;
        }

        String[] parts = applicationVersion.split("\\.");
        boolean valid = parts.length == 2
            && parts[0].length() >= 1
            && parts[0].length() <= 5
            && parts[1].length() >= 1
            && parts[1].length() <= 5;
        if (!valid) {
            throw new FormatException("The format of the version in the metadata is wrong (" + applicationVersion
                + "). Should be in the format and a range from '0.0' to '99999.99999'");
        }
    }

    private static String loadDefaultApplicationVersion() {
        Properties properties = new Properties();
        try (InputStream stream = Metadata.class.getResourceAsStream(VERSION_RESOURCE)) {
            if (stream == null) {
                return FALLBACK_VERSION;
            }
            properties.load(stream);
            return convertMavenVersion(properties.getProperty(VERSION_PROPERTY));
        } catch (IOException | NumberFormatException | FormatException exception) {
            return FALLBACK_VERSION;
        }
    }

    private static String convertMavenVersion(String version) {
        if (version == null) {
            return FALLBACK_VERSION;
        }

        Matcher matcher = MAVEN_VERSION_PATTERN.matcher(version);
        if (!matcher.matches()) {
            return FALLBACK_VERSION;
        }

        int major = Integer.parseInt(matcher.group(1));
        int minor = matcher.group(2) == null ? 0 : Integer.parseInt(matcher.group(2));
        int patch = matcher.group(3) == null ? 0 : Integer.parseInt(matcher.group(3));
        return parseVersion(major, minor, patch, 0);
    }
}
