/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.internal.structures;

/**
 * Class to manage XML document paths.
 * <p>
 * Java equivalent of the .NET internal class {@code NanoXLSX.Internal.Structures.DocumentPath}.
 * </p>
 *
 * @implNote Framework-internal class. Not intended for use outside the writer module.
 * @author Raphael Stoeckli
 */
public class DocumentPath {

    // ### P R I V A T E  F I E L D S ###

    private String filename;
    private String path;

    // ### G E T T E R S  &  S E T T E R S ###

    /**
     * Gets the file name of the document
     *
     * @return file name
     */
    public String getFilename() {
        return filename;
    }

    /**
     * Sets the file name of the document
     *
     * @param filename file name to set
     */
    public void setFilename(String filename) {
        this.filename = filename;
    }

    /**
     * Gets the path of the document
     *
     * @return path
     */
    public String getPath() {
        return path;
    }

    /**
     * Sets the path of the document
     *
     * @param path path to set
     */
    public void setPath(String path) {
        this.path = path;
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Constructor with defined file name and path
     *
     * @param filename File name of the document
     * @param path     Path of the document
     */
    public DocumentPath(String filename, String path) {
        this.filename = filename;
        this.path = path;
    }

    // ### M E T H O D S ###

    /**
     * Returns the full path of the document, using the forward slash as directory separator (XLSX convention).
     * A leading slash is always prepended.
     *
     * @return Full path with leading slash
     */
    public String getFullPath() {
        char last = path.charAt(path.length() - 1);
        if (last == '/' || last == '\\') {
            return '/' + path + filename;
        }
        return '/' + path + '/' + filename;
    }

}
