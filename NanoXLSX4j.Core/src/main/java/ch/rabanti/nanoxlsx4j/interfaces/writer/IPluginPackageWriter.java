//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces.writer;

/**
 * Interface for plug-ins that register and write additional OOXML package parts
 * during the XLSX creation process.
 * <p>
 * Java equivalent of the {@code internal} .NET interface
 * {@code NanoXLSX.Interfaces.Writer.IPluginPackageWriter}.
 * </p>
 *
 * @implNote Framework-internal interface. Not intended for use by external plugin implementors.
 */
public interface IPluginPackageWriter extends IPluginWriter {

    /**
     * Returns the sort order used when registering this package part
     * (lower values are registered first).
     *
     * @return the order number
     */
    int getOrderNumber();

    /**
     * Returns the relative path of the OOXML package part within the ZIP archive
     * (e.g. {@code "xl/worksheets/"}).
     *
     * @return the relative package-part path
     */
    String getPackagePartPath();

    /**
     * Returns the file name of the package part
     * (e.g. {@code "sheet1.xml"}).
     *
     * @return the package-part file name
     */
    String getPackagePartFileName();

    /**
     * Returns the OOXML content type of the package part
     * (e.g. {@code "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet+xml"}).
     *
     * @return the content-type string
     */
    String getContentType();

    /**
     * Returns the relationship type URI for this package part.
     *
     * @return the relationship type string
     */
    String getRelationshipType();

    /**
     * Returns whether this package part resides in the root directory of the ZIP archive.
     * If {@code false}, the part is placed in the {@code xl/} sub-directory hierarchy.
     *
     * @return {@code true} if this is a root-level package part
     */
    boolean isRootPackagePart();
}
