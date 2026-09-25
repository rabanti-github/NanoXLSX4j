/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.internal.interfaces.writer;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.internal.interfaces.Plugin;

import java.util.List;

/**
 * Interface, used by classes to register package parts at the beginning of the writer process
 *
 * <p>Remarks: All collection entries must contain the <b>same number of elements</b>. Package part indices must be unique within a writing operation and are compared ordinally.</p>
 */
@InternalApi
public interface PluginPackageRegistry extends Plugin {

    /**
     * Gets the workbook instance, defined by the constructor
     *
     * @return Workbook instance
     */
    Workbook getWorkbook();

    /**
     * Replaces the workbook instance, defined by the constructor
     *
     * @param workbook Workbook instance
     */
    void setWorkbook(Workbook workbook);

    /**
     * Initializes the registry for the current writing operation
     *
     * @param baseWriter Base writer instance that holds the current writing context
     */
    void init(BaseWriter baseWriter);

    /**
     * List of order numbers of the package parts (for sorting purpose during registration)
     * @return Order number
     */
    List<Integer> getOrderNumbers();
    /**
     * List of relative paths of the package parts
     * @return List of package part paths
     */
    List<String> getPackagePartPaths();
    /**
     * List of the file names of the package parts
     * @return List of package part file names
     */
    List<String> getPackagePartFileNames();
    /**
     * List of the content types of the target file of the parts (usually kind of XML)
     * @return List of target types
     */
    List<String> getContentTypes();
    /**
     * List of the schema URLs of the target file of the parts (usually kind of XML schema)
     * @return List of relationship types
     */
    List<String> getRelationshipTypes();
    /**
     * List of location statement. If true, the package part is in the root directory, otherwise in the 'xl' sub-directory (with various sub-sub-directories)
     * @return List with booleans indicating whether the package parts are root
     */
    List<Boolean> getArePackagePartsRoot();
    /**
     * List of unique index indicators that can be used from {@link PluginIndexedWriter} instances to write package parts
     * @return List of indicators
     */
    List<String> getUniquePackagePartIndices();

    /**
     * List of relationships owned by each registered package part. Entries correlate positionally with the other package part collections
     * @return List of relationships
     */
    List<List<PluginPackageRelationship>> getPackagePartRelationships();
}
