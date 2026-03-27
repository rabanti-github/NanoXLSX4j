/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.internal.structures;

import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Class to manage package parts of a XLSX file to be written.
 * <p>
 * Java equivalent of the .NET internal class {@code NanoXLSX.Internal.Structures.PackagePartDefinition}.
 * </p>
 *
 * @implNote Framework-internal class. Not intended for use outside the writer module.
 * @author Raphael Stoeckli
 */
public class PackagePartDefinition {

    // ### S T A T I C  F I E L D S ###

    /**
     * Package part definition index (for sorting), designated to the workbook
     */
    public static final int WORKBOOK_PACKAGE_PART_INDEX = 0;

    /**
     * Package part definition start index (for sorting), designated for metadata and other root parts
     */
    public static final int METADATA_PACKAGE_PART_START_INDEX = 1000;

    /**
     * Package part definition start index (for sorting), designated for worksheet parts.
     * These numbers shall not be used for other instances until {@link #POST_WORKSHEET_PACKAGE_PART_START_INDEX}
     */
    public static final int WORKSHEET_PACKAGE_PART_START_INDEX = 10000;

    /**
     * Package part definition start index (for sorting), designated for other non-root package parts
     */
    public static final int POST_WORKSHEET_PACKAGE_PART_START_INDEX = 2000000;

    // ### E N U M S ###

    /**
     * Enum to define the type of a package part
     */
    public enum PackagePartType {
        /**
         * Package part is a root part (e.g. workbook, metadata)
         */
        ROOT,
        /**
         * Package part is a worksheet
         */
        WORKSHEET,
        /**
         * Package part is a non-root part (e.g. style, sharedStrings)
         */
        OTHER
    }

    // ### P R I V A T E  F I E L D S ###

    private final DocumentPath path;
    private final PackagePartType partType;
    private final int orderNumber;
    private final String contentType;
    private final String relationshipType;

    // ### G E T T E R S ###

    /**
     * Gets the document path of the package part
     *
     * @return document path
     */
    public DocumentPath getPath() {
        return path;
    }

    /**
     * Gets the type of the package part
     *
     * @return package part type
     */
    public PackagePartType getPartType() {
        return partType;
    }

    /**
     * Gets the order number during registration.
     * The order number determines the rID assignment for the XML part.
     *
     * @return order number
     */
    public int getOrderNumber() {
        return orderNumber;
    }

    /**
     * Gets the content type of the target file of the part
     *
     * @return content type string
     */
    public String getContentType() {
        return contentType;
    }

    /**
     * Gets the schema URL of the target file of the part
     *
     * @return relationship type string
     */
    public String getRelationshipType() {
        return relationshipType;
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Constructor with all fields
     *
     * @param type                Type of the package part
     * @param orderNumber         Order number during registration
     * @param fileNameInPackage   Relative file name of the target file, without path
     * @param pathInPackage       Relative path to the file within the package
     * @param contentType         Content type of the target file
     * @param relationshipType    Schema URL of the target file
     */
    public PackagePartDefinition(PackagePartType type, int orderNumber, String fileNameInPackage,
                                 String pathInPackage, String contentType, String relationshipType) {
        this(type, orderNumber, new DocumentPath(fileNameInPackage, pathInPackage), contentType, relationshipType);
    }

    /**
     * Constructor with a pre-built document path
     *
     * @param type             Type of the package part
     * @param orderNumber      Order number during registration
     * @param documentPath     Document path with all relevant file and path information
     * @param contentType      Content type of the target file
     * @param relationshipType Schema URL of the target file
     */
    public PackagePartDefinition(PackagePartType type, int orderNumber, DocumentPath documentPath,
                                 String contentType, String relationshipType) {
        this.partType = type;
        this.orderNumber = orderNumber;
        this.path = documentPath;
        this.contentType = contentType;
        this.relationshipType = relationshipType;
    }

    // ### M E T H O D S ###

    /**
     * Returns the zero-based index of the worksheet represented by this package part.
     *
     * @return Zero-based worksheet index
     * @implNote This method should only be called if {@link #getPartType()} is {@link PackagePartType#WORKSHEET}.
     * It will return invalid indices otherwise.
     */
    public int getWorksheetIndex() {
        return this.orderNumber - WORKSHEET_PACKAGE_PART_START_INDEX;
    }

    /**
     * Sorts a list of package part definitions by their order number (ascending)
     *
     * @param packagePartDefinitions List to sort
     * @return New sorted list
     */
    public static List<PackagePartDefinition> sort(List<PackagePartDefinition> packagePartDefinitions) {
        return packagePartDefinitions.stream()
                .sorted(Comparator.comparingInt(PackagePartDefinition::getOrderNumber))
                .collect(Collectors.toList());
    }

}
