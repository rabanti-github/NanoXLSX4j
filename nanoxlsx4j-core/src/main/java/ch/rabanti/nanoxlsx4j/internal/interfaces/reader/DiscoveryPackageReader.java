/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.reader;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.internal.discovery.RelationshipInfo;

/**
 * Defines a queued reader that is dispatched to package parts by their discovered OOXML relationship type.
 *
 * <p>Remarks: <p>
 * The package relationship discovery step runs before this reader. For every internal relationship whose
 * {@link DocumentReader#getDocumentType()} exactly matches the relationship type and whose target exists in the
 * package, the reader is initialized with a fresh stream of that target and then executed once.
 * {@link #getCurrentRelationshipInfo()} identifies the relationship for the current execution.
 * </p>
 * <p>
 * This contract is intended for relationship targets with non-fixed names. For example, a reader whose document type is
 * {@code http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink} is dispatched to
 * {@code xl/externalLinks/externalLink1.xml}, {@code externalLink2.xml}, and every other matching internal target
 * discovered in the package.
 * </p>
 * <p>
 * This reader does not perform relationship discovery and does not receive the corresponding {@code *.rels} stream.
 * Relationship parts have already been parsed into the discovery catalog. The inherited
 * {@link PluginPackageReader#getStreamEntryName()} is not used for discovery-based dispatch and should return <see
 * langword="null"/>.
 * </p>
 * </p>
 */
@InternalApi
public interface DiscoveryPackageReader extends PluginPackageReader, DocumentReader {

    /**
     * Gets the discovered relationship whose resolved target stream is supplied for the current execution.
     *
     * @return Current relationship info
     */
    RelationshipInfo getCurrentRelationshipInfo();

    /**
     * Sets the discovered relationship whose resolved target stream is supplied for the current execution.
     *
     * @param relationshipInfo Current relationship info
     */
    void setCurrentRelationshipInfo(RelationshipInfo relationshipInfo);

}
