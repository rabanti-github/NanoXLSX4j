/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.discovery;

/**
 * Describes one relationship discovered in an OOXML package.
 *
 * @param id                   Relationship identifier
 * @param type                 Relationship type URI
 * @param target               Unmodified relationship target
 * @param targetMode           Internal or external resource target
 * @param relationshipPartPath ZIP entry path of the relationship part
 * @param sourcePartPath       Normalized path of the source part
 * @param resolvedTargetPath   Normalized ZIP entry path of an internal target, or null for an external target
 */
public record RelationshipInfo(
        String id,
        String type,
        String target,
        TargetMode targetMode,
        String relationshipPartPath,
        String sourcePartPath,
        String resolvedTargetPath
) {

    /**
     * Enum describing the target mode of the XML document
     */
    public enum TargetMode {
        /**
         * Target is internal
         */
        INTERNAL,
        /**
         * Target is external
         */
        EXTERNAL
    }

    /**
     * Initializes a relationship description.
     *
     * @param id                   Relationship identifier
     * @param type                 Relationship type URI
     * @param target               Unmodified relationship target
     * @param targetMode           Internal or external resource target
     * @param relationshipPartPath ZIP entry path of the relationship part
     * @param sourcePartPath       Normalized path of the source part
     * @param resolvedTargetPath   Normalized ZIP entry path of an internal target, or null for an external target
     */
    public RelationshipInfo(
            String id, String type, String target, TargetMode targetMode, String relationshipPartPath,
            String sourcePartPath, String resolvedTargetPath
    ) {
        this.id = id;
        this.type = type;
        this.target = target;
        this.targetMode = targetMode;
        this.relationshipPartPath = relationshipPartPath;
        this.sourcePartPath = sourcePartPath;
        this.resolvedTargetPath = resolvedTargetPath;
    }

}
