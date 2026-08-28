/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.discovery;

/**
 * Describes a relationship entry or part that could not be discovered in tolerant reader mode.
 *
 * @param relationshipPartPath ZIP entry path of the affected relationship part
 * @param relationshipId       Affected relationship identifier, or null if the complete relationship part is affected
 * @param reason               Reason why discovery could not retain the relationship data
 */
public record RelationshipDiscoveryIssue(String relationshipPartPath, String relationshipId, String reason) {

    /**
     * Initializes a discovery issue.
     *
     * @param relationshipPartPath ZIP entry path of the affected relationship part
     * @param relationshipId       Affected relationship identifier, or null if the complete relationship part is
     *                             affected
     * @param reason               Reason why discovery could not retain the relationship data
     */
    public RelationshipDiscoveryIssue(String relationshipPartPath, String relationshipId, String reason) {
        this.relationshipPartPath = relationshipPartPath;
        this.relationshipId = relationshipId;
        this.reason = reason;
    }

}
