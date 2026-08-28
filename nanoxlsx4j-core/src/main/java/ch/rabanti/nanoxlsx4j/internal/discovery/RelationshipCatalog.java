/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.discovery;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public class RelationshipCatalog {

    private final List<RelationshipInfo> relationships = new ArrayList<>();
    private final List<RelationshipDiscoveryIssue> issues = new ArrayList<>();
    private final Map<String, Map<String, RelationshipInfo>> relationshipsBySource = new HashMap<>();
    private final List<RelationshipInfo> readOnlyRelationships;
    private final List<RelationshipDiscoveryIssue> readOnlyIssues;

    /**
     * Gets all successfully discovered relationships in deterministic discovery order.
     *
     * @return List of discovered successful relationships
     */
    public List<RelationshipInfo> getRelationships() {
        return readOnlyRelationships;
    }

    /**
     * Gets all issues recorded while discovery was operating in tolerant mode.
     *
     * @return List of Relationships with issues
     */
    public List<RelationshipDiscoveryIssue> getOnlyIssues() {
        return readOnlyIssues;
    }

    /**
     * Gets whether discovery completed without skipping an invalid relationship entry or part.
     *
     * @return True if the discovery was finished without problems
     */
    public boolean isComplete() {
        return issues.isEmpty();
    }

    /**
     * Initializes an empty relationship catalog.
     */
    public RelationshipCatalog() {
        readOnlyRelationships = Collections.unmodifiableList(relationships);
        readOnlyIssues = Collections.unmodifiableList(issues);
    }

    /**
     * Adds a relationship unless the same source part already contains its identifier.
     *
     * @param relationship Relationship info object (cannot be null)
     * @return True if the relationship was added; otherwise false.
     */
    public boolean tryAdd(RelationshipInfo relationship) {
        String sourcePartPath = "";
        if (relationship.sourcePartPath() != null) {
            sourcePartPath = relationship.sourcePartPath();
        }

        Map<String, RelationshipInfo> sourceRelationships;
        if (relationshipsBySource.containsKey(sourcePartPath)) {
            sourceRelationships = relationshipsBySource.get(sourcePartPath);
        } else {
            sourceRelationships = new HashMap<>();
            relationshipsBySource.put(sourcePartPath, sourceRelationships);
        }

        if (sourceRelationships.containsKey(relationship.id())) {
            return false;
        }
        sourceRelationships.put(relationship.id(), relationship);
        relationships.add(relationship);
        return true;
    }

    /**
     * Adds an issue encountered during tolerant discovery.
     *
     * @param issue Issue object (cannot be null)
     */
    public void addIssue(RelationshipDiscoveryIssue issue) {
        issues.add(issue);
    }

    /// <summary>
    /// Gets a relationship by its source part and source-local identifier.
    /// </summary>
    /// <param name="relationshipId">rID of the relationship</param>
    /// <param name="sourcePartPath">URI of the source path</param>
    /// <returns>Returns the relationship info object</returns>
    public RelationshipInfo getBySourceAndId(String sourcePartPath, String relationshipId) {
        String normalizedSourcePartPath = "";
        if (sourcePartPath != null) {
            normalizedSourcePartPath = sourcePartPath;
        }
        if (relationshipsBySource.containsKey(normalizedSourcePartPath) && relationshipId != null) {
            Map<String, RelationshipInfo> sourceRelationships = relationshipsBySource.get(normalizedSourcePartPath);
            if (sourceRelationships.containsKey(relationshipId)) {
                return sourceRelationships.get(relationshipId);
            }
        }
        return null;
    }

    /// <summary>
    /// Gets all relationships whose type exactly matches the supplied URI.
    /// </summary>
    /// <returns>Read-only list of relationships by type</returns>
    public List<RelationshipInfo> getByType(String documentType) {
        List<RelationshipInfo> matches = new ArrayList<>();
        for (RelationshipInfo relationship : relationships) {
            if (Objects.equals(relationship.type(), documentType)) {
                matches.add(relationship);
            }
        }
        return Collections.unmodifiableList(matches);
    }

}
