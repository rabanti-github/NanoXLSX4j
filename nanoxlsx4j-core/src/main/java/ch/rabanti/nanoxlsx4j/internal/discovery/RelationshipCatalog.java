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

public class RelationshipCatalog {

    private static final List<RelationshipInfo> relationships = new ArrayList<>();
    private static final List<RelationshipDiscoveryIssue> issues = new ArrayList<>();
    private static final Map<String, Map<String, RelationshipInfo>> relationshipsBySource = new HashMap<>(); //
    // string comparer?
    private final List<RelationshipInfo> readOnlyRelationships;
    private final List<RelationshipDiscoveryIssue> readOnlyIssues;

    /// <summary>
    /// Initializes an empty relationship catalog.
    /// </summary>
    public RelationshipCatalog() {
        readOnlyRelationships = Collections.unmodifiableList(relationships);
        readOnlyIssues = Collections.unmodifiableList(issues);
    }
}
