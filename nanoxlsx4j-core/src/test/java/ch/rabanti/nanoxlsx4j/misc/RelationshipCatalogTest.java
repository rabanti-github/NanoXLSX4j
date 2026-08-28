/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.misc;

import ch.rabanti.nanoxlsx4j.internal.discovery.RelationshipCatalog;
import ch.rabanti.nanoxlsx4j.internal.discovery.RelationshipDiscoveryIssue;
import ch.rabanti.nanoxlsx4j.internal.discovery.RelationshipInfo;
import ch.rabanti.nanoxlsx4j.internal.discovery.RelationshipInfo.TargetMode;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.Locale;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class RelationshipCatalogTest {

    private static final String WORKSHEET_TYPE =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    @Test
    @DisplayName("Relationship IDs are scoped to their source part")
    public void sourceScopedRelationshipIdTest() {
        RelationshipCatalog catalog = new RelationshipCatalog();
        RelationshipInfo workbookRelationship = createRelationship(
            "rId1", WORKSHEET_TYPE, "xl/workbook.xml", "xl/worksheets/sheet1.xml");
        RelationshipInfo worksheetRelationship = createRelationship(
            "rId1", "http://example.org/drawing", "xl/worksheets/sheet1.xml", "xl/drawings/drawing1.xml");

        assertTrue(catalog.tryAdd(workbookRelationship));
        assertTrue(catalog.tryAdd(worksheetRelationship));
        assertEquals(2, catalog.getRelationships().size());
        assertSame(workbookRelationship, catalog.getBySourceAndId("xl/workbook.xml", "rId1"));
        assertSame(worksheetRelationship,
            catalog.getBySourceAndId("xl/worksheets/sheet1.xml", "rId1"));
    }

    @Test
    @DisplayName("Duplicate relationship IDs retain the first source-local entry")
    public void duplicateRelationshipIdTest() {
        RelationshipCatalog catalog = new RelationshipCatalog();
        RelationshipInfo first = createRelationship(
            "rId1", WORKSHEET_TYPE, "xl/workbook.xml", "xl/worksheets/sheet1.xml");
        RelationshipInfo duplicate = createRelationship(
            "rId1", WORKSHEET_TYPE, "xl/workbook.xml", "xl/worksheets/sheet2.xml");

        assertTrue(catalog.tryAdd(first));
        assertFalse(catalog.tryAdd(duplicate));
        assertEquals(1, catalog.getRelationships().size());
        assertSame(first, catalog.getBySourceAndId("xl/workbook.xml", "rId1"));
    }

    @Test
    @DisplayName("Relationship type lookup is ordinal and case-sensitive")
    public void exactDocumentTypeLookupTest() {
        RelationshipCatalog catalog = new RelationshipCatalog();
        RelationshipInfo relationship = createRelationship(
            "rId1", WORKSHEET_TYPE, "xl/workbook.xml", "xl/worksheets/sheet1.xml");
        catalog.tryAdd(relationship);

        assertEquals(1, catalog.getByType(WORKSHEET_TYPE).size());
        assertTrue(catalog.getByType(WORKSHEET_TYPE.toUpperCase(Locale.ROOT)).isEmpty());
    }

    @Test
    @DisplayName("Discovery issues mark a relationship catalog as incomplete")
    public void discoveryIssueTest() {
        RelationshipCatalog catalog = new RelationshipCatalog();
        assertTrue(catalog.isComplete());

        RelationshipDiscoveryIssue issue = new RelationshipDiscoveryIssue(
            "xl/_rels/workbook.xml.rels", "rId1", "Duplicate relationship identifier");
        catalog.addIssue(issue);

        assertFalse(catalog.isComplete());
        assertEquals(1, catalog.getOnlyIssues().size());
        assertSame(issue, catalog.getOnlyIssues().get(0));
    }

    private static RelationshipInfo createRelationship(
            String id, String type, String sourcePartPath, String targetPath) {
        return new RelationshipInfo(
            id,
            type,
            targetPath,
            TargetMode.INTERNAL,
            "xl/_rels/workbook.xml.rels",
            sourcePartPath,
            targetPath);
    }
}
