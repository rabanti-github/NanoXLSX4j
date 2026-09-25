/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces.writer;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.internal.discovery.RelationshipInfo;

/**
 * Interface used by package registry plug-ins to define a relationship (.rels) owned by a registered package part
 */
@InternalApi
public interface PluginPackageRelationship {
    /**
     * Gets the XML identifier of the relationship (rId). The identifier must be unique within the owning package part
     *
     * @return Relationship ID (rId)
     */
    String getRelationshipId();

    /**
     * Gets the absolute URI that identifies the role of the relationship
     *
     * @return relationship types
     */
    String getRelationshipType();

    /**
     * Gets the URI of the relationship target
     *
     * @return Relationship target
     */
    String getTarget();

    /**
     * Gets whether the target is internal or external to the package
     *
     * @return Target mode
     */
    RelationshipInfo.TargetMode getTargetMode();
}
