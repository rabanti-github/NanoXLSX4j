/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Objects;
import java.util.Optional;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

/**
 * Class represents an abstract style component
 */
public abstract class AbstractStyle implements Comparable<AbstractStyle> {

    @AppendAnnotation(ignore = true)
    private Optional<Integer> internalId = Optional.empty();

    /**
     * Gets the internal ID for sorting purpose in the Excel style document (nullable)
     *
     * @return Internal ID
     */
    public Optional<Integer> getInternalId() {
        return internalId;
    }

    /**
     * Sets the internal ID for sorting purpose in the Excel style document (nullable)
     *
     * @param internalId Internal ID
     */
    public void setInternalId(Optional<Integer> internalId) {
        this.internalId = Objects.requireNonNull(internalId, "internalId");
    }

    /**
     * Abstract method to copy a component (dereferencing)
     *
     * @return Returns a copied component
     */
    public abstract AbstractStyle copy();

    /**
     * Internal method to copy altered fields from a source object. The decision whether a field is copied is dependent
     * on an untouched reference object
     *
     * @param source    Style or subclass of Style that extends AbstractStyle
     * @param reference Source object with properties to copy
     * @param <T>       Reference object to decide whether the fields from the source objects are altered or not
     */
    <T extends AbstractStyle> void copyFields(T source, T reference) {
        if (source == null ||
                !this.getClass().equals(source.getClass()) && !this.getClass().equals(reference.getClass())) {
            throw new StyleException(
                    "The objects of the source, target and reference for style appending are not of the same type");
        }
        boolean ignore;
        Field[] infos = this.getClass().getDeclaredFields();
        Field sourceInfo;
        Field referenceInfo;
        Annotation[] annotations;
        try {
            for (Field info : infos) {
                annotations = info.getDeclaredAnnotationsByType(AppendAnnotation.class);
                if (annotations.length > 0) {
                    ignore = false;
                    for (Annotation annotation : annotations) {
                        if (((AppendAnnotation) annotation).ignore() ||
                                ((AppendAnnotation) annotation).nestedProperty()) {
                            ignore = true;
                            break;
                        }
                    }
                    if (ignore) {
                        continue;
                    } // skip field
                }
                sourceInfo = source.getClass().getDeclaredField(info.getName());
                sourceInfo.setAccessible(true); // Necessary to access private field
                referenceInfo = reference.getClass().getDeclaredField(info.getName());
                referenceInfo.setAccessible(true); // Necessary to access private field
                if (!sourceInfo.get(source).equals(referenceInfo.get(reference))) {
                    info.setAccessible(true); // Necessary to access private field
                    info.set(this, sourceInfo.get(source));
                }
            }
        } catch (Exception ex) {
            throw new StyleException(
                    "The field of the source object could not be copied to the target object: " + ex.getMessage());
        }
    }

    /**
     * Method to compare two objects for sorting purpose
     *
     * @param o Other object to compare with this object
     * @return True if both objects are equal, otherwise false
     */
    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        AbstractStyle that = (AbstractStyle) o;
        return Objects.equals(this.hashCode(), that.hashCode());
    }

    /**
     * Method to compare two objects for sorting purpose
     *
     * @param o Other object to compare with this object
     * @return -1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.
     */
    @Override
    public int compareTo(AbstractStyle o) {
        if (this.internalId.isEmpty()) {
            return -1;
        } else if (o == null || o.getInternalId().isEmpty()) {
            return 1;
        } else {
            return this.internalId.get().compareTo(o.getInternalId().get());
        }
    }

    /**
     * Append a JSON property for debug purpose (used in the ToString methods) to the passed string builder
     *
     * @param sb        String builder
     * @param name      Property name
     * @param value     Property value
     * @param terminate If true, no comma and newline will be appended
     */
    static void addPropertyAsJson(StringBuilder sb, String name, Object value, boolean terminate) {
        sb.append("\"").append(name).append("\": ");
        if (value == null) {
            sb.append("__null__");
        } else {
            sb.append("\"").append(value.toString().replace("\"", "\\\"")).append("\"");
        }
        if (!terminate) {
            sb.append(",\n");
        }
    }

}
