/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2019
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;

/**
 * Class represents an abstract style component
 *
 * @author Raphael Stoeckli
 */
public abstract class AbstractStyle implements Comparable<AbstractStyle> {
// ### P R I V A T E  F I E L D S ###

    @AppendAnnotation(ignore = true)
    private Integer internalID = null;

// ### G E T T E R S  &  S E T T E R S ###

    /**
     * Gets the internal ID for sorting purpose in the Excel style document
     *
     * @return Internal ID (can be null for random order)
     */
    public Integer getInternalID() {
        return internalID;
    }

    /**
     * Sets the internal ID for sorting purpose in the Excel style document
     *
     * @param internalID Internal ID (can be null for random order)
     */
    public void setInternalID(Integer internalID) {
        this.internalID = internalID;
    }


    // ### M E T H O D S ###
    /**
     * Abstract method definition to calculate the hash of the component
     * @return Returns the hash of the component as string
     */
    //abstract String calculateHash();


    /**
     * Abstract method to copy a component (dereferencing)
     *
     * @return Returns a copied component
     */
    public abstract AbstractStyle copy();

    /**
     * Method to compare two objects for sorting purpose
     *
     * @param other Other object to compare with this object
     * @return True if both objects are equal, otherwise false
     */
    public boolean equals(AbstractStyle other) {
        return this.hashCode() == other.hashCode();
    }

    /**
     * Method to compare two objects for sorting purpose
     *
     * @param o Other object to compare with this object
     * @return -1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.
     */
    @Override
    public int compareTo(AbstractStyle o) {
        // return this.hash.compareTo(o.getHash());
        if (this.internalID == null) {
            return -1;
        } else if (o.getInternalID() == null) {
            return 1;
        } else {
            return this.internalID.compareTo(o.getInternalID());
        }
    }

    /**
     * Method to cast values of the components to string values for the hash calculation
     *
     * @param o         Value to cast
     * @param sb        StringBuilder reference to put the casted object in
     * @param delimiter Delimiter character to append after the casted value
     */
    static void castValue(Object o, StringBuilder sb, Character delimiter) {
        if (o == null) {
            sb.append('#');
        } else if (o instanceof Boolean) {
            if ((boolean) o == true) {
                sb.append(1);
            } else {
                sb.append(0);
            }
        } else if (o instanceof Integer) {
            sb.append((int) o);
        } else if (o instanceof Double) {
            sb.append((double) o);
        } else if (o instanceof Float) {
            sb.append((float) o);
        } else if (o instanceof String) {
            if (o.toString().equals("#")) {
                sb.append("_#_");
            } else {
                sb.append((String) o);
            }
        } else if (o instanceof Long) {
            sb.append((long) o);
        } else if (o instanceof Character) {
            sb.append((char) o);
        } else {
            sb.append(o);
        }
        if (delimiter != null) {
            sb.append(delimiter);
        }
    }

    /**
     * Internal method to copy altered fields from a source object. The decision whether a field is copied is dependent on a untouched reference object
     *
     * @param source    Style or sub-class of Style that extends AbstractStyle
     * @param reference Source object with properties to copy
     * @param <T>       Reference object to decide whether the fields from the source objects are altered or not
     */
    <T extends AbstractStyle> void copyFields(T source, T reference) {
        if (this.getClass().equals(source.getClass()) == false && this.getClass().equals(reference.getClass()) == false) {
            throw new StyleException("CopyFieldException", "The objects of the source, target and reference for style appending are not of the same type");
        }
        boolean ignore;
        Field[] infos = this.getClass().getDeclaredFields();
        Field sourceInfo, referenceInfo;
        Annotation[] annotations;
        try {
            for (Field info : infos) {
                annotations = info.getDeclaredAnnotationsByType(AppendAnnotation.class);
                if (annotations.length > 0) {
                    ignore = false;
                    for (Annotation annotation : annotations) {
                        if (((AppendAnnotation) annotation).ignore() == true || ((AppendAnnotation) annotation).nestedProperty() == true) {
                            ignore = true;
                            break;
                        }
                    }
                    if (ignore == true) {
                        continue;
                    } // skip field
                }
                sourceInfo = source.getClass().getDeclaredField(info.getName());
                sourceInfo.setAccessible(true); // Necessary to access private field
                referenceInfo = reference.getClass().getDeclaredField(info.getName());
                referenceInfo.setAccessible(true); // Necessary to access private field
                if (sourceInfo.get(source).equals(referenceInfo.get(reference)) == false) {
                    info.setAccessible(true); // Necessary to access private field
                    info.set(this, sourceInfo.get(source));
                }
            }
        } catch (Exception ex) {
            throw new StyleException("CopyFieldException", "The field of the source object could not be copied to the target object: " + ex.getMessage());
        }
    }

}
