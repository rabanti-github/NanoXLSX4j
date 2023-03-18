/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.styles;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Annotation designated to control the copying of style properties
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface AppendAnnotation {

	/**
	 * Indicates whether the property annotated with the attribute is ignored during
	 * the copying of properties
	 *
	 * @return True if property is not relevant to copy styles (default false)
	 */
	boolean ignore() default false;

	/**
	 * Indicates whether the property annotated with the attribute is a nested
	 * property. Nested properties are ignored during the copying of properties but
	 * can be broken down to its sub-properties
	 *
	 * @return True if the style property is nested (default false)
	 */
	boolean nestedProperty() default false;

}
