/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.registry;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Declares a service provider as a replacement NanoXLSX4j plug-in.
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface NanoXlsxPlugin {

    /**
     * Gets the UUID of the component replaced by the plug-in.
     *
     * @return Replacement UUID
     */
    String pluginUuid();

    /**
     * Gets the precedence of the plug-in. Higher values take precedence.
     *
     * @return Plug-in order, defaulting to zero
     */
    int order() default 0;
}
