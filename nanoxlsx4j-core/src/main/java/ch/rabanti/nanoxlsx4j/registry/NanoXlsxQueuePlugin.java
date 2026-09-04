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
import java.lang.annotation.Repeatable;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Declares a service provider as a queued NanoXLSX4j plug-in.
 */
@Documented
@Repeatable(NanoXlsxQueuePlugins.class)
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface NanoXlsxQueuePlugin {

    /**
     * Gets the UUID of the queued plug-in.
     *
     * @return Plug-in UUID
     */
    String pluginUuid();

    /**
     * Gets the UUID of the queue receiving the plug-in.
     *
     * @return Queue UUID
     */
    String queueUuid();

    /**
     * Gets the execution order. Lower values are returned first.
     *
     * @return Plug-in order, defaulting to zero
     */
    int order() default 0;
}
