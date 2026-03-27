//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
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
 * Marks a class as a <strong>Replacing Plugin</strong> for NanoXLSX4j.
 * <p>
 * A replacing plugin substitutes the default implementation that is registered
 * under the given {@link #plugInUuid()}. If multiple replacing plugins declare the
 * same UUID, the one with the highest {@link #order()} value wins. Ties are
 * resolved arbitrarily.
 * </p>
 *
 * <h3>Usage</h3>
 * <ol>
 *   <li>Annotate your class with {@code @NanoXlsxPlugin} and supply the target UUID
 *       from {@link PluginUUID}.</li>
 *   <li>Implement the corresponding reader/writer interface from
 *       {@code ch.rabanti.nanoxlsx4j.registry}.</li>
 *   <li>Implement {@link NanoXlsxPluggable} (the marker interface required by
 *       the {@link java.util.ServiceLoader} discovery mechanism).</li>
 *   <li>Register the class in
 *       {@code META-INF/services/ch.rabanti.nanoxlsx4j.registry.NanoXlsxPluggable}
 *       inside your JAR.</li>
 * </ol>
 *
 * @see NanoXlsxQueuePlugin
 * @see PluginLoader
 * @see PluginUUID
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface NanoXlsxPlugin {

    /**
     * The UUID of the plugin slot this class replaces.
     * Use one of the constants defined in {@link PluginUUID}.
     *
     * @return the target plugin UUID
     */
    String plugInUuid();

    /**
     * Execution priority used to resolve conflicts when multiple replacing plugins
     * declare the same UUID. The plugin with the <strong>highest</strong> order
     * value is selected. Default is {@code 0}.
     *
     * @return the priority order
     */
    int order() default 0;
}
