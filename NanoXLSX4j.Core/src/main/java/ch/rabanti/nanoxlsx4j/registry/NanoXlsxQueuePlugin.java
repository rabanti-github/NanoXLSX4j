//STATUS: VERIFY
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
 * Marks a class as a <strong>Queue Plugin</strong> for NanoXLSX4j.
 * <p>
 * Queue plugins are executed sequentially before or after a standard reader /
 * writer operation. Unlike {@link NanoXlsxPlugin replacing plugins}, multiple
 * queue plugins can be active for the same UUID at the same time. They are
 * invoked in ascending {@link #order()} sequence via
 * {@link PluginLoader#getNextQueuePlugin}.
 * </p>
 *
 * <h3>Usage</h3>
 * <ol>
 *   <li>Annotate your class with {@code @NanoXlsxQueuePlugin} and supply the
 *       target queue UUID from {@link PluginUUID} (e.g.
 *       {@link PluginUUID#WRITER_PREPEND_QUEUE} or one of the inline-queue
 *       constants).</li>
 *   <li>Implement {@link NanoXlsxPluggable} (the marker interface required by
 *       the {@link java.util.ServiceLoader} discovery mechanism).</li>
 *   <li>Register the class in
 *       {@code META-INF/services/ch.rabanti.nanoxlsx4j.registry.NanoXlsxPluggable}
 *       inside your JAR.</li>
 * </ol>
 *
 * @see NanoXlsxPlugin
 * @see PluginLoader
 * @see PluginUUID
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface NanoXlsxQueuePlugin {

    /**
     * The UUID of the queue this plugin participates in.
     * Use one of the constants defined in {@link PluginUUID}.
     *
     * @return the target queue UUID
     */
    String plugInUuid();

    /**
     * The queue UUID for plug-ins that are not replacing a specific base plug-in, but defined as additional resource, e.g. executed before or after the writer / reader base plug-ins
     * Use one of the queue constants defined in {@link PluginUUID}.
     *
     * @return the target queue UUID
     */
    String queueUuid();


    /**
     * Execution order within the queue. Plugins are invoked in
     * <strong>ascending</strong> order (lower value = earlier execution).
     * Default is {@code 0}.
     *
     * @return the execution order
     */
    int order() default 0;
}
