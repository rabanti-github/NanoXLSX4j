//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.registry;

/**
 * Marker interface for all NanoXLSX4j plugins (both replacing and queue).
 * <p>
 * Every class that is annotated with {@link NanoXlsxPlugin} or
 * {@link NanoXlsxQueuePlugin} must also implement this interface so that the
 * Java {@link java.util.ServiceLoader} mechanism can discover it at runtime.
 * </p>
 *
 * <h3>Registration</h3>
 * To register a plugin, add a file named<br>
 * {@code META-INF/services/ch.rabanti.nanoxlsx4j.registry.NanoXlsxPluggable}<br>
 * to your JAR and list the fully-qualified class name(s) of your plugin
 * implementation(s), one per line.
 *
 * <h3>Example</h3>
 * <pre>{@code
 * @NanoXlsxPlugin(uuid = PluginUUID.SHARED_STRINGS_READER)
 * public class MySharedStringsReader implements NanoXlsxPluggable, ISharedStringsReader {
 *     // custom implementation
 * }
 * }</pre>
 *
 * @see NanoXlsxPlugin
 * @see NanoXlsxQueuePlugin
 * @see PluginLoader
 */
public interface NanoXlsxPluggable {
    // Marker interface — no methods required.
}
