/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.registry;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.InputStream;
import java.net.URL;
import java.net.URLClassLoader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.jar.JarEntry;
import java.util.jar.JarOutputStream;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import ch.rabanti.nanoxlsx4j.exceptions.PluginLoadingException;
import ch.rabanti.nanoxlsx4j.internal.registry.PluginLoader.QueuePlugin;
import ch.rabanti.nanoxlsx4j.registry.NanoXlsxPlugin;
import ch.rabanti.nanoxlsx4j.registry.NanoXlsxQueuePlugin;
import ch.rabanti.nanoxlsx4j.registry.PluginProvider;

class PluginLoaderTest {

    private static final String REPLACEMENT_UUID = "TEST_REPLACEMENT";
    private static final String QUEUE_UUID = "TEST_QUEUE";

    @AfterEach
    void dispose() {
        PluginLoader.resetForTesting();
    }

    @Test
    @DisplayName("Test of the plug-in handling initializer (dummy; should not crash or initialize twice)")
    void initializeTest() {
        assertTrue(PluginLoader.initializePluginsForTesting(List.of()));
        assertFalse(PluginLoader.initializePluginsForTesting(List.of(ReplacementPlugin.class)));
    }

    @Test
    @DisplayName("Plug-in initialization is thread-safe and occurs only once")
    void initializeIsThreadSafeTest() throws Exception {
        int threadCount = 16;
        CountDownLatch ready = new CountDownLatch(threadCount);
        CountDownLatch start = new CountDownLatch(1);
        ExecutorService executor = Executors.newFixedThreadPool(threadCount);
        try {
            List<Future<Boolean>> results = new ArrayList<>();
            for (int i = 0; i < threadCount; i++) {
                results.add(executor.submit(() -> {
                    ready.countDown();
                    start.await();
                    return PluginLoader.initializePluginsForTesting(List.of(ReplacementPlugin.class));
                }));
            }
            ready.await();
            start.countDown();

            int initializedCount = 0;
            for (Future<Boolean> result : results) {
                if (result.get()) {
                    initializedCount++;
                }
            }
            assertEquals(1, initializedCount);
        } finally {
            executor.shutdownNow();
        }
    }

    @Test
    @DisplayName("Plug-ins are discovered from a ServiceLoader provider JAR")
    void discoversPluginWithServiceLoaderTest(@TempDir Path tempDirectory) throws Exception {
        Path providerJar = createProviderJar(tempDirectory);
        try (URLClassLoader classLoader = new ProviderClassLoader(
                new URL[] {providerJar.toUri().toURL()}, getClass().getClassLoader())) {
            assertTrue(PluginLoader.initialize(classLoader));

            PluginProvider plugin = PluginLoader.getPlugin("SERVICE_PLUGIN", PluginProvider.class, null);

            assertEquals(ServiceDiscoveredPlugin.class.getName(), plugin.getClass().getName());
            assertSame(classLoader, plugin.getClass().getClassLoader());
        }
    }

    @Test
    @DisplayName("A malformed ServiceLoader declaration fails without initializing the registry")
    void malformedServiceDeclarationFailsTest(@TempDir Path tempDirectory) throws Exception {
        Path providerJar = createServiceJar(tempDirectory, "missing.Provider");
        try (URLClassLoader classLoader = new URLClassLoader(
                new URL[] {providerJar.toUri().toURL()}, getClass().getClassLoader())) {
            assertThrows(PluginLoadingException.class, () -> PluginLoader.initialize(classLoader));
        }

        assertTrue(PluginLoader.initializePluginsForTesting(List.of()));
    }

    @Test
    @DisplayName("A missing replacement returns the supplied fallback instance")
    void returnsFallbackForMissingReplacementTest() {
        PluginLoader.initializePluginsForTesting(List.of());
        PluginProvider fallback = new PlainPlugin();

        assertSame(fallback, PluginLoader.getPlugin("UNKNOWN", PluginProvider.class, fallback));
    }

    @Test
    @DisplayName("The replacement plug-in with the highest order wins")
    void highestReplacementOrderWinsTest() {
        PluginLoader.initializePluginsForTesting(List.of(ReplacementPlugin.class, HighOrderReplacementPlugin.class));

        PluginProvider plugin = PluginLoader.getPlugin(REPLACEMENT_UUID, PluginProvider.class, null);

        assertInstanceOf(HighOrderReplacementPlugin.class, plugin);
    }

    @Test
    @DisplayName("Replacement lookups create a fresh plug-in instance")
    void replacementInstancesAreFreshTest() {
        PluginLoader.initializePluginsForTesting(List.of(ReplacementPlugin.class));

        PluginProvider first = PluginLoader.getPlugin(REPLACEMENT_UUID, PluginProvider.class, null);
        PluginProvider second = PluginLoader.getPlugin(REPLACEMENT_UUID, PluginProvider.class, null);

        assertNotSame(first, second);
    }

    @Test
    @DisplayName("Equal-order replacements are resolved by provider class name")
    void replacementTieUsesProviderNameTest() {
        PluginLoader.initializePluginsForTesting(List.of(ZReplacementPlugin.class, AReplacementPlugin.class));

        PluginProvider plugin = PluginLoader.getPlugin(REPLACEMENT_UUID, PluginProvider.class, null);

        assertInstanceOf(AReplacementPlugin.class, plugin);
    }

    @Test
    @DisplayName("An incompatible replacement type causes plug-in loading to fail")
    void incompatibleReplacementTypeFailsTest() {
        PluginLoader.initializePluginsForTesting(List.of(ReplacementPlugin.class));

        assertThrows(PluginLoadingException.class,
            () -> PluginLoader.getPlugin(REPLACEMENT_UUID, SpecializedPlugin.class, null));
    }

    @Test
    @DisplayName("Queue plug-ins are ordered deterministically and duplicate registrations are collapsed")
    void queueIsOrderedAndDuplicateRegistrationIsCollapsedTest() {
        PluginLoader.initializePluginsForTesting(List.of(
            LaterQueuePlugin.class,
            ZQueuePlugin.class,
            AQueuePlugin.class,
            AQueuePlugin.class));

        List<QueuePlugin<PluginProvider>> plugins =
            PluginLoader.getQueuePlugins(QUEUE_UUID, PluginProvider.class);

        assertTrue(PluginLoader.hasQueuePlugins(QUEUE_UUID));
        assertEquals(List.of("A_PLUGIN", "Z_PLUGIN", "LATER_PLUGIN"),
            plugins.stream().map(QueuePlugin::pluginUuid).toList());
        assertEquals(3, plugins.size());
    }

    @Test
    @DisplayName("Getting queue plug-ins skips entries that do not implement the requested type")
    void getQueuePluginsSkipsIncompatibleTypeTest() {
        PluginLoader.initializePluginsForTesting(List.of(AQueuePlugin.class, SpecializedQueuePlugin.class));

        List<QueuePlugin<SpecializedPlugin>> plugins =
            PluginLoader.getQueuePlugins(QUEUE_UUID, SpecializedPlugin.class);

        assertEquals(1, plugins.size());
        assertInstanceOf(SpecializedQueuePlugin.class, plugins.getFirst().plugin());
    }

    @Test
    @DisplayName("Queue results are immutable and contain fresh plug-in instances")
    void queueResultsAreImmutableAndInstancesAreFreshTest() {
        PluginLoader.initializePluginsForTesting(List.of(AQueuePlugin.class));

        List<QueuePlugin<PluginProvider>> first =
            PluginLoader.getQueuePlugins(QUEUE_UUID, PluginProvider.class);
        List<QueuePlugin<PluginProvider>> second =
            PluginLoader.getQueuePlugins(QUEUE_UUID, PluginProvider.class);

        assertNotSame(first.getFirst().plugin(), second.getFirst().plugin());
        assertThrows(UnsupportedOperationException.class,
            () -> first.add(new QueuePlugin<>("OTHER", new PlainPlugin())));
        assertFalse(PluginLoader.hasQueuePlugins("UNKNOWN"));
        assertTrue(PluginLoader.getQueuePlugins("UNKNOWN", PluginProvider.class).isEmpty());
    }

    @Test
    @DisplayName("Failed initialization does not publish a partial plug-in registry")
    void failedInitializationDoesNotPublishRegistryTest() {
        assertThrows(PluginLoadingException.class,
            () -> PluginLoader.initializePluginsForTesting(List.of(UnannotatedPlugin.class)));

        assertTrue(PluginLoader.initializePluginsForTesting(List.of(ReplacementPlugin.class)));
        assertInstanceOf(ReplacementPlugin.class,
            PluginLoader.getPlugin(REPLACEMENT_UUID, PluginProvider.class, null));
    }

    @Test
    @DisplayName("A blank replacement plug-in UUID is rejected")
    void rejectsBlankReplacementUuidTest() {
        assertThrows(PluginLoadingException.class,
            () -> PluginLoader.initializePluginsForTesting(List.of(BlankReplacementPlugin.class)));
    }

    @Test
    @DisplayName("A blank queue UUID is rejected")
    void rejectsBlankQueueUuidTest() {
        assertThrows(PluginLoadingException.class,
            () -> PluginLoader.initializePluginsForTesting(List.of(BlankQueuePlugin.class)));
    }

    @Test
    @DisplayName("A non-public plug-in provider is rejected")
    void rejectsNonPublicProviderTest() {
        assertThrows(PluginLoadingException.class,
            () -> PluginLoader.initializePluginsForTesting(List.of(NonPublicPlugin.class)));
    }

    @Test
    @DisplayName("An abstract plug-in provider is rejected")
    void rejectsAbstractProviderTest() {
        assertThrows(PluginLoadingException.class,
            () -> PluginLoader.initializePluginsForTesting(List.of(AbstractPlugin.class)));
    }

    @Test
    @DisplayName("A provider without a public no-argument constructor is rejected")
    void rejectsProviderWithoutPublicNoArgumentConstructorTest() {
        assertThrows(PluginLoadingException.class,
            () -> PluginLoader.initializePluginsForTesting(List.of(MissingConstructorPlugin.class)));
    }

    @Test
    @DisplayName("A plug-in constructor failure is wrapped as PluginLoadingException")
    void wrapsConstructorFailureTest() {
        PluginLoader.initializePluginsForTesting(List.of(ThrowingPlugin.class));

        assertThrows(PluginLoadingException.class,
            () -> PluginLoader.getPlugin("THROWING", PluginProvider.class, null));
    }

    // Test plugins

    @NanoXlsxPlugin(pluginUuid = REPLACEMENT_UUID)
    public static class ReplacementPlugin implements PluginProvider {
        public ReplacementPlugin() {
        }
    }

    @NanoXlsxPlugin(pluginUuid = REPLACEMENT_UUID, order = 10)
    public static class HighOrderReplacementPlugin implements PluginProvider {
        public HighOrderReplacementPlugin() {
        }
    }

    @NanoXlsxPlugin(pluginUuid = REPLACEMENT_UUID, order = 5)
    public static class AReplacementPlugin implements PluginProvider {
        public AReplacementPlugin() {
        }
    }

    @NanoXlsxPlugin(pluginUuid = REPLACEMENT_UUID, order = 5)
    public static class ZReplacementPlugin implements PluginProvider {
        public ZReplacementPlugin() {
        }
    }

    @NanoXlsxQueuePlugin(pluginUuid = "A_PLUGIN", queueUuid = QUEUE_UUID)
    public static class AQueuePlugin implements PluginProvider {
        public AQueuePlugin() {
        }
    }

    @NanoXlsxQueuePlugin(pluginUuid = "Z_PLUGIN", queueUuid = QUEUE_UUID)
    public static class ZQueuePlugin implements PluginProvider {
        public ZQueuePlugin() {
        }
    }

    @NanoXlsxQueuePlugin(pluginUuid = "LATER_PLUGIN", queueUuid = QUEUE_UUID, order = 1)
    public static class LaterQueuePlugin implements PluginProvider {
        public LaterQueuePlugin() {
        }
    }

    @NanoXlsxQueuePlugin(pluginUuid = "SPECIALIZED", queueUuid = QUEUE_UUID)
    public static class SpecializedQueuePlugin implements SpecializedPlugin {
        public SpecializedQueuePlugin() {
        }
    }

    public interface SpecializedPlugin extends PluginProvider {
    }

    public static class PlainPlugin implements PluginProvider {
        public PlainPlugin() {
        }
    }

    public static class UnannotatedPlugin implements PluginProvider {
        public UnannotatedPlugin() {
        }
    }

    @NanoXlsxPlugin(pluginUuid = " ")
    public static class BlankReplacementPlugin implements PluginProvider {
        public BlankReplacementPlugin() {
        }
    }

    @NanoXlsxQueuePlugin(pluginUuid = "PLUGIN", queueUuid = "")
    public static class BlankQueuePlugin implements PluginProvider {
        public BlankQueuePlugin() {
        }
    }

    @NanoXlsxPlugin(pluginUuid = "NON_PUBLIC")
    private static class NonPublicPlugin implements PluginProvider {
        public NonPublicPlugin() {
        }
    }

    @NanoXlsxPlugin(pluginUuid = "ABSTRACT")
    public abstract static class AbstractPlugin implements PluginProvider {
        public AbstractPlugin() {
        }
    }

    @NanoXlsxPlugin(pluginUuid = "MISSING_CONSTRUCTOR")
    public static class MissingConstructorPlugin implements PluginProvider {
        public MissingConstructorPlugin(String value) {
        }
    }

    @NanoXlsxPlugin(pluginUuid = "THROWING")
    public static class ThrowingPlugin implements PluginProvider {
        public ThrowingPlugin() {
            throw new IllegalStateException("Expected constructor failure");
        }
    }

    // Helper methods

    /**
     * Method to create a test jar file, representing a plugin
     * @param tempDirectory Directory to create the file
     * @return Full path to the jar file
     * @throws Exception Thrown in nay case of an error
     */
    private static Path createProviderJar(Path tempDirectory) throws Exception {
        String providerClassName = ServiceDiscoveredPlugin.class.getName();
        String providerClassEntry = providerClassName.replace('.', '/') + ".class";
        Path jarPath = tempDirectory.resolve("test-plugin.jar");

        try (InputStream classBytes = ServiceDiscoveredPlugin.class.getClassLoader()
                .getResourceAsStream(providerClassEntry);
                JarOutputStream jar = new JarOutputStream(Files.newOutputStream(jarPath))) {
            if (classBytes == null) {
                throw new IllegalStateException("Unable to find test provider class bytes");
            }
            jar.putNextEntry(new JarEntry(providerClassEntry));
            classBytes.transferTo(jar);
            jar.closeEntry();

            jar.putNextEntry(new JarEntry("META-INF/services/" + PluginProvider.class.getName()));
            jar.write((providerClassName + System.lineSeparator()).getBytes(StandardCharsets.UTF_8));
            jar.closeEntry();
        }
        return jarPath;
    }

    /**
     * Method to create a faulty test plugin jar
     * @param tempDirectory Directory to create the file
     * @param providerClassName Name of the class (not implelemntig a plugin)
     * @return Full path of the jar file
     * @throws Exception Thrown in nay case of an error
     */
    private static Path createServiceJar(Path tempDirectory, String providerClassName) throws Exception {
        Path jarPath = tempDirectory.resolve("malformed-plugin.jar");
        try (JarOutputStream jar = new JarOutputStream(Files.newOutputStream(jarPath))) {
            jar.putNextEntry(new JarEntry("META-INF/services/" + PluginProvider.class.getName()));
            jar.write((providerClassName + System.lineSeparator()).getBytes(StandardCharsets.UTF_8));
            jar.closeEntry();
        }
        return jarPath;
    }

    private static final class ProviderClassLoader extends URLClassLoader {

        private ProviderClassLoader(URL[] urls, ClassLoader parent) {
            super(urls, parent);
        }

        @Override
        protected synchronized Class<?> loadClass(String name, boolean resolve) throws ClassNotFoundException {
            if (!name.equals(ServiceDiscoveredPlugin.class.getName())) {
                return super.loadClass(name, resolve);
            }

            Class<?> providerClass = findLoadedClass(name);
            if (providerClass == null) {
                providerClass = findClass(name);
            }
            if (resolve) {
                resolveClass(providerClass);
            }
            return providerClass;
        }
    }
}
