//STATUS: VERIFY
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.registry;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.ServiceLoader;
import java.util.concurrent.atomic.AtomicReference;

/**
 * Central orchestrator for the NanoXLSX4j plugin system.
 * <p>
 * Manages two plugin categories:
 * </p>
 * <ul>
 *   <li><b>Replacing plugins</b> ({@link NanoXlsxPlugin}) — substitute the default
 *       implementation of a specific reader/writer component. Only one plugin is
 *       active per UUID; ties are resolved by highest {@link NanoXlsxPlugin#order()}
 *       value.</li>
 *   <li><b>Queue plugins</b> ({@link NanoXlsxQueuePlugin}) — executed sequentially
 *       before or after a standard read/write operation. Multiple plugins can be
 *       active per UUID; they are iterated in ascending
 *       {@link NanoXlsxQueuePlugin#order()} order via
 *       {@link #getNextQueuePlugin}.</li>
 * </ul>
 *
 * <h3>Discovery</h3>
 * <p>
 * Plugins are discovered via the standard Java {@link ServiceLoader} mechanism.
 * Every plugin class must:
 * </p>
 * <ol>
 *   <li>Implement {@link NanoXlsxPluggable}.</li>
 *   <li>Be annotated with {@link NanoXlsxPlugin} or {@link NanoXlsxQueuePlugin}.</li>
 *   <li>Be listed in
 *       {@code META-INF/services/ch.rabanti.nanoxlsx4j.registry.NanoXlsxPluggable}
 *       inside its JAR.</li>
 * </ol>
 * <p>Call {@link #initialize()} once at application start-up to trigger discovery.
 * The method is idempotent and thread-safe.</p>
 *
 * <h3>Usage (reader/writer side)</h3>
 * <pre>{@code
 * // Replacing plugin — fall back to default if none registered
 * ISharedStringsReader reader = PluginLoader.getInstance()
 *         .getPlugin(PluginUUID.SHARED_STRINGS_READER, new SharedStringsReader());
 *
 * // Queue — iterate until null
 * AtomicReference<String> currentUUID = new AtomicReference<>(null);
 * Object plugin;
 * while ((plugin = PluginLoader.getInstance()
 *         .getNextQueuePlugin(PluginUUID.WRITER_PREPEND_QUEUE, currentUUID.get(), currentUUID)) != null) {
 *     // invoke plugin
 * }
 * }</pre>
 *
 * @see NanoXlsxPlugin
 * @see NanoXlsxQueuePlugin
 * @see PluginUUID
 */
public final class PluginLoader {

    // -------------------------------------------------------------------------
    // Singleton
    // -------------------------------------------------------------------------

    private static final PluginLoader INSTANCE = new PluginLoader();

    /** Returns the singleton {@code PluginLoader} instance. */
    public static PluginLoader getInstance() {
        return INSTANCE;
    }

    private PluginLoader() {}

    // -------------------------------------------------------------------------
    // State
    // -------------------------------------------------------------------------

    private final Object lock = new Object();
    private volatile boolean initialized = false;

    /** UUID → best replacing-plugin instance metadata. */
    private final Map<String, PluginInstance> pluginClasses = new HashMap<>();

    /** UUID → ordered list of queue-plugin instance metadata. */
    private final Map<String, List<PluginInstance>> queuePluginClasses = new HashMap<>();

    // -------------------------------------------------------------------------
    // Lifecycle
    // -------------------------------------------------------------------------

    /**
     * Discovers and registers all plugins available on the current class path.
     * <p>
     * Uses {@link ServiceLoader} to find all {@link NanoXlsxPluggable} implementations,
     * then groups them by annotation type and UUID. Safe to call from multiple threads;
     * subsequent calls after the first successful initialization are no-ops and return
     * {@code false}.
     * </p>
     *
     * @return {@code true} if this call performed the initialization;
     *         {@code false} if already initialized
     */
    public boolean initialize() {
        synchronized (lock) {
            if (initialized) {
                return false;
            }
            ServiceLoader<NanoXlsxPluggable> loader =
                    ServiceLoader.load(NanoXlsxPluggable.class);
            List<Class<?>> discovered = new ArrayList<>();
            for (NanoXlsxPluggable p : loader) {
                discovered.add(p.getClass());
            }
            registerAll(discovered);
            initialized = true;
            return true;
        }
    }

    /**
     * Manually injects a list of plugin types, bypassing {@link ServiceLoader}
     * discovery. Intended for unit tests.
     * <p>
     * Clears any previously registered plugins before registering the supplied
     * types, and marks the loader as initialized.
     * </p>
     *
     * @param types plugin classes to register; must be annotated with
     *              {@link NanoXlsxPlugin} or {@link NanoXlsxQueuePlugin}
     */
    public void injectPlugins(List<Class<?>> types) {
        synchronized (lock) {
            pluginClasses.clear();
            queuePluginClasses.clear();
            registerAll(types);
            initialized = true;
        }
    }

    /**
     * Clears all registered plugins and resets the initialization flag.
     * After calling this method {@link #initialize()} may be called again.
     */
    public void disposePlugins() {
        synchronized (lock) {
            pluginClasses.clear();
            queuePluginClasses.clear();
            initialized = false;
        }
    }

    // -------------------------------------------------------------------------
    // Plugin retrieval — Replacing
    // -------------------------------------------------------------------------

    /**
     * Returns an instance of the replacing plugin registered under {@code uuid},
     * cast to {@code T}. If no replacing plugin is registered for that UUID the
     * supplied {@code fallback} instance is returned unchanged.
     *
     * @param <T>      the expected plugin type
     * @param uuid     the plugin UUID (see {@link PluginUUID})
     * @param fallback the default instance to use when no plugin is registered
     * @return the active plugin instance or {@code fallback}
     * @throws ClassCastException if the registered plugin cannot be cast to {@code T}
     */
    @SuppressWarnings("unchecked")
    public <T> T getPlugin(String uuid, T fallback) {
        PluginInstance instance = pluginClasses.get(uuid);
        if (instance == null) {
            return fallback;
        }
        try {
            return (T) instance.type.getDeclaredConstructor().newInstance();
        } catch (ReflectiveOperationException e) {
            return fallback;
        }
    }

    // -------------------------------------------------------------------------
    // Plugin retrieval — Queue
    // -------------------------------------------------------------------------

    /**
     * Returns the next plugin in the queue identified by {@code uuid}.
     * <p>
     * Pass {@code null} as {@code lastPluginUUID} to start from the beginning of
     * the queue. After each call the internal UUID of the returned plugin is written
     * into {@code currentPluginUUID} so it can be passed as {@code lastPluginUUID}
     * in the subsequent call. Returns {@code null} when the queue is exhausted.
     * </p>
     *
     * @param <T>               the expected plugin type
     * @param uuid              the queue UUID (see {@link PluginUUID})
     * @param lastPluginUUID    the {@link PluginInstance#instanceUUID} returned by
     *                          the previous call, or {@code null} to start from the head
     * @param currentPluginUUID receives the instance UUID of the returned plugin
     * @return the next plugin instance, or {@code null} when the queue is exhausted
     * @throws ClassCastException if a queue entry cannot be cast to {@code T}
     */
    @SuppressWarnings("unchecked")
    public <T> T getNextQueuePlugin(String uuid, String lastPluginUUID,
                                    AtomicReference<String> currentPluginUUID) {
        List<PluginInstance> queue = queuePluginClasses.get(uuid);
        if (queue == null || queue.isEmpty()) {
            currentPluginUUID.set(null);
            return null;
        }

        int startIndex = 0;
        if (lastPluginUUID != null) {
            for (int i = 0; i < queue.size(); i++) {
                if (queue.get(i).instanceUUID.equals(lastPluginUUID)) {
                    startIndex = i + 1;
                    break;
                }
            }
        }

        if (startIndex >= queue.size()) {
            currentPluginUUID.set(null);
            return null;
        }

        PluginInstance next = queue.get(startIndex);
        currentPluginUUID.set(next.instanceUUID);
        try {
            return (T) next.type.getDeclaredConstructor().newInstance();
        } catch (ReflectiveOperationException e) {
            currentPluginUUID.set(null);
            return null;
        }
    }

    // -------------------------------------------------------------------------
    // Internal helpers
    // -------------------------------------------------------------------------

    private void registerAll(List<Class<?>> types) {
        for (Class<?> type : types) {
            NanoXlsxPlugin replacing = type.getAnnotation(NanoXlsxPlugin.class);
            if (replacing != null) {
                registerReplacingPlugin(replacing.plugInUuid(), replacing.order(), type);
            }
            NanoXlsxQueuePlugin queue = type.getAnnotation(NanoXlsxQueuePlugin.class);
            if (queue != null) {
                registerQueuePlugin(queue.queueUuid(), queue.order(), type);
            }
        }
    }

    private void registerReplacingPlugin(String uuid, int order, Class<?> type) {
        PluginInstance existing = pluginClasses.get(uuid);
        if (existing == null || order > existing.order) {
            pluginClasses.put(uuid, new PluginInstance(uuid, order, type));
        }
    }

    private void registerQueuePlugin(String uuid, int order, Class<?> type) {
        List<PluginInstance> queue =
                queuePluginClasses.computeIfAbsent(uuid, k -> new ArrayList<>());
        queue.add(new PluginInstance(uuid, order, type));
        queue.sort(null); // PluginInstance implements Comparable by order ascending
    }

    // -------------------------------------------------------------------------
    // Inner class
    // -------------------------------------------------------------------------

    /**
     * Metadata holder for a single registered plugin.
     */
    static final class PluginInstance implements Comparable<PluginInstance> {

        /** UUID of the plugin slot (copied from the annotation). */
        final String uuid;

        /**
         * Unique identifier for this particular registration entry, used to
         * track position in queue iteration.
         */
        final String instanceUUID;

        /** Priority for conflict resolution (replacing) or execution order (queue). */
        final int order;

        /** The concrete plugin class. */
        final Class<?> type;

        PluginInstance(String uuid, int order, Class<?> type) {
            this.uuid = uuid;
            this.order = order;
            this.type = type;
            // Unique key: combine slot UUID + class name so two registrations of
            // the same slot are always distinguishable.
            this.instanceUUID = uuid + "::" + type.getName();
        }

        @Override
        public int compareTo(PluginInstance other) {
            return Integer.compare(this.order, other.order);
        }
    }
}
