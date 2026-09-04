/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.registry;

import java.lang.reflect.Constructor;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.ServiceConfigurationError;
import java.util.ServiceLoader;
import java.util.Set;

import ch.rabanti.nanoxlsx4j.annotations.InternalApi;
import ch.rabanti.nanoxlsx4j.exceptions.PluginLoadingException;
import ch.rabanti.nanoxlsx4j.registry.NanoXlsxPlugin;
import ch.rabanti.nanoxlsx4j.registry.NanoXlsxQueuePlugin;
import ch.rabanti.nanoxlsx4j.registry.Plugin;

/**
 * Discovers and resolves NanoXLSX4j plug-ins.
 * <p>
 * This class is shared with the Reader and Writer modules but is not supported application API.
 */
@InternalApi
public final class PluginLoader {

    private static final Object LOCK = new Object();
    private static final Comparator<PluginDefinition> QUEUE_ORDER = Comparator
        .comparingInt(PluginDefinition::order)
        .thenComparing(PluginDefinition::pluginUuid)
        .thenComparing(definition -> definition.type().getName());

    private static volatile PluginRegistry registry;

    private PluginLoader() {
    }

    /**
     * Initializes plug-in discovery with the standard service-loading context.
     *
     * @return True if this call initialized the registry; false when it was already initialized
     */
    @InternalApi
    public static boolean initialize() {
        return initialize(ServiceLoader.load(Plugin.class));
    }

    /**
     * Initializes plug-in discovery with an explicit class loader.
     *
     * @param classLoader Class loader containing service-provider declarations
     * @return True if this call initialized the registry; false when it was already initialized
     */
    @InternalApi
    public static boolean initialize(ClassLoader classLoader) {
        Objects.requireNonNull(classLoader, "classLoader");
        return initialize(ServiceLoader.load(Plugin.class, classLoader));
    }

    /**
     * Gets a new replacement plug-in instance, or the supplied fallback if no replacement was registered.
     *
     * @param pluginUuid UUID of the replacement point
     * @param requestedType Required plug-in type
     * @param fallback Fallback instance
     * @param <T> Required plug-in type
     * @return New replacement instance or fallback
     */
    @InternalApi
    public static <T extends Plugin> T getPlugin(String pluginUuid, Class<T> requestedType, T fallback) {
        requireInitialized();
        Objects.requireNonNull(pluginUuid, "pluginUuid");
        Objects.requireNonNull(requestedType, "requestedType");

        PluginDefinition definition = registry.replacements().get(pluginUuid);
        if (definition == null) {
            return fallback;
        }
        if (!requestedType.isAssignableFrom(definition.type())) {
            throw new PluginLoadingException("Plug-in " + definition.type().getName()
                + " registered for " + pluginUuid + " does not implement " + requestedType.getName());
        }
        return requestedType.cast(createInstance(definition));
    }

    /**
     * Indicates whether a queue contains at least one plug-in.
     *
     * @param queueUuid Queue UUID
     * @return True when the queue contains a plug-in
     */
    @InternalApi
    public static boolean hasQueuePlugins(String queueUuid) {
        requireInitialized();
        Objects.requireNonNull(queueUuid, "queueUuid");
        List<PluginDefinition> definitions = registry.queues().get(queueUuid);
        return definitions != null && !definitions.isEmpty();
    }

    /**
     * Gets newly created compatible plug-ins in deterministic queue order.
     *
     * @param queueUuid Queue UUID
     * @param requestedType Required plug-in type
     * @param <T> Required plug-in type
     * @return Immutable list of UUID and plug-in instance pairs
     */
    @InternalApi
    public static <T extends Plugin> List<QueuePlugin<T>> getQueuePlugins(String queueUuid, Class<T> requestedType) {
        requireInitialized();
        Objects.requireNonNull(queueUuid, "queueUuid");
        Objects.requireNonNull(requestedType, "requestedType");

        List<PluginDefinition> definitions = registry.queues().get(queueUuid);
        if (definitions == null) {
            return List.of();
        }

        List<QueuePlugin<T>> instances = new ArrayList<>();
        for (PluginDefinition definition : definitions) {
            if (requestedType.isAssignableFrom(definition.type())) {
                T plugin = requestedType.cast(createInstance(definition));
                instances.add(new QueuePlugin<>(definition.pluginUuid(), plugin));
            }
        }
        return List.copyOf(instances);
    }

    static boolean initializePluginsForTesting(Collection<Class<? extends Plugin>> pluginTypes) {
        Objects.requireNonNull(pluginTypes, "pluginTypes");
        if (registry != null) {
            return false;
        }
        return initializeRegistry(buildRegistry(pluginTypes));
    }

    static void resetForTesting() {
        synchronized (LOCK) {
            registry = null;
        }
    }

    private static boolean initialize(ServiceLoader<Plugin> serviceLoader) {
        if (registry != null) {
            return false;
        }

        List<Class<? extends Plugin>> pluginTypes = new ArrayList<>();
        try {
            serviceLoader.stream().forEach(provider -> pluginTypes.add(provider.type()));
        } catch (ServiceConfigurationError | RuntimeException error) {
            throw new PluginLoadingException("Failed to discover NanoXLSX4j plug-ins", error);
        }
        return initializeRegistry(buildRegistry(pluginTypes));
    }

    private static boolean initializeRegistry(PluginRegistry candidate) {
        if (registry != null) {
            return false;
        }
        synchronized (LOCK) {
            if (registry != null) {
                return false;
            }
            registry = candidate;
            return true;
        }
    }

    private static PluginRegistry buildRegistry(Collection<Class<? extends Plugin>> pluginTypes) {
        Map<String, PluginDefinition> replacements = new HashMap<>();
        Map<String, List<PluginDefinition>> queues = new HashMap<>();
        Set<QueueDefinitionKey> queueDefinitions = new HashSet<>();

        for (Class<? extends Plugin> pluginType : pluginTypes) {
            validateProvider(pluginType);

            NanoXlsxPlugin replacement = pluginType.getAnnotation(NanoXlsxPlugin.class);
            NanoXlsxQueuePlugin[] queueAnnotations = pluginType.getAnnotationsByType(NanoXlsxQueuePlugin.class);
            if (replacement == null && queueAnnotations.length == 0) {
                throw new PluginLoadingException("Plug-in provider " + pluginType.getName()
                    + " has no NanoXLSX4j plug-in annotation");
            }

            if (replacement != null) {
                validateUuid(replacement.pluginUuid(), "pluginUuid", pluginType);
                PluginDefinition candidate = new PluginDefinition(
                    replacement.pluginUuid(), replacement.order(), pluginType);
                replacements.merge(replacement.pluginUuid(), candidate, PluginLoader::preferredReplacement);
            }

            for (NanoXlsxQueuePlugin queueAnnotation : queueAnnotations) {
                validateUuid(queueAnnotation.pluginUuid(), "pluginUuid", pluginType);
                validateUuid(queueAnnotation.queueUuid(), "queueUuid", pluginType);
                QueueDefinitionKey key = new QueueDefinitionKey(
                    queueAnnotation.queueUuid(), queueAnnotation.pluginUuid(), pluginType);
                if (queueDefinitions.add(key)) {
                    queues.computeIfAbsent(queueAnnotation.queueUuid(), ignored -> new ArrayList<>())
                        .add(new PluginDefinition(
                            queueAnnotation.pluginUuid(), queueAnnotation.order(), pluginType));
                }
            }
        }

        Map<String, List<PluginDefinition>> immutableQueues = new HashMap<>();
        for (Map.Entry<String, List<PluginDefinition>> entry : queues.entrySet()) {
            entry.getValue().sort(QUEUE_ORDER);
            immutableQueues.put(entry.getKey(), List.copyOf(entry.getValue()));
        }
        return new PluginRegistry(Map.copyOf(replacements), Map.copyOf(immutableQueues));
    }

    private static PluginDefinition preferredReplacement(PluginDefinition current, PluginDefinition candidate) {
        if (candidate.order() > current.order()) {
            return candidate;
        }
        if (candidate.order() < current.order()) {
            return current;
        }
        return candidate.type().getName().compareTo(current.type().getName()) < 0 ? candidate : current;
    }

    private static void validateProvider(Class<? extends Plugin> pluginType) {
        if (pluginType == null) {
            throw new PluginLoadingException("Plug-in provider type must not be null");
        }

        int modifiers = pluginType.getModifiers();
        if (!Modifier.isPublic(modifiers) || Modifier.isAbstract(modifiers) || pluginType.isInterface()) {
            throw new PluginLoadingException("Plug-in provider must be a public concrete class: "
                + pluginType.getName());
        }
        if (!Plugin.class.isAssignableFrom(pluginType)) {
            throw new PluginLoadingException("Plug-in provider does not implement Plugin: " + pluginType.getName());
        }

        try {
            Constructor<? extends Plugin> constructor = pluginType.getConstructor();
            if (!Modifier.isPublic(constructor.getModifiers())) {
                throw new PluginLoadingException("Plug-in provider must have a public no-argument constructor: "
                    + pluginType.getName());
            }
        } catch (NoSuchMethodException | SecurityException exception) {
            throw new PluginLoadingException("Plug-in provider must have a public no-argument constructor: "
                + pluginType.getName(), exception);
        }
    }

    private static void validateUuid(String uuid, String elementName, Class<? extends Plugin> pluginType) {
        if (uuid.isBlank()) {
            throw new PluginLoadingException("Annotation element " + elementName + " must not be blank on "
                + pluginType.getName());
        }
    }

    private static Plugin createInstance(PluginDefinition definition) {
        try {
            return definition.type().getConstructor().newInstance();
        } catch (ReflectiveOperationException | LinkageError exception) {
            throw new PluginLoadingException("Failed to instantiate plug-in " + definition.type().getName(),
                exception);
        }
    }

    private static void requireInitialized() {
        if (registry == null) {
            throw new IllegalStateException("PluginLoader has not been initialized");
        }
    }

    /**
     * UUID and instance of a resolved queue plug-in.
     *
     * @param pluginUuid Plug-in UUID
     * @param plugin Newly created plug-in instance
     * @param <T> Plug-in type
     */
    @InternalApi
    public record QueuePlugin<T extends Plugin>(String pluginUuid, T plugin) {
    }

    private record PluginDefinition(String pluginUuid, int order, Class<? extends Plugin> type) {
    }

    private record QueueDefinitionKey(String queueUuid, String pluginUuid, Class<? extends Plugin> type) {
    }

    private record PluginRegistry(
            Map<String, PluginDefinition> replacements,
            Map<String, List<PluginDefinition>> queues) {
    }
}
