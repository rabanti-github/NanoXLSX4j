/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;

/**
 * Repository for styles managed at runtime.
 * <p>
 * The repository deduplicates styles by their hash code and decouples the shared style references used by cells and
 * columns from individual workbooks. If a style with the same hash code is already present, {@link #addStyle(Style)}
 * returns the existing repository reference.
 * <p>
 * The repository is safe for concurrent single-map operations. The contained {@link Style} instances remain mutable,
 * however, and must not be changed after insertion because doing so can make their current hash code differ from the
 * key under which they are stored.
 */
public final class StyleRepository {

    private final ConcurrentMap<Integer, Style> styles;

    /** Creates the singleton repository. */
    private StyleRepository() {
        styles = new ConcurrentHashMap<>();
    }

    /**
     * Gets the singleton repository instance.
     * <p>
     * Initialization is lazy and thread-safe through the initialization-on-demand holder idiom.
     *
     * @return singleton repository instance
     */
    public static StyleRepository getInstance() {
        return InstanceHolder.INSTANCE;
    }

    /**
     * Gets the currently managed styles, keyed by their hash code.
     * <p>
     * The returned map is the live repository map. Mutating it directly affects the repository and should be limited
     * to maintenance and infrastructure code.
     *
     * @return live map of managed styles
     */
    public Map<Integer, Style> getStyles() {
        return styles;
    }

    /**
     * Adds a style to the repository and returns the canonical repository reference.
     * <p>
     * Deduplication is based on {@link Style#hashCode()}. Hash collisions are therefore treated as identical styles,
     * matching the NanoXLSX repository contract. Adding {@code null} has no effect and returns {@code null}.
     *
     * @param style style to add
     * @return the existing style with the same hash code, the added style, or {@code null} when the argument is null
     */
    public Style addStyle(Style style) {
        if (style == null) {
            return null;
        }

        Style existing = styles.putIfAbsent(style.hashCode(), style);
        return existing == null ? style : existing;
    }

    /**
     * Removes all styles from the repository.
     * <p>
     * Do not call this maintenance method while worksheet or workbook data is being processed. Existing cells and
     * columns may still reference removed styles, and concurrent additions may occur after the clear operation.
     */
    public void flushStyles() {
        styles.clear();
    }

    /** Holder used for lazy, thread-safe singleton initialization. */
    private static final class InstanceHolder {

        private static final StyleRepository INSTANCE = new StyleRepository();

        private InstanceHolder() {
            // No instances
        }
    }
}
