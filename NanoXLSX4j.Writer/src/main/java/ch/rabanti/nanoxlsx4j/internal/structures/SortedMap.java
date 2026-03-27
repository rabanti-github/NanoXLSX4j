/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.internal.structures;

import ch.rabanti.nanoxlsx4j.interfaces.IFormattableText;
import ch.rabanti.nanoxlsx4j.interfaces.ISortedMap;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Class to manage key–value pairs ({@link IFormattableText} / index string).
 * Entries are kept in the order they were added.
 * <p>
 * Java equivalent of the .NET internal class {@code NanoXLSX.Internal.SortedMap} (namespace
 * {@code NanoXLSX.Internal.Structures}).
 * </p>
 * <p>
 * This class is not compatible with the {@link Map} interface.
 * </p>
 *
 * @implNote Framework-internal class. Not intended for use outside the writer module.
 * @author Raphael Stoeckli
 */
public class SortedMap implements ISortedMap {

    // ### P R I V A T E  F I E L D S ###

    private final List<IFormattableText> keys;
    private final List<String> indexEntries;
    private final Map<IFormattableText, Integer> index;
    private int count;

    // ### G E T T E R S ###

    /**
     * Returns the number of entries currently in the map
     *
     * @return entry count
     */
    @Override
    public int getCount() {
        return count;
    }

    /**
     * Gets the keys of the map in insertion order
     *
     * @return list of {@link IFormattableText} keys
     */
    public List<IFormattableText> getKeys() {
        return this.keys;
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Default constructor
     */
    public SortedMap() {
        this.keys = new ArrayList<>();
        this.indexEntries = new ArrayList<>();
        this.index = new HashMap<>();
        this.count = 0;
    }

    // ### M E T H O D S ###

    /**
     * Adds a key–value pair to the map, or returns the existing index string if the key is already present.
     * <p>
     * Key equality is determined by {@link IFormattableText#equals(Object)} and
     * {@link IFormattableText#hashCode()} on the concrete implementation (e.g. {@link PlainText}).
     * </p>
     *
     * @param text           {@link IFormattableText} used as key
     * @param referenceIndex Index string to associate with the text
     * @return The resolved reference index (either newly added or pre-existing)
     */
    @Override
    public String add(IFormattableText text, String referenceIndex) {
        if (index.containsKey(text)) {
            return indexEntries.get(index.get(text));
        }
        index.put(text, count);
        count++;
        keys.add(text);
        indexEntries.add(referenceIndex);
        return referenceIndex;
    }

}
