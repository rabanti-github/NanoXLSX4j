/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * Class representing a reduced sorted map (key / value). This class is not
 * compatible with the Map interface
 *
 * @author Raphael Stoeckli
 */
class SortedMap {

	// ### P R I V A T E F I E L D S ###
	private int count;
	private final HashMap<String, Integer> index;
	private final List<String> keyEntries;
	private final List<String> valueEntries;

	// ### C O N S T R U C T O R S ###

	/**
	 * Default constructor
	 */
	public SortedMap() {
		this.keyEntries = new ArrayList<>();
		this.valueEntries = new ArrayList<>();
		this.index = new HashMap<>();
		this.count = 0;
	}

	// ### M E T H O D S ###

	/**
	 * Method to add a key value pair
	 *
	 * @param key
	 *            key as string
	 * @param value
	 *            value as string
	 * @return returns the resolved string (either added or returned from an
	 *         existing entry)
	 */
	public String add(String key, String value) {
		if (index.containsKey(key)) {
			return valueEntries.get(index.get(key));
		}
		index.put(key, count);
		count++;
		keyEntries.add(key);
		valueEntries.add(value);
		return value;
	}

	/**
	 * Gets the keys of the map as list
	 *
	 * @return ArrayList of Keys
	 */
	public List<String> getKeys() {
		return this.keyEntries;
	}

	/**
	 * Gets the size of the map
	 *
	 * @return Number of entries in the map
	 */
	public int size() {
		return this.count;
	}

}
