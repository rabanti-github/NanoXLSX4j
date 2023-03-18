/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import java.util.HashMap;
import java.util.Map;

/**
 * Class to manage all styles at runtime, before writing XLSX files. The main
 * purpose is deduplication and decoupling of styles from workbooks at runtime
 *
 * @apiNote Be careful when changing style data in this class. It may lead to
 *          inconsistencies
 */
public class StyleRepository {

	private static StyleRepository instance;
	private boolean importInProgress = false;

	/**
	 * Gets the singleton instance of the repository
	 *
	 * @return StyleRepository instance
	 */
	public static StyleRepository getInstance() {
		if (instance == null) {
			instance = new StyleRepository();
		}
		return instance;
	}

	private final Map<Integer, Style> styles;

	/**
	 * Gets the currently managed styles of the repository
	 *
	 * @return Map of managed styles
	 */
	public Map<Integer, Style> getStyles() {
		return styles;
	}

	/**
	 * Sets the import state of the repository
	 * 
	 * @param importInProgress
	 *            If true certain exceptions will be suppressed and transformations
	 *            on styles are performed when a worksheet is loaded
	 */
	public void setImportInProgress(boolean importInProgress) {
		this.importInProgress = importInProgress;
	}

	/**
	 * Gets the import state of the repository
	 * 
	 * @return If true certain exceptions will be suppressed and transformations on
	 *         styles are performed when a worksheet is loaded
	 */
	public boolean isImportInProgress() {
		return importInProgress;
	}

	/**
	 * Private constructor. The class is not intended to instantiate outside the
	 * singleton
	 */
	private StyleRepository() {
		this.styles = new HashMap<>();
	}

	/**
	 * Adds a style to the repository and returns the actual reference
	 *
	 * @param style
	 *            Style to add
	 * @return Reference from the repository. If the style to add already existed,
	 *         the existing object is returned, otherwise the newly added one
	 */
	public Style addStyle(Style style) {
		if (style == null) {
			return null;
		}
		int hashCode = style.hashCode();
		if (!styles.containsKey(hashCode)) {
			styles.put(hashCode, style);
		}
		return styles.get(hashCode);
	}

	/**
	 * Empties the static repository
	 *
	 * @apiNote Do not use this maintenance method while writing data on a worksheet
	 *          or workbook. It will lead to invalid style data or even
	 *          exceptions.<br>
	 *          Only use this method after all worksheets in all workbooks are
	 *          disposed. It may free memory then.
	 */
	public void flushStyles() {
		styles.clear();
	}

}
