/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.Helper;

/**
 * Static class with common util methods, used during reading XLSX files
 */
public class ReaderUtils {

	private ReaderUtils() {
		// do not instantiate
	}

	/**
	 * Parses a bool as a binary number either based on an int (0/1) or a string
	 * expression (true/ false), independent of the culture info of the host
	 * 
	 * @param rawValue
	 *            Raw number or expression as string
	 * @return arsed boolean as number (0 = false, 1 = true)
	 */
	protected static int parseBinaryBoolean(String rawValue) {
		if (Helper.isNullOrEmpty(rawValue)) {
			return 0;
		}
		IntParser parser = new IntParser(rawValue);
		if (parser.hasValue) {
			if (parser.value >= 1) {
				return 1;
			}
			else {
				return 0;
			}
		}
		if (rawValue.equalsIgnoreCase("true")) {
			return 1;
		}
		else {
			return 0;
		}
	}

	/**
	 * Parser class to handle nullable integers
	 */
	protected static class IntParser {
		public boolean hasValue;
		public int value;

		public IntParser(String rawValue) {
			try {
				value = Integer.parseInt(rawValue);
				hasValue = true;
			}
			catch (Exception e) {
				value = 0;
				hasValue = false;
			}
		}

		public static IntParser tryParse(String rawValue) {
			return new IntParser(rawValue);
		}
	}
}
