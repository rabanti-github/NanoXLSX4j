/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

import java.io.InputStream;
import java.util.Properties;

/**
 * Final class to provide metadata for the library
 *
 * @author Raphael Stoeckli
 */
public final class Version {

	// ### C O N S T A N T S ###
	/**
	 * Application name of the library
	 */
	public static final String APPLICATION_NAME = "NanoXLSX4j";

	/**
	 * Gets the library version
	 * @return Version string of the library in the format "x.yyyy"
	 */
	public static String getVersion(){
		try{
			Properties props = new Properties();
			try (InputStream is = Version.class.getResourceAsStream("/version.properties")) {
				props.load(is);
			}
			return transformVersion(props.getProperty("build.version"));
		}
		catch (Exception ex){
			return "0.0000"; // Dummy fallback
		}
	}

	private static String transformVersion(String version) {
		String[] parts = version.split("\\.");
		int major = Integer.parseInt(parts[0]);
		int minor = Integer.parseInt(parts[1]);
		int patch = Integer.parseInt(parts[2]);
		if (major > 9999 || minor > 9999 || patch > 9999) {
			throw new FormatException("Invalid version number: " + version);
		}
		String minorPatch = String.format("%02d%02d", minor, patch);
		String versionNumber = String.format("%d.%s", major, minorPatch);
		if (versionNumber.length() > 7) {
			throw new FormatException("Invalid version number: " + version);
		}
		return versionNumber;
	}
}
