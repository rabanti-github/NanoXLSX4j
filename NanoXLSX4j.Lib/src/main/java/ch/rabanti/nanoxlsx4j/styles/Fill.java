/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

/**
 * Class representing a Fill (background) entry. The Fill entry is used to
 * define background colors and fill patterns
 *
 * @author Raphael Stoeckli
 */
public class Fill extends AbstractStyle {
	// ### C O N S T A N T S ###
	/**
	 * Default Color (foreground or background)
	 */
	public static final String DEFAULT_COLOR = "FF000000";
	/**
	 * Default index color
	 */
	public static final int DEFAULT_INDEXED_COLOR = 64;
	/**
	 * Default pattern
	 */
	public static final PatternValue DEFAULT_PATTERN_FILL = PatternValue.none;

	// ### E N U M S ###

	/**
	 * Enum for the type of the fill pattern
	 */
	public enum PatternValue {
		/**
		 * No pattern (default)
		 */
		none(0),
		/**
		 * Solid fill (for colors)
		 */
		solid(1),
		/**
		 * Dark gray fill
		 */
		darkGray(2),
		/**
		 * Medium gray fill
		 */
		mediumGray(3),
		/**
		 * Light gray fill
		 */
		lightGray(4),
		/**
		 * 6.25% gray fill
		 */
		gray0625(5),
		/**
		 * 12.5% gray fill
		 */
		gray125(6);

		private final int value;

		PatternValue(int value) {
			this.value = value;
		}

		public int getValue() {
			return value;
		}

	}

	/**
	 * Enum for the type of the fill
	 */
	public enum FillType {
		/**
		 * Color defines a pattern color
		 */
		patternColor(0),
		/**
		 * Color defines a solid fill color
		 */
		fillColor(1);

		private final int value;

		FillType(int value) {
			this.value = value;
		}

		public int getValue() {
			return value;
		}
	}

	// ### P R I V A T E F I E L D S ###
	public int indexedColor;
	public PatternValue patternFill;
	public String foregroundColor;
	public String backgroundColor;

	// ### G E T T E R S & S E T T E R S ###

	/**
	 * Gets the indexed color (Default is 64)
	 *
	 * @return Indexed color
	 */
	public int getIndexedColor() {
		return indexedColor;
	}

	/**
	 * Sets the indexed color (Default is 64)
	 *
	 * @param indexedColor
	 *            Indexed color
	 */
	public void setIndexedColor(int indexedColor) {
		this.indexedColor = indexedColor;
	}

	/**
	 * Gets the pattern type of the fill (Default is none)
	 *
	 * @return Pattern type of the fill
	 */
	public PatternValue getPatternFill() {
		return patternFill;
	}

	/**
	 * Sets the pattern type of the fill (Default is none)
	 *
	 * @param patternFill
	 *            Pattern type of the fill
	 */
	public void setPatternFill(PatternValue patternFill) {
		this.patternFill = patternFill;
	}

	/**
	 * Gets the foreground color of the fill. The value is expressed as hex string
	 * with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @return Foreground color of the fill
	 */
	public String getForegroundColor() {
		return foregroundColor;
	}

	/**
	 * Sets the foreground color of the fill. The value is expressed as hex string
	 * with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @param foregroundColor
	 *            Foreground color of the fill
	 */
	public void setForegroundColor(String foregroundColor) {
		validateColor(foregroundColor, true);
		this.foregroundColor = foregroundColor;
		if (this.patternFill == PatternValue.none) {
			this.patternFill = PatternValue.solid;
		}
	}

	/**
	 * Gets the Background color of the fill. The value is expressed as hex string
	 * with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @return Background color of the fill
	 */
	public String getBackgroundColor() {
		return backgroundColor;
	}

	/**
	 * Sets the background color of the fill. The value is expressed as hex string
	 * with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @param backgroundColor
	 *            Background color of the fill
	 */
	public void setBackgroundColor(String backgroundColor) {
		validateColor(backgroundColor, true);
		this.backgroundColor = backgroundColor;
		if (this.patternFill == PatternValue.none) {
			this.patternFill = PatternValue.solid;
		}
	}

	// ### C O N S T R U C T O R S ###

	/**
	 * Default constructor
	 */
	public Fill() {
		this.indexedColor = DEFAULT_INDEXED_COLOR;
		this.patternFill = DEFAULT_PATTERN_FILL;
		this.foregroundColor = DEFAULT_COLOR;
		this.backgroundColor = DEFAULT_COLOR;
	}

	/**
	 * Constructor with foreground and background color
	 *
	 * @param foreground
	 *            Foreground color of the fill
	 * @param background
	 *            Background color of the fill
	 */
	public Fill(String foreground, String background) {
		this.setBackgroundColor(background);
		this.setForegroundColor(foreground);
		this.indexedColor = DEFAULT_INDEXED_COLOR;
		this.patternFill = PatternValue.solid;
	}

	/**
	 * Constructor with color value and fill type
	 *
	 * @param value
	 *            Color value
	 * @param fillType
	 *            Fill type (fill or pattern)
	 */
	public Fill(String value, FillType fillType) {
		if (fillType == FillType.fillColor) {
			this.backgroundColor = DEFAULT_COLOR;
			this.setForegroundColor(value);
		}
		else {
			this.setBackgroundColor(value);
			this.foregroundColor = DEFAULT_COLOR;
		}
		this.indexedColor = DEFAULT_INDEXED_COLOR;
		this.patternFill = PatternValue.solid;
	}

	// ### M E T H O D S ###

	/**
	 * Sets the color and the depending fill type
	 *
	 * @param value
	 *            Color value
	 * @param fillType
	 *            Fill type (fill or pattern)
	 */
	public void setColor(String value, FillType fillType) {
		if (fillType == FillType.fillColor) {
			this.backgroundColor = DEFAULT_COLOR;
			this.setForegroundColor(value);
		}
		else {
			this.setBackgroundColor(value);
			this.foregroundColor = DEFAULT_COLOR;
		}
		this.patternFill = PatternValue.solid;
	}

	/**
	 * Override toString method
	 *
	 * @return String of a class instance
	 */
	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder();
		sb.append("\"Fill\": {\n");
		addPropertyAsJson(sb, "BackgroundColor", backgroundColor);
		addPropertyAsJson(sb, "ForegroundColor", foregroundColor);
		addPropertyAsJson(sb, "IndexedColor", indexedColor);
		addPropertyAsJson(sb, "PatternFill", patternFill);
		addPropertyAsJson(sb, "HashCode", this.hashCode(), true);
		sb.append("\n}");
		return sb.toString();
	}

	/**
	 * Method to copy the current object to a new one
	 *
	 * @return Copy of the current object without the internal ID
	 */
	@Override
	public Fill copy() {
		Fill copy = new Fill();
		copy.setBackgroundColor(this.backgroundColor);
		copy.setForegroundColor(this.foregroundColor);
		copy.setIndexedColor(this.indexedColor);
		copy.setPatternFill(this.patternFill);
		return copy;
	}

	/**
	 * Override method to calculate the hash of this component
	 *
	 * @return Calculated hash as string
	 * @implNote Note that autogenerated hashcode algorithms may cause collisions.
	 *           Do not use 0 as fallback value for every field
	 */
	@Override
	public int hashCode() {
		int result = indexedColor;
		result = 31 * result + (patternFill != null ? patternFill.hashCode() : 1);
		result = 31 * result + (foregroundColor != null ? foregroundColor.hashCode() : 2);
		result = 31 * result + (backgroundColor != null ? backgroundColor.hashCode() : 4);
		return result;
	}

	// ### S T A T I C F U N C T I O N S ###

	/**
	 * Gets the pattern name from the enum
	 *
	 * @param pattern
	 *            Enum to process
	 * @return The valid value of the pattern as String
	 */
	public static String getPatternName(PatternValue pattern) {
		String output;
		switch (pattern) {
			case solid:
				output = "solid";
				break;
			case darkGray:
				output = "darkGray";
				break;
			case mediumGray:
				output = "mediumGray";
				break;
			case lightGray:
				output = "lightGray";
				break;
			case gray0625:
				output = "gray0625";
				break;
			case gray125:
				output = "gray125";
				break;
			default:
				output = "none";
				break;
		}
		return output;
	}

	/**
	 * Validates the passed string, whether it is a valid RGB value that can be used
	 * for Fills or Fonts
	 *
	 * @param hexCode
	 *            Hex string to check
	 * @param useAlpha
	 *            If true, two additional characters (total 8) are expected as alpha
	 *            value
	 * @throws StyleException
	 *             thrown if an invalid hex value is passed
	 */
	public static void validateColor(String hexCode, boolean useAlpha) {
		int length;
		String expression;
		length = useAlpha ? 8 : 6;
		if (hexCode == null || hexCode.length() != length) {
			throw new StyleException("The value '" + hexCode + "' is invalid. A valid value must contain six hex characters");
		}
		if (!hexCode.matches("[a-fA-F0-9]{6,8}")) {
			throw new StyleException("The expression '" + hexCode + "' is a valid hex value");
		}
	}

	/**
	 * Validates the passed string, whether it is a valid RGB value that can be used
	 * for Fills or Fonts
	 *
	 * @param hexCode
	 *            Hex string to check
	 * @param useAlpha
	 *            If true, two additional characters (total 8) are expected as alpha
	 *            value
	 * @param allowEmpty
	 *            Optional parameter that allows null or empty as valid values
	 * @throws StyleException
	 *             thrown if an invalid hex value is passed
	 */
	public static void validateColor(String hexCode, boolean useAlpha, boolean allowEmpty) {
		if ((hexCode == null || hexCode.isEmpty())) {
			if (allowEmpty) {
				return;
			}
			throw new StyleException("The color expression was null or empty");
		}
		validateColor(hexCode, useAlpha);
	}

}
