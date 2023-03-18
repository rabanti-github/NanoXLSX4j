/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

/**
 * Class representing a Border entry. The Border entry is used to define frames
 * and cell borders
 *
 * @author Raphael Stoeckli
 */
public class Border extends AbstractStyle {

	// ### C O N S T A N T S ###
	/**
	 * Default border style as constant
	 */
	public static final StyleValue DEFAULT_BORDER_STYLE = StyleValue.none;

	/**
	 * Default border color as constant
	 */
	public static final String DEFAULT_BORDER_COLOR = "";

	// ### E N U M S ###

	/**
	 * Enum for the border style
	 */
	public enum StyleValue {
		/**
		 * no border
		 */
		none(0),
		/**
		 * hair border
		 */
		hair(1),
		/**
		 * dotted border
		 */
		dotted(2),
		/**
		 * dashed border with double-dots
		 */
		dashDotDot(3),
		/**
		 * dash-dotted border
		 */
		dashDot(4),
		/**
		 * dashed border
		 */
		dashed(5),
		/**
		 * thin border
		 */
		thin(6),
		/**
		 * medium-dashed border with double-dots
		 */
		mediumDashDotDot(7),
		/**
		 * slant dash-dotted border
		 */
		slantDashDot(8),
		/**
		 * medium dash-dotted border
		 */
		mediumDashDot(9),
		/**
		 * medium dashed border
		 */
		mediumDashed(10),
		/**
		 * medium border
		 */
		medium(11),
		/**
		 * thick border
		 */
		thick(12),
		/**
		 * double border
		 */
		s_double(13);

		private final int value;

		StyleValue(int value) {
			this.value = value;
		}

		public int getValue() {
			return value;
		}
	}

	/**
	 * Gets the border style name from the enum
	 *
	 * @param style
	 *            Enum to process
	 * @return The valid value of the border style as String
	 */
	public static String getStyleName(StyleValue style) {
		String output = "";
		switch (style) {
			case hair:
				output = "hair";
				break;
			case dotted:
				output = "dotted";
				break;
			case dashDotDot:
				output = "dashDotDot";
				break;
			case dashDot:
				output = "dashDot";
				break;
			case dashed:
				output = "dashed";
				break;
			case thin:
				output = "thin";
				break;
			case mediumDashDotDot:
				output = "mediumDashDotDot";
				break;
			case slantDashDot:
				output = "slantDashDot";
				break;
			case mediumDashDot:
				output = "mediumDashDot";
				break;
			case mediumDashed:
				output = "mediumDashed";
				break;
			case medium:
				output = "medium";
				break;
			case thick:
				output = "thick";
				break;
			case s_double:
				output = "double";
				break;
			// Default / none is already handled (ignored)
		}
		return output;
	}

	// ### P R I V A T E F I E L D S ###
	private StyleValue leftStyle;
	private StyleValue rightStyle;
	private StyleValue topStyle;
	private StyleValue bottomStyle;
	private StyleValue diagonalStyle;
	private boolean diagonalDown;
	private boolean diagonalUp;
	private String leftColor;
	private String rightColor;
	private String topColor;
	private String bottomColor;
	private String diagonalColor;

	// ### G E T T E R S & S E T T E R S ###

	/**
	 * Gets the style of left cell border
	 *
	 * @return Style of left cell border
	 */
	public StyleValue getLeftStyle() {
		return leftStyle;
	}

	/**
	 * Sets the style of left cell border
	 *
	 * @param leftStyle
	 *            Style of left cell border
	 */
	public void setLeftStyle(StyleValue leftStyle) {
		this.leftStyle = leftStyle;
	}

	/**
	 * Gets the style of right cell border
	 *
	 * @return Style of right cell border
	 */
	public StyleValue getRightStyle() {
		return rightStyle;
	}

	/**
	 * Sets the style of right cell border
	 *
	 * @param rightStyle
	 *            Style of right cell border
	 */
	public void setRightStyle(StyleValue rightStyle) {
		this.rightStyle = rightStyle;
	}

	/**
	 * Gets the style of top cell border
	 *
	 * @return Style of top cell border
	 */
	public StyleValue getTopStyle() {
		return topStyle;
	}

	/**
	 * Sets the style of top cell border
	 *
	 * @param topStyle
	 *            Style of top cell border
	 */
	public void setTopStyle(StyleValue topStyle) {
		this.topStyle = topStyle;
	}

	/**
	 * Gets the style of bottom cell border
	 *
	 * @return Style of bottom cell border
	 */
	public StyleValue getBottomStyle() {
		return bottomStyle;
	}

	/**
	 * Sets the style of bottom cell border
	 *
	 * @param bottomStyle
	 *            Style of bottom cell border
	 */
	public void setBottomStyle(StyleValue bottomStyle) {
		this.bottomStyle = bottomStyle;
	}

	/**
	 * Gets the style of the diagonal lines
	 *
	 * @return Style of the diagonal lines
	 */
	public StyleValue getDiagonalStyle() {
		return diagonalStyle;
	}

	/**
	 * Sets the style of the diagonal lines
	 *
	 * @param diagonalStyle
	 *            Style of the diagonal lines
	 */
	public void setDiagonalStyle(StyleValue diagonalStyle) {
		this.diagonalStyle = diagonalStyle;
	}

	/**
	 * Gets the downwards diagonal line
	 *
	 * @return If true, the downwards diagonal line is used
	 */
	public boolean isDiagonalDown() {
		return diagonalDown;
	}

	/**
	 * Sets the downwards diagonal line
	 *
	 * @param diagonalDown
	 *            If true, the downwards diagonal line is used
	 */
	public void setDiagonalDown(boolean diagonalDown) {
		this.diagonalDown = diagonalDown;
	}

	/**
	 * Gets the upwards diagonal line
	 *
	 * @return If true, the upwards diagonal line is used
	 */
	public boolean isDiagonalUp() {
		return diagonalUp;
	}

	/**
	 * Sets the upwards diagonal line
	 *
	 * @param diagonalUp
	 *            If true, the upwards diagonal line is used
	 */
	public void setDiagonalUp(boolean diagonalUp) {
		this.diagonalUp = diagonalUp;
	}

	/**
	 * Gets the color code of the left border. The value is expressed as hex string
	 * with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @return Color code (ARGB)
	 */
	public String getLeftColor() {
		return leftColor;
	}

	/**
	 * Sets the color code of the left border. The value is expressed as hex string
	 * with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @param leftColor
	 *            Color code (ARGB)
	 */
	public void setLeftColor(String leftColor) {
		Fill.validateColor(leftColor, true, true);
		this.leftColor = leftColor;
	}

	/**
	 * Gets the color code of the right border. The value is expressed as hex string
	 * with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @return Color code (ARGB)
	 */
	public String getRightColor() {
		return rightColor;
	}

	/**
	 * Sets the color code of the right border. The value is expressed as hex string
	 * with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @param rightColor
	 *            Color code (ARGB)
	 */
	public void setRightColor(String rightColor) {
		Fill.validateColor(rightColor, true, true);
		this.rightColor = rightColor;
	}

	/**
	 * Gets the color code of the top border. The value is expressed as hex string
	 * with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @return Color code (ARGB)
	 */
	public String getTopColor() {
		return topColor;
	}

	/**
	 * Sets the color code of the top border. The value is expressed as hex string
	 * with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @param topColor
	 *            Color code (ARGB)
	 */
	public void setTopColor(String topColor) {
		Fill.validateColor(topColor, true, true);
		this.topColor = topColor;
	}

	/**
	 * Gets the color code of the bottom border. The value is expressed as hex
	 * string with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @return Color code (ARGB)
	 */
	public String getBottomColor() {
		return bottomColor;
	}

	/**
	 * Sets the color code of the bottom border. The value is expressed as hex
	 * string with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @param bottomColor
	 *            Color code (ARGB)
	 */
	public void setBottomColor(String bottomColor) {
		Fill.validateColor(bottomColor, true, true);
		this.bottomColor = bottomColor;
	}

	/**
	 * Gets the color code of the diagonal lines. The value is expressed as hex
	 * string with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @return Color code (ARGB)
	 */
	public String getDiagonalColor() {
		return diagonalColor;
	}

	/**
	 * Sets the color code of the diagonal lines. The value is expressed as hex
	 * string with the format AARRGGBB. AA (Alpha) is usually FF
	 *
	 * @param diagonalColor
	 *            Color code (ARGB)
	 */
	public void setDiagonalColor(String diagonalColor) {
		Fill.validateColor(diagonalColor, true, true);
		this.diagonalColor = diagonalColor;
	}

	// ### C O N S T R U C T O R S ###

	/**
	 * Default constructor
	 */
	public Border() {
		this.bottomColor = DEFAULT_BORDER_COLOR;
		this.topColor = DEFAULT_BORDER_COLOR;
		this.leftColor = DEFAULT_BORDER_COLOR;
		this.rightColor = DEFAULT_BORDER_COLOR;
		this.diagonalColor = DEFAULT_BORDER_COLOR;
		this.leftStyle = DEFAULT_BORDER_STYLE;
		this.rightStyle = DEFAULT_BORDER_STYLE;
		this.topStyle = DEFAULT_BORDER_STYLE;
		this.bottomStyle = DEFAULT_BORDER_STYLE;
		this.diagonalStyle = DEFAULT_BORDER_STYLE;
		this.diagonalDown = false;
		this.diagonalUp = false;
	}

	// ### M E T H O D S ###

	/**
	 * Method to determine whether the object has no values but the default values
	 * (means: is empty and must not be processed)
	 *
	 * @return True if empty, otherwise false
	 */
	public boolean isEmpty() {
		boolean state = DEFAULT_BORDER_COLOR.equals(this.bottomColor);
        if (!DEFAULT_BORDER_COLOR.equals(this.topColor)) {
			state = false;
		}
		if (!DEFAULT_BORDER_COLOR.equals(this.leftColor)) {
			state = false;
		}
		if (!DEFAULT_BORDER_COLOR.equals(this.rightColor)) {
			state = false;
		}
		if (!DEFAULT_BORDER_COLOR.equals(this.diagonalColor)) {
			state = false;
		}
		if (this.leftStyle != DEFAULT_BORDER_STYLE) {
			state = false;
		}
		if (this.rightStyle != DEFAULT_BORDER_STYLE) {
			state = false;
		}
		if (this.topStyle != DEFAULT_BORDER_STYLE) {
			state = false;
		}
		if (this.bottomStyle != DEFAULT_BORDER_STYLE) {
			state = false;
		}
		if (this.diagonalStyle != DEFAULT_BORDER_STYLE) {
			state = false;
		}
		if (this.diagonalDown) {
			state = false;
		}
		if (this.diagonalUp) {
			state = false;
		}
		return state;
	}

	/**
	 * Override toString method
	 *
	 * @return String of a class
	 */
	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder();
		sb.append("\"Border\": {\n");
		addPropertyAsJson(sb, "BottomStyle", bottomStyle);
		addPropertyAsJson(sb, "DiagonalColor", diagonalColor);
		addPropertyAsJson(sb, "DiagonalDown", diagonalDown);
		addPropertyAsJson(sb, "DiagonalStyle", diagonalStyle);
		addPropertyAsJson(sb, "DiagonalUp", diagonalUp);
		addPropertyAsJson(sb, "LeftColor", leftColor);
		addPropertyAsJson(sb, "LeftStyle", leftStyle);
		addPropertyAsJson(sb, "RightColor", rightColor);
		addPropertyAsJson(sb, "RightStyle", rightStyle);
		addPropertyAsJson(sb, "TopColor", topColor);
		addPropertyAsJson(sb, "TopStyle", topStyle);
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
	public Border copy() {
		Border copy = new Border();
		copy.setBottomColor(this.bottomColor);
		copy.setBottomStyle(this.bottomStyle);
		copy.setDiagonalColor(this.diagonalColor);
		copy.setDiagonalDown(this.diagonalDown);
		copy.setDiagonalStyle(this.diagonalStyle);
		copy.setDiagonalUp(this.diagonalUp);
		copy.setLeftColor(this.leftColor);
		copy.setLeftStyle(this.leftStyle);
		copy.setRightColor(this.rightColor);
		copy.setRightStyle(this.rightStyle);
		copy.setTopColor(this.topColor);
		copy.setTopStyle(this.topStyle);
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
		int result = leftStyle != null ? leftStyle.hashCode() : 1;
		result = 31 * result + (rightStyle != null ? rightStyle.hashCode() : 2);
		result = 31 * result + (topStyle != null ? topStyle.hashCode() : 4);
		result = 31 * result + (bottomStyle != null ? bottomStyle.hashCode() : 8);
		result = 31 * result + (diagonalStyle != null ? diagonalStyle.hashCode() : 16);
		result = 31 * result + (diagonalDown ? 1 : 32);
		result = 31 * result + (diagonalUp ? 1 : 64);
		result = 31 * result + (leftColor != null ? leftColor.hashCode() : 128);
		result = 31 * result + (rightColor != null ? rightColor.hashCode() : 256);
		result = 31 * result + (topColor != null ? topColor.hashCode() : 512);
		result = 31 * result + (bottomColor != null ? bottomColor.hashCode() : 1024);
		result = 31 * result + (diagonalColor != null ? diagonalColor.hashCode() : 2048);
		return result;
	}
}
