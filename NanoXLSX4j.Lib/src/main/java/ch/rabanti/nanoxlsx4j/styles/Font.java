/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;

/**
 * Class representing a Font entry. The Font entry is used to define text
 * formatting
 *
 * @author Raphael Stoeckli
 */
public class Font extends AbstractStyle {
	// ### E N U M S ###
	/**
	 * Default font family as constant
	 */
	public static final String DEFAULT_FONT_NAME = "Calibri";

	/**
	 * Maximum possible font size
	 */
	public static final float MIN_FONT_SIZE = 1f;

	/**
	 * Minimum possible font size
	 */
	public static final float MAX_FONT_SIZE = 409f;

	/**
	 * Default font size
	 */
	public static final float DEFAULT_FONT_SIZE = 11f;

	/**
	 * Default font family
	 */
	public static final String DEFAULT_FONT_FAMILY = "2";

	/**
	 * Default font scheme
	 */
	public static final SchemeValue DEFAULT_FONT_SCHEME = SchemeValue.minor;

	/**
	 * Default vertical alignment
	 */
	public static final VerticalAlignValue DEFAULT_VERTICAL_ALIGN = VerticalAlignValue.none;

	/**
	 * Enum for the vertical alignment of the text from base line
	 */
	public enum VerticalAlignValue {
		// baseline, // Maybe not used in Excel
		/**
		 * Text will be rendered as subscript
		 */
		subscript(1),
		/**
		 * Text will be rendered as superscript
		 */
		superscript(2),
		/**
		 * Text will be rendered normal
		 */
		none(0);

		private final int value;

		VerticalAlignValue(int value) {
			this.value = value;
		}

		public int getValue() {
			return value;
		}
	}

	/**
	 * Enum for the font scheme
	 */
	public enum SchemeValue {
		/**
		 * Font scheme is major
		 */
		major(1),
		/**
		 * Font scheme is minor (default)
		 */
		minor(2),
		/**
		 * No Font scheme is used
		 */
		none(0);

		private final int value;

		SchemeValue(int value) {
			this.value = value;
		}
	}

	/**
	 * Enum for the style of the underline property of a stylized text
	 */
	public enum UnderlineValue {
		/**
		 * Text contains a single underline
		 */
		u_single,
		/**
		 * Text contains a double underline
		 */
		u_double,
		/**
		 * Text contains a single, accounting underline
		 */
		singleAccounting,
		/**
		 * Text contains a double, accounting underline
		 */
		doubleAccounting,
		/**
		 * Text contains no underline (default)
		 */
		none,
	}

	// ### P R I V A T E F I E L D S ###
	private boolean bold;
	private boolean italic;
	private boolean strike;
	private UnderlineValue underline;
	private float size;

	// OOXML: Chp.18.8.29
	private String name;

	// TODO: v3> Refactor to enum according to specs (18.18.94)
	// OOXML: Chp.18.8.18 and 18.18.94
	private String family;

	// TODO: V3> Refactor to enum according to specs
	// OOXML: Chp.18.8.3 and 20.1.6.2(p2839ff)
	private int colorTheme;
	private String colorValue;
	private SchemeValue scheme;
	private VerticalAlignValue verticalAlign;

	// TODO: v3> Refactor to enum according to specs
	// OOXML: Chp.19.2.1.13
	private String charset;

	// ### G E T T E R S & S E T T E R S ###

	/**
	 * Gets the bold parameter of the font
	 *
	 * @return If true, the font is bold
	 */
	public boolean isBold() {
		return bold;
	}

	/**
	 * Sets the bold parameter of the font
	 *
	 * @param bold
	 *            If true, the font is bold
	 */
	public void setBold(boolean bold) {
		this.bold = bold;
	}

	/**
	 * Gets the italic parameter of the font
	 *
	 * @return If true, the font is italic
	 */
	public boolean isItalic() {
		return italic;
	}

	/**
	 * Sets the italic parameter of the font
	 *
	 * @param italic
	 *            If true, the font is italic
	 */
	public void setItalic(boolean italic) {
		this.italic = italic;
	}

	/**
	 * Gets whether the font is struck through
	 *
	 * @return If true, the font is declared as strike-through
	 */
	public boolean isStrike() {
		return strike;
	}

	/**
	 * Sets whether the font is struck through
	 *
	 * @param strike
	 *            If true, the font is declared as strike-through
	 */
	public void setStrike(boolean strike) {
		this.strike = strike;
	}

	/**
	 * Gets the underline style of the font
	 *
	 * @return Underline value
	 * @apiNote If set to {@link UnderlineValue#none} no underline will be applied
	 *          (default)
	 */
	public UnderlineValue getUnderline() {
		return underline;
	}

	/**
	 * Sets the underline style of the font
	 *
	 * @param underline
	 *            Underline value
	 * @apiNote If set to {@link UnderlineValue#none} no underline will be applied
	 *          (default)
	 */
	public void setUnderline(UnderlineValue underline) {
		this.underline = underline;
	}

	/**
	 * Gets the font size. Valid range is from 1.0 to 409.0
	 *
	 * @return Font size
	 */
	public float getSize() {
		return size;
	}

	/**
	 * Sets the Font size. Valid range is from 1 to 409
	 *
	 * @param size
	 *            Font size
	 */
	public void setSize(float size) {
		if (size < MIN_FONT_SIZE) {
			this.size = MIN_FONT_SIZE;
		}
		else if (size > MAX_FONT_SIZE) {
			this.size = MAX_FONT_SIZE;
		}
		else {
			this.size = size;
		}
	}

	/**
	 * Gets the font name (Default is Calibri)
	 *
	 * @return Font name
	 */
	public String getName() {
		return name;
	}

	/**
	 * Sets the font name (Default is Calibri)
	 *
	 * @param name
	 *            Font name
	 * @throws StyleException
	 *             thrown if the name is null or empty.
	 * @apiNote Note that the font name is not validated whether it is a valid or
	 *          existing font. The font name may not exceed more than 31 characters
	 */
	public void setName(String name) {
		if ((name == null || name.isEmpty()) && !StyleRepository.getInstance().isImportInProgress()) {
			throw new StyleException("The font name was null or empty");
		}
		this.name = name;
	}

	/**
	 * Gets the font family (Default is 2)
	 *
	 * @return Font family
	 */
	public String getFamily() {
		return family;
	}

	/**
	 * Sets the font family (Default is 2 = Swiss)
	 *
	 * @param family
	 *            Font family
	 */
	public void setFamily(String family) {
		this.family = family;
	}

	/**
	 * Gets the font color theme (Default is 1 = Light)
	 *
	 * @return Font color theme
	 */
	public int getColorTheme() {
		return colorTheme;
	}

	/**
	 * Sets the font color theme (Default is 1)
	 *
	 * @param colorTheme
	 *            Font color theme
	 * @throws StyleException
	 *             thrown if the number is below 0
	 */
	public void setColorTheme(int colorTheme) {
		if (colorTheme < 0) {
			throw new StyleException("The color theme number " + colorTheme + " is invalid. Should be >=0");
		}
		this.colorTheme = colorTheme;
	}

	/**
	 * Gets the Font color (default is empty)
	 *
	 * @return Font color
	 */
	public String getColorValue() {
		return colorValue;
	}

	/**
	 * Sets the color code of the font color. The value is expressed as hex string
	 * with the format AARRGGBB. AA (Alpha) is usually FF.<br>
	 * To omit the color, an empty string can be set. Empty is also default.
	 *
	 * @param colorValue
	 *            Font color
	 * @throws StyleException
	 *             thrown if the passed ARGB value is not valid
	 */
	public void setColorValue(String colorValue) {
		Fill.validateColor(colorValue, true, true);
		this.colorValue = colorValue;
	}

	/**
	 * Gets the font scheme (Default is minor)
	 *
	 * @return Font scheme
	 */
	public SchemeValue getScheme() {
		return scheme;
	}

	/**
	 * Sets the Font scheme (Default is minor)
	 *
	 * @param scheme
	 *            Font scheme
	 */
	public void setScheme(SchemeValue scheme) {
		this.scheme = scheme;
	}

	/**
	 * Gets the alignment of the font (Default is none)
	 *
	 * @return Alignment of the font
	 */
	public VerticalAlignValue getVerticalAlign() {
		return verticalAlign;
	}

	/**
	 * Sets the Alignment of the font (Default is none)
	 *
	 * @param verticalAlign
	 *            Alignment of the font
	 */
	public void setVerticalAlign(VerticalAlignValue verticalAlign) {
		this.verticalAlign = verticalAlign;
	}

	/**
	 * Gets the charset of the Font (Default is empty)
	 *
	 * @return Charset of the Font
	 */
	public String getCharset() {
		return charset;
	}

	/**
	 * Sets the charset of the Font (Default is empty)
	 *
	 * @param charset
	 *            Charset of the Font
	 */
	public void setCharset(String charset) {
		this.charset = charset;
	}

	/**
	 * Gets whether this object is the default font
	 *
	 * @return In true the font is equals the default font
	 */
	public boolean isDefaultFont() {
		Font temp = new Font();
		return this.equals(temp);
	}

	// ### C O N S T R U C T O R S ###

	/**
	 * Default constructor
	 */
	public Font() {
		this.size = DEFAULT_FONT_SIZE;
		this.name = DEFAULT_FONT_NAME;
		this.family = DEFAULT_FONT_FAMILY;
		this.colorTheme = 1;
		this.colorValue = "";
		this.charset = "";
		this.scheme = DEFAULT_FONT_SCHEME;
		this.verticalAlign = DEFAULT_VERTICAL_ALIGN;
		this.underline = UnderlineValue.none;
	}

	/**
	 * Override toString method
	 *
	 * @return String of a class instance
	 */
	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder();
		sb.append("\"Font\": {\n");
		addPropertyAsJson(sb, "Bold", bold);
		addPropertyAsJson(sb, "Charset", charset);
		addPropertyAsJson(sb, "ColorTheme", colorTheme);
		addPropertyAsJson(sb, "ColorValue", colorValue);
		addPropertyAsJson(sb, "VerticalAlign", verticalAlign);
		addPropertyAsJson(sb, "Family", family);
		addPropertyAsJson(sb, "Italic", italic);
		addPropertyAsJson(sb, "Name", name);
		addPropertyAsJson(sb, "Scheme", scheme);
		addPropertyAsJson(sb, "Size", size);
		addPropertyAsJson(sb, "Strike", strike);
		addPropertyAsJson(sb, "Underline", underline);
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
	public Font copy() {
		Font copy = new Font();
		copy.setBold(this.bold);
		copy.setCharset(this.charset);
		copy.setColorTheme(this.colorTheme);
		copy.setColorValue(this.colorValue);
		copy.setVerticalAlign(this.verticalAlign);
		copy.setFamily(this.family);
		copy.setItalic(this.italic);
		copy.setName(this.name);
		copy.setScheme(this.scheme);
		copy.setSize(this.size);
		copy.setStrike(this.strike);
		copy.setUnderline(this.underline);
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
		int result = (size != +0.0f ? Float.floatToIntBits(size) : 1);
		result = 31 * result + (name != null ? name.hashCode() : 2);
		result = 31 * result + (family != null ? family.hashCode() : 4);
		result = 31 * result + colorTheme;
		result = 31 * result + (colorValue != null ? colorValue.hashCode() : 8);
		result = 31 * result + (scheme != null ? scheme.hashCode() : 16);
		result = 31 * result + (verticalAlign != null ? verticalAlign.hashCode() : 32);
		result = 31 * result + (bold ? 1 : 64);
		result = 31 * result + (italic ? 1 : 128);
		result = 31 * result + (underline != null ? underline.hashCode() : 256);
		result = 31 * result + (strike ? 1 : 512);
		result = 31 * result + (charset != null ? charset.hashCode() : 1024);
		return result;
	}
}
