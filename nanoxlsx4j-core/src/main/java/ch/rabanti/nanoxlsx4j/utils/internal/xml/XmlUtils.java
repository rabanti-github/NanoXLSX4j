/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils.internal.xml;

/**
 * Provides low-level helpers for XML values used while reading or writing XLSX package parts.
 */
public final class XmlUtils {

    private XmlUtils() {
    }

    /**
     * Replaces characters that are not legal in XML 1.0 with spaces. Markup characters such as {@code <}, {@code >},
     * and {@code &} are retained because XML writers escape them. Valid supplementary Unicode code points through
     * {@code U+10FFFF} are retained, while isolated UTF-16 surrogate code units are replaced.
     *
     * @param input Input value, or {@code null}
     * @return Sanitized value; an empty string when the input is {@code null}
     */
    public static String sanitizeXmlValue(String input) {
        if (input == null) {
            return "";
        }

        StringBuilder sanitized = null;
        int unchangedStart = 0;
        for (int index = 0; index < input.length(); ) {
            int codePoint = input.codePointAt(index);
            int codePointLength = Character.charCount(codePoint);
            if (!isValidXmlCharacter(codePoint)) {
                if (sanitized == null) {
                    sanitized = new StringBuilder(input.length());
                }
                sanitized.append(input, unchangedStart, index);
                sanitized.append(' ');
                unchangedStart = index + codePointLength;
            }
            index += codePointLength;
        }

        if (sanitized == null) {
            return input;
        }
        sanitized.append(input, unchangedStart, input.length());
        return sanitized.toString();
    }

    /**
     * Returns whether a Unicode code point is allowed by the XML 1.0 {@code Char} production.
     *
     * @param codePoint Unicode code point
     * @return {@code true} if the code point is legal in XML 1.0; otherwise {@code false}
     */
    private static boolean isValidXmlCharacter(int codePoint) {
        return codePoint == 0x9
                || codePoint == 0xA
                || codePoint == 0xD
                || codePoint >= 0x20 && codePoint <= 0xD7FF
                || codePoint >= 0xE000 && codePoint <= 0xFFFD
                || codePoint >= 0x10000 && codePoint <= 0x10FFFF;
    }
}
