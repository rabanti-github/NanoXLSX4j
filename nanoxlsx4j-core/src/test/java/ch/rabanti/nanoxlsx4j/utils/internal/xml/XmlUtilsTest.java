/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils.internal.xml;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.ValueSource;

class XmlUtilsTest {

    @DisplayName("Test of the SanitizeXmlValue function")
    @ParameterizedTest
    @CsvSource(value = {
            "NULL, ''",
            "'', ''",
            "' ', ' '",
            "'This is a test & string', 'This is a test & string'",
            "'This is a <tag>', 'This is a <tag>'",
            "'This is a >tag<', 'This is a >tag<'",
            "'This is a \"quoted\" text', 'This is a \"quoted\" text'"
    }, nullValues = "NULL")
    void sanitizeXmlValueTest(String input, String expectedOutput) {
        String result = XmlUtils.sanitizeXmlValue(input);

        assertEquals(expectedOutput, result);
    }

    @DisplayName("Test of the SanitizeXmlValue function with special characters")
    @ParameterizedTest
    @ValueSource(strings = {"\001", "\002", "\020", "\037"})
    void sanitizeXmlValueSpecialCharactersTest(String input) {
        String result = XmlUtils.sanitizeXmlValue(input);

        assertEquals(" ", result);
    }
}
