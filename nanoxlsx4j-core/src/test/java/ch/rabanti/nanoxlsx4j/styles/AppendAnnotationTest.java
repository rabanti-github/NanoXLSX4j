/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

class AppendAnnotationTest {

    @ParameterizedTest
    @DisplayName("Test of the AppendAttribute, applied to a dummy class")
    @CsvSource({
            "appendProperty, true, false, false",
            "appendPropertyNonIgnore, true, false, false",
            "ignoreProperty, true, true, false",
            "nestedProperty, true, false, true",
            "nonNestedProperty, true, false, false",
            "undefinedProperty, false, false, false"
    })
    void appendAttributeTest1(String propertyName, boolean expectedAttribute, boolean expectedIgnore,
            boolean expectedNested) {
        Field[] fields = DummyClass.class.getDeclaredFields();
        boolean propertyFound = false;
        for (Field field : fields) {
            if (field.getName().equals(propertyName)) {
                Annotation[] annotations = field.getAnnotations();
                boolean attributeFound = false;
                for (Annotation annotation : annotations) {
                    if (annotation instanceof AppendAnnotation appendAnnotation) {
                        assertTrue(expectedAttribute);
                        assertEquals(expectedIgnore, appendAnnotation.ignore());
                        assertEquals(expectedNested, appendAnnotation.nestedProperty());
                        attributeFound = true;
                    }
                }
                propertyFound = true;
                if (expectedAttribute) {
                    assertTrue(attributeFound);
                }
            }
        }
        if (expectedAttribute) {
            assertTrue(propertyFound);
        }
    }

    private static class DummyClass {

        @AppendAnnotation
        private int appendProperty;

        @AppendAnnotation(ignore = false)
        private int appendPropertyNonIgnore;

        @AppendAnnotation(ignore = true)
        private int ignoreProperty;

        @AppendAnnotation(nestedProperty = true)
        private int nestedProperty;

        @AppendAnnotation(nestedProperty = false)
        private int nonNestedProperty;

        private int undefinedProperty;
    }
}
