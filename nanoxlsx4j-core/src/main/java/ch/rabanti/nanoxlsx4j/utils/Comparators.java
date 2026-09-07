/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.utils;

/**
 * Class providing general comparator methods.
 */
public class Comparators {

    private static final float FLOAT_THRESHOLD = 0.00001f;
    private static final double DOUBLE_THRESHOLD = 1e-12;

    private Comparators() {
        // Do not instantiate
    }

    /**
     * Compares whether the content of two password character arrays is equal.
     *
     * @param value1 Password character array one
     * @param value2 Password character array two
     * @return True if the content of the two arrays is equal, otherwise false
     */
    public static boolean compareSecureStrings(char[] value1, char[] value2) {
        boolean value1Empty = value1 == null || value1.length == 0;
        boolean value2Empty = value2 == null || value2.length == 0;
        if (value1Empty != value2Empty) {
            return false;
        }
        if (value1Empty) {
            return true;
        }
        if (value1.length != value2.length) {
            return false;
        }
        for (int i = 0; i < value1.length; i++) {
            if (value1[i] != value2[i]) {
                return false;
            }
        }
        return true;
    }

    /**
     * Compares two dimensions, for example a column width or row height.
     *
     * @param dimension1 Nullable dimension one
     * @param dimension2 Nullable dimension two
     * @return 1 if dimension one is greater, -1 if dimension one is smaller, otherwise 0
     */
    public static int compareDimensions(Float dimension1, Float dimension2) {
        float value1 = dimension1 == null ? -Float.MAX_VALUE : dimension1;
        float value2 = dimension2 == null ? -Float.MAX_VALUE : dimension2;
        if (Math.abs(value1 - value2) < FLOAT_THRESHOLD) {
            return 0;
        } else if (value1 > value2) {
            return 1;
        } else {
            return -1;
        }
    }

    /**
     * Checks whether the passed double value is considered zero using a defined threshold.
     *
     * @param value Value to check
     * @return True if zero, otherwise false
     */
    public static boolean isZero(double value) {
        return Math.abs(value) < DOUBLE_THRESHOLD;
    }

    /**
     * Checks whether the passed float value is considered zero using a defined threshold.
     *
     * @param value Value to check
     * @return True if zero, otherwise false
     */
    public static boolean isZero(float value) {
        return Math.abs(value) < FLOAT_THRESHOLD;
    }
}
