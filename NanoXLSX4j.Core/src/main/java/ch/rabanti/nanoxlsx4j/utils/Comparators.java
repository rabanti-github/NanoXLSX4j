/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.utils;

/**
 * Class providing general comparator methods
 *
 * @author Raphael Stoeckli
 */
public class Comparators {

    /**
     * Threshold used when floats are compared
     */
    private static final float FLOAT_THRESHOLD = 0.00001f;

    /**
     * Threshold used when doubles are compared
     */
    private static final double DOUBLE_THRESHOLD = 1e-12;

    private Comparators() {
        // Prevents class instantiation
    }

    /**
     * Compares two dimensions (e.g. column width, row height).
     * A threshold is applied for equality comparison.
     *
     * @param dimension1 Nullable dimension 1
     * @param dimension2 Nullable dimension 2
     * @return 1 if dimension1 is greater than dimension2, -1 if smaller, 0 if equal.
     * If dimension1 is null, Float.MIN_VALUE is used; same for dimension2.
     */
    public static int compareDimensions(Float dimension1, Float dimension2) {
        if (dimension1 == null) {
            dimension1 = Float.MIN_VALUE;
        }
        if (dimension2 == null) {
            dimension2 = Float.MIN_VALUE;
        }
        if (Math.abs(dimension1 - dimension2) < FLOAT_THRESHOLD) {
            return 0;
        }
        else if (dimension1 > dimension2) {
            return 1;
        }
        else {
            return -1;
        }
    }

    /**
     * Checks whether the passed double value is considered as zero using a defined threshold
     *
     * @param value Value to check
     * @return True if zero (within threshold), otherwise false
     */
    public static boolean isZero(double value) {
        return Math.abs(value) < DOUBLE_THRESHOLD;
    }

    /**
     * Checks whether the passed float value is considered as zero using a defined threshold
     *
     * @param value Value to check
     * @return True if zero (within threshold), otherwise false
     */
    public static boolean isZero(float value) {
        return Math.abs(value) < FLOAT_THRESHOLD;
    }

}
