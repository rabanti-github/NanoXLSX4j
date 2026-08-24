/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.exceptions;

/**
 * Exception for errors related to styles.
 */
public class StyleException extends RuntimeException {

    /**
     * Creates a new style exception.
     */
    public StyleException() {
        super();
    }

    /**
     * Creates a new style exception with the specified message.
     *
     * @param message the detail message
     */
    public StyleException(String message) {
        super(message);
    }

    /**
     * Creates a new style exception with the specified message and cause.
     *
     * @param message the detail message
     * @param cause the cause of this exception
     */
    public StyleException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Creates a new style exception with the specified cause.
     *
     * @param cause the cause of this exception
     */
    public StyleException(Throwable cause) {
        super(cause);
    }
}
