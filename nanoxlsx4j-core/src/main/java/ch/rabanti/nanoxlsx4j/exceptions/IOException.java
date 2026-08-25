/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.exceptions;

public class IOException extends RuntimeException {

    /**
     * Default constructor
     */
    public IOException() {
        super();
    }

    /**
     * Constructor with passed message
     *
     * @param message Message of the exception
     */
    public IOException(String message) {
        super(message);
    }

    /**
     * Constructor with passed message and inner exception
     *
     * @param message Message of the exception
     * @param cause   Inner exception
     */
    public IOException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructor with inner exception
     *
     * @param cause Inner exception
     */
    public IOException(Throwable cause) {
        super(cause);
    }

}
