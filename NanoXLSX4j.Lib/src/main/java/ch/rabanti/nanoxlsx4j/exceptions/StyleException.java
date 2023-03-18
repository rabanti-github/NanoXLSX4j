/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.exceptions;

/**
 * Class for exceptions regarding Styles
 *
 * @author Raphael Stoeckli
 */
public class StyleException extends RuntimeException {

	/**
	 * Default constructor
	 */
	public StyleException() {
		super();
	}

	/**
	 * Constructor with passed message
	 *
	 * @param message
	 *            Message of the exception
	 */
	public StyleException(String message) {
		super(message);
	}

	/**
	 * Constructor with passed message
	 *
	 * @param message
	 *            Message of the exception
	 * @param inner
	 *            Inner exception
	 */
	public StyleException(String message, Exception inner) {
		super(message);
	}
}
