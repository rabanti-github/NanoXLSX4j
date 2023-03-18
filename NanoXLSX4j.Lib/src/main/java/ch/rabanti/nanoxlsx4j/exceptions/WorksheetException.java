/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.exceptions;

/**
 * Class for exceptions regarding worksheets
 *
 * @author Raphael Stoeckli
 */
public class WorksheetException extends RuntimeException {

	/**
	 * Default constructor
	 */
	public WorksheetException() {
		super();
	}

	/**
	 * Constructor with passed message
	 *
	 * @param message
	 *            Message of the exception
	 */
	public WorksheetException(String message) {
		super(message);
	}

}
