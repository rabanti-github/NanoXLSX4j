/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.exceptions;

/**
 * Class for exceptions regarding stream or save error incidents
 *
 * @author Raphael Stoeckli
 */
public class IOException extends Exception {

	private final Exception innerException;

	/**
	 * Gets the inner exception
	 *
	 * @return Inner exception
	 */
	public Exception getInnerException() {
		return innerException;
	}

	/**
	 * Default constructor
	 */
	public IOException() {
		super();
		this.innerException = null;
	}

	/**
	 * Constructor with passed message
	 *
	 * @param message
	 *            Message of the exception
	 */
	public IOException(String message) {
		super(message);
		this.innerException = null;
	}

	/**
	 * Constructor with passed message and inner exception
	 *
	 * @param message
	 *            Message of the exception
	 * @param inner
	 *            Inner exception
	 */
	public IOException(String message, Exception inner) {
		super(message);
		this.innerException = inner;
	}

}
