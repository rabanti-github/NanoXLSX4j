//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.exceptions;

/**
 * Class for exceptions regarding plug-ins and their configuration. These exceptions should only occur on faulty
 * configured plug-in packages.
 * <p>
 * Java equivalent of the .NET {@code PackageException}. Renamed to {@code PluginException} to reflect the plug-in
 * architecture context and to avoid ambiguity with the Java {@code package} keyword.
 * </p>
 *
 * @author Raphael Stoeckli
 */
public class PluginException extends Exception {

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
    public PluginException() {
        super();
        this.innerException = null;
    }

    /**
     * Constructor with passed message
     *
     * @param message Message of the exception
     */
    public PluginException(String message) {
        super(message);
        this.innerException = null;
    }

    /**
     * Constructor with passed message and inner exception
     *
     * @param message Message of the exception
     * @param inner   Inner exception
     */
    public PluginException(String message, Exception inner) {
        super(message);
        this.innerException = inner;
    }
}
