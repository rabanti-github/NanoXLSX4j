/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.exceptions;

/**
 * Class for exceptions regarding Styles
 * @author Raphael Stoeckli
 */
public class StyleException extends RuntimeException{

    public static final String MISSING_REFERENCE = "A reference is missing in the style definition";
    public static final String GENERAL = "A general style exception occurred";
    public static final String NOT_SUPPORTED = "A not supported style component could not be handled";

    private final String exceptionTitle;

    /**
     * Gets the title of the exception
     * @return Title as string
     */
    public String getExceptionTitle() {
        return exceptionTitle;
    }
    
    
    /**
     * Default constructor
     */    
    public StyleException()
    {
        super();
        this.exceptionTitle = GENERAL;
    }
    
    /**
     * Constructor with passed message
     * @param title Title of the exception
     * @param message Message of the exception
     */    
    public StyleException(String title, String message)
    {
        super(title + ": " + message);
        this.exceptionTitle = title;
    }

    /**
     * Constructor with passed message
     * @param title Title of the exception
     * @param message Message of the exception
     * @param inner Inner exception
     */
    public StyleException(String title, String message, Exception inner)
    {
        super(title + ": " + message);
        this.exceptionTitle = title;
    }
}
