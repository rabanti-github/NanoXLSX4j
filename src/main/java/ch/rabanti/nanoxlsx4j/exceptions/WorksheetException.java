/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.exceptions;

/**
 * Class for exceptions regarding worksheets
 * @author Raphael Stoeckli
 */
public class WorksheetException extends RuntimeException{

    public static final String GENERAL = "A general worksheet exception occurred";

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
    public WorksheetException()
    {
        super();
        this.exceptionTitle = GENERAL;
    }
    
    /**
     * Constructor with passed message
     * @param title Title of the exception
     * @param message Message of the exception
     */    
    public WorksheetException(String title, String message)
    {
        super(title + ": " + message);
        this.exceptionTitle = title;
    }
    
    
}
