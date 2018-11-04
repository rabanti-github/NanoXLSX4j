/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.exceptions;

/**
 * Class for exceptions regarding stream or save error incidents
 * @author Raphael Stoeckli
 */
public class IOException extends Exception{
    
    private Exception innerException;
    private String exceptionTitle;

     /**
     * Gets the inner exception
     * @return Inner exception
     */
    public Exception getInnerException() {
        return innerException;
    }
    
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
    public IOException()
    {
        super();
    }
    
    /**
     * Constructor with passed message
     * @param title Title of the exception
     * @param message Message of the exception
     */    
    public IOException(String title, String message)
    {
        super(title + ": " + message);
        this.exceptionTitle = title;
    }
    
    /**
     * Constructor with passed message and inner exception
     * @param title Title of the exception
     * @param message Message of the exception
     * @param inner Inner exception
     */    
    public  IOException(String title, String message, Exception inner)
    {
        super(title + ": " + message);
        this.innerException = inner;
        this.exceptionTitle = title;
    }
    
    
}
