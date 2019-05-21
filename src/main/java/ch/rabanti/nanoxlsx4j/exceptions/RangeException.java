/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2019
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.exceptions;

/**
 * Class for exceptions regarding range incidents (e.g out-of-range)
 * @author Raphael Stoeckli
 */
public class RangeException extends RuntimeException{
    
    private String exceptionTitle;
 
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
    public RangeException()
    {
        super();
    }
    
    /**
     * Constructor with passed message
     * @param title Title of the exception
     * @param message Message of the exception
     */    
    public RangeException(String title, String message)
    {
        super(title + ": " + message);
        this.exceptionTitle = title;
    }
    
    
}
