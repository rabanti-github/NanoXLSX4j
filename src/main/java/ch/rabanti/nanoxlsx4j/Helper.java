/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

import java.util.Calendar;
import java.util.Date;

/**
 * Class for shared used (static) methods
 * @author Raphael Stoeckli
 */
public class Helper {
        
// ### C O N S T A N T S ###    
    private static final long ROOT_TICKS;
    static
    {
        Calendar rootCalendar = Calendar.getInstance();
        rootCalendar.set(1899, Calendar.DECEMBER, 30,0,0,0);
        ROOT_TICKS = rootCalendar.getTimeInMillis();
    }
    
// ### S T A T I C   M E T H O D S ###    
    /**
     * Method to calculate the OA date (OLE automation) of the passed date.<br>
     * OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond this date cannot be handled by Excel under normal circumstances and will throw a FormatException
     * @param date Date to convert
     * @exception FormatException Throws a FormatException if the passed date cannot be translated to the OADate format
     * @return OA date
     */
    public static double getOADate(Date date)
    {
        Calendar dateCal = Calendar.getInstance();
        dateCal.setTime(date);
        long currentTicks = dateCal.getTimeInMillis();
        double d = ((double)(dateCal.get(Calendar.SECOND) + (dateCal.get(Calendar.MINUTE) * 60) + (dateCal.get(Calendar.HOUR_OF_DAY) * 3600)) / 86400) + Math.floor((currentTicks - ROOT_TICKS) / (86400000));
        if (d < 0)
        {
            throw new FormatException("FormatException","The date is not in a valid range for Excel. Dates before 1900-01-01 are not allowed.");
        }
        return d;
    }

    /**
     * Method to calculate a common Date from the OA date (OLE automation) format<br>
     * OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond this date cannot be handled by Excel under normal circumstances and will throw a FormatException
     * @param oaDate OA date number
     * @return Converted date
     */
    public static Date getDateFromOA(double oaDate)
    {
        double  remainder = oaDate - (long)oaDate;
        double hours = remainder * 24;
        double minutes = (hours - (long)hours) * 60;
        double seconds = (minutes- (long)minutes) * 60;
        Calendar dateCal = Calendar.getInstance();
        dateCal.setTimeInMillis(ROOT_TICKS);
        dateCal.add(Calendar.DATE,(int)oaDate);
        dateCal.add(Calendar.HOUR,(int)hours);
        dateCal.add(Calendar.MINUTE,(int)minutes);
        dateCal.add(Calendar.SECOND,(int)seconds + 1); // Rounding error
        return dateCal.getTime();
    }

    /**
     * Method of a string to check whether its reference is null or the content is empty
     * @param value value / reference to check
     * @return True if the passed parameter is null or empty, otherwise false
     */
    public static boolean isNullOrEmpty(String value)
    {
        if (value == null) { return true; }
        return value.isEmpty();
    }
   
    
}
