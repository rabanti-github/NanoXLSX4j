/*
 * NanoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.demo.testing;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;

import java.util.Date;

/**
 * Class for performance tests
 */
public class Performance {

    /**
     * Method to test and measure the performance if a huge number of date vales is inserted as new cells.
     *
     * @param fileName     Filename of the output
     * @param sheetName    Worksheet name
     * @param numberOfRows Number of generated rows
     */
    public static void dateStressTest(String fileName, String sheetName, int numberOfRows) {

        Workbook wb = new Workbook(fileName, sheetName);
        for (int i = 0; i < numberOfRows; i++) {
            wb.WS.value(new Date());
            wb.WS.value("XYZ", BasicStyles.Bold());
            wb.WS.down();
        }
        try {
            wb.save();
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
    }

}