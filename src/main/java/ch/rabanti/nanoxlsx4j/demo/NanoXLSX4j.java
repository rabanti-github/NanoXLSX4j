/*
 * NanoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.demo;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.BasicFormulas;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.CellXf;
import ch.rabanti.nanoxlsx4j.styles.Fill;
import ch.rabanti.nanoxlsx4j.styles.Font;
import ch.rabanti.nanoxlsx4j.styles.Style;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * Demo Program for NanoXLSX4j
 *
 * @author Raphael Stoeckli
 */
public class NanoXLSX4j {

    private static String outputFolder = "./demoFiles/";                        // Output folder for all demo files

    /**
     * Method to run all demos
     *
     * @param args the command line arguments (not used)
     */
    public static void main(String[] args) {

        /** PROVIDING OUTPUT FOLDER **/
        if (!Files.exists(Paths.get(outputFolder)))                     // Check existence of output folder
        {
            File dir = new File(outputFolder);                                  // Create new folder if not existing
            dir.mkdirs();
        }

        /** PERFORMANCE TESTING **/
        // Performance.dateStressTest(outputFolder + "stressTest.xlsx", "Dates", 40000); // Only uncomment this to test the library performance
        /* *********************** */

        /** DEMOS **/
        // bug();

        // basicDemo();
        read();
        /*
        shortenerDemo();
        streamDemo();
        demo1();
        demo2();
        demo3();
        demo4();
        demo5();
        demo6();
        demo7();
        demo8();
        demo9();
        demo10();
        */
    }

    /*
    private static void bug(){


        Workbook data = new Workbook(false);

        Observable mCheckPermissionUseCase = Observable.empty();

        mCheckPermissionUseCase
                .execute(Permission.WriteExternalStorage)
                .doOnTerminate(() -> {
                    if (view != null) {
                        view.progressBarLoading(false);
                    }
                })
                .compose(bindToLifecycle())
                .subscribe(result -> {
                            view.storagePermissionGranted();
                            try (FileOutputStream out = context.openFileOutput(fileName, Context.MODE_PRIVATE)) {
                                data.saveAsStream(out);
                                File file = new File(context.getFilesDir(), fileName);
                                view.openShareFileIntent(file, fileFormat);
                            } catch (Exception ex) {
                                // noop
                            }


                });
    }
*/


    /**
     * This is a very basic demo (adding three values and save the workbook)
     */
    private static void basicDemo() {

        Workbook workbook = new Workbook(outputFolder + "basic.xlsx", "Sheet1"); // Create new workbook
        workbook.getCurrentWorksheet().addNextCell("Test");                 // Add cell A1
        workbook.getCurrentWorksheet().addNextCell("Test2");                // Add cell B1
        workbook.getCurrentWorksheet().addNextCell("Test3");                // Add cell C1
        try {
            workbook.save();
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }


    /**
     * This is a demo to read the previously created basic.xlsx file
     */
    private static void read() {
        try {
            Workbook wb = Workbook.load(outputFolder + "basic.xlsx");
            System.out.println("contains worksheet name: " + wb.getCurrentWorksheet().getSheetName());
            wb.getCurrentWorksheet().getCells().forEach((k, v) -> {
                System.out.println("Cell address: " + k + ": content:'" + v.getValue() + "'");
            });

            // The same as stream
            FileInputStream fi = new FileInputStream(outputFolder + "basic.xlsx");
            Workbook wb2 = Workbook.load(fi);
            fi.close();
            System.out.println("contains worksheet name: " + wb2.getCurrentWorksheet().getSheetName());
            wb2.getCurrentWorksheet().getCells().forEach((k, v) -> {
                System.out.println("Cell address: " + k + ": content:'" + v.getValue() + "'");
            });
        } catch (Exception ex) {
            System.out.println("Error: " + ex);
        }

    }

    /**
     * This method shows the shortened style of writing cells
     */
    private static void shortenerDemo() {
        Workbook wb = new Workbook(outputFolder + "shortenerDemo.xlsx", "Sheet1"); // Create a workbook (important: A worksheet must be created as well)
        wb.WS.value("Some Text");                                                     // Add cell A1
        wb.WS.value(58.55, BasicStyles.DoubleUnderline());                      // Add a formatted value to cell B1
        wb.WS.right(2);                                               // Move to cell E1
        wb.WS.value(true);                                                            // Add cell E1
        wb.addWorksheet("Sheet2");                                              // Add a new worksheet
        wb.getCurrentWorksheet().setCurrentCellDirection(Worksheet.CellDirection.RowToRow); // Change the cell direction
        wb.WS.value("This is another text");                                           // Add cell A1
        wb.WS.formula("=A1");                                                          // Add a formula in Cell A2
        wb.WS.down();                                                                  // Go to cell A4
        wb.WS.value("Formatted Text", BasicStyles.Bold());                        // Add a formatted value to cell A4
        try {
            wb.save();                                                      // Save the workbook
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
    }

    /**
     * This method shows how to save a workbook as stream
     */
    private static void streamDemo() {
        Workbook workbook = new Workbook(true);                             // Create new workbook without file name
        workbook.getCurrentWorksheet().addNextCell("This is an example");   // Add cell A1
        workbook.getCurrentWorksheet().addNextCellFormula("=A1");           // Add formula in cell B1
        workbook.getCurrentWorksheet().addNextCell(123456789);              // Add cell C1
        FileOutputStream fs;                                                // Define a stream
        try {
            fs = new FileOutputStream(outputFolder + "stream.xlsx");        // Create a file output stream (could be whatever output stream you want)
            workbook.saveAsStream(fs);                                      // Save the workbook into the stream
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }
    }

    /**
     * This method shows the usage of AddNextCell with several data types and formulas. Furthermore, the several types of Addresses are demonstrated
     */
    private static void demo1() {
        Workbook workbook = new Workbook(outputFolder + "test1.xlsx", "Sheet1");  // Create new workbook
        workbook.getCurrentWorksheet().addNextCell("Test");                 // Add cell A1
        workbook.getCurrentWorksheet().addNextCell(123);                    // Add cell B1
        workbook.getCurrentWorksheet().addNextCell(true);                   // Add cell C1
        workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 2
        workbook.getCurrentWorksheet().addNextCell(123.456d);               // Add cell A2
        workbook.getCurrentWorksheet().addNextCell(123.789f);               // Add cell B2
        workbook.getCurrentWorksheet().addNextCell(new Date());             // Add cell C2
        workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 3
        workbook.getCurrentWorksheet().addNextCellFormula("B1*22");         // Add cell A3 as formula (B1 times 22)
        workbook.getCurrentWorksheet().addNextCellFormula("ROUNDDOWN(A2,1)"); // Add cell B3 as formula (Floor A2 with one decimal place)
        workbook.getCurrentWorksheet().addNextCellFormula("PI()");          // Add cell C3 as formula (Pi = 3.14.... )
        workbook.addWorksheet("Addresses");                                                    // Add new worksheet
        workbook.getCurrentWorksheet().setCurrentCellDirection(Worksheet.CellDirection.Disabled);    // Disable automatic addressing
        workbook.getCurrentWorksheet().addCell("Default", 0, 0);       // Add a value
        Address address = new Address(1, 0, Cell.AddressType.Default);                  // Create Address with default behavior
        workbook.getCurrentWorksheet().addCell(address.toString(), 1, 0);    // Add the string of the address
        workbook.getCurrentWorksheet().addCell("Fixed Column", 0, 1); // Add a value
        address = new Address(1, 1, Cell.AddressType.FixedColumn);                      // Create Address with fixed column
        workbook.getCurrentWorksheet().addCell(address.toString(), 1, 1);   // Add the string of the address
        workbook.getCurrentWorksheet().addCell("Fixed Row", 0, 2);    // Add a value
        address = new Address(1, 2, Cell.AddressType.FixedRow);                         // Create Address with fixed row
        workbook.getCurrentWorksheet().addCell(address.toString(), 1, 2);   // Add the string of the address
        workbook.getCurrentWorksheet().addCell("Fixed Row and Column", 0, 3); // Add a value
        address = new Address(1, 3, Cell.AddressType.FixedRowAndColumn);                // Create Address with fixed row and column
        workbook.getCurrentWorksheet().addCell(address.toString(), 1, 3);   // Add the string of the address
        try {
            workbook.save();                                                    // Save the workbook
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * This demo shows the usage of several data types, the method AddCell, more than one worksheet and the SaveAs method
     */
    private static void demo2() {
        Workbook workbook = new Workbook(false);                            // Create new workbook
        workbook.addWorksheet("Sheet1");                                    // Add a new Worksheet and set it as current sheet
        workbook.getCurrentWorksheet().addNextCell("月曜日");                // Add cell A1 (Unicode)
        workbook.getCurrentWorksheet().addNextCell(-987);                   // Add cell B1
        workbook.getCurrentWorksheet().addNextCell(false);                  // Add cell C1
        workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 2
        workbook.getCurrentWorksheet().addNextCell(-123.456d);              // Add cell A2
        workbook.getCurrentWorksheet().addNextCell(-123.789f);              // Add cell B2
        workbook.getCurrentWorksheet().addNextCell(new Date());             // Add cell C3
        workbook.addWorksheet("Sheet2");                                    // Add a new Worksheet and set it as current sheet
        workbook.getCurrentWorksheet().addCell("ABC", "A1");                // Add cell A1
        workbook.getCurrentWorksheet().addCell(779, 2, 1);                  // Add cell C2 (zero based addresses: column 2=C, row 1=2)
        workbook.getCurrentWorksheet().addCell(false, 3, 2);                // Add cell D3 (zero based addresses: column 3=D, row 2=3)
        workbook.getCurrentWorksheet().addNextCell(0);                      // Add cell E3 (direction: column to column)
        List<Object> values = new ArrayList<>();                            // Create a List of mixed values
        values.add("V1");
        values.add(true);
        values.add(16.8);
        workbook.getCurrentWorksheet().addCellRange(values, "A4:C4");       // Add a cell range to A4 - C4
        try {
            workbook.saveAs(outputFolder + "test2.xlsx");                       // Save the workbook
            workbook.saveAs(outputFolder + "test2.xlsx");                       // Save the workbook
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * This demo shows the usage of CellDirection when using AddNextCell, reading of the current cell address, and retrieving of cell values
     */
    private static void demo3() {
        Workbook workbook = new Workbook(outputFolder + "test3.xlsx", "Sheet1"); // Create new workbook
        workbook.getCurrentWorksheet().setCurrentCellDirection(Worksheet.CellDirection.RowToRow);  // Change the cell direction
        workbook.getCurrentWorksheet().addNextCell(1);                      // Add cell A1
        workbook.getCurrentWorksheet().addNextCell(2);                      // Add cell A2
        workbook.getCurrentWorksheet().addNextCell(3);                      // Add cell A3
        workbook.getCurrentWorksheet().addNextCell(4);                      // Add cell A4
        int row = workbook.getCurrentWorksheet().getCurrentRowNumber();    // Get the row number (will be 4 = row row 5)
        int col = workbook.getCurrentWorksheet().getCurrentColumnNumber(); // Get the column number (will be 0 = column A)
        workbook.getCurrentWorksheet().addNextCell("This cell has the row number " + (row + 1) + " and column number " + (col + 1));
        workbook.getCurrentWorksheet().goToNextColumn();                    // Go to Column B
        workbook.getCurrentWorksheet().addNextCell("A");                    // Add cell B1
        workbook.getCurrentWorksheet().addNextCell("B");                    // Add cell B2
        workbook.getCurrentWorksheet().addNextCell("C");                    // Add cell B3
        workbook.getCurrentWorksheet().addNextCell("D");                    // Add cell B4
        workbook.getCurrentWorksheet().removeCell("A2");                  // Delete cell A2
        workbook.getCurrentWorksheet().removeCell(1, 1);  // Delete cell B2
        workbook.getCurrentWorksheet().goToNextRow(3);              // Move 3 rows down
        Object value = workbook.getCurrentWorksheet().getCell(1, 2).getValue();  // Gets the value of cell B3
        workbook.getCurrentWorksheet().addNextCell("Value of B3 is: " + value);
        workbook.getCurrentWorksheet().setCurrentCellDirection(Worksheet.CellDirection.Disabled);   // Disable automatic cell addressing
        workbook.getCurrentWorksheet().addCell("Text A", 3, 0);       // Add manually placed value
        workbook.getCurrentWorksheet().addCell("Text B", 4, 1);       // Add manually placed value
        workbook.getCurrentWorksheet().addCell("Text C", 3, 2);       // Add manually placed value
        try {
            workbook.save();                                                    // Save the workbook
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * This demo shows the usage of several styles, column widths and row heights
     */
    private static void demo4() {
        Workbook workbook = new Workbook(outputFolder + "test4.xlsx", "Sheet1"); // Create new workbook
        List<Object> values = new ArrayList<>();                            // Create a List of values
        values.add("Header1");
        values.add("Header2");
        values.add("Header3");
        workbook.getCurrentWorksheet().addCellRange(values, new Address(0, 0), new Address(2, 0));         // Add a cell range to A4 - C4
        workbook.getCurrentWorksheet().getCells().get("A1").setStyle(BasicStyles.Bold());                // Assign predefined basic style to cell
        workbook.getCurrentWorksheet().getCells().get("B1").setStyle(BasicStyles.Bold());                // Assign predefined basic style to cell
        workbook.getCurrentWorksheet().getCells().get("C1").setStyle(BasicStyles.Bold());                // Assign predefined basic style to cell
        workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 2
        workbook.getCurrentWorksheet().addNextCell(new Date());            // Add cell A2
        workbook.getCurrentWorksheet().addNextCell(2);                      // Add cell B2
        workbook.getCurrentWorksheet().addNextCell(3);                      // Add cell C2
        workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 3
        workbook.getCurrentWorksheet().addNextCell(new Date());             // Add cell A3
        workbook.getCurrentWorksheet().addNextCell("B");                    // Add cell B3
        workbook.getCurrentWorksheet().addNextCell("C");                    // Add cell C3

        Style s = new Style();                                                 // Create new style
        s.getFill().setColor("FF22FF11", Fill.FillType.fillColor);       // Set fill color
        s.getFont().setUnderline(Font.UnderlineValue.u_double);                // Set double underline
        s.getCellXf().setHorizontalAlign(CellXf.HorizontalAlignValue.center);  // Set alignment

        Style s2 = s.copyStyle();                                           // Copy the previously defined style
        s2.getFont().setItalic(true);                                       // Change an attribute of the copied style

        workbook.getCurrentWorksheet().getCells().get("B2").setStyle(s);    // Assign style to cell
        workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 3
        workbook.getCurrentWorksheet().addNextCell(new Date(115, 9, 3));    // Add cell B1
        workbook.getCurrentWorksheet().addNextCell(true);                   // Add cell B2
        workbook.getCurrentWorksheet().addNextCell(false, s2);              // Add cell B3 with style in the same step
        workbook.getCurrentWorksheet().getCells().get("C2").setStyle(BasicStyles.BorderFrame());        // Assign predefined basic style to cell

        Style s3 = BasicStyles.Strike();                                    // Create a style from a predefined style
        s3.getCellXf().setTextRotation(45);                                 // Set text rotation
        s3.getCellXf().setVerticalAlign(CellXf.VerticalAlignValue.center);  // Set alignment

        workbook.getCurrentWorksheet().getCells().get("B4").setStyle(s3);   // Assign style to cell

        workbook.getCurrentWorksheet().setColumnWidth(0, 20f);              // Set column width
        workbook.getCurrentWorksheet().setColumnWidth(1, 15f);              // Set column width
        workbook.getCurrentWorksheet().setColumnWidth(2, 25f);              // Set column width
        workbook.getCurrentWorksheet().setRowHeight(0, 20);                 // Set row height
        workbook.getCurrentWorksheet().setRowHeight(1, 30);                 // Set row height
        try {
            workbook.save();                                                    // Save the workbook
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * This demo shows the usage of cell ranges, adding and removing styles, and meta data
     */
    private static void demo5() {
        Workbook workbook = new Workbook(outputFolder + "test5.xlsx", "Sheet1"); // Create new workbook
        List<Object> values = new ArrayList<>();                            // Create a List of values
        values.add("Header1");
        values.add("Header2");
        values.add("Header3");
        workbook.getCurrentWorksheet().setActiveStyle(BasicStyles.BorderFrameHeader());  // Assign predefined basic style as active style
        workbook.getCurrentWorksheet().addCellRange(values, "A1:C1"); // Add cell range

        values = new ArrayList<>();                                         // Create a List of values
        values.add("Cell A2");
        values.add("Cell B2");
        values.add("Cell C2");
        workbook.getCurrentWorksheet().setActiveStyle(BasicStyles.BorderFrame());  // Assign predefined basic style as active style
        workbook.getCurrentWorksheet().addCellRange(values, "A2:C2"); // Add cell range

        values = new ArrayList<>();                                         // Create a List of values
        values.add("Cell A3");
        values.add("Cell B3");
        values.add("Cell C3");
        workbook.getCurrentWorksheet().addCellRange(values, "A3:C3"); // Add cell range

        values = new ArrayList<>();                                         // Create a List of values
        values.add("Cell A4");
        values.add("Cell B4");
        values.add("Cell C4");
        workbook.getCurrentWorksheet().clearActiveStyle();                  // Clear the active style
        workbook.getCurrentWorksheet().addCellRange(values, "A4:C4");       // Add cell range

        workbook.getWorkbookMetadata().setTitle("Test 5");                           // Add meta data to workbook
        workbook.getWorkbookMetadata().setSubject("This is the 5th PicoXLSX test");  // Add meta data to workbook
        workbook.getWorkbookMetadata().setCreator("PicoXLSX");                       // Add meta data to workbook
        workbook.getWorkbookMetadata().setKeywords("Keyword1;Keyword2;Keyword3");    // Add meta data to workbook
        try {
            workbook.save();                                                    // Save the workbook
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * This demo shows the usage of merging cells, protecting cells, worksheet password protection and workbook protection
     */
    private static void demo6() {
        Workbook workbook = new Workbook(outputFolder + "test6.xlsx", "Sheet1");  // Create new workbook
        workbook.getCurrentWorksheet().addNextCell("Merged1");             // Add cell A1
        workbook.getCurrentWorksheet().mergeCells("A1:C1");                 // Merge cells from A1 to C1
        workbook.getCurrentWorksheet().goToNextRow();                       // Go to next row
        workbook.getCurrentWorksheet().addNextCell(false);                  // Add cell A2
        workbook.getCurrentWorksheet().mergeCells("A2:D2");                 // Merge cells from A2 to D1
        workbook.getCurrentWorksheet().goToNextRow();                       // Go to next row
        workbook.getCurrentWorksheet().addNextCell("22.2d");                // Add cell A3
        workbook.getCurrentWorksheet().mergeCells("A3:E4");                 // Merge cells from A3 to E4
        workbook.addWorksheet("Protected");                                 // Add a new worksheet
        workbook.getCurrentWorksheet().addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.sort);               // Allow to sort sheet (worksheet is automatically set as protected)
        workbook.getCurrentWorksheet().addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.insertRows);         // Allow to insert rows
        workbook.getCurrentWorksheet().addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.selectLockedCells);  // Allow to select cells (locked cells caused automatically to select unlocked cells)
        workbook.getCurrentWorksheet().addNextCell("Cell A1");              // Add cell A1
        workbook.getCurrentWorksheet().addNextCell("Cell B1");              // Add cell B1
        workbook.getCurrentWorksheet().getCells().get("A1").setCellLockedState(false, true); // Set the locking state of cell A1 (not locked but value is hidden when cell selected)
        workbook.addWorksheet("PWD-Protected");                             // Add a new worksheet
        workbook.getCurrentWorksheet().addCell("This worksheet is password protected. The password is:", 0, 0);  // Add cell A1
        workbook.getCurrentWorksheet().addCell("test123", 0, 1);            // Add cell A2
        workbook.getCurrentWorksheet().setSheetProtectionPassword("test123"); // Set the password "test123"
        workbook.setWorkbookProtection(true, true, true, null);             // Set workbook protection (windows locked, structure locked, no password)
        try {
            workbook.save();                                                    // Save the workbook            
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * This demo shows the usage of hiding rows and columns, auto-filter and worksheet name sanitizing
     */
    private static void demo7() {
        Workbook workbook = new Workbook(false);                            // Create new workbook without worksheet
        String invalidSheetName = "Sheet?1";                                // ? is not allowed in the names of worksheets
        String sanitizedSheetName = Worksheet.sanitizeWorksheetName(invalidSheetName, workbook); // Method to sanitize a worksheet name (replaces ? with _)
        workbook.addWorksheet(sanitizedSheetName);                          // Add new worksheet
        Worksheet ws = workbook.getCurrentWorksheet();                      // Create reference (shortening)
        List<Object> values = new ArrayList<>();                            // Create a List of values
        values.add("Cell A1");                                              // set a value
        values.add("Cell B1");                                              // set a value
        values.add("Cell C1");                                              // set a value
        values.add("Cell D1");                                              // set a value
        ws.addCellRange(values, "A1:D1");                                   // Insert cell range
        values = new ArrayList<>();                                         // Create a List of values
        values.add("Cell A2");                                              // set a value
        values.add("Cell B2");                                              // set a value
        values.add("Cell C2");                                              // set a value
        values.add("Cell D2");                                              // set a value
        ws.addCellRange(values, "A2:D2");                                   // Insert cell range
        values = new ArrayList<>();                                         // Create a List of values
        values.add("Cell A3");                                              // set a value
        values.add("Cell B3");                                              // set a value
        values.add("Cell C3");                                              // set a value
        values.add("Cell D3");                                              // set a value
        ws.addCellRange(values, "A3:D3");                                   // Insert cell range
        ws.addHiddenColumn("C");                                            // Hide column C
        ws.addHiddenRow(1);                                                 // Hider row 2 (zero-based: 1)
        ws.setAutoFilter(1, 3);                                             // Set auto-filter for column B to D
        try {
            workbook.saveAs(outputFolder + "test7.xlsx");                       // Save the workbook            
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * This demo shows the usage of cell and worksheet selection, auto-sanitizing of worksheet names
     */
    private static void demo8() {
        Workbook workbook = new Workbook(outputFolder + "test8.xlsx", "Sheet*1", true);  // Create new workbook with invalid sheet name (*); Auto-Sanitizing will replace * with _
        workbook.getCurrentWorksheet().addNextCell("Test");                // Add cell A1
        workbook.getCurrentWorksheet().setSelectedCells("A5:B10");        // Set the selection to the range A5:B10
        workbook.addWorksheet("Sheet2");                    // Create new worksheet
        workbook.getCurrentWorksheet().addNextCell("Test2");                // Add cell A1
        Range range = new Range(new Address(1, 1), new Address(3, 3));        // Create a cell range for the selection B2:D4
        workbook.getCurrentWorksheet().setSelectedCells(range);        // Set the selection to the range
        workbook.addWorksheet("Sheet2", true);                // Create new worksheet with already existing name; The name will be changed to Sheet21 due to auto-sanitizing (appending of 1)
        workbook.getCurrentWorksheet().addNextCell("Test3");                // Add cell A1
        workbook.getCurrentWorksheet().setSelectedCells(new Address(2, 2), new Address(4, 4)); // Set the selection to the range C3:E5
        workbook.setSelectedWorksheet(1);                    // Set the second Tab as selected (zero-based: 1)
        try {
            workbook.save();                                                    // Save the workbook            
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * This demo shows the usage of basic Excel formulas
     */
    private static void demo9() {
        Workbook workbook = new Workbook(outputFolder + "test9.xlsx", "sheet1");  // Create a new workbook
        List<Object> numbers = Arrays.asList(1.15d, 2.225d, 13.8d, 15d, 15.1d, 17.22d, 22d, 107.5d, 128d); // Create a list of numbers
        List<Object> texts = Arrays.asList("value 1", "value 2", "value 3", "value 4", "value 5", "value 6", "value 7", "value 8", "value 9"); // Create a list of strings (for vlookup)
        workbook.WS.value("Numbers", BasicStyles.Bold());                   // Add a header with a basic style
        workbook.WS.value("Values", BasicStyles.Bold());                    // Add a header with a basic style
        workbook.WS.value("Formula type", BasicStyles.Bold());              // Add a header with a basic style
        workbook.WS.value("Formula value", BasicStyles.Bold());             // Add a header with a basic style
        workbook.WS.value("(See also worksheet2)");                         // Add a note
        workbook.getCurrentWorksheet().addCellRange(numbers, "A2:A10");     // Add the numbers as range
        workbook.getCurrentWorksheet().addCellRange(texts, "B2:B10");       // Add the values as range

        workbook.getCurrentWorksheet().setCurrentCellAddress("D2");         // Set the "cursor" to D2
        Cell c;                                                             // Create an empty cell object (reusable)
        c = BasicFormulas.Average(new Range("A2:A10"));                     // Define an average formula
        workbook.getCurrentWorksheet().addCell("Average", "C2");            // Add the description of the formula to the worksheet
        workbook.getCurrentWorksheet().addCell(c, "D2");                    // Add the formula to the worksheet

        c = BasicFormulas.Ceil(new Address("A2"), 0);                       // Define a ceil formula
        workbook.getCurrentWorksheet().addCell("Ceil", "C3");               // Add the description of the formula to the worksheet
        workbook.getCurrentWorksheet().addCell(c, "D3");                    // Add the formula to the worksheet

        c = BasicFormulas.Floor(new Address("A2"), 0);                      // Define a floor formula
        workbook.getCurrentWorksheet().addCell("Floor", "C4");              // Add the description of the formula to the worksheet
        workbook.getCurrentWorksheet().addCell(c, "D4");                    // Add the formula to the worksheet

        c = BasicFormulas.Round(new Address("A3"), 1);                      // Define a round formula with one digit after the comma
        workbook.getCurrentWorksheet().addCell("Round", "C5");              // Add the description of the formula to the worksheet
        workbook.getCurrentWorksheet().addCell(c, "D5");                    // Add the formula to the worksheet

        c = BasicFormulas.Max(new Range("A2:A10"));                         // Define a max formula
        workbook.getCurrentWorksheet().addCell("Max", "C6");                // Add the description of the formula to the worksheet
        workbook.getCurrentWorksheet().addCell(c, "D6");                    // Add the formula to the worksheet

        c = BasicFormulas.Min(new Range("A2:A10"));                         // Define a min formula
        workbook.getCurrentWorksheet().addCell("Min", "C7");                // Add the description of the formula to the worksheet
        workbook.getCurrentWorksheet().addCell(c, "D7");                    // Add the formula to the worksheet

        c = BasicFormulas.Median(new Range("A2:A10"));                      // Define a median formula
        workbook.getCurrentWorksheet().addCell("Median", "C8");             // Add the description of the formula to the worksheet
        workbook.getCurrentWorksheet().addCell(c, "D8");                    // Add the formula to the worksheet

        c = BasicFormulas.Sum(new Range("A2:A10"));                         // Define a sum formula
        workbook.getCurrentWorksheet().addCell("Sum", "C9");                // Add the description of the formula to the worksheet
        workbook.getCurrentWorksheet().addCell(c, "D9");                    // Add the formula to the worksheet

        c = BasicFormulas.VLookup(13.8d, new Range("A2:B10"), 2, true);     // Define a vlookup formula (look for the value of the number 13.8)
        workbook.getCurrentWorksheet().addCell("Vlookup", "C10");           // Add the description of the formula to the worksheet
        workbook.getCurrentWorksheet().addCell(c, "D10");                   // Add the formula to the worksheet

        workbook.addWorksheet("sheet2");                                    // Create a new worksheet
        c = BasicFormulas.VLookup(workbook.getWorksheets().get(0), new Address("B4"), workbook.getWorksheets().get(0), new Range("B2:C10"), 2, true); // Define a vlookup formula in worksheet1 (look for the text right of the (value of) cell B4)
        workbook.WS.value(c);                                               // Add the formula to the worksheet

        c = BasicFormulas.Median(workbook.getWorksheets().get(0), new Range("A2:A10"));             // Define a median formula in worksheet1
        workbook.WS.value(c);                                               // Add the formula to the worksheet

        try {
            workbook.save();                                                  // Save the workbook
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * This demo shows the usage of style appending
     */
    private static void demo10() {
        Workbook wb = new Workbook(outputFolder + "demo10.xlsx", "styleAppending");  // Create a new workbook

        Style style = new Style();                                                              // Create a new style
        style.append(BasicStyles.Bold());                                                       // Append a basic style (bold)
        style.append(BasicStyles.Underline());                                                  // Append a basic style (underline)
        style.append(BasicStyles.font("Arial Black", 20));                    // Append a basic style (custom Font)

        wb.WS.value("THIS IS A TEST", style);                                             // Add text and the appended style
        wb.WS.down();                                                                           // Go to a new row

        Style chainedStyle = new Style()                                                        // Create a new style...
                .append(BasicStyles.Underline())                                                // ... and append another part (chaining underline)
                .append(BasicStyles.colorizedText("FF00FF"))                               // ... and append another part (chaining colorized text)
                .append(BasicStyles.colorizedBackground("AAFFAA"));                        // ... and append another part (chaining colorized background)

        wb.WS.value("Another test", chainedStyle);                                        // Add text and the appended style
        try {
            wb.save();                                                                          // Save the workbook
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

}
