# NanoXLSX4j

NanoXLSX4j is a small Java library to create and read XLSX files (Microsoft Excel 2007 or newer) in an easy and native way. The library is originated form [PicoXLSX4j](https://github.com/rabanti-github/PicoXLSX4j) and has basic support of reading spreadsheets. PicoXLSX4j and thus NanoXLSX4j are direct ports of [PicoXLSX for C#](https://github.com/rabanti-github/PicoXLSX)

* No need for an installation of Microsoft Office
* No need for Office interop/DCOM or other bridging libraries
* No need for 3rd party libraries
* Pure usage of standard JRE

Project website: [https://picoxlsx.rabanti.ch](https://picoxlsx.rabanti.ch)

See the **[Change Log](https://github.com/rabanti-github/NanoXLSX4j/blob/master/Changelog.md)** for recent updates.

## Reader Support

Currently, only basic reader functionality is available:

* Reading and casting of cell values into the appropriate data types
* Reading of several worksheets in one workbook with worksheet names
* Limited processing of styles (when reading) at the moment
* No support of other objects than spreadsheet data at the moment

**Note: Styles in loaded files are only considering number formats (to determine date and time values), as well as custom formats. The scope of reader functionality may change with future versions.**

## Requirements

PicoXLSX4j was initially created with Java 8 and currently build with OpenJDK 11
The only requirement for development is an up-to-date Java environment (OpenJDK 11 or higher recommended)

## Installation

### As JAR

Simply place the NanoXLSX4j jar file (e.g. **nanoxlsx4j-1.3.0.jar**) into the lib folder of your project and create a library reference to it in your IDE.

### As source files

Place all .java files from the NanoXLSX4j source folder into your project. The folder structure defines the packages. Please use refactoring if you want to relocate the files.

### Maven

Add the following information to your POM file within the ```<dependencies>``` tag:

```xml
<dependency>
    <groupId>ch.rabanti</groupId>
    <artifactId>nanoxlsx4j</artifactId>
    <version>1.3.0</version>
</dependency>
```

**Important:** The version number may change.
Please see the version number of Maven Central [![Maven Central](https://maven-badges.herokuapp.com/maven-central/ch.rabanti/nanoxlsx4j/badge.svg)](https://maven-badges.herokuapp.com/maven-central/ch.rabanti/nanoxlsx4j)
 or check the [Change Log](https://github.com/rabanti-github/NanoXLSX4j/blob/master/Changelog.md) for the most recent version. The keywords ```LATEST```  and ```RELEASE``` are only valid in Maven 2, not 3 and newer.

## Usage

### Quick Start (shortened syntax)

```java
 Workbook workbook = new Workbook("myWorkbook.xlsx", "Sheet1");         // Create new workbook with a worksheet called Sheet1
 workbook.WS.value("Some Data");                                        // Add cell A1
 workbook.WS.formula("=A1");                                            // Add formula to cell B1
 workbook.WS.down();                                                    // Go to row 2
 workbook.WS.value(new Date(), BasicStyles.Bold());                     // Add formatted value to cell A2
 try{
   workbook.save();                                                     // Save the workbook as myWorkbook.xlsx
 } catch (Exception ex) {}
```

### Quick Start (regular syntax)

```java
 Workbook workbook = new Workbook("myWorkbook.xlsx", "Sheet1");       // Create new workbook with a worksheet called Sheet1
 workbook.getCurrentWorksheet().addNextCell("Some Data");             // Add cell A1
 workbook.getCurrentWorksheet().addNextCell(42);                      // Add cell B1
 workbook.getCurrentWorksheet().goToNextRow();                        // Go to row 2
 workbook.getCurrentWorksheet().addNextCell(new Date());              // Add cell A2
 try {
   workbook.Save();                                                   // Save the workbook as myWorkbook.xlsx
 } catch (Exception ex) {}
```

### Quick Start (read)

```java
try
{
    Workbook wb = Workbook.load("basic.xlsx");                      // Read the workbook
    System.out.println("contains worksheet name: " + wb.getCurrentWorksheet().getSheetName());
    wb.getCurrentWorksheet().getCells().forEach((k, v) -> {         // Iterate through data
        System.out.println("Cell address: " + k + ": content:'" + v.getValue() + "'");
    });
}
catch (Exception ex)
{
 System.out.println("Error: " + ex);
}
```

## Further References

See the full <b>API-Documentation</b> at: [https://rabanti-github.github.io/NanoXLSX4j/](https://rabanti-github.github.io/NanoXLSX4j/).
The [Demo class](https://github.com/rabanti-github/NanoXLSX4j/blob/master/src/main/java/ch/rabanti/nanoxlsx4j/demo/NanoXLSX4j.java) contains 14 simple use cases. You can also look at the full API documentation or the Javadoc annotations in the particular .java files.
