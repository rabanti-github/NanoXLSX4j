# NanoXLSX4j

[![Maven Central](https://maven-badges.herokuapp.com/maven-central/ch.rabanti/nanoxlsx4j/badge.svg)](https://maven-badges.herokuapp.com/maven-central/ch.rabanti/nanoxlsx4j) ![GitHub license](https://img.shields.io/github/license/rabanti-github/picoXlsx4j.svg)[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX4j.svg?type=shield)](https://app.fossa.com/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX4j?ref=badge_shield)


NanoXLSX4j is a small Java library to create and read XLSX files (Microsoft Excel 2007 or newer) in an easy and native way. The library is originated form [PicoXLSX4j](https://github.com/rabanti-github/PicoXLSX4j) and has basic support of reading spreadsheets. PicoXLSX4j and thus NanoXLSX4j are direct ports of [PicoXLSX for C#](https://github.com/rabanti-github/PicoXLSX)

* No need for an installation of Microsoft Office
* No need for Office interop/DCOM or other bridging libraries
* No need for 3rd party libraries
* Pure usage of standard JRE

Project website: [https://picoxlsx.rabanti.ch](https://picoxlsx.rabanti.ch)

## Modules

* NanoXLSX4l.Lib : Contains the actual library as basis for the Maven artefact
* NanoXLSX4j.Demo : Contains usage examples of the library

See the **[Change Log](https://github.com/rabanti-github/NanoXLSX4j/blob/master/Changelog.md)** for recent updates.

## What's new in version 2.x

There are some additional functions for workbooks and worksheets, as well as support of further data types.
The biggest change is the full capable reader support for workbook, worksheet and style information. Also, all features are now fully unit tested. This means, that nanoXLSX is no longer in Beta state. Some key features are:

* Full reader support for styles
* Copy functions for worksheets
* Advance import options for the reader
* Several additional checks, exception handling and updated documentation

## Road map
Version 2.x of NanoXLSX4j was completely overhauled and a high number of (partially parametrized) unit tests with a code coverage of >97% were written to improve the quality of the library.
However, it is not planned as a LTS version. The upcoming v3.x is supposed to introduce some important functions, like in-line cell formatting, better formula handling and additional worksheet features.
Furthermore, it is planned to introduce more modern OOXML features like the SHA256 implementation of worksheet passwords.

## Reader Support

The reader of NanoXLS4j follows the principle of "What You Can Write Is What You Can Read". Therefore, all information about workbooks, worksheets, cells and styles that can be written into an XLSX file by NanoXLSX, can also be read by it.
There are some limitations:

* A workbook or worksheet password cannot be recovered, only its hash
* Information that is not supported by the library will be discarded
* There are some approximations for floating point numbers. These values (e.g. pane split widths) may deviate from the originally written values
* Numeric values are cast to the appropriate Java types with the best effort. There are import options available to enforce specific types
* No support of other objects than spreadsheet data at the moment
* Due to the potential high complexity, custom number format codes are currently not automatically escaped on writing or un-escaped on reading

## Requirements

NanoXLSX4j was initially created with Java 8 and currently build with OpenJDK 11
The only requirement for development is an up-to-date Java environment (OpenJDK 11 or higher recommended)

## Installation

### As JAR

Simply place the NanoXLSX4j jar file (e.g. **nanoxlsx4j-2.5.2.jar**) into the lib folder of your project and create a library reference to it in your IDE.

### As source files

Place all .java files from the NanoXLSX4j [source folder](https://github.com/rabanti-github/NanoXLSX4j/tree/master/NanoXLSX4j.Lib/src/main/java/ch/rabanti/nanoxlsx4j) into your project. The folder structure defines the packages. Please use refactoring if you want to relocate the files.

### Maven

Add the following information to your POM file within the ```<dependencies>``` tag:

```xml
<dependency>
    <groupId>ch.rabanti</groupId>
    <artifactId>nanoxlsx4j</artifactId>
    <version>2.5.5</version>
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
The [Demo module](https://github.com/rabanti-github/NanoXLSX4j/tree/master/NanoXLSX4j.Demo) contains 14 simple use cases. You can also look at the full API documentation or the Javadoc annotations in the particular .java files.

Hint: You will find most certainly any function, and the way how to use it, in the [Unit Test Section](https://github.com/rabanti-github/NanoXLSX4j/tree/master/NanoXLSX4j.Lib/src/test/java/ch/rabanti/nanoxlsx4j)



## License
[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX4j.svg?type=large)](https://app.fossa.com/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX4j?ref=badge_large)