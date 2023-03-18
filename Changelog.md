# Change Log

## v2.1.1

---
Release Date: **18.03.2023**

- Fixed a bug when a workbook contains charts instead of worksheets. Bug fix provided by Iivari Mokelainen for NanoXLSX (.NET) and ported to Java
- Minor code maintenance

## v2.1.0

---
Release Date: **12.11.2022**

- Added a several methods in the Worksheet class to add multiple ranges of selected cells
- Fixed a bug in the reader function to read worksheets with multiple ranges of selected cells
- Fixed a bug in several readers to cope (internally) with bools, represented by numbers and textual expressions
- Renamed APPLICATIONNAME to APPLICATION_NAME in Version class
- Removed internal escaping of custom number format codes for now
- Updated example in demo
- Code maintenance

Note: It seems that newer versions of Excel may store boolean attributes internally now as texts (true/false) and not anymore as numbers (1/0).
      This release adds compatibility to read this newer format but will currently store files still in the old format

Note 2: The incomplete internal escaping of custom number format codes was removed due to the potential high complexity.
        Escaping must be performed currently by hand, according to OOXML specs: Part 1 - Fundamentals And Markup Language Reference, Chapter 18.8.31

## v2.0.1

---
Release Date: **04.10.2022**

- Code maintenance
- Documentation update
- Updated code formatter (reverted unreadable parametrized  unit tests)

## v2.0.1

---
Release Date: **03.10.2022**

- Fixed a bug in the functions to write and read custom number formats
- Fixed behavior of empty cells and added re-evaluation if values are set by the setValue setter
- Fixed a bug in the functions to write and read font values (styles)
- Adapted and added further tests
- Updated documentation
- Added some internal notes to prepare the development of the next mayor version

Note:
- When defining a custom number format, now the CustomFormatCode property must always be defined as well, since an empty value leads to an invalid Workbook
- When a cell is now created (by constructor) with the type EMPTY, any passed value will be discarded in this cell

## v2.0.0

---
Release Date: **03.09.2022 - Major Release**

### Workbook and Shortener

- Added a list of MRU colors that can be defined in the Workbook class (methods addMruColor, clearMruColors)
- Added a getter for the workbook protection password hash (will be filled when loading a workbook)
- Added the method setSelectedWorksheet by name in the Workbook class
- Added two methods getWorksheet by name or index in the Workbook class
- Added the methods copyWorksheetIntoThis and copyWorksheetTo with several overloads in the Workbook class
- Added the function removeWorksheet by index with the option of resetting the current worksheet, in the Workbook class
- Added the function setCurrentWorksheet by index in the Workbook class
- Added the function setSelectedWorksheet by name in the Workbook class
- Added a Shortener-Class constructor with a workbook reference
- The shortener functions down and right have now an option to keep row and column positions
- Added two shortener functions up and left
- Made several style assigning methods deprecated in the Workbook class (will be removed in future versions)

### Worksheet

- Added a getter for the worksheet protection password hash (will be filled when loading a workbook)
- Added the methods getFirstDataColumnNumber, getFirstDataColumnNumber, getFirstDataRowNumber, getFirstRowNumber, getLastDataColumnNumber, getFirstCellAddress, getFirstDataCellAddress, getLastDataColumnNumber, getLastDataRowNumber, getLastRowNumber, getLastCellAddress,  getLastCellAddress and getLastDataCellAddress
- Added the methods getRow and getColumns by address string or index
- Added the method copy to copy a worksheet (deep copy)
- Added a constructor with only the worksheet name as parameter
- Added and option in goToNextColumn and goToNextRow to either keep the current row or column
- Added the methods removeRowHeight and removeAllowedActionOnSheetProtection
- Renamed columnAddress and rowAddress to columnNumber and rowNumber in the addNextCell, addCellFormula and removeCell methods
- Added several validations for worksheet data

### Cells, Rows and Columns

- In Cell, the address can now have reference modifiers ($)
- The worksheet reference in the Cell constructor was removed. Assigning to a worksheet is now managed automatically by the worksheet when adding a cell
- Added a field cellAddressType in Cell
- Cells can now have null as value, interpreted as empty
- Added a new overloaded function resolveCellCoordinate to resolve the address type as well
- Added validateColumnNumber and validateRowNumber in Cell
- In Address, the constructor with string and address type now only needs a string, since reference modifiers ($) are resolved automatically
- Address objects are now comparable
- Implemented better address validation
- Range start and end addresses are swapped automatically, if reversed

### Styles

- Font has now an enum of possible underline values (e.g. instead of a bool)
- CellXf supports now indentation
- A new, internal style repository was introduced to streamline the style management
- Color (RGB) values are now validated (Fill class has a function validateColor)
- Style components have now more appropriate default values
- MRU colors are now not collected from defined style colors but from the MRU list in the workbook object
- The toString function of Styles and all sub parts will now give a complete outline of all elements
- Fixed several issues with style comparison
- Several style default values were introduced as constants

### Formulas

- Added types as formula values: byte, Double, Short, Integer, Long, BigDecimal
- Added several validity checks

### Reader

- Added default values for dates and times in the import options
- Added global casting import options: AllNumbersToDouble, AllNumbersToBigDecimal, AllNumbersToInt, EverythingToString
- Added column casting import options: Numeric, Double, BigDecimal
- Added global import options: EnforcePhoneticCharacterImport, EnforceEmptyValuesAsString, DateTimeFormat, TemporalCultureInfo
- Added a meta data reader
- All style elements that can be written can also be read
- All workbook elements that can be written can also be read (exception: passwords cannot be recovered)
- All worksheet elements that can be written can also be read (exception: passwords cannot be recovered)
- Better handling of dates and times, especially with invalid (too low and too high numbers) values

### Misc
- Added a unit test project with several thousand, partially parametrized test cases
- Added several constants for boundary dates in the Helper class
- Added several functions for pane splitting in the Helper class
- Exposed the (legacy) password generation method in the Helper class
- Updated documentation among the whole project
- Overhauled the whole writer
- Removed lot of dead code for better maintenance

## v1.3.0

---
Release Date: **02.09.2020**

- Added style reader to resolve dates and times properly
- Added new data type TIME, represented by LocalTime objects in reader and writer
- Added time (LocalTime) examples to the demos
- Added a check to ensure dates are not set beyond 9999-12-31 (limitation of OAdate)
- Updated documentation
- Fixed some code formatting issues

## v1.2.8

---
Release Date: **01.12.2019**

- Fixed a bug of reorganized worksheets (when deleted in Excel)

## v1.2.7

---
Release Date: **21.05.2019**

- Maintenance update
- Code cleanup
- Removed dist folder, since binaries are available through releases, compilation or maven central

## v1.2.6

---
Release Date: **07.12.2018**

- Improved the performance of adding stylized cells by factor 10 to 100
- Code reformatting

## v1.2.5

---
Release Date: **04.11.2018**

- Renamed packages LowLevel to XlsxWriter, style to styles, exception to exceptions
- Fixed a bug in the style handling of merged cells. Bug fix provided by David Courtel for PicoXLSX (C#)
- Fixed typos
- Documentation update

## v1.2.4

---
Release Date: **24.08.2018**

**Note**: Due to some refactoring (see below) in this version, changes of existing code may be necessary. However, most introduced changes are on a rather low level or can be fixed by search&replace

- Fixed a bug in the calculation of OA Dates (internal format)
- Fixed a bug regarding formulas in the reader
- Added support for dates in the reader
- Documentation Update

## v1.2.3

 ---
Release Date: **19.08.2018**

- Fixed a bug in the Font style class
- Fixed typos

## v1.2.2

---
Release Date: **06.07.2018**

- Fixed a bug in the reader of the sharedStrings table (follow-up bug of v.1.2.1 - Please update)

## v1.2.1

---
Release Date: **03.07.2018**

- Fixed a bug in the reader of the sharedStrings table (html entities were truncated)

## v1.2.0

---
Release Date: **03.07.2018**

- Added address types (no fixed rows and columns, fixed rows, fixed columns, fixed rows and columns)
- Added new CellDirection Disabled, if the addresses of the cells are defined manually (addNextCell will override the current cell in this case)
- Altered Demo 3 to to demonstrate disabling of automatic cell addressing
- Extended Demo 1 to demonstrate the new address types
- Minor, internal changes

## v1.1.0

---
Release Date: **08.06.2018**

- Added style appending (builder / method chaining)
- Added new basic styles colorizedText, colorizedBackground and font as functions
- Added a new constructor for Workbooks without file name to handle stream-only workbooks more logical
- Added the functions hasCell, getLastColumnNumber and getLastRowNumber in the Worksheet class
- Renamed the function SetColor in the class Fill (Style) to setColor, to follow conventions. Minor refactoring in existing projects may be possible
- Fixed a bug when overriding a worksheet name with sanitizing
- Added new demo for the introduced style features
- Internal optimizations and fixes

## v1.0.1

---

Release Date: **31.05.2018**

- Fixed versioning issue
- Fixed a bug in the processing of column widths. Bug fix provided by Johan Lindvall for PicoXLSX, adapted for PicoXLSX4j and NanoXLSX4j
- Added numeric data types Byte, BigDecimal, and Short (proposal by Johan Lindvall for PicoXLSX)
- Changed the behavior of cell type casting. User defined cell types will now only be overwritten if the type is DEFAULT (proposal by Johan Lindvall for PicoXLSX)

## v1.0.0

---

Release Date: **27.05.2018**

- Initial release
