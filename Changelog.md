# Change Log

## v1.4.0

---
Release Date: **xx.06.2021**

- Added style reader to resolve dates and times properly
- Added new data type TIME, represented by LocalTime objects in reader and writer
- Added time (LocalTime) examples to the demos
- Added a check to ensure dates are not set beyond 9999-12-31 (limitation of OAdate)
- Updated documentation
- Fixed some code formatting issues

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
