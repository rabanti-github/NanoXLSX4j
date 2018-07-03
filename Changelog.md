# Change Log

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
