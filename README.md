# ExcelToCSV-Command-Line-Utility
##TLDR
A simple Excel-to-CSV command line converter using OpenXML (no Excel COM required).

## Motivation
I often have to work with Excel files that lock up if multiple workflows are trying to access them at the same time. Once solution is to convert the individual sheets to CSVs instead. In another repository, I have PowerShell files that use the Excel COM object to do it, but I can't run it on machines that don't have Excel installed. The solution was to build a simple command line utility using the OpenXML standard (as implemented by Microsoft in C#). 

## Usage / Help
```
ExcelToCSV Commmand Line Utility
Assembly Version: 1.0.66.5
Product Version: 1.0.66.5-alpha.2

Description:
  Converts Excel (.xlsx) files to CSVs (.csv) without having Excel installed.

Usage:
  ExcelToCSV <file> [options]

Arguments:
  <file>  Absolute or relative path to Excel (.xlsx) file to convert. []

Options:
  --output <output>                                  Specifies an absolute or relative directory where CSVs will save.
  --sheets <sheets>                                  Sheet names to include. If not specified, will include all sheets.
  --rename <rename>                                  List of names for output files. Must be equal to the number of
                                                     sheets available OR the number of sheets selected with '--sheets'.
  --indexed, /hidden, /nullErrors, /removeEmptyRows  Will add a row number column in first position without a header to
                                                     each sheet.
  --hidden                                           Will include hidden and veryhidden sheets.
  --nullErrors                                       Will convert any Excel error to an empty string.
  --removeEmptyRows                                  Will remove empty rows from the worksheet
  --version                                          Show version information
  -?, -h, --help                                     Show help and usage information
```

## Licensing and Repurposing
This software is provided AS IS on an MIT license. In short, use it, modify it, repurpose it, whatever you want to do with it, but do it at your own risk. 

