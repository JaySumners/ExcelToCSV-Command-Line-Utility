using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelToCSV.Utilities;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelToCSV.Models;

internal class WorksheetPartModel
{
    #region Properties

    #region Initialized (Internal) Properties
    internal string FilePath { get; init; }
    internal string SheetName { get; init; }
    internal string SheetRename { get; init; }
    internal bool Indexed { get; init; }
    internal bool NullErrors { get; init; }
    internal bool RemoveEmptyRows { get; init; }
    #endregion

    #region Constant Properties
    private const string NULL_CELL = "";
    #endregion

    #region Static Properties
    private static List<char> _escapableStrings = [',', '"', (char)13, (char)10];
    #endregion

    #region Private Properties
    private readonly ILogger<WorksheetPartModel> _log;

    private readonly WorksheetPart _worksheetPart;
    //private readonly IEnumerable<SharedStringItem> _sharedStrings;
    private readonly Dictionary<int, string?> _sharedStringsDict = [];
    private readonly Dictionary<uint, string> _stylesDict = [];

    private RangeReferenceModel _rangeReference;

    private int _rowsCovered = 0;
    private int _rowsWritten = 0;
    private (int left, int top) _cursorPosition = (0, 0);
    #endregion

    #endregion

    #region Constructors
    public WorksheetPartModel(
        WorksheetPart worksheetPart,
        //IEnumerable<SharedStringItem> sharedStrings,
        Dictionary<int, string?> sharedStringsDict,
        Dictionary<uint, string> stylesDict,
        string outputDirectory,
        string sheetName,
        string sheetRename,
        bool indexed,
        bool nullErrors,
        bool removeEmptyRows
    )
    {
        _worksheetPart = worksheetPart;
        _sharedStringsDict = sharedStringsDict;
        _stylesDict = stylesDict;

        SheetName = sheetName;
        SheetRename = sheetRename;
        Indexed = indexed;
        NullErrors = nullErrors;
        RemoveEmptyRows = removeEmptyRows;

        _log = LoggingUtility.GetLogger<WorksheetPartModel>();

        string newFileName = PathUtility.ReplaceInvalidFileNameChars($"{sheetRename}.csv");
        FilePath = PathUtility.ReplaceInvalidPathCharacters(Path.Combine(outputDirectory, newFileName)) ?? string.Empty;
    }
    #endregion

    #region Methods
    private static string GetAttributeValue(OpenXmlReader reader, string localName)
    {
        return reader.Attributes?.Where(a => a.LocalName == localName).FirstOrDefault().Value ?? string.Empty;
    }
    private void DeleteFileIfExists()
    {
        if (File.Exists(this.FilePath))
        {
            File.Delete(this.FilePath);
        }
    }
    private StreamWriter GetOutputWriter()
    {
        int bufferSize = 128 * 1024; //KB

        return new StreamWriter(
            this.FilePath,
            append: true,
            encoding: Encoding.UTF8,
            bufferSize: bufferSize)
        {
            AutoFlush = false
        };
    }
    private void SetSheetDimensions(OpenXmlReader reader)
    {
        #region Sheet Dimensions
        while (reader.Read())
        {
            if (reader.ElementType == typeof(SheetDimension))
            {
                string? reference = GetAttributeValue(reader, "ref") ?? "A1:A1";
                _rangeReference = new(reference);
                break;
            }
        }
        #endregion
    }

    #region Parsing
    internal void ParseWorksheet((int left, int right) padding)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();

        #region Parsing Rows / Sheet
        Console.Write("".PadLeft(padding.left));
        Console.Write(this.SheetName.PadRight(padding.right));

        Console.Write("Setting up output file...".PadRight(padding.right));
        DeleteFileIfExists();
        using StreamWriter writer = GetOutputWriter();
        Console.WriteLine("done!");

        Console.Write("".PadLeft(padding.left));
        Console.Write("".PadRight(padding.right));
        Console.Write("Parsing Rows ");

        using OpenXmlReader wsPartReader = OpenXmlReader.Create(_worksheetPart);

        SetSheetDimensions(wsPartReader);

        _cursorPosition = Console.GetCursorPosition();
        Console.Write(_rowsCovered);

        #region Sheet Data
        while (wsPartReader.Read())
        {
            if (wsPartReader.ElementType == typeof(SheetData))
            {
                // Look for Rows
                if (wsPartReader.ReadFirstChild())
                {
                    //ParseRows, but left here so we can report easier.
                    do
                    {
                        if (wsPartReader.ElementType == typeof(Row))
                        {
                            string parsedRow = ParseRow(wsPartReader, writer);

                            WriteRow(parsedRow, writer);
                        }
                    } while (wsPartReader.ReadNextSibling()); //Rows
                }

                break;
            }
        }
        #endregion

        Console.SetCursorPosition(_cursorPosition.left, _cursorPosition.top);
        Console.WriteLine($"({_rowsWritten}/{_rowsCovered}) rows (Written/Read).");

        // Maybe an option here
        //DeleteFileIfExists();

        writer.Flush();

        stopwatch.Stop();
        Console.Write("".PadLeft(padding.left + padding.right));
        Console.WriteLine($"Elapsed Time (s): {Math.Round(stopwatch.Elapsed.TotalSeconds, 2)}");

        Console.WriteLine();
        #endregion
    }
    private void WriteRow(string parsedRow, StreamWriter writer)
    {
        if (!(string.IsNullOrEmpty(parsedRow) && this.RemoveEmptyRows))
        {
            writer.WriteLine(parsedRow);
            _rowsWritten++;
        }

        Console.SetCursorPosition(_cursorPosition.left, _cursorPosition.top);

        _rowsCovered++;
        double percentComplete = (_rowsCovered / (double)_rangeReference.EndCellReference.RowNumber);
        int completeHashes = (int)Math.Floor(percentComplete * 10);
        Console.Write($"[{new string('#', completeHashes)}{new string('-', 10 - completeHashes)}]");
        Console.Write($" {percentComplete:P0}");
    }
    private static string EscapeString(string str)
    {
        if (string.IsNullOrEmpty(str))
        {
            return str;
        }

        if (_escapableStrings.Any(str.Contains))
        {
            str = $"\"{str}\"";
        }

        return str.Replace("\"", "\"\"");
    }
    private void AddEmptyRows(int rowId, StreamWriter writer)
    {
        List<string> blankRow = Enumerable.Repeat(NULL_CELL, _rangeReference.EndCellReference.ColumnNumber).ToList();
        List<string> rowStrings;

        for (int i = _rowsCovered + 1; i < rowId; i++)
        {
            rowStrings = blankRow;

            if (this.Indexed)
            {
                rowId = Math.Min(i, _rowsWritten + 1);
                rowStrings = rowStrings.Prepend($"\"{rowId}\"").ToList();
            }

            string parsedRow = string.Join(",", rowStrings);

            WriteRow(parsedRow, writer);
        }
    }
    private string ParseRow(OpenXmlReader wsPartReader, StreamWriter writer)
    {
        #region Parse Cells
        List<string> rowStrings = [];

        string rowIdString = GetAttributeValue(wsPartReader, "r");

        if (!int.TryParse(rowIdString, out int rowId))
        {
            throw new Exception("Row does not have an row number.");
        }

        if (!this.RemoveEmptyRows)
        {
            AddEmptyRows(rowId, writer);
        }

        if (wsPartReader.ReadFirstChild())
        {
            rowStrings = ParseCells(wsPartReader);
        }
        #endregion

        #region Empty Row (Early Return)
        bool emptyRow = !rowStrings.Where(s => s != NULL_CELL).Any();

        if (emptyRow && this.RemoveEmptyRows)
        {
            return string.Empty;
        }
        #endregion

        #region Complete Line
        rowStrings.AddRange(Enumerable.Repeat(NULL_CELL, _rangeReference.EndCellReference.ColumnNumber - rowStrings.Count));
        #endregion

        #region Indexed
        if (this.Indexed)
        {
            rowId = Math.Min(rowId, _rowsWritten + 1);
            rowStrings = rowStrings.Prepend($"\"{rowId}\"").ToList();
        }
        #endregion

        return string.Join(",", rowStrings);
    }
    private List<string> ParseCells(OpenXmlReader wsPartReader)
    {
        List<string> rowStrings = [];

        #region ParseCells
        // Row has children, likely cells
        do
        {
            if (wsPartReader.ElementType == typeof(Cell))
            {
                (CellReferenceModel cellReference, string cellValue) = ParseCell(wsPartReader);

                #region Add Null Cells in Row
                if ((rowStrings.Count + 1) < cellReference.ColumnNumber)
                {
                    rowStrings.AddRange(Enumerable.Repeat(NULL_CELL, cellReference.ColumnNumber - (rowStrings.Count + 1)));
                }
                #endregion

                rowStrings.Add(string.IsNullOrEmpty(cellValue) ? NULL_CELL : cellValue);
            }
        } while (wsPartReader.ReadNextSibling()); //Cells
        #endregion

        return rowStrings;
    }
    private (CellReferenceModel reference, string value) ParseCell(OpenXmlReader wsPartReader)
    {
        CellReferenceModel cellReference = new(GetAttributeValue(wsPartReader, "r"));
        string cellDataTypeString = GetAttributeValue(wsPartReader, "t");
        string cellStyleString = GetAttributeValue(wsPartReader, "s");
        string cellValue = NULL_CELL;

        if (wsPartReader.ReadFirstChild())
        {
            cellValue = ParseCellValues(wsPartReader, cellDataTypeString, cellStyleString);
        }

        return (reference: cellReference, value: cellValue);
    }
    private string ParseCellValues(OpenXmlReader wsPartReader, string cellDataTypeString, string cellStyleString)
    {
        string cellValue = NULL_CELL;

        do
        {
            if (wsPartReader.ElementType == typeof(CellValue))
            {
                cellValue = ParseCellValue(wsPartReader.GetText(), cellDataTypeString, cellStyleString);
                //Do no break. We need to read through until the siblings are done.
            }
        } while (wsPartReader.ReadNextSibling());

        return EscapeString(cellValue);
    }
    private string ParseCellValue(string? cellValue, string cellDataTypeString, string cellStyleString)
    {
        #region Return Empty Checks
        if (string.IsNullOrEmpty(cellValue))
        {
            return NULL_CELL;
        }

        if (this.NullErrors && FormatUtility.IsExcelError(cellValue))
        {
            return NULL_CELL;
        }
        #endregion

        #region Share String
        // "s" is the CellValues.SharedString value
        if (cellDataTypeString == "s")
        {
            if (int.TryParse(cellValue, out int itemId))
            {
                //SharedStringItem item = GetSharedStringItemById(itemId);
                //cellValue = item.Text?.Text ?? item.InnerText ?? item.InnerXml ?? string.Empty;
                _ = _sharedStringsDict.TryGetValue(itemId, out cellValue);
            }

            return cellValue ?? NULL_CELL;
        }
        #endregion

        #region Styles
        if (uint.TryParse(cellStyleString, out uint numberFormatId))
        {
            if (_stylesDict.TryGetValue(numberFormatId, out string? formatCode))
            {
                cellValue = FormatUtility.Format(cellValue, numberFormatId, formatCode ?? string.Empty);
            }
        }
        #endregion

        return cellValue;
    }
    #endregion

    #endregion
}
