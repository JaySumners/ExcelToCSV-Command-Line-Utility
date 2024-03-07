using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelToCSV.Models;
using ExcelToCSV.Utilities;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace ExcelToCSV;

internal class App
{
    #region Properties

    #region Init Properties
    public string FilePath { get; init; }
    public string OutputDirectory { get; init; }
    public IEnumerable<string> SheetNames { get; init; }
    public IEnumerable<string> SheetRenames { get; init; }
    public bool Indexed { get; init; }
    public bool Hidden { get; init; }
    public bool NullErrors { get; init; }
    public bool RemoveEmptyRows { get; init; }
    #endregion

    #region Private Properties
    private readonly ILogger<App> _log;
    private string _validatedOutputDirectory = string.Empty;
    private WorkbookPart? _workbookPart;
    private Dictionary<string, (string name, string rename)> _worksheetsDict = [];
    private IEnumerable<SharedStringItem> _sharedStrings = [];
    private Dictionary<int, string?> _sharedStringsDict = [];
    private Dictionary<uint, string> _stylesDict = [];
    private List<WorksheetPartModel> _worksheetPartModels = [];

    #region V1 Properties
    private Dictionary<string, Worksheet> _worksheetsDictV1 = [];
    private IEnumerable<(string name, string rename, Worksheet worksheet)> _worksheetsToParse = [];
    #endregion

    #endregion

    #endregion

    #region Constructors
    public App(
        string filePath,
        string outputDirectory,
        IEnumerable<string> sheetNames,
        IEnumerable<string> sheetRenames,
        bool indexed,
        bool hidden,
        bool nullErrors,
        bool removeEmptyRows
    )
    {
        FilePath = filePath;
        OutputDirectory = outputDirectory;
        SheetNames = sheetNames;
        SheetRenames = sheetRenames;
        Indexed = indexed;
        Hidden = hidden;
        NullErrors = nullErrors;
        RemoveEmptyRows = removeEmptyRows;

        _log = LoggingUtility.GetLogger<App>();
    }
    #endregion

    #region Methods

    public void Run()
    {
        try
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            Console.WriteLine($"Parsing file: {this.FilePath}");
            Console.WriteLine($"  Setting output directory...");
            CreateAndValidateOutputDirectory();

            Console.WriteLine($"  Setting workbook level properties...");
            using SpreadsheetDocument document = OpenSpreadsheet();
            SetOperations(document);

            Console.WriteLine($"  Saving worksheets...");
            var padding = WriteTableHeader();

            foreach (WorksheetPartModel _worksheetPartModel in _worksheetPartModels)
            {
                _worksheetPartModel.ParseWorksheet(padding);
            }

            stopwatch.Stop();
            Console.WriteLine($"  Total Elapsed (s): {stopwatch.Elapsed.TotalSeconds}");
        }
        catch (Exception e)
        {
            _log.LogError(0, e, "An unhandled exception occurred. Message: {Message}", e.Message);
            throw;
        }
    }

    #region Reporters
    private (int left, int right) WriteTableHeader()
    {
        string header1 = "Worksheet";
        string header2 = "Information";

        int leftPadding = 4;
        int rightPadding = Math.Max(_worksheetsDict.Values.Select(v => v.name.Length).Max(), header1.Length) + 2;

        Console.WriteLine();
        Console.Write("".PadLeft(leftPadding));
        Console.Write(header1.PadRight(rightPadding));
        Console.WriteLine(header2);

        Console.Write("".PadLeft(leftPadding));
        Console.Write(new string('-', header1.Length));
        Console.Write("".PadLeft(rightPadding - header1.Length));
        Console.WriteLine(new string('-', header2.Length));

        return (left: leftPadding, right: rightPadding);
    }
    #endregion

    #region Static Helpers
    private static string GetAttributeValue(OpenXmlReader reader, string localName)
    {
        return reader.Attributes?.Where(a => a.LocalName == localName).FirstOrDefault().Value ?? string.Empty;
    }
    #endregion

    #region File Operations
    private SpreadsheetDocument OpenSpreadsheet()
    {
        return SpreadsheetDocument.Open(this.FilePath, isEditable: false);
    }
    private void CreateAndValidateOutputDirectory()
    {
        string? outputDirectory = this.OutputDirectory;

        if (string.IsNullOrEmpty(outputDirectory))
        {
            string? newDirectoryName = PathUtility.ReplaceInvalidFileNameChars(Path.GetFileNameWithoutExtension(this.FilePath));
            string? sourceDirectory = PathUtility.ReplaceInvalidPathCharacters(Path.GetDirectoryName(this.FilePath) ?? "");

            if (newDirectoryName is null || sourceDirectory is null)
            {
                throw new Exception($"Valid directory name could not be generated from '{this.FilePath}'. Please specify an output directory.");
            }

            outputDirectory = Path.Combine(sourceDirectory, newDirectoryName);
        }

        try
        {
            _ = Directory.CreateDirectory(outputDirectory);
        }
        catch (Exception ex)
        {
            throw new Exception($"Directory '{outputDirectory}' cannot be created. Message: '{ex.Message}'");
        }

        this._validatedOutputDirectory = outputDirectory;
    }
    #endregion

    #region Setters
    private void SetOperations(SpreadsheetDocument document)
    {
        SetWorkbookPart(document);
        SetSharedStrings();
        SetWorksheetsDict();
        SetStylesDict();
        SetWorksheetPartModels(); // Maybe we should start parsing here rather than setting up a bunch of things.
    }
    private void SetWorkbookPart(SpreadsheetDocument document)
    {
        _workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();
    }
    private void SetWorksheetsDict()
    {
        if (_workbookPart is null)
        {
            throw new Exception("WorkbookPart cannot be null!");
        }

        //using SpreadsheetDocument document = OpenSpreadsheet();

        #region All Worksheets Dictionary (Hidden Flag Implemented)
        Dictionary<string, string> worksheetPartIds = [];

        using OpenXmlReader wbPartReader = OpenXmlReader.Create(_workbookPart);

        while (wbPartReader.Read())
        {
            string workSheetName = string.Empty;

            if (wbPartReader.ElementType == typeof(Sheet))
            {
                string partIdString = wbPartReader.Attributes.Where(a => a.LocalName == "id").FirstOrDefault().Value ?? string.Empty;
                string sheetName = wbPartReader.Attributes.Where(a => a.LocalName == "name").FirstOrDefault().Value ?? string.Empty;

                if (!string.IsNullOrEmpty(partIdString))
                {
                    string? stateString = wbPartReader.Attributes.Where(a => a.LocalName == "state").FirstOrDefault().Value ?? string.Empty;
                    bool hidden = stateString == "hidden" || stateString == "veryHidden";

                    // Hidden Flag Implementation
                    if (!hidden || this.Hidden)
                    {
                        worksheetPartIds.Add(sheetName, partIdString);
                    }
                }
            }
        }
        #endregion

        #region Worksheet Names Dictionary (SheetNames and SheetRenames Implemented)
        IEnumerable<string> worksheetsAvailable = worksheetPartIds.Keys;
        IEnumerable<string> worksheetNames = this.SheetNames.Any() ? this.SheetNames : worksheetsAvailable;
        IEnumerable<string> worksheetRenames = this.SheetRenames.Any() ? this.SheetRenames : worksheetNames;

        // Error: There are specified sheet names that are not available.
        if (this.SheetNames.Any())
        {
            IEnumerable<string> missingNames = this.SheetNames.Except(worksheetsAvailable);

            if (missingNames.Any())
            {
                throw new Exception($"The following specified sheets were not found: {string.Join(", ", missingNames)}");
            }
        }

        // Error: SheetNames and SheetRenames must be the same length
        if (worksheetNames.Count() != worksheetRenames.Count())
        {
            throw new ArgumentException($"If '--rename' is supplied it must equal the number of sheets in the workbook OR the number of sheets spcified with '--sheets'.");
        }

        IEnumerable<(string name, string rename)> worksheetNameRenames = worksheetNames
            .Zip(
                worksheetRenames,
                (name, rename) => (name: name, rename: PathUtility.ReplaceInvalidFileNameChars(rename))
            );
        #endregion

        //_worksheetsDictV2
        foreach (var worksheetNameRename in worksheetNameRenames)
        {
            _worksheetsDict.Add(worksheetPartIds[worksheetNameRename.name], worksheetNameRename);
        }
    }
    private void SetSharedStrings()
    {
        if (_workbookPart is null)
        {
            throw new Exception("WorkbookPart cannot be null!");
        }

        using OpenXmlReader reader = OpenXmlReader.Create(_workbookPart.SharedStringTablePart!);

        int idx = 0;
        string? itemText;

        while (reader.Read())
        {
            if (reader.ElementType == typeof(SharedStringItem))
            {
                itemText = string.Empty;

                if (reader.ReadFirstChild())
                {
                    do
                    {
                        if (reader.ElementType == typeof(Text))
                        {
                            itemText = reader.GetText();
                        }
                    } while (reader.ReadNextSibling());
                }

                _sharedStringsDict.Add(idx, itemText);
                idx++;
            }
        }
    }
    private void SetStylesDict()
    {
        if (_workbookPart is null)
        {
            throw new Exception("WorkbookPart cannot be null!");
        }

        WorkbookStylesPart? wbStylesPart = _workbookPart.WorkbookStylesPart;

        if (wbStylesPart is not null)
        {
            using OpenXmlReader wbStylesPartReader = OpenXmlReader.Create(wbStylesPart);

            while (wbStylesPartReader.Read())
            {
                // We could just look for the element we want rather than these do statments
                // but it seems better to work down the tree than jump the fence (just in case).
                if (wbStylesPartReader.ElementType == typeof(Stylesheet))
                {
                    wbStylesPartReader.ReadFirstChild();
                    do
                    {
                        if (wbStylesPartReader.ElementType == typeof(NumberingFormats))
                        {
                            wbStylesPartReader.ReadFirstChild();
                            do
                            {
                                if (wbStylesPartReader.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.NumberingFormat))
                                {
                                    string numFormatIdString = GetAttributeValue(wbStylesPartReader, "numFmtId");
                                    string formatCodeString = GetAttributeValue(wbStylesPartReader, "formatCode");

                                    if (uint.TryParse(numFormatIdString, out uint numFormatId))
                                    {
                                        _stylesDict.Add(numFormatId, formatCodeString);
                                    }
                                }
                            } while (wbStylesPartReader.ReadNextSibling());
                        }
                    } while (wbStylesPartReader.ReadNextSibling());
                }
            }
        }
    }
    private void SetWorksheetPartModels()
    {
        if (_workbookPart is null)
        {
            throw new Exception("WorkbookPart cannot be null!");
        }

        IEnumerable<WorksheetPart> wsParts = this._workbookPart.WorksheetParts;

        List<WorksheetPartModel> worksheetPartModels = [];

        foreach (var wsPart in wsParts)
        {
            string wsRelationshipId = _workbookPart.GetIdOfPart(wsPart) ?? string.Empty;
            if (_worksheetsDict.TryGetValue(wsRelationshipId, out (string name, string rename) sheetNames))
            {
                WorksheetPartModel worksheetPartModel = new(
                    worksheetPart: wsPart,
                    sharedStringsDict: _sharedStringsDict,
                    stylesDict: _stylesDict,
                    outputDirectory: _validatedOutputDirectory,
                    sheetName: sheetNames.name,
                    sheetRename: sheetNames.rename,
                    indexed: this.Indexed,
                    nullErrors: this.NullErrors,
                    removeEmptyRows: this.RemoveEmptyRows
                );

                this._worksheetPartModels.Add(worksheetPartModel);
            }
        }
    }
    #endregion

    #endregion
}
