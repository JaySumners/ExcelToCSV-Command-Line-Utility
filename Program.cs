using ExcelToCSV.Utilities;
using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Builder;
using System.CommandLine.Invocation;
using System.CommandLine.Parsing;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelToCSV;

internal class Program
{
    static void Main(string[] args)
    {
        Assembly assembly = Assembly.GetExecutingAssembly();
        string? assemblyVersion = assembly.GetName().Version?.ToString();
        string? productVersion = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;

        Console.WriteLine($"ExcelToCSV Commmand Line Utility");
        Console.WriteLine($"Assembly Version: {assemblyVersion}");
        Console.WriteLine($"Product Version: {productVersion}");
        Console.WriteLine();

        var parser =
            BuildCommandLine()
            .UseDefaults()
            .Build();

        parser.Invoke(args);
    }

    internal static CommandLineBuilder BuildCommandLine()
    {
        #region Argument Parsing

        #region FileArgument
        Argument<string> fileArgument = new(
            name: "file",
            description: "Absolute or relative path to Excel (.xlsx) file to convert.",
            isDefault: true,
            parse: result =>
            {
                if (result.Tokens.Count == 0)
                {
                    result.ErrorMessage = "Argument '<file>' must be supplied.";
                    return string.Empty;
                }

                string? filePath = result.Tokens.Single().Value;

                if (string.IsNullOrEmpty(filePath))
                {
                    result.ErrorMessage = "Argument '<file>' is null or empty";
                    return string.Empty;
                }

                if (PathUtility.FileNameContainsInvalidChars(filePath))
                {
                    result.ErrorMessage = "Argument '<file>' contains invalid path or filename characters.";
                    return string.Empty;
                }

                if ((!File.Exists(filePath)))
                {
                    result.ErrorMessage = $"'{filePath}' does not exist or is not a valid file path!";
                    return string.Empty;
                }

                if (Path.GetExtension(filePath) != ".xlsx")
                {
                    result.ErrorMessage = $"Only '.xlsx' files are supported. Entered: '{Path.GetExtension(filePath)}'.";
                    return string.Empty;
                }

                return filePath;
            }
        )
        {
            Arity = ArgumentArity.ExactlyOne
        };
        fileArgument.LegalFilePathsOnly();
        #endregion

        #region Output Directory Option
        Option<string> outputDirectoryOption = new(
            name: "--output",
            description: "Specifies an absolute or relative directory where CSVs will save.",
            parseArgument: result =>
            {
                if (result.Tokens.Count == 0)
                {
                    return string.Empty;
                }

                string? outputPath = result.Tokens.Single().Value;

                if (string.IsNullOrEmpty(outputPath))
                {
                    result.ErrorMessage = $"Option '--output' has an empty value.";
                    return string.Empty;
                }

                if (PathUtility.PathContainsInvalidChars(outputPath))
                {
                    result.ErrorMessage = $"Option '--output' contains invalid path characters.";
                    return string.Empty;
                }

                return outputPath;
            }
        );
        outputDirectoryOption.LegalFilePathsOnly();
        outputDirectoryOption.AddAlias("/output");
        #endregion

        #region Sheet Names Option
        Option<IEnumerable<string>> sheetSelectionOption = new(
            name: "--sheets",
            description: "Sheet names to include. If not specified, will include all sheets."
        )
        {
            AllowMultipleArgumentsPerToken = true
        };
        sheetSelectionOption.AddAlias("/sheets");
        #endregion

        #region Rename Sheets Option
        Option<IEnumerable<string>> renameOption = new(
            name: "--rename",
            description: "List of names for output files. Must be equal to the number of sheets available OR the number of sheets selected with '--sheets'."
        )
        {
            AllowMultipleArgumentsPerToken = true
        };
        renameOption.AddAlias("/rename");
        #endregion

        #region Index Column Option
        Option<bool> indexColumnOption = new(
            name: "--indexed",
            description: "Will add a row number column in first position without a header to each sheet.",
            parseArgument: _ => true
        )
        {
            Arity = ArgumentArity.Zero
        };
        indexColumnOption.AddAlias("/indexed");
        #endregion

        #region Hidden Sheets Option
        Option<bool> hiddenSheetsOption = new(
            name: "--hidden",
            description: "Will include hidden and veryhidden sheets.",
            parseArgument: _ => true
        )
        {
            Arity = ArgumentArity.Zero
        };
        indexColumnOption.AddAlias("/hidden");
        #endregion

        #region Null Errors Option
        Option<bool> nullErrorsOption = new(
            name: "--nullErrors",
            description: "Will convert any Excel error to an empty string.",
            parseArgument: _ => true
        )
        {
            Arity = ArgumentArity.Zero
        };
        indexColumnOption.AddAlias("/nullErrors");
        #endregion

        #region RemoveEmptyRows Option
        Option<bool> removeEmptyRowsOption = new(
            name: "--removeEmptyRows",
            description: "Will remove empty rows from the worksheet",
            parseArgument: _ => true
        )
        {
            Arity = ArgumentArity.Zero
        };
        indexColumnOption.AddAlias("/removeEmptyRows");
        #endregion

        #endregion

        #region Root Command
        RootCommand rootCommand = new("Converts Excel (.xlsx) files to CSVs (.csv) without having Excel installed.");

        rootCommand.AddArgument(fileArgument);

        rootCommand.AddOption(outputDirectoryOption);
        rootCommand.AddOption(sheetSelectionOption);
        rootCommand.AddOption(renameOption);
        rootCommand.AddOption(indexColumnOption);
        rootCommand.AddOption(hiddenSheetsOption);
        rootCommand.AddOption(nullErrorsOption);
        rootCommand.AddOption(removeEmptyRowsOption);
        #endregion

        rootCommand.SetHandler(
            (InvocationContext context) =>
            {
                App app = new(
                    filePath: context.ParseResult.GetValueForArgument(fileArgument),
                    outputDirectory: context.ParseResult.GetValueForOption(outputDirectoryOption) ?? string.Empty,
                    sheetNames: context.ParseResult.GetValueForOption(sheetSelectionOption) ?? [],
                    sheetRenames: context.ParseResult.GetValueForOption(renameOption) ?? [],
                    indexed: context.ParseResult.GetValueForOption(indexColumnOption),
                    hidden: context.ParseResult.GetValueForOption(hiddenSheetsOption),
                    nullErrors: context.ParseResult.GetValueForOption(nullErrorsOption),
                    removeEmptyRows: context.ParseResult.GetValueForOption(removeEmptyRowsOption)
                );

                /*
                Console.WriteLine("Testing Version 1:");
                for(int i = 0; i < 5; i++)
                {
                    Console.WriteLine($"ROUND {i}");
                    app.RunV1();
                }
                */

                // Shared String may be an issue
                // We should time individual things and see how long
                // they take
                app.Run();
            }
        );

        return new CommandLineBuilder(rootCommand);
    }
}
