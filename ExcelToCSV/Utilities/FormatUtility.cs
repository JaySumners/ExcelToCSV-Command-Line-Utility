using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelToCSV.Utilities;

internal static class FormatUtility
{
    #region Properties
    private static readonly List<string> _errorValues =
        [
            "#N/A",
            "#REF!",
            "#VALUE!",
            "#NAME?",
            "#DIV/O!",
            "#DIV/0!",
            "#NULL!",
            "#NUM!"
        ];

    private static readonly List<uint> _exponentialIds = [11, 48];

    private static readonly List<uint> _dateTimeIds =
        [
            //14 - 22
            14,
            15,
            16,
            17,
            18,
            19,
            20,
            21,
            22,
            //27 - 36
            27,
            28,
            29,
            30,
            31,
            32,
            33,
            34,
            35,
            36,
            //45 - 47
            45,
            46,
            47,
            //50 - 58
            50,
            51,
            52,
            53,
            54,
            55,
            56,
            57,
            58,
            //81
            81
        ];

    private static readonly char[] _dateTimeChars = ['d', 'm', 'y', 'h', 's'];

    #endregion

    #region Constructors
    #endregion

    #region Methods

    #region Exponentials (Scientific)
    private static bool IsExponentialCode(string formatCode)
    {
        return formatCode.Contains('E');
    }
    private static string FormatExponential(string cellValue)
    {
        string formattedString = cellValue;

        if (decimal.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal cellDecimalValue))
        {
            formattedString = cellDecimalValue.ToString();
        }

        return formattedString;
    }
    #endregion

    #region DateTime
    internal static bool IsDateTimeCode(string formatCode)
    {
        char[] formatCodeChars = formatCode.ToCharArray();
        return (formatCodeChars.Intersect(_dateTimeChars).Any());
    }
    private static string FormatDateTime(string cellValue)
    {
        string formattedString = cellValue;

        if (double.TryParse(cellValue, out double cellNumericValue))
        {
            DateTime cellAsDateTime = DateTime.FromOADate(cellNumericValue);
            formattedString = cellAsDateTime.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
        }

        return formattedString;
    }
    #endregion

    #region Public Facing
    internal static bool IsExcelError(string cellValue)
    {
        return _errorValues.Contains(cellValue);
    }
    internal static string Format(string cellValue, UInt32Value formatId, string formatCode)
    {
        if (formatId <= 163)
        {
            // Rests on the idea that we've properly mapped those earlier ones.
            // We may be able to just get rid of this at some point.
            return Format(cellValue, formatId);
        }

        return Format(cellValue, formatCode);
    }
    internal static string Format(string cellValue, UInt32Value formatId)
    {
        return formatId switch
        {
            var id when _exponentialIds.Contains(id) => FormatExponential(cellValue),
            var id when _dateTimeIds.Contains(id) => FormatDateTime(cellValue),
            _ => cellValue
        };
    }
    internal static string Format(string cellValue, string formatCode)
    {
        return formatCode switch
        {
            var code when IsExponentialCode(code) => FormatExponential(cellValue),
            var code when IsDateTimeCode(code) => FormatDateTime(cellValue),
            _ => cellValue
        };
    }
    #endregion
    #endregion
}
