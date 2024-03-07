using System.Linq;

namespace ExcelToCSV.Models;

internal readonly record struct CellReferenceModel(string A1Reference)
{
    #region Properties
    //internal string A1Reference; //Implicit
    internal readonly string ColumnReference => GetColumnReference();
    internal readonly int ColumnNumber => GetColumnNumber();
    internal readonly int RowNumber => GetRowNumber();
    #endregion

    #region Constructors
    #endregion

    #region Static Methods
    private static int ConvertCharToIndex(char c)
    {
        // Bitwise shift will return 1-26 regardless of lower/uppercase
        // Alternative is (int)(c.ToUpper()) - 26;
        return (int)c & 0x1F;
    }
    private readonly string GetColumnReference()
    {
        return new string(this.A1Reference.ToCharArray().Where(char.IsLetter).ToArray());
    }
    private readonly int GetColumnNumber()
    {
        return this.A1Reference
            .ToCharArray()
            .Where(char.IsLetter)
            .Reverse()
            .Select((c, i) => i == 0 ? ConvertCharToIndex(c) : ConvertCharToIndex(c) * i * 26)
            .Sum();
    }
    private readonly int GetRowNumber()
    {
        string rowNumString = new(this.A1Reference.ToCharArray().Where(char.IsNumber).ToArray());

        if (int.TryParse(rowNumString, out int rowNumInt))
        {
            return rowNumInt;
        }

        return 0;
    }
    #endregion
}
