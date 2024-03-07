using System;

namespace ExcelToCSV.Models;

internal readonly record struct RangeReferenceModel
{
    #region Properties
    internal readonly string A1RangeReference;
    internal readonly CellReferenceModel StartCellReference;
    internal readonly CellReferenceModel EndCellReference;
    #endregion

    #region Constructors
    public RangeReferenceModel(string A1RangeReference)
    {
        if (!A1RangeReference.Contains(':'))
        {
            throw new Exception("'A1RangeReference' must have a colon (':').");
        }

        this.A1RangeReference = A1RangeReference;

        string[] a1References = this.A1RangeReference.Split(":");
        StartCellReference = new CellReferenceModel(a1References[0]);
        EndCellReference = new CellReferenceModel(a1References[1]);
    }
    #endregion

    #region Methods
    #endregion
}
