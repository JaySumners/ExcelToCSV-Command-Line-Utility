using System.IO;
using System.Linq;

namespace ExcelToCSV.Utilities;

internal static class PathUtility
{
    #region Properties
    #endregion

    #region Constructors
    #endregion

    #region Methods
    internal static bool FileNameContainsInvalidChars(string filePath)
    {
        string fileNameNoExt = Path.GetFileNameWithoutExtension(filePath);

        return
            (Path.GetInvalidFileNameChars().Any(fileNameNoExt.Contains))
            || (Path.GetInvalidPathChars().Any(filePath.Contains));
    }
    internal static bool PathContainsInvalidChars(string path)
    {
        return (Path.GetInvalidPathChars().Any(path.Contains));
    }
    internal static string ReplaceInvalidFileNameChars(string fileName)
    {
        return string.Join("_", fileName.Split(Path.GetInvalidFileNameChars()));
    }
    internal static string ReplaceInvalidPathCharacters(string path)
    {
        return string.Join("_", path.Split(Path.GetInvalidPathChars()));
    }
    #endregion
}
