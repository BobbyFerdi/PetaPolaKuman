using ClosedXML.Excel;
using Travewell.Library.Extensions;

namespace PetaPolaKuman.Utilities;

public static class Extensions
{
    public static int ExcelColumnNameToNumber(this string columnName)
    {
        if (columnName.IsNullOrEmpty()) throw new ArgumentNullException(nameof(columnName));

        columnName = columnName.ToUpperInvariant();

        var sum = 0;

        for (var i = 0; i < columnName.Length; i++)
        {
            sum *= 26;
            sum += (columnName[i] - 'A' + 1);
        }

        return sum;
    }

    public static string GetCellValue(this IXLRangeRow row, string cell) => row.Cell(cell.ExcelColumnNameToNumber()).GetString();

    public static string GetCellValue(this IXLRangeRow row, int cellNumber) => row.Cell(cellNumber).GetString();
}