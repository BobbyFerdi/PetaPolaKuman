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
            sum += columnName[i] - 'A' + 1;
        }

        return sum;
    }

    public static string GetCellValue(this IXLRangeRow row, string cell) => row.Cell(cell.ExcelColumnNameToNumber()).GetString();

    public static string GetCellValue(this IXLRangeRow row, int cellNumber) => row.Cell(cellNumber).GetString();

    public static IXLCell SetCellValue(this IXLWorksheet sheet, string cell, string value) => sheet.Cell(cell).SetValue(value);

    public static IXLCell SetCellValue(this IXLWorksheet sheet, int row, int cell, string value) => sheet.Cell(row, cell).SetValue(value);

    public static IXLCell SetCellValue(this IXLWorksheet sheet, string cell, double value) => sheet.Cell(cell).SetValue(value);

    public static IXLCell SetCellValue(this IXLWorksheet sheet, int row, int cell, double value) => sheet.Cell(row, cell).SetValue(value);

    public static List<string> DistinctAndOrder(this List<string> input) => input.DistinctBy(o => o.ToLower()).OrderBy(o => o).ToList();

    public static string OrganismTranslator(this string input) => input.ToLower() switch
    {
        "staphylococcus hemolyticus" => "Staphylococcus haemolyticus",
        "eschericia coli" => "Escherichia coli",
        "staphyloccous hominis" => "Staphylococcus hominis",
        _ => input,
    };
}