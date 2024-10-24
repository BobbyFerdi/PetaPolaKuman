using ClosedXML.Excel;

namespace PetaPolaKuman.Utilities;

public class Formatter
{
    public static XLColor GetNumberColor(double rate) =>
        rate switch
        {
            <= 50 => XLColor.Red,
            > 50 and <= 75 => XLColor.Yellow,
            > 75 => XLColor.Green,
            _ => null
        };
}