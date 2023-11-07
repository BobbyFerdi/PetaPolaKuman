using ClosedXML.Excel;

namespace PetaPolaKuman.Utilities;

public class Formatter
{
    public static XLColor GetNumberColor(double rate)
    {
        if (rate <= 50) return XLColor.Red;
        if (rate > 50 && rate <= 75) return XLColor.Yellow;
        if (rate > 75) return XLColor.Green;

        return null;
    }
}