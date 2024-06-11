using System.Drawing;
using ClosedXML.Excel;

namespace ReportCreator.Helpers;

internal static class ExcelHelper
{
    public static string GetLetterByNumber(int index)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        var value = "";

        if (index >= letters.Length)
            value += letters[index / letters.Length - 1];

        value += letters[index % letters.Length];
        return value;
    }

    public static XLColor GetFontColorByHex(string hex)
    {
        var color = ColorTranslator.FromHtml(hex);
        int r = Convert.ToInt16(color.R);
        int g = Convert.ToInt16(color.G);
        int b = Convert.ToInt16(color.B);
        var yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000;
        return (yiq >= 128) ? XLColor.Black : XLColor.White; 
    }
}