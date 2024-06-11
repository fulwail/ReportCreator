using ClosedXML.Excel;

namespace ReportCreator.Models;

public class ReportCellStyleDto
{
    public XLColor?  BackgroundColor { get; set; }
    public XLColor?  FontColor { get; set; }
}