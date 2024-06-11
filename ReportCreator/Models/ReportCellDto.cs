using ClosedXML.Excel;

namespace ReportCreator.Models;

public class ReportCellDto
{
    public ReportCellDto()
    {
        Style = new ReportCellStyleDto();
    }

    public XLCellValue Value { get; set; }

    public int Colspan { get; set; }
    public int Rowspan { get; set; }

    public ReportCellStyleDto Style { get; set; }
}