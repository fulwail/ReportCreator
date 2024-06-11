namespace ReportCreator.Models;

public class ReportHeaderDto
{
    public string Description { get; set; }
    public int Level  { get; set; }
    public int Colspan { get; set; }
    public int Rowspan { get; set; }
    public int OffsetCols { get; set; }
    public int OffsetRows { get; set; }
}