namespace ReportCreator.Models;

public class ReportDto
{
    public string Title { get; set; }
    public List<ReportHeaderDto> Headers { get; set; }
    public IEnumerable<ReportRowDto> Rows { get; set; }
}