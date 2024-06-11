using ReportCreator.Models;

namespace ReportCreator;

public interface IReportCreatorService
{
    Task<ReportResultDto> CreateReport(IEnumerable<ReportDto> reportDtos, string fileName);
}