using ReportCreator.Models;

namespace ReportCreator;

public abstract class BaseReportDataMapper<T,TEnum> 
{

    protected abstract string Title { get; }
    protected abstract IEnumerable<ReportRowDto> GetReportData(IEnumerable<T> exportTasks, TEnum groupingType,
        DateTime[] filteredPeriods);
    protected abstract  List<ReportHeaderDto> GetHeaders(IEnumerable<T> entities, TEnum groupingType,
        DateTime[] filteredPeriods);

    public ReportDto MapToReportData(IEnumerable<T> entities, TEnum groupingType, DateTime[] filteredPeriods=null)
    {
        var entitiesArray = entities as T[] ?? entities.ToArray();
        return new ReportDto
        {
            Headers = GetHeaders(entitiesArray,groupingType,filteredPeriods),
            Rows = GetReportData(entitiesArray,groupingType,filteredPeriods),
            Title = Title
        };
    }
}