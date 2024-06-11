using ClosedXML.Excel;
using ReportCreator.Helpers;
using ReportCreator.Models;

namespace ReportCreator;

public class ExcelReportCreatorService : IReportCreatorService
{
    private const int MaxColumnWidth = 50;
    private const int MinColumnWidth = 10;

    public async Task<ReportResultDto> CreateReport(IEnumerable<ReportDto> reportDtos, string title)
    {
        var workBook = new XLWorkbook();
        foreach (var reportDto in reportDtos)
        {
            var workSheet = workBook.Worksheets.Add(reportDto.Title ?? "Лист 1");
            var offsetRows = string.IsNullOrEmpty(reportDto.Title) ? 1 : 3;
            ApplyStylesForReport(workSheet);
            if (!string.IsNullOrEmpty(reportDto.Title))
                CreateReportTitle(reportDto.Title, workSheet);

            CreateReportTable(reportDto.Headers, reportDto.Rows, workSheet, offsetRows);
            ApplyStylesForTable(workSheet, offsetRows);
        }


        await using var ms = new MemoryStream();
        workBook.SaveAs(ms);

        return new ReportResultDto
        {
            ReportContent = ms.ToArray(),
            FileName = $"{title}.xlsx"
        };
    }

    private static void ApplyStylesForReport(IXLRangeBase xlWorksheet)
    {
        xlWorksheet.Style.Font.SetFontSize(12);
        xlWorksheet.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        xlWorksheet.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        xlWorksheet.Style.Alignment.SetWrapText(true);
    }

    private static void CreateReportTitle(string title, IXLWorksheet xlWorksheet)
    {
        xlWorksheet.Cell("A1").SetValue(title);
        xlWorksheet.Cell("A1").Style.Font.SetBold(true);
        xlWorksheet.Cell("A1").Style.Font.SetFontSize(16);
        xlWorksheet.Cell("A1").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
        xlWorksheet.Range($"A1:M1").Row(1).Merge();
    }


    private static void CreateReportTable(IReadOnlyList<ReportHeaderDto> headers, IEnumerable<ReportRowDto> rows,
        IXLWorksheet xlWorksheet, int offsetRows)
    {
        #region Работа с заголовком

        var levels = new Dictionary<int, LevelCell>();
        var lastIndex = 0;
        LevelCell currentLevel;
        for (var index = 0; index < headers.Count; index++)
        {
            var currentHeader = headers[index];

            if (!levels.ContainsKey(currentHeader.Level))
            {
                currentLevel = new LevelCell()
                {
                    Level = currentHeader.Level,
                    Index = 0
                };
                levels.Add(currentHeader.Level, currentLevel);
            }
            else
            {
                currentLevel = levels[currentHeader.Level];
            }

            var cellPosition = string.Join("", ExcelHelper.GetLetterByNumber(currentLevel.Index), offsetRows + currentLevel.Level);
            var levelIndex = currentLevel.Index;
            while (xlWorksheet.Cell(cellPosition).IsMerged() || xlWorksheet.Cell(cellPosition).IsEmpty() == false)
            {
                levelIndex += 1;
                cellPosition = string.Join("", ExcelHelper.GetLetterByNumber(levelIndex), offsetRows + currentLevel.Level);
            }

            if (currentHeader.Colspan != 0 || currentHeader.Rowspan != 0)
            {
                var cellToMergePosition = string.Join("",
                    ExcelHelper.GetLetterByNumber((currentHeader.Colspan != 0 ? currentHeader.Colspan - 1 : 0) + levelIndex),
                    offsetRows + (currentHeader.Rowspan != 0 ? currentHeader.Rowspan - 1 : 0) + currentLevel.Level);

                xlWorksheet.Range(cellPosition, cellToMergePosition).Merge();
                levels[currentHeader.Level].Index += (currentHeader.Colspan != 0 ? currentHeader.Colspan - 1 : 0) + 1;
            }
            else
            {
                levels[currentHeader.Level].Index += 1;
            }


            xlWorksheet.Cell(cellPosition).SetValue(currentHeader.Description).Style.Font.Bold = true;
        }

        #endregion

        var rowIndex = xlWorksheet.LastRowUsed().RowNumber() + 1;
        foreach (var row in rows)
        {
            for (var index = 0; index < row.Cells.Count; index++)
            {
                var element = row.Cells[index];
                var column = ExcelHelper.GetLetterByNumber(index);
                var cell = xlWorksheet.Cell($"{column}{rowIndex}");

                if (cell.IsMerged() == false)
                    cell.SetValue(element.Value);

                if (element.Rowspan != 0 && cell.IsMerged() == false)
                {
                    var cellToMerge = xlWorksheet.Cell($"{column}{rowIndex + element.Rowspan - 1}");
                    xlWorksheet.Range(cell, cellToMerge).Merge();
                }

                if (element.Style.BackgroundColor != null)
                    cell.Style.Fill.BackgroundColor = element.Style.BackgroundColor;
                if (element.Style.FontColor != null)
                    cell.Style.Font.FontColor = element.Style.FontColor;
            }

            rowIndex++;
        }
    }

    private static void ApplyStylesForTable(IXLWorksheet xlWorksheet, int offsetRows)
    {
        var offset = offsetRows == 1 ? 0 : 1;
        var lastRowIndex = xlWorksheet.RowsUsed().Count() + offset;
        var lastColumnLetter = ExcelHelper.GetLetterByNumber(xlWorksheet.ColumnsUsed().Count() - 1);
        var range = xlWorksheet.Range($"A{offsetRows}:{lastColumnLetter}{lastRowIndex}");
        range.Cells().Style.Alignment.SetWrapText(true);
        range.Rows().Style.Alignment.SetWrapText(true);
        range.Columns().Style.Alignment.SetWrapText(true);
        range.Cells().Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
        range.Cells().Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
        range.Cells().Style.Border.SetLeftBorder(XLBorderStyleValues.Thin);
        range.Cells().Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
        xlWorksheet.Columns().AdjustToContents(1, MinColumnWidth, MaxColumnWidth);
    }

   
}