using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace ConvertPro.Services;

/// <summary>
/// Excel → PDF 转换服务。
/// 使用 ClosedXML 读取 .xlsx，QuestPDF 生成 .pdf。
/// 将每个工作表渲染为 PDF 页面。
/// </summary>
public class ExcelToPdfService : IConversionService
{
    public string ConversionType => "excel2pdf";
    public bool SupportsBatch => true;

    public Task<List<ConversionResult>> ConvertAsync(
        List<string> inputFiles,
        string outputDir,
        string? options,
        IProgress<ConversionProgress> progress,
        CancellationToken ct)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        var results = new List<ConversionResult>();

        for (int i = 0; i < inputFiles.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var file = inputFiles[i];

            progress?.Report(new ConversionProgress(
                inputFiles.Count, i, Path.GetFileName(file),
                (double)i / inputFiles.Count * 100));

            try
            {
                var pdfPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(file) + ".pdf");

                ConvertExcelToPdf(file, pdfPath);
                results.Add(new ConversionResult(true, pdfPath,
                    OutputSize: new FileInfo(pdfPath).Length));
            }
            catch (OperationCanceledException)
            {
                results.Add(new ConversionResult(false, "", "已取消"));
                break;
            }
            catch (Exception ex)
            {
                results.Add(new ConversionResult(false, "", ex.Message));
            }

            progress?.Report(new ConversionProgress(
                inputFiles.Count, i + 1, "",
                (double)(i + 1) / inputFiles.Count * 100));
        }

        return Task.FromResult(results);
    }

    private void ConvertExcelToPdf(string xlsxPath, string pdfPath)
    {
        using var wb = new XLWorkbook(xlsxPath);

        QuestPDF.Fluent.Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4.Landscape());
                page.Margin(1, Unit.Centimetre);

                page.Content().Column(col =>
                {
                    foreach (var ws in wb.Worksheets)
                    {
                        var usedRange = ws.RangeUsed();
                        if (usedRange == null) continue;

                        col.Item().Text(ws.Name)
                            .FontSize(14).Bold().FontColor(Colors.Blue.Darken2);

                        col.Item().Table(table =>
                        {
                            table.ColumnsDefinition(c =>
                            {
                                for (int i = 0; i < usedRange.ColumnCount(); i++)
                                    c.RelativeColumn();
                            });

                            // 表头
                            table.Header(header =>
                            {
                                for (int c = 0; c < usedRange.ColumnCount(); c++)
                                {
                                    header.Cell()
                                        .Background(Colors.Grey.Lighten3)
                                        .Padding(4)
                                        .Text(usedRange.Cell(1, c + 1)
                                            .GetString())
                                        .FontSize(9).Bold();
                                }
                            });

                            // 数据行
                            for (int r = 2; r <= usedRange.RowCount(); r++)
                            {
                                var rowIdx = r; // capture
                                for (int c = 0; c < usedRange.ColumnCount(); c++)
                                {
                                    var colIdx = c;
                                    table.Cell()
                                        .BorderBottom(1)
                                        .BorderColor(Colors.Grey.Lighten2)
                                        .Padding(3)
                                        .Text(usedRange.Cell(rowIdx, colIdx + 1)
                                            .GetString())
                                        .FontSize(8);
                                }
                            }
                        });

                        col.Item().PaddingTop(10);
                    }
                });
            });
        }).GeneratePdf(pdfPath);
    }
}
