using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace ConvertPro.Services;

/// <summary>
/// Word → Excel 转换服务。
/// 用 OpenXml 提取 .docx 中所有 Table（含文档段落作为附注 sheet），
/// 用 ClosedXML 写入 .xlsx：每个 Word 表格 → 一个 worksheet。
/// </summary>
public class WordToExcelService : IConversionService
{
    public string ConversionType => "word2excel";
    public bool SupportsBatch => true;

    public Task<List<ConversionResult>> ConvertAsync(
        List<string> inputFiles,
        string outputDir,
        string? options,
        IProgress<ConversionProgress> progress,
        CancellationToken ct)
    {
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
                var xlsxPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(file) + ".xlsx");

                ConvertWordToExcel(file, xlsxPath);
                results.Add(new ConversionResult(true, xlsxPath,
                    OutputSize: new FileInfo(xlsxPath).Length));
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

    private static void ConvertWordToExcel(string docxPath, string xlsxPath)
    {
        using var wdoc = WordprocessingDocument.Open(docxPath, false);
        var body = wdoc.MainDocumentPart?.Document.Body;
        if (body == null)
            throw new InvalidOperationException("Word 文档无正文");

        // 收集所有表格
        var tables = body.Elements<W.Table>().ToList();

        using var wb = new XLWorkbook();

        if (tables.Count == 0)
        {
            // 没有表格：把段落写到 sheet "文档内容"
            var ws = wb.AddWorksheet("文档内容");
            var paragraphs = body.Elements<W.Paragraph>()
                .Select(p => p.InnerText?.Trim() ?? "")
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();

            ws.Cell(1, 1).Value = "段落";
            ws.Cell(1, 1).Style.Font.Bold = true;
            for (int i = 0; i < paragraphs.Count; i++)
                ws.Cell(i + 2, 1).Value = paragraphs[i];
            ws.Columns().AdjustToContents();
        }
        else
        {
            for (int t = 0; t < tables.Count; t++)
            {
                var sheetName = $"表格{t + 1}";
                if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31); // Excel sheet 名长度上限
                var ws = wb.AddWorksheet(sheetName);

                var rows = tables[t].Elements<W.TableRow>().ToList();
                for (int r = 0; r < rows.Count; r++)
                {
                    var cells = rows[r].Elements<W.TableCell>().ToList();
                    for (int c = 0; c < cells.Count; c++)
                    {
                        var text = cells[c].InnerText?.Trim() ?? "";
                        var cell = ws.Cell(r + 1, c + 1);
                        // 尝试解析数字
                        if (double.TryParse(text, out var num))
                            cell.Value = num;
                        else
                            cell.Value = text;
                    }
                }

                // 首行加粗
                if (rows.Count > 0)
                {
                    var firstRow = ws.Row(1);
                    firstRow.Style.Font.Bold = true;
                    firstRow.Style.Fill.BackgroundColor = XLColor.LightGray;
                }

                ws.Columns().AdjustToContents();
            }
        }

        if (wb.Worksheets.Count == 0)
            wb.AddWorksheet("空").Cell(1, 1).Value = "(文档无内容)";

        wb.SaveAs(xlsxPath);
    }
}
