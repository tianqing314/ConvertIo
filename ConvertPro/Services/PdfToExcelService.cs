using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using UglyToad.PdfPig;

namespace ConvertPro.Services;

/// <summary>
/// PDF → Excel 转换服务。
/// 使用 PdfPig 提取 PDF 文本行 → ClosedXML 生成 .xlsx。
/// 
/// 限制：当前实现提取所有文本行到单个工作表。
/// 表格自动识别和结构化提取建议使用 Aspose.PDF 或 Tabula。
/// </summary>
public class PdfToExcelService : IConversionService
{
    public string ConversionType => "pdf2excel";
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

                ExtractPdfToExcel(file, xlsxPath);
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

    private void ExtractPdfToExcel(string pdfPath, string xlsxPath)
    {
        var rows = new List<List<string>>();

        using (var pdf = PdfDocument.Open(pdfPath))
        {
            foreach (var page in pdf.GetPages())
            {
                var text = page.Text?.Trim();
                if (string.IsNullOrEmpty(text)) continue;

                var lines = text.Split('\n')
                    .Select(l => l.Trim())
                    .Where(l => l.Length > 0)
                    .ToList();

                foreach (var line in lines)
                {
                    // 简单按空格/制表符分割为列
                    var cells = line.Split([' ', '\t'],
                            StringSplitOptions.RemoveEmptyEntries)
                        .ToList();
                    if (cells.Count > 0)
                        rows.Add(cells);
                }
            }
        }

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("提取内容");
        for (int r = 0; r < rows.Count; r++)
        {
            for (int c = 0; c < rows[r].Count; c++)
            {
                ws.Cell(r + 1, c + 1).Value = rows[r][c];
            }
        }

        wb.SaveAs(xlsxPath);
    }
}
