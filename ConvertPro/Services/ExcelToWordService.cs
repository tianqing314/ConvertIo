using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace ConvertPro.Services;

/// <summary>
/// Excel → Word 转换服务。
/// 用 ClosedXML 读取 .xlsx，用 OpenXml 生成 .docx：
/// 每个 worksheet → 一级标题 + 表格（保留首行加粗表头）。
/// </summary>
public class ExcelToWordService : IConversionService
{
    public string ConversionType => "excel2word";
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
                var docxPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(file) + ".docx");

                ConvertExcelToWord(file, docxPath);
                results.Add(new ConversionResult(true, docxPath,
                    OutputSize: new FileInfo(docxPath).Length));
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

    private static void ConvertExcelToWord(string xlsxPath, string docxPath)
    {
        using var wb = new XLWorkbook(xlsxPath);

        using var wdoc = WordprocessingDocument.Create(docxPath, WordprocessingDocumentType.Document);
        var mainPart = wdoc.AddMainDocumentPart();
        mainPart.Document = new W.Document(new W.Body());
        var body = mainPart.Document.Body!;

        bool firstSheet = true;
        foreach (var ws in wb.Worksheets)
        {
            var usedRange = ws.RangeUsed();
            if (usedRange == null) continue;

            // 第一个 sheet 前不加段落分隔
            if (!firstSheet)
            {
                body.AppendChild(new W.Paragraph());
            }
            firstSheet = false;

            // 工作表标题（Heading1）
            var headingPara = new W.Paragraph(
                new W.Run(
                    new W.RunProperties(new W.Bold()),
                    new W.Text(ws.Name))
            );
            headingPara.ParagraphProperties = new W.ParagraphProperties
            {
                ParagraphStyleId = new W.ParagraphStyleId { Val = "Heading1" }
            };
            body.AppendChild(headingPara);

            int rowCount = usedRange.RowCount();
            int colCount = usedRange.ColumnCount();

            if (rowCount == 0 || colCount == 0) continue;

            // 创建 Word 表格
            var table = new W.Table(
                new W.TableProperties(
                    new W.TableStyle { Val = "TableGrid" },
                    new W.TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" },
                    new W.TableBorders(
                        TopBorder(), BottomBorder(), LeftBorder(), RightBorder(),
                        InsideHorizontalBorder(), InsideVerticalBorder())));

            // 表格列宽（等分）
            var grid = new W.TableGrid();
            for (int c = 0; c < colCount; c++)
                grid.AppendChild(new W.GridColumn { Width = (2000 / colCount).ToString() });
            table.AppendChild(grid);

            for (int r = 1; r <= rowCount; r++)
            {
                var row = new W.TableRow();
                for (int c = 1; c <= colCount; c++)
                {
                    var cell = usedRange.Cell(r, c);
                    var text = cell.GetFormattedString() ?? "";
                    var isHeader = r == 1;

                    var cellEl = new W.TableCell(
                        new W.Paragraph(
                            new W.Run(
                                new W.RunProperties(isHeader ? new W.Bold() : new W.RunProperties()),
                                new W.Text(text))));

                    if (isHeader)
                    {
                        cellEl.AppendChild(new W.TableCellProperties(
                            new W.Shading { Fill = "EEEEEE", Val = ShadingPatternValues.Clear }));
                    }

                    row.AppendChild(cellEl);
                }
                table.AppendChild(row);
            }

            body.AppendChild(table);
            body.AppendChild(new W.Paragraph());
        }

        mainPart.Document.Save();
    }

    private static W.TopBorder TopBorder() => new() { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" };
    private static W.BottomBorder BottomBorder() => new() { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" };
    private static W.LeftBorder LeftBorder() => new() { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" };
    private static W.RightBorder RightBorder() => new() { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" };
    private static W.InsideHorizontalBorder InsideHorizontalBorder() => new() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" };
    private static W.InsideVerticalBorder InsideVerticalBorder() => new() { Val = BorderValues.Single, Size = 4, Color = "DDDDDD" };
}
