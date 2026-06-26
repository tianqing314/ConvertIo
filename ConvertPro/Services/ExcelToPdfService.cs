using System;
using System.Collections.Generic;
using System.Globalization;
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
/// 改进：保留列宽比例、单元格字体/背景色/对齐/数字格式、合并单元格跨越。
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
                        RenderWorksheet(col, ws);
                        col.Item().PaddingTop(10);
                    }
                });
            });
        }).GeneratePdf(pdfPath);
    }

    // ================================================================
    // 单个工作表渲染
    // ================================================================
    private void RenderWorksheet(ColumnDescriptor col, IXLWorksheet ws)
    {
        var usedRange = ws.RangeUsed();
        if (usedRange == null) return;

        // 工作表标题
        col.Item().Text(ws.Name)
            .FontSize(14).Bold().FontColor(Colors.Blue.Darken2);

        // 计算每列的相对权重（基于 Excel 列宽）
        int colCount = usedRange.ColumnCount();
        int rowCount = usedRange.RowCount();

        var colWidths = new double[colCount];
        double totalWidth = 0;
        for (int c = 0; c < colCount; c++)
        {
            // Excel 列宽单位约等于"字符数"，ClosedXML 已直接返回
            var w = ws.Column(usedRange.FirstColumn().ColumnNumber() + c).Width;
            if (w <= 0) w = 8.43; // Excel 默认列宽
            colWidths[c] = w;
            totalWidth += w;
        }

        // 收集合并单元格信息：把 (startRow, startCol) → (endRow, endCol) 映射建好
        // 渲染时遇到合并区域内的非起始单元格直接跳过；起始单元格按列跨度合并渲染
        var mergedMap = new Dictionary<(int r, int c), (int rowSpan, int colSpan)>();
        var coveredSet = new HashSet<(int r, int c)>();
        foreach (var mrange in ws.MergedRanges)
        {
            int r1 = mrange.FirstRow().RowNumber() - usedRange.FirstRow().RowNumber() + 1;
            int r2 = mrange.LastRow().RowNumber() - usedRange.FirstRow().RowNumber() + 1;
            int c1 = mrange.FirstColumn().ColumnNumber() - usedRange.FirstColumn().ColumnNumber() + 1;
            int c2 = mrange.LastColumn().ColumnNumber() - usedRange.FirstColumn().ColumnNumber() + 1;
            if (r1 < 1 || c1 < 1) continue;
            mergedMap[(r1, c1)] = (r2 - r1 + 1, c2 - c1 + 1);
            for (int r = r1; r <= r2; r++)
                for (int c = c1; c <= c2; c++)
                    if (r != r1 || c != c1) coveredSet.Add((r, c));
        }

        col.Item().Table(table =>
        {
            // 用相对权重设置列宽，让宽列在 PDF 中也宽
            table.ColumnsDefinition(c =>
            {
                for (int i = 0; i < colCount; i++)
                    c.RelativeColumn((float)(colWidths[i] / Math.Max(totalWidth, 1)));
            });

            // 不再用 header 单独画表头（用户可能没有"表头"概念），
            // 统一按行渲染：第一行加粗 + 背景色，便于视觉区分
            for (int r = 1; r <= rowCount; r++)
            {
                var rowIdx = r;
                for (int c = 1; c <= colCount; c++)
                {
                    var colIdx = c;

                    // 跳过被合并单元格覆盖的非起始位置
                    if (coveredSet.Contains((rowIdx, colIdx))) continue;

                    // 拿到 Cell
                    var cell = usedRange.Cell(rowIdx, colIdx);
                    var text = cell.GetFormattedString() ?? string.Empty;
                    var style = cell.Style;

                    // 计算合并跨度
                    int colSpan = 1, rowSpan = 1;
                    if (mergedMap.TryGetValue((rowIdx, colIdx), out var span))
                    {
                        colSpan = span.colSpan;
                        rowSpan = span.rowSpan;
                    }

                    // 字体属性
                    var textColor = XlColorToHex(style.Font.FontColor) ?? "#000000";
                    var fontSize = (float)(style.Font.FontSize > 0 ? style.Font.FontSize : 9);
                    var isBold = style.Font.Bold || rowIdx == 1; // 首行强制加粗
                    var bgHex = XlColorToHex(style.Fill.BackgroundColor);
                    // 首行无背景色时给个默认浅灰，便于视觉区分
                    var defaultBgHex = rowIdx == 1 ? "EEEEEE" : "FFFFFF";

                    // ===== 构建 cell =====
                    // Step 1: ColumnSpan/RowSpan 是 cell 特有方法，返回 ITableCellContainer
                    var cellBase = table.Cell();
                    if (colSpan > 1) cellBase = cellBase.ColumnSpan((uint)colSpan);
                    if (rowSpan > 1) cellBase = cellBase.RowSpan((uint)rowSpan);

                    // Step 2: 之后的 .Border()/.AlignXxx()/.Background()/.Padding()/.Text() 都返回 IContainer
                    // 不能再赋值回 cellBase（ITableCellContainer）。用 IContainer 局部变量承接
                    IContainer styled = style.Alignment.Horizontal switch
                    {
                        XLAlignmentHorizontalValues.Center => cellBase.AlignCenter(),
                        XLAlignmentHorizontalValues.Right => cellBase.AlignRight(),
                        _ => cellBase.AlignLeft()
                    };

                    // QuestPDF 的 .Bold() 不接受参数，是切换样式而非 setter
                    var textDesc = styled
                        .Border(0.5f)
                        .BorderColor(Colors.Grey.Lighten1)
                        .Background(bgHex ?? defaultBgHex)
                        .Padding(3)
                        .Text(text)
                        .FontColor(textColor)
                        .FontSize(fontSize);

                    if (isBold) textDesc.Bold();
                }
            }
        });
    }

    // ================================================================
    // ClosedXML XLColor → QuestPDF 接受的 hex 字符串（不带 #）
    // ================================================================
    private static string? XlColorToHex(XLColor color)
    {
        if (color == null) return null;
        try
        {
            if (color.ColorType == XLColorType.Color)
            {
                // Color.ToArgb() 返回 0xAARRGGBB，去掉 AA 透明度只保留 RRGGBB
                return color.Color.ToArgb().ToString("X8").Substring(2);
            }
            // Indexed 和 Theme 颜色映射复杂，留空让调用方用默认色
            return null;
        }
        catch
        {
            return null;
        }
    }
}
