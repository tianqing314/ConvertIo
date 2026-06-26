using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Tabula;
using Tabula.Extractors;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis;

namespace ConvertPro.Services;

/// <summary>
/// PDF → Excel 转换服务。
/// 使用 Tabula-sharp（基于 PdfPig）自动识别 PDF 表格结构，输出到 .xlsx。
///
/// 两种识别模式：
/// - Lattice（SpreadsheetExtractionAlgorithm）：适用于有边框线的表格，准确率高
/// - Stream  （BasicExtractionAlgorithm）：适用于无边框表格，基于空白对齐推断
///
/// 输出策略：
/// - mode=table：仅输出识别到的表格（每个表格一个工作表）
/// - mode=all  ：表格 + 文本回退（无表格的页面把文本行写入工作表）
/// </summary>
public class PdfToExcelService : IConversionService
{
    public string ConversionType => "pdf2excel";
    public bool SupportsBatch => true;

    public async Task<List<ConversionResult>> ConvertAsync(
        List<string> inputFiles,
        string outputDir,
        string? options,
        IProgress<ConversionProgress> progress,
        CancellationToken ct)
    {
        // 解析 options：默认 mode=all，兼容旧调用方
        var mode = "all";
        if (!string.IsNullOrWhiteSpace(options))
        {
            try
            {
                using var opts = JsonDocument.Parse(options);
                if (opts.RootElement.TryGetProperty("mode", out var mEl) &&
                    mEl.ValueKind == JsonValueKind.String)
                {
                    var v = mEl.GetString()?.Trim().ToLowerInvariant();
                    if (v == "table" || v == "all") mode = v;
                }
            }
            catch { /* 忽略 options 解析失败，按默认走 */ }
        }

        // Tabula 是 CPU 密集型，文件之间相互独立，可并行处理
        // 限并发 3，避免一次性启动过多 PdfDocument.Open 占用太多内存
        var results = new ConcurrentBag<ConversionResult>();
        int completedCount = 0;
        int totalCount = inputFiles.Count;

        var parallelOpts = new ParallelOptions
        {
            MaxDegreeOfParallelism = Math.Min(3, Environment.ProcessorCount),
            CancellationToken = ct
        };

        await Parallel.ForEachAsync(inputFiles, parallelOpts, async (file, token) =>
        {
            token.ThrowIfCancellationRequested();

            // 报告"开始处理此文件"
            int started = Interlocked.Increment(ref completedCount);
            progress?.Report(new ConversionProgress(
                totalCount, started - 1, Path.GetFileName(file),
                (double)(started - 1) / totalCount * 100));

            try
            {
                var xlsxPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(file) + ".xlsx");

                // Tabula 是 CPU 工作，用 Task.Run 让它真正并行
                await Task.Run(() => ExtractWithTabula(file, xlsxPath, mode), token);

                results.Add(new ConversionResult(true, xlsxPath,
                    OutputSize: new FileInfo(xlsxPath).Length));
            }
            catch (OperationCanceledException)
            {
                results.Add(new ConversionResult(false, "", "已取消"));
            }
            catch (Exception ex)
            {
                // Tabula 失败时回退到纯文本提取，保证可用性
                try
                {
                    var xlsxPath = Path.Combine(outputDir,
                        Path.GetFileNameWithoutExtension(file) + ".xlsx");
                    await Task.Run(() => FallbackTextExtract(file, xlsxPath), token);
                    results.Add(new ConversionResult(true, xlsxPath,
                        OutputSize: new FileInfo(xlsxPath).Length,
                        ErrorMessage: "Tabula 表格识别失败，已回退到纯文本提取: " + ex.Message));
                }
                catch (Exception ex2)
                {
                    results.Add(new ConversionResult(false, "", ex2.Message));
                }
            }

            progress?.Report(new ConversionProgress(
                totalCount, started, "",
                (double)started / totalCount * 100));
        });

        // 保持原始顺序（ConcurrentBag 不保证顺序）
        var ordered = inputFiles
            .Select(f => Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(f) + ".xlsx"))
            .Select(p => results.FirstOrDefault(r => r.OutputPath == p))
            .Where(r => r != null)
            .Select(r => r!)
            .ToList();

        return ordered;
    }

    // ================================================================
    // Tabula 真正的表格识别
    // ================================================================
    private void ExtractWithTabula(string pdfPath, string xlsxPath, string mode)
    {
        // ClipPaths=true 是 Tabula 要求的，用于正确处理路径
        var parsingOptions = new ParsingOptions { ClipPaths = true };

        using var pdf = PdfDocument.Open(pdfPath, parsingOptions);
        using var wb = new XLWorkbook();

        var tableSheets = 0;     // 已写入的表格 sheet 数
        var textFallbackSheets = 0; // 回退文本 sheet 数
        var pageCount = pdf.GetPages().Count();

        int pageIndex = 1;
        foreach (var page in pdf.GetPages())
        {
            PageArea pageArea;
            try
            {
                pageArea = ObjectExtractor.Extract(pdf, page.Number);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[Tabula] 第 {page.Number} 页提取失败: {ex.Message}");
                pageIndex++;
                continue;
            }

            // 先尝试 Lattice（带边框表格），再尝试 Stream（无边框表格）
            var tables = new List<Table>();
            try
            {
                var lattice = new SpreadsheetExtractionAlgorithm();
                var latticeTables = lattice.Extract(pageArea);
                if (latticeTables != null && latticeTables.Count > 0)
                    tables.AddRange(latticeTables);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[Tabula] 第 {page.Number} 页 Lattice 算法失败: {ex.Message}");
            }

            if (tables.Count == 0)
            {
                try
                {
                    var stream = new BasicExtractionAlgorithm();
                    var streamTables = stream.Extract(pageArea);
                    if (streamTables != null && streamTables.Count > 0)
                        tables.AddRange(streamTables);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[Tabula] 第 {page.Number} 页 Stream 算法失败: {ex.Message}");
                }
            }

            if (tables.Count > 0)
            {
                foreach (var table in tables)
                {
                    tableSheets++;
                    var sheetName = MakeSheetName("表格", tableSheets, pageCount, page.Number);
                    var ws = wb.Worksheets.Add(sheetName);
                    WriteTableToSheet(ws, table);
                    FormatHeaderRow(ws);
                    AutoFitColumns(ws);
                }
            }
            else if (mode == "all")
            {
                // 无表格：把整页文本写入回退 sheet
                var pageText = page.Text?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(pageText))
                {
                    pageIndex++;
                    continue;
                }

                textFallbackSheets++;
                var sheetName = MakeSheetName("文本", textFallbackSheets, pageCount, page.Number);
                var wsText = wb.Worksheets.Add(sheetName);
                var lines = pageText.Split('\n')
                    .Select(l => l.Trim())
                    .Where(l => l.Length > 0)
                    .ToList();
                for (int li = 0; li < lines.Count; li++)
                {
                    var cells = lines[li].Split([' ', '\t'],
                        StringSplitOptions.RemoveEmptyEntries);
                    for (int ci = 0; ci < cells.Length; ci++)
                        wsText.Cell(li + 1, ci + 1).Value = cells[ci];
                }
                AutoFitColumns(wsText);
            }

            pageIndex++;
        }

        // 极端情况：啥都没识别到，写一个空 sheet 避免保存失败
        if (wb.Worksheets.Count == 0)
        {
            var ws = wb.Worksheets.Add("空文档");
            ws.Cell(1, 1).Value = "(未识别到任何内容)";
        }

        wb.SaveAs(xlsxPath);
    }

    // ================================================================
    // 把一个 Tabula Table 写入 ClosedXML 工作表
    // ================================================================
    private static void WriteTableToSheet(IXLWorksheet ws, Table table)
    {
        int row = 1;
        foreach (var r in table.Rows)
        {
            int col = 1;
            // Tabula-sharp 的 Row 本身就是 IReadOnlyList<Cell>，直接迭代
            foreach (var cell in r)
            {
                // Cell 可能为 null（合并单元格的空隙）；Cell.GetText() 是方法不是属性
                var text = cell?.GetText()?.Trim() ?? string.Empty;
                // 尝试按数字写入，否则按字符串
                if (!string.IsNullOrEmpty(text))
                {
                    if (double.TryParse(text, out var num))
                        ws.Cell(row, col).Value = num;
                    else
                        ws.Cell(row, col).Value = text;
                }
                col++;
            }
            row++;
        }
    }

    // 把第一行加粗、加底色，模拟表头
    private static void FormatHeaderRow(IXLWorksheet ws)
    {
        var firstRow = ws.FirstRow();
        if (firstRow == null) return;
        firstRow.Style.Font.Bold = true;
        firstRow.Style.Fill.BackgroundColor = XLColor.FromHtml("#4F81BD");
        firstRow.Style.Font.FontColor = XLColor.White;
    }

    private static void AutoFitColumns(IXLWorksheet ws)
    {
        foreach (var col in ws.ColumnsUsed())
        {
            try { col.AdjustToContents(); } catch { /* 忽略个别列宽度计算异常 */ }
        }
    }

    // 工作表名长度 ≤ 31 字符（Excel 限制），且不重复
    private static string MakeSheetName(string prefix, int idx, int pageCount, int pageNum)
    {
        var name = pageCount > 1 ? $"{prefix}{idx}_p{pageNum}" : $"{prefix}{idx}";
        if (name.Length > 31) name = name.Substring(0, 31);
        return name;
    }

    // ================================================================
    // 回退：Tabula 完全失败时的纯文本提取（与旧实现相同）
    // ================================================================
    private void FallbackTextExtract(string pdfPath, string xlsxPath)
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
                ws.Cell(r + 1, c + 1).Value = rows[r][c];
        }
        wb.SaveAs(xlsxPath);
    }
}
