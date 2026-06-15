using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace ConvertPro.Services;

/// <summary>
/// PDF → Word 转换服务。
/// 使用 PdfPig 提取 PDF 文本，OpenXml 生成 .docx。
/// 
/// 限制：仅提取纯文本和基本段落结构，不保留图片、表格、复杂排版。
/// 如需完整格式保真的 PDF→Word 转换，请集成 Aspose.PDF 或 Spire.PDF。
/// </summary>
public class PdfToWordService : IConversionService
{
    public string ConversionType => "pdf2word";
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

                ExtractPdfToDocx(file, docxPath);

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

    private void ExtractPdfToDocx(string pdfPath, string docxPath)
    {
        var lines = new List<string>();

        using (var pdf = PdfDocument.Open(pdfPath))
        {
            foreach (var page in pdf.GetPages())
            {
                var pageText = page.Text?.Trim();
                if (!string.IsNullOrEmpty(pageText))
                {
                    // 按段落分割
                    var paragraphLines = pageText.Split('\n')
                        .Select(l => l.Trim())
                        .Where(l => l.Length > 0);
                    lines.AddRange(paragraphLines);
                }
            }
        }

        // 创建 .docx
        using var doc = WordprocessingDocument.Create(docxPath,
            WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();

        foreach (var line in lines)
        {
            var para = new Paragraph();
            var run = new Run();
            // 检测是否可能是标题（短文本、以数字或"第"开头）
            var isHeadingLike = line.Length < 50 &&
                (char.IsDigit(line.FirstOrDefault()) ||
                 line.StartsWith("第") ||
                 line.StartsWith("一") ||
                 line.StartsWith("二") ||
                 line.StartsWith("三"));

            if (isHeadingLike)
            {
                run.AppendChild(new RunProperties(
                    new Bold(), new FontSize { Val = "28" }));
            }

            run.AppendChild(new Text(line));
            para.AppendChild(run);
            body.AppendChild(para);
        }

        mainPart.Document.AppendChild(body);
        mainPart.Document.Save();
    }
}
