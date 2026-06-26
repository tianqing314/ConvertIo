using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;

namespace ConvertPro.Services;

/// <summary>
/// Word → PPT 转换服务（AI 增强）。
///
/// 流程与 PdfToPptService 对称：
/// 1. 用 OpenXml 提取 .docx 全文（段落 + 表格文本）
/// 2. 复用 PdfToPptService 的 AI 大纲 + 文生图 + PptRenderer 管线
///
/// 回退：若 AI 未配置或失败，按段落分页生成基础 PPT。
/// </summary>
public class WordToPptService : IConversionService
{
    public string ConversionType => "word2ppt";
    public bool SupportsBatch => true;

    public async Task<List<ConversionResult>> ConvertAsync(
        List<string> inputFiles,
        string outputDir,
        string? options,
        IProgress<ConversionProgress> progress,
        CancellationToken ct)
    {
        // 解析 options.template
        var templateId = "cyber_blue";
        if (!string.IsNullOrWhiteSpace(options))
        {
            try
            {
                using var opts = JsonDocument.Parse(options);
                if (opts.RootElement.TryGetProperty("template", out var tEl) &&
                    tEl.ValueKind == JsonValueKind.String)
                {
                    var v = tEl.GetString()?.Trim();
                    if (!string.IsNullOrEmpty(v)) templateId = v;
                }
            }
            catch { /* 忽略 options 解析失败 */ }
        }
        var template = PptTemplates.Get(templateId);

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
                var pptxPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(file) + ".pptx");

                progress?.Report(new ConversionProgress(
                    inputFiles.Count, i, Path.GetFileName(file) + " · 提取内容…",
                    (double)i / inputFiles.Count * 100 + 5));

                var docText = ExtractWordText(file);

                ConversionResult result;

                if (AiProviderManager.AnyAvailable)
                {
                    progress?.Report(new ConversionProgress(
                        inputFiles.Count, i, Path.GetFileName(file) + " · AI 生成大纲…",
                        (double)i / inputFiles.Count * 100 + 20));

                    try
                    {
                        var content = await PdfToPptService.GenerateContentWithAiAsync(docText, ct);

                        Dictionary<string, byte[]>? images = null;
                        if (ImageGenManager.AnyAvailable)
                        {
                            progress?.Report(new ConversionProgress(
                                inputFiles.Count, i, Path.GetFileName(file) + " · AI 生成配图…",
                                (double)i / inputFiles.Count * 100 + 50));
                            try
                            {
                                images = await PdfToPptService.GenerateImagesAsync(content, progress,
                                    inputFiles.Count, i, ct);
                            }
                            catch (OperationCanceledException) { throw; }
                            catch (Exception imgEx)
                            {
                                System.Diagnostics.Debug.WriteLine(
                                    $"[WordToPpt] 文生图失败: {imgEx.Message}");
                            }
                        }

                        progress?.Report(new ConversionProgress(
                            inputFiles.Count, i, Path.GetFileName(file) + " · 渲染幻灯片…",
                            (double)i / inputFiles.Count * 100 + 85));

                        var imagesArg = images;
                        await Task.Run(() => PptRenderer.Build(pptxPath, template, content, imagesArg), ct);

                        result = new ConversionResult(true, pptxPath,
                            OutputSize: new FileInfo(pptxPath).Length);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception aiEx)
                    {
                        progress?.Report(new ConversionProgress(
                            inputFiles.Count, i, Path.GetFileName(file) + " · AI 失败，回退基础模式…",
                            (double)i / inputFiles.Count * 100 + 80));

                        BuildFallbackPptx(docText, pptxPath, template);
                        result = new ConversionResult(true, pptxPath,
                            ErrorMessage: "AI 生成失败，已回退到基础模式（按段落分页）: " +
                                          PdfToPptService.Truncate(aiEx.Message, 200),
                            OutputSize: new FileInfo(pptxPath).Length);
                    }
                }
                else
                {
                    progress?.Report(new ConversionProgress(
                        inputFiles.Count, i, Path.GetFileName(file) + " · 基础模式（未配置 AI）…",
                        (double)i / inputFiles.Count * 100 + 30));

                    BuildFallbackPptx(docText, pptxPath, template);
                    result = new ConversionResult(true, pptxPath,
                        ErrorMessage: "AI 未配置，使用基础模式。" +
                                      "请在「设置」填写 API Key 以启用 AI 生成。",
                        OutputSize: new FileInfo(pptxPath).Length);
                }

                results.Add(result);
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

        return results;
    }

    // ================================================================
    // 提取 .docx 文本：段落 + 表格
    // 用 MainDocumentPart 根元素遍历，避免 Body.Descendants() 顺序错乱
    // ================================================================
    private static string ExtractWordText(string docxPath)
    {
        var sb = new StringBuilder();
        using var wdoc = WordprocessingDocument.Open(docxPath, false);
        var body = wdoc.MainDocumentPart?.Document.Body;
        if (body == null) return "";

        foreach (var child in body.ChildElements)
            {
                if (child is Paragraph p)
                {
                    var text = p.InnerText?.Trim();
                    if (!string.IsNullOrEmpty(text))
                    {
                        // 检测标题样式
                        var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                        if (!string.IsNullOrEmpty(styleId) && styleId.Contains("Heading", StringComparison.OrdinalIgnoreCase))
                            sb.AppendLine($"## {text}");
                        else
                            sb.AppendLine(text);
                    }
                }
            else if (child is Table tbl)
            {
                foreach (var row in tbl.Elements<TableRow>())
                {
                    var cells = row.Elements<TableCell>()
                        .Select(c => c.InnerText?.Trim() ?? "")
                        .Where(s => !string.IsNullOrEmpty(s));
                    sb.AppendLine(string.Join(" | ", cells));
                }
                sb.AppendLine();
            }
        }

        var full = sb.ToString().Trim();
        if (string.IsNullOrEmpty(full)) full = "(Word 文档无可提取文本)";
        return full;
    }

    // ================================================================
    // 回退模式：按段落分章节/页，用 PptRenderer 输出基础 PPT
    // ================================================================
    private static void BuildFallbackPptx(string docText, string pptxPath, PptTemplate tpl)
    {
        // 把文本按段落切，每 5 段拼成一张幻灯片
        var paragraphs = docText.Split('\n')
            .Select(s => s.Trim())
            .Where(s => !string.IsNullOrEmpty(s))
            .ToList();

        if (paragraphs.Count == 0) paragraphs.Add("(文档无内容)");

        var slides = new List<PptSlide>();
        for (int i = 0; i < paragraphs.Count; i += 5)
        {
            var batch = paragraphs.Skip(i).Take(5).ToList();
            var title = batch[0].Length > 30 ? batch[0].Substring(0, 30) + "…" : batch[0];
            slides.Add(new PptSlide(title, batch, Layout: "text"));
        }

        var content = new PptContent(
            Title: "文档演示",
            Subtitle: "",
            Sections: new List<PptSection>
            {
                new PptSection("内容", slides)
            },
            Summary: new List<string>());

        PptRenderer.Build(pptxPath, tpl, content, null);
    }
}
