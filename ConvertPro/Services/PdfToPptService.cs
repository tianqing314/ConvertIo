using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using UglyToad.PdfPig;
using D = DocumentFormat.OpenXml.Drawing;

namespace ConvertPro.Services;

/// <summary>
/// PDF → PPT 转换服务（AI 增强）。
///
/// 流程：
/// 1. 用 PdfPig 提取 PDF 全文（带页码标记）
/// 2. 调用 AI（DeepSeek / MiMo）分析内容并生成结构化 JSON（PptContent）
/// 3. 用 PptRenderer 把 JSON 填充到用户选定的科技极简模板，输出 .pptx
///
/// 回退：若 AI 未配置或调用失败，退化为旧逻辑（每页文字 → 一张幻灯片），
///      并在 ConversionResult.ErrorMessage 中给出提示。
/// </summary>
public class PdfToPptService : IConversionService
{
    public string ConversionType => "pdf2ppt";
    public bool SupportsBatch => true;

    // 单次喂给 AI 的 PDF 文本上限（字符数），避免超出上下文窗口
    private const int MaxAiInputChars = 20000;

    public async Task<List<ConversionResult>> ConvertAsync(
        List<string> inputFiles,
        string outputDir,
        string? options,
        IProgress<ConversionProgress> progress,
        CancellationToken ct)
    {
        // 解析 options.template（用户在 pdf2ppt 页选的模板 id）
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
            catch { /* 忽略 options 解析失败，按默认模板走 */ }
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

                // ===== 1. 提取 PDF 文本 =====
                progress?.Report(new ConversionProgress(
                    inputFiles.Count, i, Path.GetFileName(file) + " · 提取内容…",
                    (double)i / inputFiles.Count * 100 + 5));

                var pdfText = ExtractPdfText(file);

                // ===== 2. AI 生成 or 回退 =====
                ConversionResult result;

                if (AiProviderManager.AnyAvailable)
                {
                    progress?.Report(new ConversionProgress(
                        inputFiles.Count, i, Path.GetFileName(file) + " · AI 生成大纲…",
                        (double)i / inputFiles.Count * 100 + 20));

                    try
                    {
                        var content = await GenerateContentWithAiAsync(pdfText, ct);

                        // 文生图（若图模型已配置）
                        Dictionary<string, byte[]>? images = null;
                        if (ImageGenManager.AnyAvailable)
                        {
                            progress?.Report(new ConversionProgress(
                                inputFiles.Count, i, Path.GetFileName(file) + " · AI 生成配图…",
                                (double)i / inputFiles.Count * 100 + 50));

                            try
                            {
                                images = await GenerateImagesAsync(content, progress,
                                    inputFiles.Count, i, ct);
                            }
                            catch (OperationCanceledException) { throw; }
                            catch (Exception imgEx)
                            {
                                // 文生图失败不致命，继续渲染（无图）
                                System.Diagnostics.Debug.WriteLine(
                                    $"[PdfToPpt] 文生图失败: {imgEx.Message}");
                            }
                        }

                        progress?.Report(new ConversionProgress(
                            inputFiles.Count, i, Path.GetFileName(file) + " · 渲染幻灯片…",
                            (double)i / inputFiles.Count * 100 + 85));

                        // CPU 密集型渲染放到线程池，避免阻塞 UI 线程
                        var imagesArg = images;
                        await Task.Run(() => PptRenderer.Build(pptxPath, template, content, imagesArg), ct);

                        result = new ConversionResult(true, pptxPath,
                            OutputSize: new FileInfo(pptxPath).Length);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception aiEx)
                    {
                        // AI 阶段失败：回退到旧模式，但标记警告
                        progress?.Report(new ConversionProgress(
                            inputFiles.Count, i, Path.GetFileName(file) + " · AI 失败，回退基础模式…",
                            (double)i / inputFiles.Count * 100 + 80));

                        await Task.Run(() => BuildFallbackPptx(file, pptxPath), ct);
                        result = new ConversionResult(true, pptxPath,
                            ErrorMessage: "AI 生成失败，已回退到基础模式（每页文本→幻灯片）: " +
                                          Truncate(aiEx.Message, 200),
                            OutputSize: new FileInfo(pptxPath).Length);
                    }
                }
                else
                {
                    // AI 未配置：直接走回退模式
                    progress?.Report(new ConversionProgress(
                        inputFiles.Count, i, Path.GetFileName(file) + " · 基础模式（未配置 AI）…",
                        (double)i / inputFiles.Count * 100 + 30));

                    await Task.Run(() => BuildFallbackPptx(file, pptxPath), ct);
                    result = new ConversionResult(true, pptxPath,
                        ErrorMessage: "AI 未配置，使用基础模式（效果一般）。" +
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
    // PDF 文本提取：拼接所有页，带页码标记
    // ================================================================
    private static string ExtractPdfText(string pdfPath)
    {
        var sb = new StringBuilder();
        using var pdf = PdfDocument.Open(pdfPath);
        foreach (var page in pdf.GetPages())
        {
            var text = page.Text?.Trim();
            if (string.IsNullOrEmpty(text)) continue;
            sb.AppendLine($"--- 第 {page.Number} 页 ---");
            sb.AppendLine(text);
            sb.AppendLine();
        }
        var full = sb.ToString().Trim();
        if (string.IsNullOrEmpty(full)) full = "(PDF 无可提取文本)";
        return full;
    }

    // ================================================================
    // AI 生成 PptContent（可被 WordToPptService 等复用）
    // ================================================================
    public static async Task<PptContent> GenerateContentWithAiAsync(
        string pdfText, CancellationToken ct)
    {
        // 截断超长内容
        var input = pdfText;
        if (input.Length > MaxAiInputChars)
        {
            input = input.Substring(0, MaxAiInputChars) +
                    "\n\n[注：原文较长，已截断，请基于以上内容生成大纲]";
        }

        var provider = AiProviderManager.Current;
        var model = AiProviderManager.GetCurrentModel();

        var systemPrompt = BuildSystemPrompt();
        var userPrompt = BuildUserPrompt(input);

        // jsonOutput=true 让 AI 返回纯 JSON（DeepSeek/MiMo 均支持 response_format）
        var raw = await provider.ChatAsync(systemPrompt, userPrompt,
            model: model, jsonOutput: true, ct: ct);

        var json = ExtractJson(raw);
        var content = PptContent.FromJson(json);

        // 基本校验：标题/章节不能全空
        if (string.IsNullOrWhiteSpace(content.Title) &&
            (content.Sections == null || content.Sections.Count == 0))
        {
            throw new InvalidOperationException("AI 返回的内容为空（无标题且无章节）");
        }

        // 兜底：确保集合非 null
        content = content with
        {
            Title = string.IsNullOrWhiteSpace(content.Title) ? "演示文稿" : content.Title.Trim(),
            Subtitle = content.Subtitle?.Trim() ?? "",
            Sections = content.Sections ?? new List<PptSection>(),
            Summary = content.Summary ?? new List<string>()
        };

        return content;
    }

    private static string BuildSystemPrompt()
    {
        return """
你是一个专业的演示文稿内容策划师。用户会给你一份文档的文本内容（来自 PDF 或 Word），你需要分析内容并生成结构化的 PPT 大纲，用于后续填充到精美模板中，并配 AI 文生图。

【硬性要求】
1. 只输出合法 JSON，不要任何额外文字、解释或 markdown 代码块标记（如 ```json）。
2. JSON 结构严格如下：
{
  "title": "主标题",
  "subtitle": "副标题",
  "coverImagePrompt": "封面配图的英文 prompt（可空字符串）",
  "sections": [{
    "title": "章节标题",
    "imagePrompt": "章节分隔页配图英文 prompt（可空字符串）",
    "slides": [{
      "title": "页面标题",
      "bullets": ["要点1", "要点2"],
      "note": "可选备注",
      "layout": "text|left_image|top_image|quote|chart",
      "icon": "单个emoji或空字符串",
      "imagePrompt": "本页配图英文 prompt（仅 left_image/top_image 必填，其它可空）",
      "chart": {
        "type": "bar|pie|line",
        "title": "图表标题",
        "labels": ["标签1","标签2"],
        "series": [{"name": "系列1", "values": [10, 20]}]
      }
    }]
  }],
  "summary": ["总结要点1", "总结要点2"]
}
3. title：演示文稿主标题，精炼有力，不超过 20 字。
4. subtitle：一句话副标题概括，可为空字符串 ""。
5. sections：3-6 个章节，按逻辑组织内容。
6. 每个 slide：title 为页面标题；bullets 为 3-6 条要点；note 可省略或为空字符串。
7. 每条 bullet 不超过 30 字，简洁、信息密度高，不要写成完整段落。
8. summary：3-5 条总结要点。
9. layout 选择规则：
   - 重要观点/金句 → "quote"
   - 数据展示（含数字对比/统计） → "chart"，并填 chart 字段
   - 概念说明/产品介绍 → "left_image"
   - 流程/案例/场景 → "top_image"
   - 普通要点列表 → "text"
10. icon：与该页内容相关的单个 emoji 字符，如 📊 🚀 💡 🎯 ⚙️ 📈 🏆 🔬 📋 🌐，可为空。
11. imagePrompt 规则：
    - 必须为英文，描述具体可生成的视觉场景，例如 "modern technology circuit board, blue glow, dark background"
    - 不超过 80 个英文单词
    - 大约 30-50% 的 left_image/top_image 页应指定 imagePrompt（避免每页都生成图太慢）
    - coverImagePrompt：封面图，1-2 个英文短句，描绘整体主题视觉
    - section.imagePrompt：章节分隔页配图（可选，约 50% 章节有图）
    - layout=text/quote/chart 的 slide，imagePrompt 应为空字符串
12. chart 字段（仅 layout="chart" 时填）：
    - type: bar（柱状图，适合对比）/ pie（饼图，适合占比）/ line（折线图，适合趋势）
    - title: 图表标题
    - labels: X轴或饼图分块标签，2-6 项
    - series: 数值系列，至少 1 个
    - 必须忠实于原文数据，禁止编造数字
13. 全部使用简体中文（imagePrompt 字段除外，必须英文）。
14. 内容必须忠实于原文，不要编造原文中没有的事实或数据。

请直接输出 JSON。
""";
    }

    private static string BuildUserPrompt(string pdfText)
    {
        return $"以下是文档的文本内容，请分析并生成结构化 PPT 大纲 JSON：\n\n{pdfText}";
    }

    // ================================================================
    // 从 AI 回复中提取 JSON（容忍 markdown 代码块包裹或前后多余文字）
    // ================================================================
    public static string ExtractJson(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            throw new InvalidOperationException("AI 返回空内容");

        var s = raw.Trim();

        // 情况1：被 ```json ... ``` 或 ``` ... ``` 包裹
        var fenceMatch = Regex.Match(s, "```(?:json)?\\s*([\\s\\S]*?)```",
            RegexOptions.IgnoreCase);
        if (fenceMatch.Success)
        {
            s = fenceMatch.Groups[1].Value.Trim();
        }

        // 情况2：前后有多余文字，截取第一个 { 到最后一个 }
        var firstBrace = s.IndexOf('{');
        var lastBrace = s.LastIndexOf('}');
        if (firstBrace >= 0 && lastBrace > firstBrace)
        {
            s = s.Substring(firstBrace, lastBrace - firstBrace + 1);
        }

        return s;
    }

    public static string Truncate(string s, int max) =>
        string.IsNullOrEmpty(s) ? s : (s.Length > max ? s.Substring(0, max) + "..." : s);

    // ================================================================
    // 并行调用文生图 API（public 以便 WordToPptService 复用）
    // ================================================================
    public static async Task<Dictionary<string, byte[]>> GenerateImagesAsync(
        PptContent content,
        IProgress<ConversionProgress> progress,
        int totalCount, int fileIndex, CancellationToken ct)
    {
        // 1. 收集所有去重的 imagePrompt
        var prompts = new HashSet<string>();
        if (!string.IsNullOrWhiteSpace(content.CoverImagePrompt))
            prompts.Add(content.CoverImagePrompt.Trim());

        if (content.Sections != null)
        {
            foreach (var section in content.Sections)
            {
                if (!string.IsNullOrWhiteSpace(section.ImagePrompt))
                    prompts.Add(section.ImagePrompt.Trim());

                if (section.Slides != null)
                {
                    foreach (var slide in section.Slides)
                    {
                        if (!string.IsNullOrWhiteSpace(slide.ImagePrompt))
                            prompts.Add(slide.ImagePrompt.Trim());
                    }
                }
            }
        }

        if (prompts.Count == 0) return new Dictionary<string, byte[]>();

        // 2. 并行下载，限制并发 2
        var result = new Dictionary<string, byte[]>();
        var provider = ImageGenManager.Current;
        var model = ImageGenManager.GetCurrentModel();

        using var sem = new SemaphoreSlim(2, 2);
        var tasks = prompts.Select(async prompt =>
        {
            await sem.WaitAsync(ct);
            try
            {
                var bytes = await provider.GenerateImageAsync(prompt, model, ct);
                lock (result)
                {
                    result[prompt] = bytes;
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                // 单张失败不致命
                System.Diagnostics.Debug.WriteLine(
                    $"[PdfToPpt] 图片生成失败 [{prompt.Substring(0, Math.Min(40, prompt.Length))}...]: {ex.Message}");
            }
            finally
            {
                sem.Release();
            }
        }).ToList();

        var totalImgs = prompts.Count;
        var completed = 0;
        while (tasks.Any(t => !t.IsCompleted))
        {
            await Task.WhenAny(tasks);
            completed = totalImgs - tasks.Count(t => !t.IsCompleted);
            progress?.Report(new ConversionProgress(
                totalCount, fileIndex,
                Path.GetFileName("") + $" · 生成配图 {completed}/{totalImgs}…",
                (double)fileIndex / totalCount * 100 + 50 + (double)completed / totalImgs * 30));
        }

        await Task.WhenAll(tasks);
        return result;
    }

    // ================================================================
    // 回退模式：旧逻辑（每页文本 → 一张幻灯片）
    // 当 AI 不可用或失败时使用
    // ================================================================
    private void BuildFallbackPptx(string pdfPath, string pptxPath)
    {
        var pages = new List<string>();
        using (var pdf = PdfDocument.Open(pdfPath))
        {
            foreach (var page in pdf.GetPages())
            {
                var text = page.Text?.Trim() ?? $"第 {page.Number} 页";
                pages.Add(text);
            }
        }
        if (pages.Count == 0) pages.Add("(PDF 无文字内容)");

        using var pptDoc = PresentationDocument.Create(pptxPath,
            PresentationDocumentType.Presentation);

        var presPart = pptDoc.AddPresentationPart();
        presPart.Presentation = new Presentation();
        presPart.Presentation.AppendChild(new SlideSize
        {
            Cx = 12192000, Cy = 6858000,
            Type = SlideSizeValues.Screen16x9
        });

        var masterPart = presPart.AddNewPart<SlideMasterPart>();
        var master = new SlideMaster(
            new CommonSlideData(new ShapeTree()),
            new SlideLayoutIdList());
        masterPart.SlideMaster = master;

        var masterId = new SlideMasterId
        {
            Id = 2147483648,
            RelationshipId = presPart.GetIdOfPart(masterPart)
        };
        presPart.Presentation.AppendChild(new SlideMasterIdList(masterId));

        var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
        layoutPart.SlideLayout = new SlideLayout(
            new CommonSlideData(
                new ShapeTree(new Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties
                        { Id = 1, Name = "Placeholder" },
                        new NonVisualShapeDrawingProperties()),
                    new ShapeProperties(),
                    new TextBody(
                        new D.BodyProperties(),
                        new D.ListStyle())))));
        layoutPart.SlideLayout.Save();

        var layoutId = new SlideLayoutId
        {
            Id = 2147483649,
            RelationshipId = masterPart.GetIdOfPart(layoutPart)
        };
        master.SlideLayoutIdList!.Append(layoutId);
        masterPart.SlideMaster.Save();

        var slideIdList = new SlideIdList();
        uint slideId = 256;

        for (int pi = 0; pi < pages.Count; pi++)
        {
            var content = pages[pi];
            if (content.Length > 2000) content = content[..2000];

            var slidePart = presPart.AddNewPart<SlidePart>();
            var shapeTree = new ShapeTree();

            var titleShape = BuildTextBox(
                $"第 {pi + 1} 页", 1,
                500000, 100000, 11000000, 600000,
                2400, true, "4f46e5");

            var lines = content.Split('\n')
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .Take(25)
                .ToList();
            var bodyText = string.Join("\n", lines);
            if (bodyText.Length > 5000) bodyText = bodyText[..5000];
            if (string.IsNullOrWhiteSpace(bodyText))
                bodyText = $"(第 {pi + 1} 页 — 无可提取文本)";

            var bodyShape = BuildTextBox(
                bodyText, 2,
                500000, 900000, 11000000, 5700000,
                1400, false, "333333");

            shapeTree.Append(titleShape);
            shapeTree.Append(bodyShape);

            slidePart.Slide = new Slide(new CommonSlideData(shapeTree));
            slidePart.AddPart(layoutPart);
            slidePart.Slide.Save();

            slideIdList.AppendChild(new SlideId
            {
                Id = slideId++,
                RelationshipId = presPart.GetIdOfPart(slidePart)
            });
        }

        presPart.Presentation.AppendChild(slideIdList);
        presPart.Presentation.Save();
    }

    private static Shape BuildTextBox(
        string text, int id,
        long x, long y, long cx, long cy,
        int fontSize, bool bold, string colorHex)
    {
        var paragraphs = text.Split('\n').Select(line =>
        {
            var rp = new D.RunProperties
            {
                FontSize = fontSize,
                Language = "zh-CN"
            };
            var run = new D.Run(rp, new D.Text(line.Trim()));
            return new D.Paragraph(run);
        });

        var textBody = new D.TextBody(
            new D.BodyProperties { Wrap = D.TextWrappingValues.Square },
            new D.ListStyle());
        foreach (var p in paragraphs) textBody.Append(p);

        var spPr = new ShapeProperties(
            new D.Transform2D(
                new D.Offset { X = x, Y = y },
                new D.Extents { Cx = cx, Cy = cy }));

        var pg = new D.PresetGeometry();
        pg.Preset = D.ShapeTypeValues.Rectangle;
        pg.Append(new D.AdjustValueList());
        spPr.Append(pg);

        spPr.Append(new D.SolidFill(
            new D.RgbColorModelHex { Val = "FFFFFF" }));
        spPr.Append(new D.Outline(
            new D.NoFill()));

        var shape = new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties
                { Id = (uint)id, Name = $"Text{id}" },
                new NonVisualShapeDrawingProperties()),
            spPr);

        shape.Append(textBody);
        return shape;
    }
}
