using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using UglyToad.PdfPig;
using D = DocumentFormat.OpenXml.Drawing;

namespace ConvertPro.Services;

/// <summary>
/// PDF → PPT 转换服务。
/// 使用 PdfPig 提取 PDF 每页文本，OpenXml 创建 .pptx 幻灯片。
/// 每页 PDF 文字转为一张幻灯片（含标题文本框 + 内容文本框）。
/// </summary>
public class PdfToPptService : IConversionService
{
    public string ConversionType => "pdf2ppt";
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
                var pptxPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(file) + ".pptx");
                BuildPptx(file, pptxPath);
                results.Add(new ConversionResult(true, pptxPath,
                    OutputSize: new FileInfo(pptxPath).Length));
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

    private void BuildPptx(string pdfPath, string pptxPath)
    {
        // 提取 PDF 每页文本
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

        // --- 1. Presentation part ---
        var presPart = pptDoc.AddPresentationPart();
        presPart.Presentation = new Presentation();

        // Slide size: 16:9
        presPart.Presentation.AppendChild(new SlideSize
        {
            Cx = 12192000, Cy = 6858000,
            Type = SlideSizeValues.Screen16x9
        });

        // --- 2. SlideMaster (required!) ---
        var masterPart = presPart.AddNewPart<SlideMasterPart>();
        var master = new SlideMaster(
            new CommonSlideData(new ShapeTree()),
            new SlideLayoutIdList());
        masterPart.SlideMaster = master;

        var masterId = new SlideMasterId
        {
            Id = 2147483648, // min SlideId
            RelationshipId = presPart.GetIdOfPart(masterPart)
        };
        presPart.Presentation.AppendChild(new SlideMasterIdList(masterId));

        // --- 3. SlideLayout (required!) ---
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

        // Add layout to master
        var layoutId = new SlideLayoutId
        {
            Id = 2147483649,
            RelationshipId = masterPart.GetIdOfPart(layoutPart)
        };
        master.SlideLayoutIdList!.Append(layoutId);
        masterPart.SlideMaster.Save();

        // --- 4. Slides ---
        var slideIdList = new SlideIdList();
        uint slideId = 256;

        for (int pi = 0; pi < pages.Count; pi++)
        {
            var content = pages[pi];
            if (content.Length > 2000) content = content[..2000];

            var slidePart = presPart.AddNewPart<SlidePart>();

            // Build shape tree
            var shapeTree = new ShapeTree();

            // Title shape
            var titleShape = BuildTextBox(
                $"第 {pi + 1} 页", 1,
                500000, 100000, 11000000, 600000,
                2400, true, "4f46e5");

            // Content shape (first few lines)
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

            // Reference layout
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

        // Set shape fill and outline
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
