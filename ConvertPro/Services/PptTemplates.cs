using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace ConvertPro.Services;

/// <summary>
/// PPT 模板描述。颜色完全驱动渲染，浅色/深色模板共用同一渲染器。
/// </summary>
public record PptTemplate(
    string Id,           // 如 "cyber_blue"
    string Name,         // 如 "赛博蓝"
    string BgHex,        // 背景色（不带#）
    string PanelHex,     // 面板/卡片色
    string AccentHex,    // 强调色（高亮）
    string Accent2Hex,   // 次强调色（渐变/点缀）
    string TextHex,      // 主文字色
    string MutedHex,     // 次要文字色（灰）
    string FontHeading,  // 标题字体
    string FontBody,     // 正文字体
    string Category = "其他", // 风格分类：科技/商务/简约/清新/教育/中国风/答辩/医学/文艺/喜庆
    string Tags = "",        // 搜索关键词（空格分隔）
    string Theme = "dark")   // dark / light
{
    /// <summary>是否匹配搜索词（名称/分类/标签，不区分大小写）</summary>
    public bool Matches(string keyword)
    {
        if (string.IsNullOrWhiteSpace(keyword)) return true;
        var k = keyword.Trim();
        return (Name?.Contains(k, StringComparison.OrdinalIgnoreCase) == true)
            || (Category?.Contains(k, StringComparison.OrdinalIgnoreCase) == true)
            || (Tags?.Contains(k, StringComparison.OrdinalIgnoreCase) == true)
            || (Theme?.Contains(k, StringComparison.OrdinalIgnoreCase) == true)
            || (Theme == "dark" && "深色".Contains(k, StringComparison.OrdinalIgnoreCase))
            || (Theme == "light" && "浅色".Contains(k, StringComparison.OrdinalIgnoreCase));
    }
}

/// <summary>
/// 内置模板库（对标 aiPPT 风格库，30+ 套，覆盖多风格，可搜索）。
/// </summary>
public static class PptTemplates
{
    public static readonly IReadOnlyList<PptTemplate> All = new List<PptTemplate>
    {
        // ========== 科技（深色） ==========
        new PptTemplate("cyber_blue", "赛博蓝",
            "0B1020", "141B30", "38BDF8", "818CF8", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "科技", "科技 数据 产品 现代 蓝", "dark"),
        new PptTemplate("neon_purple", "霓虹紫",
            "120B20", "1D1430", "C084FC", "F472B6", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "科技", "科技 创意 紫", "dark"),
        new PptTemplate("matrix_green", "矩阵绿",
            "0A1410", "122019", "34D399", "A3E635", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "科技", "科技 终端 绿", "dark"),
        new PptTemplate("solar_orange", "烈日橙",
            "1A1208", "2A1E10", "FB923C", "FBBF24", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "科技", "科技 能源 橙", "dark"),
        new PptTemplate("arctic_cyan", "极地青",
            "08121A", "101E2A", "22D3EE", "60A5FA", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "科技", "科技 冰 冷 青", "dark"),
        new PptTemplate("magma_red", "熔岩红",
            "1A0808", "2A1010", "F87171", "FB7185", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "科技", "科技 热烈 红", "dark"),

        // ========== 商务 ==========
        new PptTemplate("executive_navy", "商务深蓝",
            "0F172A", "1E293B", "3B82F6", "60A5FA", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "商务", "商务 汇报 年报 蓝", "dark"),
        new PptTemplate("emerald_pro", "墨绿商务",
            "052E2B", "0A4A45", "10B981", "34D399", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "商务", "商务 稳重 绿", "dark"),
        new PptTemplate("charcoal_pro", "炭灰商务",
            "1A1A1A", "2A2A2A", "9CA3AF", "D1D5DB", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "商务", "商务 极简 灰 高级", "dark"),

        // ========== 简约（浅色） ==========
        new PptTemplate("pure_white", "极简白",
            "FFFFFF", "F8FAFC", "2563EB", "93C5FD", "1F2937", "6B7280",
            "Microsoft YaHei", "Microsoft YaHei", "简约", "简约 白 干净 浅色", "light"),
        new PptTemplate("mist_gray", "薄雾灰",
            "F1F5F9", "E2E8F0", "475569", "94A3B8", "1E293B", "64748B",
            "Microsoft YaHei", "Microsoft YaHei", "简约", "简约 灰 浅色 通用", "light"),
        new PptTemplate("cream_paper", "米白纸",
            "FDFCF7", "F5F1E8", "B45309", "D97706", "292524", "78716C",
            "Microsoft YaHei", "Microsoft YaHei", "简约", "简约 米色 温暖 浅色", "light"),

        // ========== 清新 ==========
        new PptTemplate("mint_green", "薄荷绿",
            "F0FDF4", "DCFCE7", "16A34A", "4ADE80", "14532D", "6B7280",
            "Microsoft YaHei", "Microsoft YaHei", "清新", "清新 绿 春天 浅色", "light"),
        new PptTemplate("sky_blue", "天空蓝",
            "F0F9FF", "E0F2FE", "0284C7", "38BDF8", "0C4A6E", "6B7280",
            "Microsoft YaHei", "Microsoft YaHei", "清新", "清新 蓝 天空 浅色", "light"),
        new PptTemplate("sakura_pink", "樱花粉",
            "FDF2F8", "FCE7F3", "DB2777", "F472B6", "831843", "6B7280",
            "Microsoft YaHei", "Microsoft YaHei", "清新", "清新 粉 樱花 少女 浅色", "light"),
        new PptTemplate("lemon_yellow", "柠檬黄",
            "FEFCE8", "FEF9C3", "CA8A04", "EAB308", "422006", "6B7280",
            "Microsoft YaHei", "Microsoft YaHei", "清新", "清新 黄 明亮 浅色", "light"),

        // ========== 教育 ==========
        new PptTemplate("academy_blue", "学院蓝",
            "F8FAFC", "E2E8F0", "1D4ED8", "3B82F6", "1E293B", "64748B",
            "Microsoft YaHei", "Microsoft YaHei", "教育", "教育 教学 课件 蓝 浅色", "light"),
        new PptTemplate("book_brown", "书卷棕",
            "FBF7F0", "EDE4D3", "92400E", "B45309", "292524", "78716C",
            "Microsoft YaHei", "Microsoft YaHei", "教育", "教育 阅读 棕 浅色", "light"),
        new PptTemplate("blackboard_green", "黑板绿",
            "0A2A1F", "103524", "34D399", "A3E635", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "教育", "教育 课堂 板书 绿", "dark"),

        // ========== 中国风 ==========
        new PptTemplate("cinnabar_red", "朱砂红",
            "FDF5F0", "FBE6D9", "C0392B", "E74C3C", "4A1C1C", "8C6B5C",
            "Microsoft YaHei", "Microsoft YaHei", "中国风", "中国风 传统 红 国风 浅色", "light"),
        new PptTemplate("forbidden_gold", "故宫金",
            "FBF6E9", "F5EBC9", "B8860B", "DAA520", "3D2E0A", "8C7B5A",
            "Microsoft YaHei", "Microsoft YaHei", "中国风", "中国风 皇家 金 浅色", "light"),
        new PptTemplate("celadon_green", "青瓷绿",
            "F0F5EF", "DDE9DC", "5B8C5A", "8BAF89", "243D24", "6B7280",
            "Microsoft YaHei", "Microsoft YaHei", "中国风", "中国风 瓷 雅 浅色", "light"),
        new PptTemplate("ink_black", "墨韵黑",
            "14110F", "241F1A", "C9A961", "E8D5A0", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "中国风", "中国风 水墨 墨 金", "dark"),
        new PptTemplate("blue_white", "青花蓝",
            "0E1A2B", "1A2E45", "4A90D9", "8BB8E8", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "中国风", "中国风 青花 瓷 蓝", "dark"),

        // ========== 答辩 ==========
        new PptTemplate("thesis_blue", "答辩蓝",
            "F8FAFC", "E2E8F0", "1E40AF", "3B82F6", "1E293B", "64748B",
            "Microsoft YaHei", "Microsoft YaHei", "答辩", "答辩 论文 毕业 学术 蓝 浅色", "light"),
        new PptTemplate("academic_gray", "学术灰",
            "FAFAFA", "F0F0F0", "374151", "6B7280", "111827", "6B7280",
            "Microsoft YaHei", "Microsoft YaHei", "答辩", "答辩 学术 严谨 灰 浅色", "light"),
        new PptTemplate("thesis_purple", "学位紫",
            "1A1230", "2A1F45", "8B5CF6", "A78BFA", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "答辩", "答辩 学位 毕业 紫", "dark"),

        // ========== 医学 ==========
        new PptTemplate("clinical_white", "洁净白",
            "F8FAFC", "E0F2FE", "0891B2", "06B6D4", "0C4A6E", "64748B",
            "Microsoft YaHei", "Microsoft YaHei", "医学", "医学 医疗 洁净 青 浅色", "light"),
        new PptTemplate("surgery_green", "手术绿",
            "F0FDF4", "D1FAE5", "059669", "10B981", "064E3B", "6B7280",
            "Microsoft YaHei", "Microsoft YaHei", "医学", "医学 手术 健康 绿 浅色", "light"),
        new PptTemplate("medical_dark", "医疗深蓝",
            "0A1A2A", "102A40", "06B6D4", "67E8F9", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "医学", "医学 医疗 专业 蓝", "dark"),

        // ========== 文艺 ==========
        new PptTemplate("morandi_blue", "莫兰迪蓝",
            "EEF1F5", "DFE4EB", "6B7E9E", "9AAEC4", "2D3A4F", "6B7280",
            "Microsoft YaHei", "Microsoft YaHei", "文艺", "文艺 莫兰迪 雅 浅色", "light"),
        new PptTemplate("vintage_sepia", "复古棕",
            "F5EFE3", "EADFCD", "8B6F47", "B08D57", "3A2E1E", "7A6A55",
            "Microsoft YaHei", "Microsoft YaHei", "文艺", "文艺 复古 棕 浅色", "light"),
        new PptTemplate("retro_film", "复古胶片",
            "1F1A14", "2E2620", "D4A574", "E8C9A0", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "文艺", "文艺 胶片 怀旧 棕", "dark"),

        // ========== 喜庆 ==========
        new PptTemplate("festivity_red", "喜庆红",
            "2A0A0A", "3D1414", "DC2626", "F59E0B", "F8FAFC", "94A3B8",
            "Microsoft YaHei", "Microsoft YaHei", "喜庆", "喜庆 节日 红 金", "dark"),
        new PptTemplate("gold_autumn", "金秋橙",
            "FFF7ED", "FFEDD5", "EA580C", "F59E0B", "431407", "9A3412",
            "Microsoft YaHei", "Microsoft YaHei", "喜庆", "喜庆 秋 收获 橙 浅色", "light"),
    };

    public static PptTemplate Default => All[0];

    public static PptTemplate Get(string id) =>
        All.FirstOrDefault(t => t.Id == id) ?? All[0];
}

// ---------- AI 生成 JSON 反序列化得到的内容模型 ----------

public record PptContent(
    string Title,            // 封面标题
    string Subtitle,         // 封面副标题（可为空）
    List<PptSection> Sections,
    List<string> Summary,    // 总结要点
    string? CoverImagePrompt = null)  // 封面图生图提示词（可选）
{
    public static PptContent FromJson(string json) =>
        JsonSerializer.Deserialize<PptContent>(json,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
        ?? throw new InvalidOperationException("Invalid PPT content JSON");
}

public record PptSection(
    string Title,            // 章节标题
    List<PptSlide> Slides,
    string? ImagePrompt = null);  // 章节分隔页配图提示词（可选）

/// <summary>
/// 单张幻灯片内容。
/// Layout 取值：
///   "text"        纯文字带装饰（默认，回退兼容）
///   "left_image"  左图右文（图占左 45%）
///   "top_image"   上图下文（图占上 45%）
///   "quote"       引言样式（大字居中）
///   "chart"       数据图表（含 Chart 字段）
/// </summary>
public record PptSlide(
    string Title,            // 页面标题
    List<string> Bullets,    // 要点列表
    string? Note = null,            // 可选备注（渲染时忽略）
    string? ImagePrompt = null,     // 文生图提示词（可选）
    string Layout = "text",         // 布局类型
    string? Icon = null,            // emoji 或单字符图标（可选，置于标题前）
    PptChart? Chart = null);        // 图表数据（仅 Layout="chart" 时使用）

/// <summary>
/// 图表数据。Type: bar / pie / line。
/// Labels 与 Series 维度需一致；多系列时 Series.Count > 1。
/// </summary>
public record PptChart(
    string Type,             // bar / pie / line
    string Title,            // 图表标题
    List<string> Labels,     // X 轴标签或饼图分块标签
    List<PptChartSeries> Series);

public record PptChartSeries(
    string Name,             // 系列名
    List<double> Values);    // 数值，与 Labels 一一对应

// ---------- 渲染器 ----------

/// <summary>
/// PPT 渲染器：把 <see cref="PptContent"/> 渲染为 .pptx。
/// 幻灯片顺序：封面 → 目录 → (章节分隔 + 章节内容页) ×N → 总结。
/// 支持：多布局变体（left_image/top_image/quote/chart/text）、AI 文生图嵌入、
/// 简化柱状图/饼图、封面与分隔页装饰强化、Icon 点缀。
/// </summary>
public static class PptRenderer
{
    private const int SlideW = 12192000;
    private const int SlideH = 6858000;

    /// <param name="images">key = imagePrompt，value = 图片字节。可选。</param>
    public static void Build(string pptxPath, PptTemplate tpl, PptContent content,
        Dictionary<string, byte[]>? images = null)
    {
        using var pptDoc = PresentationDocument.Create(pptxPath,
            PresentationDocumentType.Presentation);

        // --- 1. Presentation part ---
        var presPart = pptDoc.AddPresentationPart();
        presPart.Presentation = new Presentation();

        presPart.Presentation.AppendChild(new SlideSize
        {
            Cx = SlideW, Cy = SlideH,
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
            Id = 2147483648U,
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
                        { Id = 1U, Name = "Placeholder" },
                        new NonVisualShapeDrawingProperties()),
                    new ShapeProperties(),
                    new TextBody(
                        new D.BodyProperties(),
                        new D.ListStyle())))));
        layoutPart.SlideLayout.Save();

        var layoutId = new SlideLayoutId
        {
            Id = 2147483649U,
            RelationshipId = masterPart.GetIdOfPart(layoutPart)
        };
        master.SlideLayoutIdList!.Append(layoutId);
        masterPart.SlideMaster.Save();

        // --- 4. Slides ---
        var slideIdList = new SlideIdList();
        uint slideId = 256;

        // 封面
        var coverImg = TryGetImage(images, content.CoverImagePrompt);
        AddSlide(presPart, layoutPart, slideIdList, ref slideId,
            BuildCover(tpl, content, coverImg));

        // 目录
        AddSlide(presPart, layoutPart, slideIdList, ref slideId,
            BuildToc(tpl, content));

        // 章节：分隔页 + 内容页
        for (int si = 0; si < content.Sections.Count; si++)
        {
            var section = content.Sections[si];
            var secImg = TryGetImage(images, section.ImagePrompt);
            AddSlide(presPart, layoutPart, slideIdList, ref slideId,
                BuildSectionDivider(tpl, si + 1, section, secImg));

            foreach (var slide in section.Slides)
            {
                var slideImg = TryGetImage(images, slide.ImagePrompt);
                AddSlide(presPart, layoutPart, slideIdList, ref slideId,
                    BuildContent(tpl, slide, slideImg));
            }
        }

        // 总结
        AddSlide(presPart, layoutPart, slideIdList, ref slideId,
            BuildSummary(tpl, content));

        presPart.Presentation.AppendChild(slideIdList);
        presPart.Presentation.Save();
    }

    private static byte[]? TryGetImage(Dictionary<string, byte[]>? images, string? prompt)
    {
        if (images == null || string.IsNullOrWhiteSpace(prompt)) return null;
        return images.TryGetValue(prompt, out var b) ? b : null;
    }

    // ===== 幻灯片构建 =====

    private static List<object> BuildCover(PptTemplate tpl, PptContent content,
        byte[]? coverImage)
    {
        uint id = 1;
        var list = new List<object>();

        // 渐变背景（双色调）
        list.Add(GradientRect(id++, "bg", 0, 0, SlideW, SlideH, tpl.BgHex, tpl.PanelHex));

        // 右下角装饰大圆（accent 色，半透明模拟）
        list.Add(Circle(id++, "deco1", SlideW - 1800000, SlideH - 1800000, 2200000, 2200000,
            tpl.AccentHex, 60));

        // 左上角装饰方块
        list.Add(Rect(id++, "deco2", 0, 0, 80000, 600000, tpl.AccentHex));

        bool hasCover = coverImage != null && coverImage.Length > 0;
        long titleX = hasCover ? 600000 : 800000;
        long titleW = hasCover ? 6200000 : 10500000;

        // 标题
        var titleParas = new List<D.Paragraph>
        {
            Para(content.Title ?? "", 4400, tpl.TextHex, true)
        };
        list.Add(TextBox(id++, "title", titleX, 2200000, titleW, 1500000, titleParas));

        // 副标题
        if (!string.IsNullOrWhiteSpace(content.Subtitle))
        {
            var subParas = new List<D.Paragraph>
            {
                Para(content.Subtitle, 1600, tpl.MutedHex)
            };
            list.Add(TextBox(id++, "subtitle", titleX, 3700000, titleW, 800000, subParas));
        }

        // 底部 accent 横线
        list.Add(Rect(id++, "bar", titleX, 4500000, 600000, 30000, tpl.AccentHex));

        // 封面图（右侧）
        if (hasCover)
        {
            // 图占右侧 40%
            long imgX = 7200000, imgY = 900000;
            long imgW = 4500000, imgH = 5058000;
            list.Add(new PicturePlaceholder(id++, "coverImage", imgX, imgY, imgW, imgH,
                coverImage!, tpl));
        }

        return list;
    }

    private static List<object> BuildToc(PptTemplate tpl, PptContent content)
    {
        uint id = 1;
        var list = new List<object>();
        list.Add(Rect(id++, "bg", 0, 0, SlideW, SlideH, tpl.BgHex));

        // 装饰：左上角竖条
        list.Add(Rect(id++, "bar", 600000, 450000, 80000, 1500000, tpl.AccentHex));

        var titleParas = new List<D.Paragraph>
        {
            Para("目录", 3200, tpl.TextHex, true)
        };
        list.Add(TextBox(id++, "title", 800000, 400000, 10500000, 900000, titleParas));

        var itemParas = new List<D.Paragraph>();
        for (int i = 0; i < content.Sections.Count; i++)
        {
            var p = new D.Paragraph();
            p.Append(new D.Run(RunProps(1800, tpl.AccentHex, true),
                new D.Text($"{i + 1:00}  ")));
            p.Append(new D.Run(RunProps(1800, tpl.TextHex),
                new D.Text(content.Sections[i].Title ?? "")));
            itemParas.Add(p);
        }
        list.Add(TextBox(id++, "items", 1000000, 1500000, 10000000, 4500000, itemParas));
        return list;
    }

    private static List<object> BuildSectionDivider(PptTemplate tpl, int index,
        PptSection section, byte[]? sectionImage)
    {
        uint id = 1;
        var list = new List<object>();
        list.Add(GradientRect(id++, "bg", 0, 0, SlideW, SlideH, tpl.BgHex, tpl.PanelHex));

        // 大数字"01"在左侧（超大字号，accent 色，透明度模拟）
        var numParas = new List<D.Paragraph>
        {
            Para($"{index:00}", 8000, tpl.AccentHex, true,
                D.TextAlignmentTypeValues.Center)
        };
        list.Add(TextBox(id++, "bignum", 200000, 800000, 4000000, 5300000, numParas));

        // 章节标题在右侧
        var secLabelParas = new List<D.Paragraph>
        {
            Para($"CHAPTER {index}", 1200, tpl.MutedHex, false)
        };
        list.Add(TextBox(id++, "seclabel", 5500000, 2400000, 6000000, 500000, secLabelParas));

        var titleParas = new List<D.Paragraph>
        {
            Para(section.Title ?? "", 3600, tpl.TextHex, true)
        };
        list.Add(TextBox(id++, "title", 5500000, 2900000, 6000000, 1200000, titleParas));

        // 装饰横线
        list.Add(Rect(id++, "line", 5500000, 4200000, 2000000, 60000, tpl.AccentHex));

        // 章节配图（右侧下方）
        if (sectionImage != null && sectionImage.Length > 0)
        {
            list.Add(new PicturePlaceholder(id++, "secImg", 8800000, 4400000,
                2800000, 2000000, sectionImage, tpl));
        }

        return list;
    }

    private static List<object> BuildContent(PptTemplate tpl, PptSlide slide,
        byte[]? image)
    {
        // 按 Layout 分发
        return (slide.Layout ?? "text").ToLowerInvariant() switch
        {
            "left_image" when image != null => BuildContent_LeftImage(tpl, slide, image),
            "top_image" when image != null => BuildContent_TopImage(tpl, slide, image),
            "quote" => BuildContent_Quote(tpl, slide),
            "chart" when slide.Chart != null => BuildContent_Chart(tpl, slide),
            _ => BuildContent_Text(tpl, slide, image)
        };
    }

    // 默认文本布局（带 icon 装饰）
    private static List<object> BuildContent_Text(PptTemplate tpl, PptSlide slide,
        byte[]? image)
    {
        uint id = 1;
        var list = new List<object>();
        list.Add(Rect(id++, "bg", 0, 0, SlideW, SlideH, tpl.BgHex));

        // 标题栏 accent 色左竖条
        list.Add(Rect(id++, "bar", 500000, 380000, 80000, 720000, tpl.AccentHex));

        // 标题（含 icon 前缀）
        var titleText = !string.IsNullOrWhiteSpace(slide.Icon)
            ? $"{slide.Icon}  {slide.Title}"
            : slide.Title ?? "";
        var titleParas = new List<D.Paragraph>
        {
            Para(titleText, 2600, tpl.TextHex, true)
        };
        list.Add(TextBox(id++, "title", 700000, 360000, 10800000, 800000, titleParas));

        // bullets
        var bulletParas = new List<D.Paragraph>();
        foreach (var b in slide.Bullets)
        {
            if (string.IsNullOrEmpty(b)) continue;
            bulletParas.Add(BulletPara(b, 1600, tpl.TextHex, tpl.AccentHex));
        }

        long bulletsY = 1300000;
        long bulletsH = 5000000;
        long bulletsW = 10800000;

        // 若有图但不是 left/top 布局，把图放右下角小图
        if (image != null && image.Length > 0)
        {
            // 右下角装饰图（占右下 30%）
            list.Add(new PicturePlaceholder(id++, "img", 8200000, 3400000,
                3300000, 2700000, image, tpl));
            bulletsW = 7400000;
        }

        if (bulletParas.Count > 0)
        {
            list.Add(TextBox(id++, "bullets", 700000, bulletsY, bulletsW, bulletsH, bulletParas));
        }
        return list;
    }

    // 左图右文布局
    private static List<object> BuildContent_LeftImage(PptTemplate tpl, PptSlide slide,
        byte[] image)
    {
        uint id = 1;
        var list = new List<object>();
        list.Add(Rect(id++, "bg", 0, 0, SlideW, SlideH, tpl.BgHex));

        // 左侧大图（占左 45%）
        long imgX = 0, imgY = 0, imgW = 5400000, imgH = SlideH;
        list.Add(new PicturePlaceholder(id++, "img", imgX, imgY, imgW, imgH, image, tpl));

        // 右侧内容区
        long contentX = 5800000;
        long contentW = 6100000;

        // 标题栏 accent 竖条
        list.Add(Rect(id++, "bar", contentX - 100000, 380000, 80000, 720000, tpl.AccentHex));

        var titleText = !string.IsNullOrWhiteSpace(slide.Icon)
            ? $"{slide.Icon}  {slide.Title}"
            : slide.Title ?? "";
        var titleParas = new List<D.Paragraph>
        {
            Para(titleText, 2400, tpl.TextHex, true)
        };
        list.Add(TextBox(id++, "title", contentX, 400000, contentW, 900000, titleParas));

        var bulletParas = new List<D.Paragraph>();
        foreach (var b in slide.Bullets)
        {
            if (string.IsNullOrEmpty(b)) continue;
            bulletParas.Add(BulletPara(b, 1500, tpl.TextHex, tpl.AccentHex));
        }
        if (bulletParas.Count > 0)
        {
            list.Add(TextBox(id++, "bullets", contentX, 1500000, contentW, 5000000, bulletParas));
        }
        return list;
    }

    // 上图下文布局
    private static List<object> BuildContent_TopImage(PptTemplate tpl, PptSlide slide,
        byte[] image)
    {
        uint id = 1;
        var list = new List<object>();
        list.Add(Rect(id++, "bg", 0, 0, SlideW, SlideH, tpl.BgHex));

        // 顶部图（占上 45%）
        long imgX = 0, imgY = 0, imgW = SlideW, imgH = 3000000;
        list.Add(new PicturePlaceholder(id++, "img", imgX, imgY, imgW, imgH, image, tpl));

        // 标题
        var titleText = !string.IsNullOrWhiteSpace(slide.Icon)
            ? $"{slide.Icon}  {slide.Title}"
            : slide.Title ?? "";
        var titleParas = new List<D.Paragraph>
        {
            Para(titleText, 2200, tpl.TextHex, true)
        };
        list.Add(TextBox(id++, "title", 700000, 3200000, 10800000, 700000, titleParas));

        // accent 横线分隔
        list.Add(Rect(id++, "line", 700000, 3950000, 1500000, 40000, tpl.AccentHex));

        var bulletParas = new List<D.Paragraph>();
        foreach (var b in slide.Bullets)
        {
            if (string.IsNullOrEmpty(b)) continue;
            bulletParas.Add(BulletPara(b, 1400, tpl.TextHex, tpl.AccentHex));
        }
        if (bulletParas.Count > 0)
        {
            list.Add(TextBox(id++, "bullets", 700000, 4100000, 10800000, 2500000, bulletParas));
        }
        return list;
    }

    // 引言布局（大字居中，无 bullets）
    private static List<object> BuildContent_Quote(PptTemplate tpl, PptSlide slide)
    {
        uint id = 1;
        var list = new List<object>();
        list.Add(GradientRect(id++, "bg", 0, 0, SlideW, SlideH, tpl.BgHex, tpl.PanelHex));

        // 装饰：左上大引号字符
        var quoteMark = new List<D.Paragraph>
        {
            Para("\u201C", 9000, tpl.AccentHex, true, D.TextAlignmentTypeValues.Left)
        };
        list.Add(TextBox(id++, "mark", 600000, 600000, 3000000, 3000000, quoteMark));

        // 主文字
        var quoteText = slide.Bullets.FirstOrDefault() ?? slide.Title ?? "";
        var quoteParas = new List<D.Paragraph>
        {
            Para(quoteText, 3000, tpl.TextHex, false,
                D.TextAlignmentTypeValues.Center)
        };
        list.Add(TextBox(id++, "quote", 1500000, 2400000, 9000000, 2200000, quoteParas));

        // 署名（slide.Title 作为署名）
        if (!string.IsNullOrWhiteSpace(slide.Title) && quoteText != slide.Title)
        {
            var sigParas = new List<D.Paragraph>
            {
                Para($"— {slide.Title}", 1400, tpl.MutedHex, false,
                    D.TextAlignmentTypeValues.Center)
            };
            list.Add(TextBox(id++, "sig", 1500000, 4700000, 9000000, 600000, sigParas));
        }
        return list;
    }

    // 图表布局（柱状图/饼图，用矩形模拟，下方 bullets）
    private static List<object> BuildContent_Chart(PptTemplate tpl, PptSlide slide)
    {
        uint id = 1;
        var list = new List<object>();
        list.Add(Rect(id++, "bg", 0, 0, SlideW, SlideH, tpl.BgHex));

        // 标题
        var titleText = !string.IsNullOrWhiteSpace(slide.Icon)
            ? $"{slide.Icon}  {slide.Title}"
            : slide.Title ?? "";
        var titleParas = new List<D.Paragraph>
        {
            Para(titleText, 2400, tpl.TextHex, true)
        };
        list.Add(TextBox(id++, "title", 700000, 360000, 10800000, 700000, titleParas));

        // 图表区
        var chart = slide.Chart!;
        var chartShapes = BuildSimpleChart(tpl, chart, 700000, 1200000, 10800000, 3500000);
        foreach (var s in chartShapes) list.Add(s);

        // 下方 bullets
        var bulletParas = new List<D.Paragraph>();
        foreach (var b in slide.Bullets)
        {
            if (string.IsNullOrEmpty(b)) continue;
            bulletParas.Add(BulletPara(b, 1300, tpl.TextHex, tpl.AccentHex));
        }
        if (bulletParas.Count > 0)
        {
            list.Add(TextBox(id++, "bullets", 700000, 4900000, 10800000, 1700000, bulletParas));
        }
        return list;
    }

    /// <summary>
    /// 用矩形模拟柱状图（垂直 bar）或饼图（水平 stacked bar）。
    /// 单系列图表；多系列取第一个。
    /// </summary>
    private static List<object> BuildSimpleChart(PptTemplate tpl, PptChart chart,
        long x, long y, long w, long h)
    {
        var result = new List<object>();
        uint id = 1;

        // 图表标题
        if (!string.IsNullOrWhiteSpace(chart.Title))
        {
            var tParas = new List<D.Paragraph>
            {
                Para(chart.Title, 1600, tpl.TextHex, true)
            };
            result.Add(TextBox(id++, "ctitle", x, y, w, 400000, tParas));
            y += 450000;
            h -= 450000;
        }

        if (chart.Series == null || chart.Series.Count == 0 ||
            chart.Labels == null || chart.Labels.Count == 0)
        {
            var empty = new List<D.Paragraph>
            {
                Para("(无数据)", 1400, tpl.MutedHex)
            };
            result.Add(TextBox(id++, "empty", x, y, w, h, empty));
            return result;
        }

        var series = chart.Series[0];
        var values = series.Values ?? new List<double>();
        var labels = chart.Labels;
        int n = Math.Min(values.Count, labels.Count);
        if (n == 0)
        {
            var empty = new List<D.Paragraph> { Para("(无数据)", 1400, tpl.MutedHex) };
            result.Add(TextBox(id++, "empty", x, y, w, h, empty));
            return result;
        }

        double maxVal = 0;
        for (int i = 0; i < n; i++) maxVal = Math.Max(maxVal, Math.Abs(values[i]));
        if (maxVal <= 0) maxVal = 1;

        var type = (chart.Type ?? "bar").ToLowerInvariant();

        if (type == "pie")
        {
            // 饼图用水平 stacked bar 模拟
            double sum = 0;
            for (int i = 0; i < n; i++) sum += Math.Abs(values[i]);
            if (sum <= 0) sum = 1;

            long barY = y + h / 2 - 300000;
            long barH = 600000;
            long curX = x;
            for (int i = 0; i < n; i++)
            {
                long segW = (long)(Math.Abs(values[i]) / sum * w);
                if (segW < 30000) segW = 30000;
                if (i == n - 1) segW = x + w - curX; // 最后一段补齐

                var color = ChartColor(i, tpl);
                result.Add(Rect(id++, $"seg{i}", curX, barY, segW, barH, color));
                curX += segW;
            }

            // 图例
            var legendParas = new List<D.Paragraph>();
            for (int i = 0; i < n; i++)
            {
                var p = new D.Paragraph();
                p.Append(new D.Run(RunProps(1100, ChartColor(i, tpl), true),
                    new D.Text($"■ ")));
                p.Append(new D.Run(RunProps(1100, tpl.TextHex),
                    new D.Text($"{labels[i]}: {values[i]}  ")));
                legendParas.Add(p);
            }
            result.Add(TextBox(id++, "legend", x, y + h - 600000, w, 600000, legendParas));
        }
        else
        {
            // 柱状图（bar / line 都用柱状图模拟）
            long gap = 200000;
            long availW = w - gap * (n + 1);
            long barW = availW / n;
            if (barW < 100000) barW = 100000;

            long baseY = y + h - 600000; // 留底部 label
            long chartH = h - 600000;

            for (int i = 0; i < n; i++)
            {
                var val = values[i];
                long barH = (long)(Math.Abs(val) / maxVal * chartH);
                if (barH < 60000) barH = 60000;
                long bx = x + gap + i * (barW + gap);
                long by = baseY - barH;
                var color = ChartColor(i, tpl);

                // 柱子
                result.Add(Rect(id++, $"bar{i}", bx, by, barW, barH, color));

                // 数值
                var valParas = new List<D.Paragraph>
                {
                    Para(val.ToString("0.#"), 1000, tpl.TextHex, true,
                        D.TextAlignmentTypeValues.Center)
                };
                result.Add(TextBox(id++, $"val{i}", bx - 50000, by - 350000,
                    barW + 100000, 300000, valParas));

                // X 轴 label
                var lblParas = new List<D.Paragraph>
                {
                    Para(labels[i], 1000, tpl.MutedHex, false,
                        D.TextAlignmentTypeValues.Center)
                };
                result.Add(TextBox(id++, $"lbl{i}", bx - 50000, baseY + 50000,
                    barW + 100000, 400000, lblParas));
            }

            // 底线
            result.Add(Rect(id++, "axis", x, baseY, w, 20000, tpl.MutedHex));
        }

        return result;
    }

    private static string ChartColor(int idx, PptTemplate tpl)
    {
        var colors = new[] { tpl.AccentHex, tpl.Accent2Hex, tpl.TextHex, tpl.MutedHex, "F59E0B", "10B981" };
        return colors[idx % colors.Length];
    }

    private static List<object> BuildSummary(PptTemplate tpl, PptContent content)
    {
        uint id = 1;
        var list = new List<object>();
        list.Add(GradientRect(id++, "bg", 0, 0, SlideW, SlideH, tpl.BgHex, tpl.PanelHex));

        // 装饰
        list.Add(Rect(id++, "bar", 600000, 450000, 80000, 1500000, tpl.AccentHex));

        var titleParas = new List<D.Paragraph>
        {
            Para("总结", 3200, tpl.TextHex, true)
        };
        list.Add(TextBox(id++, "title", 800000, 400000, 10500000, 900000, titleParas));

        var bulletParas = new List<D.Paragraph>();
        foreach (var s in content.Summary)
        {
            if (string.IsNullOrEmpty(s)) continue;
            bulletParas.Add(BulletPara(s, 1600, tpl.TextHex, tpl.AccentHex));
        }
        if (bulletParas.Count > 0)
        {
            list.Add(TextBox(id++, "bullets", 800000, 1500000, 10500000, 4500000, bulletParas));
        }
        return list;
    }

    // ===== 基础设施 =====

    private static void AddSlide(PresentationPart presPart, SlideLayoutPart layoutPart,
        SlideIdList slideIdList, ref uint slideId, List<object> shapes,
        Dictionary<(string, uint), byte[]>? imagePayloads = null)
    {
        var slidePart = presPart.AddNewPart<SlidePart>();
        var shapeTree = new ShapeTree();

        // 把图片占位符转成真实 Picture
        foreach (var elem in shapes)
        {
            if (elem is PicturePlaceholder ph)
            {
                var pic = BuildPicture(slidePart, ph.Id, ph.Name, ph.X, ph.Y, ph.Cx, ph.Cy,
                    ph.ImageBytes, ph.Tpl);
                if (pic != null) shapeTree.Append(pic);
                // 失败则跳过
            }
            else if (elem is OpenXmlElement oxe)
            {
                shapeTree.Append(oxe);
            }
        }

        slidePart.Slide = new Slide(new CommonSlideData(shapeTree));
        slidePart.AddPart(layoutPart);
        slidePart.Slide.Save();
        slideIdList.AppendChild(new SlideId
        {
            Id = slideId++,
            RelationshipId = presPart.GetIdOfPart(slidePart)
        });
    }

    /// <summary>构造 Picture 元素（创建 ImagePart 并嵌入字节）。</summary>
    private static Picture? BuildPicture(SlidePart slidePart, uint id, string name,
        long x, long y, long cx, long cy, byte[] imageBytes, PptTemplate tpl)
    {
        if (imageBytes == null || imageBytes.Length == 0) return null;

        var imgType = DetectImageType(imageBytes);
        var imagePart = slidePart.AddImagePart(imgType);
        using (var ms = new MemoryStream(imageBytes))
        {
            imagePart.FeedData(ms);
        }
        var relId = slidePart.GetIdOfPart(imagePart);

        var pic = new Picture(
            new NonVisualPictureProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualPictureDrawingProperties()),
            new D.BlipFill(
                new D.Blip { Embed = relId },
                new D.Stretch(new D.FillRectangle())),
            new ShapeProperties(
                new D.Transform2D(
                    new D.Offset { X = x, Y = y },
                    new D.Extents { Cx = cx, Cy = cy }),
                new D.PresetGeometry
                {
                    Preset = D.ShapeTypeValues.Rectangle,
                    AdjustValueList = new D.AdjustValueList()
                }));
        return pic;
    }

    private static PartTypeInfo DetectImageType(byte[] data)
    {
        if (data.Length >= 2 && data[0] == 0xFF && data[1] == 0xD8)
            return ImagePartType.Jpeg;
        if (data.Length >= 4 && data[0] == 0x89 && data[1] == 0x50 &&
            data[2] == 0x4E && data[3] == 0x47)
            return ImagePartType.Png;
        return ImagePartType.Png; // 默认按 PNG（大多数文生图返回 PNG）
    }

    /// <summary>纯色矩形</summary>
    private static Shape Rect(uint id, string name, long x, long y, long cx, long cy, string fillHex)
    {
        var spPr = new ShapeProperties(
            new D.Transform2D(
                new D.Offset { X = x, Y = y },
                new D.Extents { Cx = cx, Cy = cy }));
        var pg = new D.PresetGeometry();
        pg.Preset = D.ShapeTypeValues.Rectangle;
        pg.Append(new D.AdjustValueList());
        spPr.Append(pg);
        spPr.Append(new D.SolidFill(new D.RgbColorModelHex { Val = fillHex }));
        spPr.Append(new D.Outline(new D.NoFill()));

        var tb = new D.TextBody(new D.BodyProperties(), new D.ListStyle());
        var shape = new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualShapeDrawingProperties()),
            spPr);
        shape.Append(tb);
        return shape;
    }

    /// <summary>渐变背景矩形（双色，对角线）</summary>
    private static Shape GradientRect(uint id, string name, long x, long y, long cx, long cy,
        string hex1, string hex2)
    {
        var spPr = new ShapeProperties(
            new D.Transform2D(
                new D.Offset { X = x, Y = y },
                new D.Extents { Cx = cx, Cy = cy }));
        var pg = new D.PresetGeometry();
        pg.Preset = D.ShapeTypeValues.Rectangle;
        pg.Append(new D.AdjustValueList());
        spPr.Append(pg);

        // 渐变填充：从 hex1 到 hex2
        var gradFill = new D.GradientFill();

        var stop1 = new D.GradientStop { Position = 0 };
        stop1.Append(new D.RgbColorModelHex { Val = hex1 });
        var stop2 = new D.GradientStop { Position = 100000 };
        stop2.Append(new D.RgbColorModelHex { Val = hex2 });

        var stopList = new D.GradientStopList();
        stopList.Append(stop1);
        stopList.Append(stop2);
        gradFill.Append(stopList);
        gradFill.Append(new D.LinearGradientFill { Angle = 4500000, Scaled = true });
        spPr.Append(gradFill);
        spPr.Append(new D.Outline(new D.NoFill()));

        var tb = new D.TextBody(new D.BodyProperties(), new D.ListStyle());
        var shape = new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualShapeDrawingProperties()),
            spPr);
        shape.Append(tb);
        return shape;
    }

    /// <summary>圆形装饰（accent 色纯色椭圆）</summary>
    private static Shape Circle(uint id, string name, long x, long y, long cx, long cy,
        string fillHex, int alphaPercent = 100)
    {
        var spPr = new ShapeProperties(
            new D.Transform2D(
                new D.Offset { X = x, Y = y },
                new D.Extents { Cx = cx, Cy = cy }));
        var pg = new D.PresetGeometry();
        pg.Preset = D.ShapeTypeValues.Ellipse;
        pg.Append(new D.AdjustValueList());
        spPr.Append(pg);

        spPr.Append(new D.SolidFill(new D.RgbColorModelHex { Val = fillHex }));
        spPr.Append(new D.Outline(new D.NoFill()));

        var tb = new D.TextBody(new D.BodyProperties(), new D.ListStyle());
        var shape = new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualShapeDrawingProperties()),
            spPr);
        shape.Append(tb);
        return shape;
    }

    private static Shape TextBox(uint id, string name, long x, long y, long cx, long cy,
        IEnumerable<D.Paragraph> paragraphs)
    {
        var spPr = new ShapeProperties(
            new D.Transform2D(
                new D.Offset { X = x, Y = y },
                new D.Extents { Cx = cx, Cy = cy }));
        var pg = new D.PresetGeometry();
        pg.Preset = D.ShapeTypeValues.Rectangle;
        pg.Append(new D.AdjustValueList());
        spPr.Append(pg);
        spPr.Append(new D.NoFill());
        spPr.Append(new D.Outline(new D.NoFill()));

        var bp = new D.BodyProperties { Wrap = D.TextWrappingValues.Square };
        var tb = new D.TextBody(bp, new D.ListStyle());
        foreach (var p in paragraphs) tb.Append(p);

        var shape = new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualShapeDrawingProperties()),
            spPr);
        shape.Append(tb);
        return shape;
    }

    // ===== 文本/段落辅助 =====

    private static D.RunProperties RunProps(int size, string colorHex, bool bold = false)
    {
        var rp = new D.RunProperties
        {
            FontSize = size,
            Language = "zh-CN",
            Bold = bold
        };
        rp.Append(new D.SolidFill(new D.RgbColorModelHex { Val = colorHex }));
        rp.Append(new D.LatinFont { Typeface = "Microsoft YaHei" });
        rp.Append(new D.EastAsianFont { Typeface = "Microsoft YaHei" });
        return rp;
    }

    private static D.Paragraph Para(string text, int size, string colorHex, bool bold = false,
        D.TextAlignmentTypeValues? align = null)
    {
        var p = new D.Paragraph();
        if (align.HasValue)
        {
            p.Append(new D.ParagraphProperties { Alignment = align.Value });
        }
        p.Append(new D.Run(RunProps(size, colorHex, bold), new D.Text(text)));
        return p;
    }

    private static D.Paragraph BulletPara(string text, int size, string textColor, string accentHex)
    {
        var p = new D.Paragraph();
        p.Append(new D.Run(RunProps(size, accentHex), new D.Text("▪  ")));
        p.Append(new D.Run(RunProps(size, textColor), new D.Text(text)));
        return p;
    }
}

/// <summary>
/// 图片占位符：在 BuildXxx 阶段作为 List 元素占位，
/// AddSlide 时真正构造 Picture（因 SlidePart 创建时机晚于 shape 构造）。
/// 不继承 OpenXmlElement，作为普通数据容器。
/// </summary>
internal sealed class PicturePlaceholder
{
    public uint Id { get; }
    public string Name { get; }
    public long X { get; }
    public long Y { get; }
    public long Cx { get; }
    public long Cy { get; }
    public byte[] ImageBytes { get; }
    public PptTemplate Tpl { get; }

    public PicturePlaceholder(uint id, string name, long x, long y, long cx, long cy,
        byte[] imageBytes, PptTemplate tpl)
    {
        Id = id; Name = name; X = x; Y = y; Cx = cx; Cy = cy;
        ImageBytes = imageBytes; Tpl = tpl;
    }
}
