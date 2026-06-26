using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertPro.Services;

public class ConversionManager
{
    private readonly Dictionary<string, IConversionService> _services = new();
    private readonly WordComConverter? _wordCom;
    private readonly LibreOfficeConverter? _loConverter;
    private readonly PythonPdf2DocxConverter? _pythonPdf2Docx;

    public ConversionManager()
    {
        _wordCom = new WordComConverter();
        _loConverter = new LibreOfficeConverter();
        _pythonPdf2Docx = new PythonPdf2DocxConverter();

        // PNG→ICO：纯 C#
        Register(new PngToIcoService());

        // PDF→Word：Python pdf2docx 优先（已安装 Python）
        Register(new MultiEnginePdfToWordService(_pythonPdf2Docx, _wordCom, _loConverter, new PdfToWordService()));

        // 其他文档转换
        Register(new WordComOrLibreOfficeService("word2pdf", "pdf", _wordCom, _loConverter, new WordToPdfService()));
        Register(new WordComOrLibreOfficeService("pdf2excel", "xlsx", null, _loConverter, new PdfToExcelService()));
        Register(new WordComOrLibreOfficeService("excel2pdf", "pdf", null, _loConverter, new ExcelToPdfService()));
        // pdf2ppt：AI 生成是首选路径（AI 不可用时内部回退到文本→幻灯片），
        // 不再走 LibreOffice（LibreOffice 的 PDF→PPT 是图元堆砌，效果差，正是要解决的问题）
        Register(new PdfToPptService());

        // ===== 互相转换矩阵补齐 =====
        // PPT→PDF：LibreOffice 直接转换，保真度好
        Register(new WordComOrLibreOfficeService("ppt2pdf", "pdf", null, _loConverter, null));
        // Word→PPT：复用 PdfToPptService 的 AI 大纲 + 文生图 + PptRenderer 管线
        Register(new WordToPptService());
        // Excel→Word：ClosedXML 读 + OpenXml 写
        Register(new ExcelToWordService());
        // Word→Excel：OpenXml 读 + ClosedXML 写
        Register(new WordToExcelService());
    }

    public void Register(IConversionService service)
    {
        _services[service.ConversionType] = service;
    }

    public async Task<List<ConversionResult>> ConvertAsync(
        string conversionType, List<string> inputFiles, string outputDir,
        string? options, IProgress<ConversionProgress> progress,
        CancellationToken ct)
    {
        if (!_services.TryGetValue(conversionType, out var service))
            return new List<ConversionResult>
            {
                new(false, "", $"不支持的转换类型: {conversionType}")
            };

        // 过滤掉空路径，避免下游服务崩溃
        var validFiles = inputFiles.Where(f => !string.IsNullOrWhiteSpace(f)).ToList();
        if (validFiles.Count == 0)
            return new List<ConversionResult>
            {
                new(false, "", "未提供有效的文件路径，请通过文件对话框选择文件。")
            };

        return await service.ConvertAsync(validFiles, outputDir,
            options, progress, ct);
    }

    public Task<string?> TryConvertWithWordComAsync(
        string inputPath, string outputDir, string conversionType)
    {
        if (_wordCom == null || !_wordCom.IsAvailable)
            return Task.FromResult<string?>(null);

        return conversionType switch
        {
            "pdf2word" => _wordCom.ConvertPdfToWordAsync(inputPath, outputDir),
            "word2pdf" => _wordCom.ConvertWordToPdfAsync(inputPath, outputDir),
            _ => Task.FromResult<string?>(null)
        };
    }

    public Task<string?> TryConvertWithLibreOfficeAsync(
        string inputPath, string outputDir, string targetExt,
        CancellationToken ct)
    {
        if (_loConverter == null || !_loConverter.IsAvailable)
            return Task.FromResult<string?>(null);

        return _loConverter.ConvertAsync(inputPath, outputDir, targetExt, ct)!;
    }

    public bool IsWordComAvailable => _wordCom?.IsAvailable ?? false;
    public bool IsLibreOfficeAvailable => _loConverter?.IsAvailable ?? false;

    /// <summary>
    /// 根据转换类型返回目标文件扩展名（不含点）。用于冲突检测预判输出文件名。
    /// </summary>
    public static string GetTargetExtension(string conversionType) => conversionType switch
    {
        "png2ico"   => "ico",
        "pdf2word"  => "docx",
        "word2pdf"  => "pdf",
        "pdf2excel" => "xlsx",
        "excel2pdf" => "pdf",
        "pdf2ppt"   => "pptx",
        "ppt2pdf"   => "pdf",
        "word2ppt"  => "pptx",
        "excel2word" => "docx",
        "word2excel" => "xlsx",
        _           => "out"
    };

    public static string GetDefaultOutputDir()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "ConvertPro_Output");
        Directory.CreateDirectory(dir);
        return dir;
    }
}
