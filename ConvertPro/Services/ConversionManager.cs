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
        Register(new WordComOrLibreOfficeService("pdf2ppt", "pptx", null, _loConverter, new PdfToPptService()));
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

    public static string GetDefaultOutputDir()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "ConvertPro_Output");
        Directory.CreateDirectory(dir);
        return dir;
    }
}

/// <summary>
/// 多级回退转换服务：Word COM → LibreOffice → 手动实现。
/// </summary>
public class WordComOrLibreOfficeService : IConversionService
{
    private readonly string _convType;
    private readonly string _targetExt;
    private readonly WordComConverter? _wordCom;
    private readonly LibreOfficeConverter? _lo;
    private readonly IConversionService? _fallback;

    public string ConversionType => _convType;
    public bool SupportsBatch => true;

    public WordComOrLibreOfficeService(string type, string targetExt,
        WordComConverter? wordCom, LibreOfficeConverter? lo,
        IConversionService? fallback = null)
    {
        _convType = type;
        _targetExt = targetExt;
        _wordCom = wordCom;
        _lo = lo;
        _fallback = fallback;
    }

    public async Task<List<ConversionResult>> ConvertAsync(
        List<string> inputFiles, string outputDir, string? options,
        IProgress<ConversionProgress> progress, CancellationToken ct)
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
                string? resultPath = null;

                // 1. 尝试 Word COM（最高保真度）
                if (_wordCom?.IsAvailable == true &&
                    (_convType == "pdf2word" || _convType == "word2pdf"))
                {
                    resultPath = _convType == "pdf2word"
                        ? await _wordCom.ConvertPdfToWordAsync(file, outputDir)
                        : await _wordCom.ConvertWordToPdfAsync(file, outputDir);
                }

                // 2. 尝试 LibreOffice
                if (resultPath == null && _lo?.IsAvailable == true)
                {
                    resultPath = await _lo.ConvertAsync(
                        file, outputDir, _targetExt, ct);
                }

                // 3. 手动回退
                if (resultPath == null && _fallback != null)
                {
                    var fallbackResults = await _fallback.ConvertAsync(
                        new List<string> { file }, outputDir, options,
                        new Progress<ConversionProgress>(), ct);
                    var fr = fallbackResults.Count > 0 ? fallbackResults[0]
                        : new ConversionResult(false, "", "转换失败");
                    results.Add(fr);
                    progress?.Report(new ConversionProgress(
                        inputFiles.Count, i + 1, "",
                        (double)(i + 1) / inputFiles.Count * 100));
                    continue;
                }

                if (resultPath != null)
                {
                    results.Add(new ConversionResult(true, resultPath,
                        OutputSize: new FileInfo(resultPath).Length));
                }
                else
                {
                    results.Add(new ConversionResult(false, "",
                        "无可用的转换引擎。请安装 Microsoft Word 或 LibreOffice。"));
                }
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
}

/// <summary>
/// PDF→Word 多引擎服务：Python pdf2docx → Word COM → LibreOffice → PdfPig。
/// </summary>
public class MultiEnginePdfToWordService : IConversionService
{
    private readonly PythonPdf2DocxConverter _python;
    private readonly WordComConverter? _wordCom;
    private readonly LibreOfficeConverter? _lo;
    private readonly IConversionService _fallback;

    public string ConversionType => "pdf2word";
    public bool SupportsBatch => true;

    public MultiEnginePdfToWordService(
        PythonPdf2DocxConverter python, WordComConverter? wordCom,
        LibreOfficeConverter? lo, IConversionService fallback)
    {
        _python = python; _wordCom = wordCom; _lo = lo; _fallback = fallback;
    }

    public async Task<List<ConversionResult>> ConvertAsync(
        List<string> inputFiles, string outputDir, string? options,
        IProgress<ConversionProgress> progress, CancellationToken ct)
    {
        var results = new List<ConversionResult>();

        for (int i = 0; i < inputFiles.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            var file = inputFiles[i];
            progress?.Report(new ConversionProgress(inputFiles.Count, i,
                Path.GetFileName(file), (double)i / inputFiles.Count * 100));

            try
            {
                string? resultPath = null;

                // 1. Python pdf2docx（最佳，已验证可以正确转换该 PDF）
                if (_python.IsAvailable)
                    resultPath = await _python.ConvertAsync(file, outputDir, ct);

                // 2. Word COM
                if (resultPath == null && _wordCom?.IsAvailable == true)
                    resultPath = await _wordCom.ConvertPdfToWordAsync(file, outputDir);

                // 3. LibreOffice
                if (resultPath == null && _lo?.IsAvailable == true)
                    resultPath = await _lo.ConvertAsync(file, outputDir, "docx", ct);

                // 4. 回退
                if (resultPath == null)
                {
                    var fr = await _fallback.ConvertAsync(
                        new List<string> { file }, outputDir, options,
                        new Progress<ConversionProgress>(), ct);
                    results.Add(fr.Count > 0 ? fr[0] : new ConversionResult(false, "", "转换失败"));
                }
                else
                {
                    results.Add(new ConversionResult(true, resultPath,
                        OutputSize: new FileInfo(resultPath).Length));
                }
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

            progress?.Report(new ConversionProgress(inputFiles.Count, i + 1, "",
                (double)(i + 1) / inputFiles.Count * 100));
        }

        return results;
    }
}
