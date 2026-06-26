using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertPro.Services;

/// <summary>
/// 多级回退转换服务：Word COM → LibreOffice → 手动实现。
/// 用于 word2pdf / pdf2excel / excel2pdf / pdf2ppt 这几类不依赖 Python 的转换。
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

                // 1. 尝试 Word COM（最高保真度，仅 pdf2word / word2pdf）
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
