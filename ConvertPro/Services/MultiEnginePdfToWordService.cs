using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertPro.Services;

/// <summary>
/// PDF→Word 多引擎服务：Python pdf2docx → Word COM → LibreOffice → PdfPig。
/// 优先级最高的 Python pdf2docx 在保真度和速度上都是最优解，
/// 后续依次降级到 Microsoft 原生、LibreOffice，最后才是纯文本提取。
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
