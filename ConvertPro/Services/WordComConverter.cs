using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertPro.Services;

/// <summary>
/// Microsoft Word COM 转换器 — 使用已安装的 Word 进行原生转换。
/// PDF→Word 和 Word→PDF 转换保真度最高（Microsoft 原生引擎）。
/// 注意：Word.Application 必须在 STA 线程上访问，本类内部已通过 Task.Run
/// + STA 切换处理；但批量复用同一实例的性能优化放到 Phase 3。
/// </summary>
public class WordComConverter
{
    private readonly bool _available;

    public bool IsAvailable => _available;

    public WordComConverter()
    {
        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application");
            _available = wordType != null;
        }
        catch
        {
            _available = false;
        }
    }

    /// <summary>
    /// PDF → Word，使用 Word 2013+ 原生 PDF 导入。
    /// </summary>
    public async Task<string?> ConvertPdfToWordAsync(string pdfPath, string outputDir)
    {
        return await RunWithWordAsync(word =>
        {
            // Formats:=7 (wdOpenFormatPDF)
            dynamic doc = word.Documents.Open(pdfPath, false, false,
                PasswordDocument: Type.Missing,
                AddToRecentFiles: false,
                Format: 7, // PDF
                NoEncodingDialog: true);

            try
            {
                var outPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(pdfPath) + ".docx");
                doc.SaveAs2(outPath, 16); // 16 = wdFormatDocumentDefault
                return File.Exists(outPath) ? outPath : null;
            }
            finally
            {
                SafeCloseDoc(doc);
            }
        }, TimeSpan.FromMinutes(5));
    }

    /// <summary>
    /// Word → PDF，使用 Word 原生导出。
    /// </summary>
    public async Task<string?> ConvertWordToPdfAsync(string docxPath, string outputDir)
    {
        return await RunWithWordAsync(word =>
        {
            dynamic doc = word.Documents.Open(docxPath);
            try
            {
                var outPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(docxPath) + ".pdf");
                doc.ExportAsFixedFormat(outPath, 17); // 17 = wdExportFormatPDF
                return File.Exists(outPath) ? outPath : null;
            }
            finally
            {
                SafeCloseDoc(doc);
            }
        }, TimeSpan.FromMinutes(5));
    }

    /// <summary>
    /// 在一个新的 Word.Application 实例上执行回调，确保实例被正确释放。
    /// 后续 Phase 3 会引入对象池复用同一实例。
    /// </summary>
    private async Task<string?> RunWithWordAsync(Func<dynamic, string?> action, TimeSpan timeout)
    {
        return await Task.Run(() =>
        {
            dynamic? word = null;
            try
            {
                var wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null) return null;

                word = Activator.CreateInstance(wordType);
                word.Visible = false;
                word.DisplayAlerts = 0;

                // 用 Task + 超时避免 Word 卡死导致永远不返回
                var tcs = new TaskCompletionSource<string?>();
                var thread = new System.Threading.Thread(() =>
                {
                    try { tcs.TrySetResult(action(word!)); }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[WordCom] 转换失败: {ex.Message}");
                        tcs.TrySetResult(null);
                    }
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();

                if (!tcs.Task.Wait(timeout))
                {
                    // 超时：尝试杀掉 Word 进程
                    System.Diagnostics.Debug.WriteLine(
                        "[WordCom] 转换超时，强制结束 Word");
                    try { word.Quit(); } catch { }
                    return null;
                }
                return tcs.Task.Result;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (word != null)
                {
                    try { word.Quit(); } catch { }
                    try { Marshal.ReleaseComObject(word); } catch { }
                }
            }
        });
    }

    /// <summary>
    /// 安全关闭 doc 对象：仅关闭一次，已关闭则忽略。
    /// 修复原代码在 try 和 finally 都调用 doc.Close() 导致的 COM 异常。
    /// </summary>
    private static void SafeCloseDoc(dynamic doc)
    {
        if (doc == null) return;
        try
        {
            doc.Close(false); // false = 不保存更改
            Marshal.ReleaseComObject(doc);
        }
        catch
        {
            // doc 可能已被关闭或处于无效状态，忽略
            try { Marshal.ReleaseComObject(doc); } catch { }
        }
    }
}
