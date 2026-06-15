using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace ConvertPro.Services;

/// <summary>
/// Microsoft Word COM 转换器 — 使用已安装的 Word 进行原生转换。
/// PDF→Word 和 Word→PDF 转换保真度最高（Microsoft 原生引擎）。
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
        return await Task.Run(() =>
        {
            dynamic? word = null;
            dynamic? doc = null;
            try
            {
                word = Activator.CreateInstance(
                    Type.GetTypeFromProgID("Word.Application")!);
                word.Visible = false;
                word.DisplayAlerts = 0;

                // Formats:=7 (wdOpenFormatPDF)
                doc = word.Documents.Open(pdfPath, false, false,
                    PasswordDocument: Type.Missing,
                    AddToRecentFiles: false,
                    Format: 7, // PDF
                    NoEncodingDialog: true);

                var outPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(pdfPath) + ".docx");
                doc.SaveAs2(outPath, 16); // 16 = wdFormatDocumentDefault
                doc.Close();

                return File.Exists(outPath) ? outPath : null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (doc != null) try { doc.Close(); } catch { }
                if (word != null)
                {
                    try { word.Quit(); } catch { }
                    Marshal.ReleaseComObject(word);
                }
            }
        });
    }

    /// <summary>
    /// Word → PDF，使用 Word 原生导出。
    /// </summary>
    public async Task<string?> ConvertWordToPdfAsync(string docxPath, string outputDir)
    {
        return await Task.Run(() =>
        {
            dynamic? word = null;
            dynamic? doc = null;
            try
            {
                word = Activator.CreateInstance(
                    Type.GetTypeFromProgID("Word.Application")!);
                word.Visible = false;
                word.DisplayAlerts = 0;

                doc = word.Documents.Open(docxPath);
                var outPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(docxPath) + ".pdf");
                doc.ExportAsFixedFormat(outPath, 17); // 17 = wdExportFormatPDF
                doc.Close();

                return File.Exists(outPath) ? outPath : null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (doc != null) try { doc.Close(); } catch { }
                if (word != null)
                {
                    try { word.Quit(); } catch { }
                    Marshal.ReleaseComObject(word);
                }
            }
        });
    }
}
