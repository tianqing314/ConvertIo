using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertPro.Services;

/// <summary>
/// Python pdf2docx 转换器 — 高质量 PDF→Word 转换。
/// 前置条件：Python 3 + pdf2docx (pip install pdf2docx)
/// </summary>
public class PythonPdf2DocxConverter
{
    private readonly bool _available;

    public bool IsAvailable => _available;

    public PythonPdf2DocxConverter()
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "python",
                Arguments = "-c \"import pdf2docx\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using var p = Process.Start(psi);
            p?.WaitForExit(5000);
            _available = p?.ExitCode == 0;
        }
        catch
        {
            _available = false;
        }
    }

    public async Task<string?> ConvertAsync(
        string pdfPath, string outputDir, CancellationToken ct)
    {
        if (!_available) return null;

        var outPath = Path.Combine(outputDir,
            Path.GetFileNameWithoutExtension(pdfPath) + ".docx");

        var script = $@"
from pdf2docx import Converter
cv = Converter(r'{pdfPath}')
cv.convert(r'{outPath}')
cv.close()
";

        var psi = new ProcessStartInfo
        {
            FileName = "python",
            Arguments = $"-c \"{script.Replace("\"", "\\\"")}\"",
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        using var process = Process.Start(psi);
        if (process == null) return null;

        var tcs = new TaskCompletionSource<bool>();
        process.Exited += (_, _) => tcs.TrySetResult(true);
        process.EnableRaisingEvents = true;

        await Task.WhenAny(tcs.Task, Task.Delay(120000, ct));

        if (ct.IsCancellationRequested)
        {
            try { process.Kill(); } catch { }
            return null;
        }

        return File.Exists(outPath) ? outPath : null;
    }
}
