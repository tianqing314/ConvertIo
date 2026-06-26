using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertPro.Services;

/// <summary>
/// Python pdf2docx 转换器 — 高质量 PDF→Word 转换。
/// 前置条件：Python 3 + pdf2docx (pip install pdf2docx)
/// 修复点：
/// 1. 不再用字符串插值拼 Python 代码（路径含单引号会断），改为写临时 .py 文件
/// 2. 改用 WaitForExitAsync + 取消令牌，避免进程竞态和超时泄漏
/// 3. 异步读取 stdout/stderr，避免管道缓冲区满死锁
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

        // 把脚本写到临时文件，避免路径里的特殊字符破坏 Python 字符串
        var scriptPath = Path.Combine(Path.GetTempPath(),
            "convertpro_pdf2docx_" + Guid.NewGuid().ToString("N") + ".py");
        var script = $@"# -*- coding: utf-8 -*-
import sys
from pdf2docx import Converter

src = sys.argv[1]
dst = sys.argv[2]

cv = Converter(src)
try:
    cv.convert(dst)
finally:
    cv.close()
";
        await File.WriteAllTextAsync(scriptPath, script, ct);

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "python",
                ArgumentList = { scriptPath, pdfPath, outPath },
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };

            var exitedTcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            EventHandler exitedHandler = (_, _) => exitedTcs.TrySetResult(true);
            process.Exited += exitedHandler;

            if (!process.Start())
                return null;

            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();

            // Python pdf2docx 大文件可能耗时几分钟，给 5 分钟
            using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
            using var linked = CancellationTokenSource.CreateLinkedTokenSource(ct, timeoutCts.Token);

            try
            {
                using (linked.Token.Register(() =>
                {
                    try { if (!process.HasExited) process.Kill(entireProcessTree: true); }
                    catch { /* ignore */ }
                }))
                {
                    await exitedTcs.Task.WaitAsync(linked.Token);
                }
            }
            catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested && !ct.IsCancellationRequested)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[Python pdf2docx] 转换超时: {Path.GetFileName(pdfPath)}");
                return null;
            }
            catch (OperationCanceledException)
            {
                return null;
            }
            finally
            {
                process.Exited -= exitedHandler;
            }

            try { await Task.WhenAll(stdoutTask, stderrTask); } catch { }

            return File.Exists(outPath) ? outPath : null;
        }
        finally
        {
            try { if (File.Exists(scriptPath)) File.Delete(scriptPath); } catch { }
        }
    }
}
