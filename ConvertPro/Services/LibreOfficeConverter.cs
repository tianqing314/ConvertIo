using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertPro.Services;

/// <summary>
/// LibreOffice 命令行转换器。
/// 使用 LibreOffice 的 headless 模式实现高保真文档转换。
/// 修复点：
/// 1. 改用 WaitForExitAsync + 取消令牌，避免进程竞态和超时泄漏
/// 2. 异步读取 stdout/stderr，避免管道缓冲区满导致死锁
/// 3. 单文件 60s 超时；取消时显式 Kill 进程
/// </summary>
public class LibreOfficeConverter
{
    private readonly string? _sofficePath;
    private readonly bool _available;

    public bool IsAvailable => _available;

    public LibreOfficeConverter()
    {
        var paths = new[]
        {
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        };

        foreach (var p in paths)
        {
            if (File.Exists(p))
            {
                _sofficePath = p;
                _available = true;
                return;
            }
        }

        _available = false;
    }

    public async Task<string?> ConvertAsync(
        string inputPath, string outputDir, string targetExt,
        CancellationToken ct)
    {
        if (!_available || _sofficePath == null) return null;

        var args = $"--headless --convert-to {targetExt} " +
                   $"--outdir \"{outputDir}\" \"{inputPath}\"";

        var psi = new ProcessStartInfo
        {
            FileName = _sofficePath,
            Arguments = args,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };

        // 启动前先订阅，避免启动后立即退出导致 Exited 事件丢失
        var exitedTcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        EventHandler exitedHandler = (_, _) => exitedTcs.TrySetResult(true);
        process.Exited += exitedHandler;

        if (!process.Start())
            return null;

        // 异步读取输出，避免管道缓冲区满导致进程 hang
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        // 整体超时 60s + 用户取消
        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(1));
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(ct, timeoutCts.Token);

        try
        {
            // 直接 await 这个 task；如果被取消或超时会抛 OperationCanceledException
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
                $"[LibreOffice] 转换超时: {Path.GetFileName(inputPath)}");
            return null;
        }
        catch (OperationCanceledException)
        {
            // 用户主动取消
            return null;
        }
        finally
        {
            process.Exited -= exitedHandler;
        }

        // 确保 stdout/stderr 都读完（避免僵尸）
        try { await Task.WhenAll(stdoutTask, stderrTask); } catch { }

        var baseName = Path.GetFileNameWithoutExtension(inputPath);
        var outputPath = Path.Combine(outputDir, baseName + "." + targetExt);
        return File.Exists(outputPath) ? outputPath : null;
    }
}
