using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertPro.Services;

/// <summary>
/// LibreOffice 命令行转换器。
/// 使用 LibreOffice 的 headless 模式实现高保真文档转换。
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

        using var process = Process.Start(psi);
        if (process == null) return null;

        var tcs = new TaskCompletionSource<bool>();
        process.Exited += (_, _) => tcs.TrySetResult(true);
        process.EnableRaisingEvents = true;

        await Task.WhenAny(tcs.Task, Task.Delay(60000, ct));

        if (ct.IsCancellationRequested)
        {
            try { process.Kill(); } catch { }
            return null;
        }

        var baseName = Path.GetFileNameWithoutExtension(inputPath);
        var outputPath = Path.Combine(outputDir, baseName + "." + targetExt);
        return File.Exists(outputPath) ? outputPath : null;
    }
}
