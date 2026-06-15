using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertPro.Services;

/// <summary>
/// PNG → ICO 真实转换服务。
/// 使用 System.Drawing 读取 PNG、缩放、生成合法的 ICO 容器文件。
/// ICO 格式完全按照 Windows 规范：header(6) + directory entries(16*n) + BITMAPINFOHEADER + BGRA pixels + AND mask。
/// </summary>
public class PngToIcoService : IConversionService
{
    public string ConversionType => "png2ico";
    public bool SupportsBatch => true;

    public async Task<List<ConversionResult>> ConvertAsync(
        List<string> inputFiles,
        string outputDir,
        string? options,
        IProgress<ConversionProgress> progress,
        CancellationToken ct)
    {
        var results = new List<ConversionResult>();

        // 解析选项：选中的 ICO 尺寸
        var sizes = ParseSizes(options);

        for (int i = 0; i < inputFiles.Count; i++)
        {
            ct.ThrowIfCancellationRequested();

            var file = inputFiles[i];
            progress?.Report(new ConversionProgress(
                inputFiles.Count, i, Path.GetFileName(file),
                (double)i / inputFiles.Count * 100));

            try
            {
                var icoPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(file) + ".ico");

                // 生成 ICO 文件
                await Task.Run(() => GenerateIcoFile(file, icoPath, sizes), ct);

                var fi = new FileInfo(icoPath);
                results.Add(new ConversionResult(true, icoPath,
                    OutputSize: fi.Exists ? fi.Length : 0));
            }
            catch (OperationCanceledException)
            {
                results.Add(new ConversionResult(false, "",
                    "已取消"));
                break;
            }
            catch (Exception ex)
            {
                results.Add(new ConversionResult(false, "",
                    ex.Message));
            }

            progress?.Report(new ConversionProgress(
                inputFiles.Count, i + 1, "",
                (double)(i + 1) / inputFiles.Count * 100));
        }

        return results;
    }

    /// <summary>
    /// 解析前端传来的 ICO 尺寸选项 JSON。
    /// 格式：{ "sizes": [16, 32, 48] }，默认 [32]。
    /// </summary>
    private int[] ParseSizes(string? options)
    {
        if (string.IsNullOrEmpty(options)) return [32];

        try
        {
            using var doc = JsonDocument.Parse(options);
            var arr = doc.RootElement.GetProperty("sizes");
            return arr.EnumerateArray().Select(e => e.GetInt32()).ToArray();
        }
        catch
        {
            return [16, 32, 48]; // 默认尺寸
        }
    }

    /// <summary>
    /// 将一张 PNG 转换为含多尺寸的 ICO 文件。
    /// 算法与 JS 原型中 generateIcoBlob 完全一致，已验证通过。
    /// </summary>
    private void GenerateIcoFile(string pngPath, string icoPath, int[] sizes)
    {
        // 读取源 PNG
        using var srcBmp = new Bitmap(pngPath);

        // 预处理：调整每个尺寸的位图
        var bitmaps = new List<(int width, byte[] pixels, int stride, int rowSize, int andRowSize)>();
        foreach (var s in sizes)
        {
            var resized = new Bitmap(s, s, PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(resized);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.DrawImage(srcBmp, 0, 0, s, s);
            g.Dispose();

            var rect = new Rectangle(0, 0, s, s);
            var bmpData = resized.LockBits(rect, ImageLockMode.ReadOnly,
                PixelFormat.Format32bppArgb);

            var pixelBytes = new byte[bmpData.Stride * s];
            System.Runtime.InteropServices.Marshal.Copy(
                bmpData.Scan0, pixelBytes, 0, pixelBytes.Length);

            var stride = bmpData.Stride;
            var rowSize = (s * 32 + 31) / 32 * 4;
            var andRowSize = (s + 31) / 32 * 4;

            bitmaps.Add((s, pixelBytes, stride, rowSize, andRowSize));

            resized.UnlockBits(bmpData);
            resized.Dispose();
        }

        // 计算 ICO 文件大小
        var headerSize = 6;
        var dirEntrySize = 16;
        var dataOffset = headerSize + dirEntrySize * sizes.Length;

        using var fs = new FileStream(icoPath, FileMode.Create);
        using var bw = new BinaryWriter(fs);

        // === ICO Header ===
        bw.Write((ushort)0);     // reserved
        bw.Write((ushort)1);     // type = 1 (ICO)
        bw.Write((ushort)sizes.Length);

        // 预写目录项占位（尺寸、大小、偏移稍后回填）
        var dirEntryPositions = new long[sizes.Length];
        for (int i = 0; i < sizes.Length; i++)
        {
            dirEntryPositions[i] = fs.Position;
            bw.Write(new byte[16]); // 占位 16 bytes
        }

        // 记录每个图像的偏移和大小
        var imageOffsets = new long[sizes.Length];
        var imageSizes = new int[sizes.Length];

        // 写入每个尺寸的图像数据
        for (int i = 0; i < sizes.Length; i++)
        {
            var s = sizes[i];
            var (width, pixelBytes, stride, rowSize, andRowSize) = bitmaps[i];
            var dataSize = 40 + rowSize * s + andRowSize * s;

            imageOffsets[i] = fs.Position;
            imageSizes[i] = dataSize;

            // BITMAPINFOHEADER (40 bytes)
            bw.Write(40);
            bw.Write(s);
            bw.Write(s * 2);
            bw.Write((ushort)1);
            bw.Write((ushort)32);
            bw.Write(0); bw.Write(0); bw.Write(0);
            bw.Write(0); bw.Write(0); bw.Write(0);

            // 像素数据：bottom-up, BGRA
            for (int y = s - 1; y >= 0; y--)
            {
                for (int x = 0; x < s; x++)
                {
                    var idx = y * stride + x * 4;
                    bw.Write(pixelBytes[idx + 0]); // B
                    bw.Write(pixelBytes[idx + 1]); // G
                    bw.Write(pixelBytes[idx + 2]); // R
                    bw.Write(pixelBytes[idx + 3]); // A
                }
                var padding = rowSize - s * 4;
                if (padding > 0) bw.Write(new byte[padding]);
            }

            // AND mask
            bw.Write(new byte[andRowSize * s]);
        }

        // 回填目录项
        for (int i = 0; i < sizes.Length; i++)
        {
            fs.Seek(dirEntryPositions[i], SeekOrigin.Begin);
            var s = sizes[i];
            var w = s >= 256 ? 0 : s;
            bw.Write((byte)w);                    // width (0 = 256)
            bw.Write((byte)w);                    // height
            bw.Write((byte)0);                    // color count
            bw.Write((byte)0);                    // reserved
            bw.Write((ushort)1);                  // planes
            bw.Write((ushort)32);                 // bpp
            bw.Write(imageSizes[i]);              // size
            bw.Write((int)imageOffsets[i]);       // offset (int = 32-bit, matches DWORD)
        }

        bw.Flush();
    }
}
