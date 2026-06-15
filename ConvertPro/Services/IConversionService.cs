using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertPro.Services;

/// <summary>
/// 转换服务接口 — 所有转换引擎必须实现此接口。
/// 方便日后替换底层库（如 Aspose ↔ Spire），只需实现新类并通过 DI 注入。
/// </summary>
public interface IConversionService
{
    /// <summary>转换类型标识，如 "png2ico"、"pdf2word"</summary>
    string ConversionType { get; }

    /// <summary>支持批量转换</summary>
    bool SupportsBatch { get; }

    /// <summary>执行转换</summary>
    /// <param name="inputFiles">输入文件的完整路径</param>
    /// <param name="outputDir">输出目录</param>
    /// <param name="options">格式特定的选项（JSON）</param>
    /// <param name="progress">进度回调</param>
    /// <param name="ct">取消令牌</param>
    /// <returns>每个输入文件的转换结果</returns>
    Task<List<ConversionResult>> ConvertAsync(
        List<string> inputFiles,
        string outputDir,
        string? options,
        IProgress<ConversionProgress> progress,
        CancellationToken ct
    );
}
