namespace ConvertPro.Services;

/// <summary>
/// 转换进度回调
/// </summary>
public record ConversionProgress(int TotalFiles, int CompletedFiles, string CurrentFile, double Percent);

/// <summary>
/// 转换结果
/// </summary>
public record ConversionResult(
    bool Success,
    string OutputPath,
    string ErrorMessage = "",
    long OutputSize = 0
);
