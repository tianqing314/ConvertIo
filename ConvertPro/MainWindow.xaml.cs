using ConvertPro.Services;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;

namespace ConvertPro
{
    public partial class MainWindow : Window
    {
        // ================================================================
        // Win32 API
        // ================================================================
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg,
            IntPtr wParam, IntPtr lParam);

        private const int WM_NCLBUTTONDOWN = 0x00A1;
        private const int HTCAPTION = 0x0002;
        private const int WM_NCCALCSIZE = 0x0083;

        // ================================================================
        // 转换引擎
        // ================================================================
        private readonly ConversionManager _conversionManager = new();
        private CancellationTokenSource? _currentCts;

        // 冲突对话框异步等待：C# 触发 JS 显示对话框，等待 JS 回传用户选择
        private TaskCompletionSource<string>? _conflictTcs;

        // ================================================================
        // 构造函数
        // ================================================================
        public MainWindow()
        {
            InitializeComponent();
            this.SourceInitialized += OnSourceInitialized;
            this.Loaded += async (s, e) => await InitializeWebViewAsync();
        }

        // ================================================================
        // Win32: 去顶部白边
        // ================================================================
        private void OnSourceInitialized(object sender, EventArgs e)
        {
            var handle = new WindowInteropHelper(this).Handle;
            HwndSource.FromHwnd(handle)?.AddHook(WndProc);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam,
            IntPtr lParam, ref bool handled)
        {
            if (msg == WM_NCCALCSIZE && wParam != IntPtr.Zero)
            {
                handled = true;
                return IntPtr.Zero;
            }
            return IntPtr.Zero;
        }

        // ================================================================
        // WebView2 初始化
        // ================================================================
        /// <summary>
        /// 查找系统上已安装的 WebView2 Runtime 文件夹
        /// </summary>
        private static string? FindWebView2RuntimeFolder()
        {
            // 候选安装根目录（x86 和 x64）
            string[] roots = [
                @"C:\Program Files (x86)\Microsoft\EdgeWebView\Application",
                @"C:\Program Files\Microsoft\EdgeWebView\Application",
            ];

            foreach (var root in roots)
            {
                if (!Directory.Exists(root)) continue;

                // 取最高版本号目录（目录名格式如 "149.0.4022.69"）
                // ⚠️ 必须用 Version 比较，字符串比较会让 "9.0.0" > "10.0.0"
                var versionDirs = Directory.GetDirectories(root)
                    .Select(Path.GetFileName)
                    .Where(v => Version.TryParse(v, out _))
                    .OrderByDescending(v => Version.Parse(v))
                    .ToList();

                foreach (var ver in versionDirs)
                {
                    var candidate = Path.Combine(root, ver);
                    if (File.Exists(Path.Combine(candidate, "msedgewebview2.exe")))
                        return candidate;
                }
            }
            return null;
        }

        private async Task InitializeWebViewAsync()
        {
            var userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ConvertPro", "WebView2");

            // 显式查找 WebView2 Runtime 文件夹，绕过注册表查找
            var browserFolder = FindWebView2RuntimeFolder();
            if (browserFolder == null)
            {
                MessageBox.Show(
                    "未找到 WebView2 Runtime，程序无法运行。\n\n" +
                    "请从以下地址下载并安装 WebView2 Runtime：\n" +
                    "https://developer.microsoft.com/en-us/microsoft-edge/webview2/",
                    "缺少 WebView2 Runtime", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
                return;
            }

            var env = await CoreWebView2Environment.CreateAsync(browserFolder, userDataFolder);
            await webView.EnsureCoreWebView2Async(env);

            var wwwroot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
            webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "app.assets", wwwroot, CoreWebView2HostResourceAccessKind.Allow);

            var settings = webView.CoreWebView2.Settings;
            settings.AreDefaultContextMenusEnabled = false;
            settings.IsStatusBarEnabled = false;
            settings.IsZoomControlEnabled = false;
            settings.IsPasswordAutosaveEnabled = false;
            settings.IsGeneralAutofillEnabled = false;

            webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

            // 注入通信脚本 + 禁用浏览器快捷键
            await webView.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(@"
                window.postAction = function(action, data) {
                    try {
                        window.chrome.webview.postMessage(
                            JSON.stringify({ action: action, data: data || {} })
                        );
                    } catch(e) {}
                };
                document.addEventListener('keydown', function(e) {
                    if (e.key === 'F5' || e.key === 'F12' ||
                        (e.ctrlKey && ('rRfFpPuU'.indexOf(e.key) !== -1)) ||
                        (e.ctrlKey && e.shiftKey && e.key === 'i')) {
                        e.preventDefault(); return false;
                    }
                }, true);
            ");

            // 引擎状态：放到首次 NavigationCompleted 后再发，避免前端脚本未就绪
            bool engineNotified = false;
            webView.CoreWebView2.NavigationCompleted += async (s, ev) =>
            {
                if (ev.IsSuccess && !engineNotified)
                {
                    engineNotified = true;
                    var engine = _conversionManager.IsWordComAvailable
                        ? "word_com"
                        : _conversionManager.IsLibreOfficeAvailable ? "libreoffice" : "fallback";
                    try
                    {
                        await webView.CoreWebView2.ExecuteScriptAsync(
                            $"window.__onEngineStatus && window.__onEngineStatus('{engine}')");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[ConvertPro] 发送引擎状态失败: {ex.Message}");
                    }

                    // 推送 AI 提供商状态给设置页渲染
                    try
                    {
                        await SendAiStatusAsync();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[ConvertPro] 发送 AI 状态失败: {ex.Message}");
                    }

                    // 推送 PPT 模板列表给 PDF→PPT 页渲染
                    try
                    {
                        await SendPptTemplatesAsync();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[ConvertPro] 发送 PPT 模板失败: {ex.Message}");
                    }

                    // 推送图模型（文生图）状态给设置页渲染
                    try
                    {
                        await SendImageGenStatusAsync();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[ConvertPro] 发送图模型状态失败: {ex.Message}");
                    }
                }
            };

            webView.CoreWebView2.Navigate("https://app.assets/index.html");
        }

        // ================================================================
        // JS → C# 消息处理
        // ================================================================
        // 事件处理器必须是 async void；为防止内部 await 后抛异常导致进程崩溃，
        // 实际工作交给 async Task 方法，这里仅做 fire-and-forget + 异常捕获。
        private void OnWebMessageReceived(object sender,
            CoreWebView2WebMessageReceivedEventArgs e)
        {
            _ = HandleWebMessageAsync(sender, e);
        }

        private async Task HandleWebMessageAsync(object sender,
            CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var json = e.TryGetWebMessageAsString();
                if (string.IsNullOrEmpty(json)) return;

                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                var action = root.GetProperty("action").GetString();

                switch (action)
                {
                    case "window_drag":
                        DragWindow();
                        break;
                    case "window_minimize":
                        WindowState = WindowState.Minimized;
                        break;
                    case "window_maximize":
                        WindowState = WindowState == WindowState.Maximized
                            ? WindowState.Normal : WindowState.Maximized;
                        break;
                    case "window_close":
                        Close();
                        break;

                    case "open_file":
                        await HandleOpenFileDialogAsync(root);
                        break;

                    case "start_convert":
                        await HandleStartConversionAsync(root);
                        break;

                    case "open_folder":
                        OpenOutputFolder();
                        break;

                    case "cancel_convert":
                        _currentCts?.Cancel();
                        break;

                    case "get_output_dir":
                        await SendOutputDirAsync();
                        break;

                    case "conflict_resolve":
                        HandleConflictResolve(root);
                        break;

                    case "ai_get_status":
                        await SendAiStatusAsync();
                        break;
                    case "ai_set_provider":
                        HandleAiSetProvider(root);
                        await SendAiStatusAsync();
                        break;
                    case "ai_set_key":
                        HandleAiSetKey(root);
                        await SendAiStatusAsync();
                        break;
                    case "ai_set_model":
                        HandleAiSetModel(root);
                        await SendAiStatusAsync();
                        break;
                    case "ai_test":
                        await HandleAiTestAsync(root);
                        break;

                    case "img_get_status":
                        await SendImageGenStatusAsync();
                        break;
                    case "img_set_provider":
                        HandleImageGenSetProvider(root);
                        await SendImageGenStatusAsync();
                        break;
                    case "img_set_key":
                        HandleImageGenSetKey(root);
                        await SendImageGenStatusAsync();
                        break;
                    case "img_set_model":
                        HandleImageGenSetModel(root);
                        await SendImageGenStatusAsync();
                        break;
                    case "img_test":
                        await HandleImageGenTestAsync(root);
                        break;

                    case "ppt_get_templates":
                        await SendPptTemplatesAsync();
                        break;
                }
            }
            catch (OperationCanceledException)
            {
                // 用户取消，不算异常，不报错
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ConvertPro] 消息异常: {ex.Message}");
            }
        }

        // ================================================================
        // 文件对话框
        // ================================================================
        private async Task HandleOpenFileDialogAsync(JsonElement root)
        {
            var convType = root.GetProperty("data").GetProperty("type").GetString() ?? "png2ico";

            var filterMap = new Dictionary<string, string>
            {
                ["png2ico"] = "PNG 图片|*.png",
                ["pdf2word"] = "PDF 文档|*.pdf",
                ["word2pdf"] = "Word 文档|*.docx;*.doc",
                ["pdf2excel"] = "PDF 文档|*.pdf",
                ["excel2pdf"] = "Excel 工作簿|*.xlsx;*.xls",
                ["pdf2ppt"] = "PDF 文档|*.pdf",
                ["ppt2pdf"] = "PowerPoint 演示文稿|*.pptx;*.ppt",
                ["word2ppt"] = "Word 文档|*.docx;*.doc",
                ["excel2word"] = "Excel 工作簿|*.xlsx;*.xls",
                ["word2excel"] = "Word 文档|*.docx;*.doc"
            };

            var filter = filterMap.GetValueOrDefault(convType, "所有文件|*.*");
            var multiSelect = true;

            var dialog = new OpenFileDialog
            {
                Filter = filter,
                Multiselect = multiSelect,
                Title = "选择要转换的文件"
            };

            if (dialog.ShowDialog() == true)
            {
                var files = dialog.FileNames.Select(f => new
                {
                    name = Path.GetFileName(f),
                    path = f,
                    size = new FileInfo(f).Length
                }).ToList();

                var filesJson = JsonSerializer.Serialize(files);
                await webView.CoreWebView2.ExecuteScriptAsync(
                    $"window.__onFilesSelected({filesJson})");
            }
        }

        // ================================================================
        // 真实转换
        // ================================================================
        private async Task HandleStartConversionAsync(JsonElement root)
        {
            var data = root.GetProperty("data");
            var convType = data.GetProperty("type").GetString() ?? "png2ico";
            var optionsJson = data.TryGetProperty("options", out var o) ? o.GetRawText() : null;

            // 解析文件路径列表
            var pathsJson = data.GetProperty("files");
            var inputFiles = pathsJson.EnumerateArray()
                .Select(f => f.GetProperty("path").GetString() ?? "")
                .ToList();

            // 读取前端传来的输出目录（用户在设置页指定的），为空则用 C# 默认值
            string outputDir;
            if (data.TryGetProperty("outputDir", out var odEl) &&
                odEl.ValueKind == JsonValueKind.String)
            {
                var od = odEl.GetString()?.Trim();
                outputDir = !string.IsNullOrWhiteSpace(od) ? od : ConversionManager.GetDefaultOutputDir();
            }
            else
            {
                outputDir = ConversionManager.GetDefaultOutputDir();
            }

            // 确保输出目录存在
            Directory.CreateDirectory(outputDir);

            // 读取冲突处理策略（Phase 2-B 会真正使用，此处先记录日志）
            var conflict = data.TryGetProperty("conflict", out var cEl) && cEl.ValueKind == JsonValueKind.String
                ? cEl.GetString() ?? "ask"
                : "ask";

            System.Diagnostics.Debug.WriteLine(
                $"[ConvertPro] 开始转换: type={convType}, files={string.Join(", ", inputFiles)}, " +
                $"outputDir={outputDir}, conflict={conflict}");

            // ===== 冲突检测 =====
            // 默认假设输出文件名为 <basename>.<targetExt>，与各 ConversionService 内部约定一致
            var targetExt = ConversionManager.GetTargetExtension(convType);
            var conflictFiles = new List<string>(); // 存在冲突的输入文件路径
            foreach (var f in inputFiles)
            {
                if (string.IsNullOrWhiteSpace(f)) continue;
                var expectedOutput = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(f) + "." + targetExt);
                if (File.Exists(expectedOutput)) conflictFiles.Add(f);
            }

            // ===== 应用冲突策略 =====
            // ask 模式：若有冲突，弹窗让用户选择（一次性应用到全部冲突文件）
            if (conflict == "ask" && conflictFiles.Count > 0)
            {
                conflict = await AskUserConflictAsync(conflictFiles);
                System.Diagnostics.Debug.WriteLine(
                    $"[ConvertPro] 用户选择冲突策略: {conflict}");
            }

            // skip 模式：从输入列表中过滤掉冲突文件
            if (conflict == "skip" && conflictFiles.Count > 0)
            {
                var skipSet = new HashSet<string>(conflictFiles, StringComparer.OrdinalIgnoreCase);
                int skippedCount = inputFiles.Count(f => skipSet.Contains(f));
                inputFiles = inputFiles.Where(f => !skipSet.Contains(f)).ToList();

                // 通知前端跳过了多少文件
                await webView.CoreWebView2.ExecuteScriptAsync(
                    $"window.__onConvertStarted && window.__onConvertStarted()");

                if (inputFiles.Count == 0)
                {
                    // 全部跳过，直接发"完成"事件
                    var skipJson = JsonSerializer.Serialize(new
                    {
                        toast = $"已跳过 {skippedCount} 个已存在文件",
                        skipped = skippedCount
                    });
                    await webView.CoreWebView2.ExecuteScriptAsync(
                        $"window.__showToast && window.__showToast({skipJson})");
                    await webView.CoreWebView2.ExecuteScriptAsync(
                        $"window.__onConvertComplete([])");
                    return;
                }

                // 部分跳过，发个 toast 提示
                var partialSkipJson = JsonSerializer.Serialize(new
                {
                    toast = $"已跳过 {skippedCount} 个已存在文件，继续转换 {inputFiles.Count} 个"
                });
                await webView.CoreWebView2.ExecuteScriptAsync(
                    $"window.__showToast && window.__showToast({partialSkipJson})");
            }

            // rename 模式：把已存在的输出文件先备份移走，转换完成后再恢复
            // 实现思路：备份 → 转换（生成新文件覆盖原位置）→ 将新文件改名为 <basename>_1.<ext>，再恢复备份到原名
            var renameBackups = new List<(string backupPath, string originalPath, string newOutputPath)>();
            if (conflict == "rename" && conflictFiles.Count > 0)
            {
                foreach (var f in conflictFiles)
                {
                    var originalPath = Path.Combine(outputDir,
                        Path.GetFileNameWithoutExtension(f) + "." + targetExt);
                    if (!File.Exists(originalPath)) continue;

                    // 找一个不冲突的新名：<basename>_1.<ext>、_2、_3...
                    var baseName = Path.GetFileNameWithoutExtension(originalPath);
                    string newOutputPath;
                    int idx = 1;
                    do
                    {
                        newOutputPath = Path.Combine(outputDir, $"{baseName}_{idx}.{targetExt}");
                        idx++;
                    } while (File.Exists(newOutputPath));

                    // 把已存在的文件先移到临时备份位置
                    var backupPath = originalPath + ".convertpro_bak";
                    File.Move(originalPath, backupPath, overwrite: false);
                    renameBackups.Add((backupPath, originalPath, newOutputPath));
                }
            }

            // 发送开始消息给前端
            await webView.CoreWebView2.ExecuteScriptAsync(
                "window.__onConvertStarted()");

            _currentCts = new CancellationTokenSource();
            var progress = new Progress<ConversionProgress>(async p =>
            {
                var pJson = JsonSerializer.Serialize(new
                {
                    total = p.TotalFiles,
                    completed = p.CompletedFiles,
                    current = p.CurrentFile,
                    percent = p.Percent
                });
                await webView.CoreWebView2.ExecuteScriptAsync(
                    $"window.__onConvertProgress({pJson})");
            });

            var results = await _conversionManager.ConvertAsync(
                convType, inputFiles, outputDir, optionsJson,
                progress, _currentCts.Token);

            _currentCts = null;

            // rename 模式后处理：把刚生成的新文件改名为 _1.<ext>，把备份恢复到原名
            for (int i = 0; i < results.Count; i++)
            {
                var r = results[i];
                if (!r.Success || renameBackups.Count == 0) continue;

                // 找匹配的备份（按输出路径匹配 originalPath）
                var match = renameBackups.FirstOrDefault(b =>
                    string.Equals(b.originalPath, r.OutputPath, StringComparison.OrdinalIgnoreCase));
                if (match.Equals(default)) continue;

                try
                {
                    // 1) 把刚转换出的新文件移到 _1.<ext>
                    if (File.Exists(match.newOutputPath))
                        File.Delete(match.newOutputPath);
                    File.Move(r.OutputPath, match.newOutputPath);

                    // 2) 把备份恢复到原名
                    File.Move(match.backupPath, match.originalPath);

                    // 3) 更新结果中的路径
                    results[i] = r with { OutputPath = match.newOutputPath };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[ConvertPro] rename 后处理失败: {ex.Message}");
                    // 清理可能的备份残留
                    try { if (File.Exists(match.backupPath)) File.Move(match.backupPath, match.originalPath, overwrite: true); }
                    catch { /* 忽略 */ }
                }
            }

            // 发送结果
            var resultList = results.Select(r => new
            {
                success = r.Success,
                path = r.OutputPath,
                name = Path.GetFileName(r.OutputPath),
                error = r.ErrorMessage,
                size = r.OutputSize
            }).ToList();

            foreach (var r in resultList)
                System.Diagnostics.Debug.WriteLine(
                    $"[ConvertPro] 结果: success={r.success}, name={r.name}, error={r.error}, path={r.path}");

            var resultJson = JsonSerializer.Serialize(resultList);
            await webView.CoreWebView2.ExecuteScriptAsync(
                $"window.__onConvertComplete({resultJson})");
        }

        private async Task SendOutputDirAsync()
        {
            var dir = ConversionManager.GetDefaultOutputDir();
            await webView.CoreWebView2.ExecuteScriptAsync(
                $"window.__onOutputDir('{dir.Replace("\\", "\\\\")}')");
        }

        // ================================================================
        // AI 提供商配置（设置页）
        // ================================================================
        private async Task SendAiStatusAsync()
        {
            var json = AiProviderManager.GetStatusJson();
            await webView.CoreWebView2.ExecuteScriptAsync(
                $"window.__onAiStatus && window.__onAiStatus({json})");
        }

        // ================================================================
        // PPT 模板列表（PDF→PPT 页）
        // ================================================================
        private async Task SendPptTemplatesAsync()
        {
            var list = PptTemplates.All.Select(t => new
            {
                id = t.Id,
                name = t.Name,
                bg = t.BgHex,
                panel = t.PanelHex,
                accent = t.AccentHex,
                accent2 = t.Accent2Hex,
                category = t.Category,
                theme = t.Theme,
                tags = t.Tags
            }).ToList();

            var defaultId = PptTemplates.Default.Id;
            var payload = new { templates = list, defaultId = defaultId };

            var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            await webView.CoreWebView2.ExecuteScriptAsync(
                $"window.__onPptTemplates && window.__onPptTemplates({json})");
        }

        private static void HandleAiSetProvider(JsonElement root)
        {
            var name = root.GetProperty("data").GetProperty("provider").GetString() ?? "deepseek";
            AiProviderManager.CurrentProviderName = name;
        }

        private static void HandleAiSetKey(JsonElement root)
        {
            var data = root.GetProperty("data");
            var provider = data.GetProperty("provider").GetString() ?? "deepseek";
            var key = data.TryGetProperty("key", out var kEl) && kEl.ValueKind == JsonValueKind.String
                ? kEl.GetString()
                : null;
            AiProviderManager.SetApiKey(provider, key);
        }

        private static void HandleAiSetModel(JsonElement root)
        {
            var data = root.GetProperty("data");
            var provider = data.GetProperty("provider").GetString() ?? "deepseek";
            var model = data.GetProperty("model").GetString() ?? "";
            AiProviderManager.SetModel(provider, model);
        }

        private async Task HandleAiTestAsync(JsonElement root)
        {
            var data = root.GetProperty("data");
            var provider = data.TryGetProperty("provider", out var pEl) && pEl.ValueKind == JsonValueKind.String
                ? pEl.GetString()
                : null;

            // 先告诉前端"测试中"
            await webView.CoreWebView2.ExecuteScriptAsync(
                "window.__onAiTest && window.__onAiTest({state:'testing'})");

            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            var result = await AiProviderManager.TestAsync(provider, cts.Token);

            var resultJson = JsonSerializer.Serialize(new { state = "done", message = result });
            await webView.CoreWebView2.ExecuteScriptAsync(
                $"window.__onAiTest && window.__onAiTest({resultJson})");
        }

        // ================================================================
        // 图模型（文生图）配置 — 用于 PDF→PPT 配图
        // ================================================================
        private async Task SendImageGenStatusAsync()
        {
            var json = ImageGenManager.GetStatusJson();
            await webView.CoreWebView2.ExecuteScriptAsync(
                $"window.__onImageGenStatus && window.__onImageGenStatus({json})");
        }

        private static void HandleImageGenSetProvider(JsonElement root)
        {
            var name = root.GetProperty("data").GetProperty("provider").GetString() ?? "zhipu";
            ImageGenManager.CurrentProviderName = name;
        }

        private static void HandleImageGenSetKey(JsonElement root)
        {
            var data = root.GetProperty("data");
            var provider = data.GetProperty("provider").GetString() ?? "zhipu";
            var key = data.TryGetProperty("key", out var kEl) && kEl.ValueKind == JsonValueKind.String
                ? kEl.GetString()
                : null;
            ImageGenManager.SetApiKey(provider, key);
        }

        private static void HandleImageGenSetModel(JsonElement root)
        {
            var data = root.GetProperty("data");
            var provider = data.GetProperty("provider").GetString() ?? "zhipu";
            var model = data.GetProperty("model").GetString() ?? "";
            ImageGenManager.SetModel(provider, model);
        }

        private async Task HandleImageGenTestAsync(JsonElement root)
        {
            var data = root.GetProperty("data");
            var provider = data.TryGetProperty("provider", out var pEl) && pEl.ValueKind == JsonValueKind.String
                ? pEl.GetString()
                : null;

            // 先告诉前端"测试中"
            await webView.CoreWebView2.ExecuteScriptAsync(
                "window.__onImageGenTest && window.__onImageGenTest({state:'testing'})");

            // 文生图测试需要更长时间（生成图片）
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(120));
            var result = await ImageGenManager.TestAsync(provider, cts.Token);

            var resultJson = JsonSerializer.Serialize(new { state = "done", message = result });
            await webView.CoreWebView2.ExecuteScriptAsync(
                $"window.__onImageGenTest && window.__onImageGenTest({resultJson})");
        }

        // ================================================================
        // 冲突对话框：C# → JS 显示弹窗 → 等待 JS 回传选择
        // ================================================================
        private async Task<string> AskUserConflictAsync(List<string> conflictFiles)
        {
            _conflictTcs = new TaskCompletionSource<string>(
                TaskCreationOptions.RunContinuationsAsynchronously);

            // 只显示第一个冲突文件名 + 冲突总数，避免列表太长
            var firstFile = Path.GetFileName(conflictFiles[0]);
            var fileJson = JsonSerializer.Serialize(new
            {
                fileName = firstFile,
                count = conflictFiles.Count
            });

            try
            {
                await webView.CoreWebView2.ExecuteScriptAsync(
                    $"window.__showConflictBatch && window.__showConflictBatch({fileJson})");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ConvertPro] 显示冲突弹窗失败: {ex.Message}");
                _conflictTcs = null;
                return "overwrite"; // 出错时默认覆盖
            }

            // 等待用户选择（最多等 5 分钟，避免永久挂起）
            using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(
                timeoutCts.Token, _currentCts?.Token ?? CancellationToken.None);

            try
            {
                using (linkedCts.Token.Register(() =>
                    _conflictTcs?.TrySetResult("skip"))) // 超时或取消时按跳过处理
                {
                    return await _conflictTcs.Task;
                }
            }
            finally
            {
                _conflictTcs = null;
            }
        }

        private void HandleConflictResolve(JsonElement root)
        {
            var choice = root.GetProperty("data").GetProperty("choice").GetString() ?? "skip";
            _conflictTcs?.TrySetResult(choice);
        }

        // ================================================================
        // 打开输出文件夹
        // ================================================================
        private void OpenOutputFolder()
        {
            var dir = Services.ConversionManager.GetDefaultOutputDir();
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = dir,
                UseShellExecute = true
            });
        }

        // ================================================================
        // 窗口拖拽
        // ================================================================
        private void DragWindow()
        {
            Dispatcher.Invoke(() =>
            {
                ReleaseCapture();
                SendMessage(new WindowInteropHelper(this).Handle,
                    WM_NCLBUTTONDOWN, HTCAPTION, IntPtr.Zero);
            });
        }
    }
}
