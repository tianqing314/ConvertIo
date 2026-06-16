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
                var versionDirs = Directory.GetDirectories(root)
                    .Select(Path.GetFileName)
                    .Where(v => Version.TryParse(v, out _))
                    .OrderByDescending(v => v, StringComparer.OrdinalIgnoreCase)
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

            webView.CoreWebView2.Navigate("https://app.assets/index.html");

            // 通知前端引擎状态
            var engine = _conversionManager.IsWordComAvailable
                ? "word_com"
                : _conversionManager.IsLibreOfficeAvailable ? "libreoffice" : "fallback";
            await webView.CoreWebView2.ExecuteScriptAsync(
                $"window.__onEngineStatus('{engine}')");
        }

        // ================================================================
        // JS → C# 消息处理
        // ================================================================
        private async void OnWebMessageReceived(object sender,
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
                        // 打开文件选择对话框
                        await HandleOpenFileDialogAsync(root);
                        break;

                    case "start_convert":
                        // 启动真实转换
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
                }
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
                ["pdf2ppt"] = "PDF 文档|*.pdf"
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

            System.Diagnostics.Debug.WriteLine(
                $"[ConvertPro] 开始转换: type={convType}, files={string.Join(", ", inputFiles)}");

            var outputDir = ConversionManager.GetDefaultOutputDir();

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
