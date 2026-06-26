using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace ConvertPro.Services;

/// <summary>
/// AI 提供商接口（DeepSeek / MiMo 等都实现了这个接口）。
/// 用于 PDF→PPT 的智能内容生成。
/// </summary>
public interface IAiProvider
{
    /// <summary>提供商标识：deepseek / mimo</summary>
    string Name { get; }

    /// <summary>展示名称</summary>
    string DisplayName { get; }

    /// <summary>API Base URL（不含 /chat/completions）</summary>
    string BaseUrl { get; }

    /// <summary>支持的模型列表</summary>
    IReadOnlyList<string> AvailableModels { get; }

    /// <summary>默认模型</summary>
    string DefaultModel { get; }

    /// <summary>API Key 是否已配置（环境变量或手动设置）</summary>
    bool IsAvailable { get; }

    /// <summary>手动覆盖 API Key（设置页调用）</summary>
    void SetApiKey(string? key);

    /// <summary>
    /// 调用 chat/completions 接口
    /// </summary>
    /// <param name="systemPrompt">系统提示词</param>
    /// <param name="userPrompt">用户消息</param>
    /// <param name="model">模型名（null 用 DefaultModel）</param>
    /// <param name="jsonOutput">是否要求 JSON 输出</param>
    /// <param name="ct">取消令牌</param>
    /// <returns>助手回复内容</returns>
    Task<string> ChatAsync(string systemPrompt, string userPrompt,
        string? model = null, bool jsonOutput = false, CancellationToken ct = default);
}

// ================================================================
// OpenAI-compatible 基类（DeepSeek、MiMo 都走这个协议）
// ================================================================
public abstract class OpenAiCompatibleProvider : IAiProvider
{
    private static readonly HttpClient HttpClient = new()
    {
        Timeout = TimeSpan.FromMinutes(5) // AI 生成可能慢，给 5 分钟
    };

    private string? _apiKey;

    protected OpenAiCompatibleProvider(string? initialKey = null)
    {
        _apiKey = initialKey;
    }

    public abstract string Name { get; }
    public abstract string DisplayName { get; }
    public abstract string BaseUrl { get; }
    public abstract IReadOnlyList<string> AvailableModels { get; }
    public abstract string DefaultModel { get; }

    public bool IsAvailable => !string.IsNullOrWhiteSpace(_apiKey);

    public void SetApiKey(string? key)
    {
        _apiKey = string.IsNullOrWhiteSpace(key) ? null : key.Trim();
    }

    public async Task<string> ChatAsync(string systemPrompt, string userPrompt,
        string? model = null, bool jsonOutput = false, CancellationToken ct = default)
    {
        if (!IsAvailable)
            throw new InvalidOperationException(
                $"[{Name}] API Key 未配置，请在设置页填写或设置环境变量 API_KEY");

        var useModel = string.IsNullOrWhiteSpace(model) ? DefaultModel : model!;
        var url = $"{BaseUrl.TrimEnd('/')}/chat/completions";

        // 构造请求体（OpenAI 兼容格式）
        var reqObj = new ChatRequest
        {
            Model = useModel,
            Messages = new List<ChatMessage>
            {
                new() { Role = "system", Content = systemPrompt },
                new() { Role = "user", Content = userPrompt }
            },
            Temperature = 0.7,
            Stream = false
        };

        if (jsonOutput)
        {
            reqObj.ResponseFormat = new ResponseFormat { Type = "json_object" };
        }

        var json = JsonSerializer.Serialize(reqObj,
            new JsonSerializerOptions
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower
            });

        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Content = new StringContent(json, Encoding.UTF8, "application/json");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
        req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        using var resp = await HttpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead, ct);

        var respText = await resp.Content.ReadAsStringAsync(ct);

        if (!resp.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(
                $"[{Name}] API 调用失败 HTTP {(int)resp.StatusCode} {resp.StatusCode}: " +
                Truncate(respText, 500));
        }

        // 解析响应
        try
        {
            var chatResp = JsonSerializer.Deserialize<ChatResponse>(respText,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            var content = chatResp?.Choices?[0]?.Message?.Content;
            if (string.IsNullOrEmpty(content))
                throw new InvalidOperationException($"[{Name}] AI 返回空内容");

            return content!;
        }
        catch (JsonException ex)
        {
            throw new InvalidOperationException(
                $"[{Name}] AI 响应解析失败: {ex.Message}。原始响应: " + Truncate(respText, 500));
        }
    }

    private static string Truncate(string s, int max) =>
        s.Length > max ? s.Substring(0, max) + "..." : s;

    // ====== 请求/响应 DTO（OpenAI 兼容格式） ======
    private sealed class ChatRequest
    {
        public string Model { get; set; } = "";
        public List<ChatMessage> Messages { get; set; } = new();
        public double? Temperature { get; set; }
        public bool Stream { get; set; }
        public ResponseFormat? ResponseFormat { get; set; }
    }

    private sealed class ChatMessage
    {
        public string Role { get; set; } = "";
        public string Content { get; set; } = "";
    }

    private sealed class ResponseFormat
    {
        public string Type { get; set; } = "";
    }

    private sealed class ChatResponse
    {
        public string? Id { get; set; }
        public List<ChatChoice>? Choices { get; set; }
    }

    private sealed class ChatChoice
    {
        public ChatMessage? Message { get; set; }
        public string? FinishReason { get; set; }
    }
}

// ================================================================
// DeepSeek 提供商
// 文档：https://api-docs.deepseek.com/zh-cn/
// ================================================================
public sealed class DeepSeekProvider : OpenAiCompatibleProvider
{
    public static readonly string[] Models = ["deepseek-v4-pro", "deepseek-v4-flash"];

    public DeepSeekProvider(string? apiKey = null) : base(apiKey) { }

    public override string Name => "deepseek";
    public override string DisplayName => "DeepSeek";
    public override string BaseUrl => "https://api.deepseek.com";
    public override IReadOnlyList<string> AvailableModels => Models;
    public override string DefaultModel => "deepseek-v4-flash";
}

// ================================================================
// 小米 MiMo 提供商
// 文档：https://mimo.mi.com/docs/zh-CN/quick-start/faq/api-integration
// ================================================================
public sealed class MiMoProvider : OpenAiCompatibleProvider
{
    public static readonly string[] Models = ["mimo-v2.5-pro", "mimo-v2.5", "mimo-v2-flash"];

    public MiMoProvider(string? apiKey = null) : base(apiKey) { }

    public override string Name => "mimo";
    public override string DisplayName => "小米 MiMo";
    public override string BaseUrl => "https://api.xiaomimimo.com/v1";
    public override IReadOnlyList<string> AvailableModels => Models;
    public override string DefaultModel => "mimo-v2.5";
}

// ================================================================
// AI 提供商管理器
// 负责加载环境变量、实例化、切换提供商、持久化配置。
//
// 配置文件：%LocalAppData%\ConvertPro\ai_config.json
// 优先级：手动设置（UI 输入）> 环境变量 API_KEY（仅 DeepSeek）
// ================================================================
public static class AiProviderManager
{
    private static readonly Dictionary<string, IAiProvider> _providers =
        new(StringComparer.OrdinalIgnoreCase);
    private static string _currentProviderName = "deepseek";

    // 每个提供商当前选用的模型（未设则回退到 DefaultModel）
    private static readonly Dictionary<string, string> _models =
        new(StringComparer.OrdinalIgnoreCase);

    // 手动设置的 API Key（用于持久化；env 变量 key 不进这里）
    private static readonly Dictionary<string, string?> _keys =
        new(StringComparer.OrdinalIgnoreCase);

    private static readonly string _configDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ConvertPro");
    private static readonly string _configFile =
        Path.Combine(_configDir, "ai_config.json");

    static AiProviderManager()
    {
        // 1. 读环境变量 API_KEY（DeepSeek 默认）
        var envApiKey = Environment.GetEnvironmentVariable("API_KEY",
            EnvironmentVariableTarget.User) ??
            Environment.GetEnvironmentVariable("API_KEY",
            EnvironmentVariableTarget.Process);

        // 2. 实例化两个提供商，DeepSeek 优先用环境变量
        var deepseek = new DeepSeekProvider(envApiKey);
        var mimo = new MiMoProvider();

        _providers[deepseek.Name] = deepseek;
        _providers[mimo.Name] = mimo;

        // 3. 加载持久化配置（手动 key 会覆盖 env；模型选择、当前 provider 也会恢复）
        LoadConfig();
    }

    /// <summary>当前活跃的提供商</summary>
    public static IAiProvider Current => _providers[_currentProviderName];

    /// <summary>当前提供商名称</summary>
    public static string CurrentProviderName
    {
        get => _currentProviderName;
        set
        {
            if (!_providers.ContainsKey(value))
                throw new ArgumentException($"未知的 AI 提供商: {value}");
            if (!string.Equals(_currentProviderName, value, StringComparison.OrdinalIgnoreCase))
            {
                _currentProviderName = value;
                SaveConfig();
            }
        }
    }

    /// <summary>所有可用提供商</summary>
    public static IReadOnlyCollection<IAiProvider> All => _providers.Values;

    /// <summary>按名取提供商</summary>
    public static IAiProvider Get(string name)
    {
        if (!_providers.TryGetValue(name, out var p))
            throw new ArgumentException($"未知的 AI 提供商: {name}");
        return p;
    }

    /// <summary>为指定提供商手动覆盖 API Key（设置页用，会持久化）</summary>
    public static void SetApiKey(string providerName, string? apiKey)
    {
        if (!_providers.TryGetValue(providerName, out var p))
            throw new ArgumentException($"未知的 AI 提供商: {providerName}");
        p.SetApiKey(apiKey);
        _keys[providerName] = apiKey;
        SaveConfig();
    }

    /// <summary>设置指定提供商的当前模型（会持久化）</summary>
    public static void SetModel(string providerName, string model)
    {
        if (!_providers.ContainsKey(providerName))
            throw new ArgumentException($"未知的 AI 提供商: {providerName}");
        _models[providerName] = model;
        SaveConfig();
    }

    /// <summary>当前提供商当前模型（未设则用 DefaultModel）</summary>
    public static string GetCurrentModel() => GetModel(_currentProviderName);

    /// <summary>指定提供商当前模型（未设则用 DefaultModel）</summary>
    public static string GetModel(string providerName) =>
        _models.TryGetValue(providerName, out var m) && !string.IsNullOrWhiteSpace(m)
            ? m
            : _providers[providerName].DefaultModel;

    /// <summary>是否有任何提供商可用（用于 UI 显示"AI 已就绪"状态）</summary>
    public static bool AnyAvailable
    {
        get
        {
            foreach (var p in _providers.Values)
                if (p.IsAvailable) return true;
            return false;
        }
    }

    // ================================================================
    // 序列化当前状态为 JSON，供前端渲染设置页 AI 区
    // ================================================================
    public static string GetStatusJson()
    {
        var providers = _providers.Values.Select(p => new
        {
            name = p.Name,
            displayName = p.DisplayName,
            baseUrl = p.BaseUrl,
            defaultModel = p.DefaultModel,
            models = p.AvailableModels,
            currentModel = GetModel(p.Name),
            isAvailable = p.IsAvailable,
            // 是否有手动设置的 key（区别于 env 变量）
            hasManualKey = _keys.TryGetValue(p.Name, out var k) && !string.IsNullOrWhiteSpace(k)
        }).ToList();

        var status = new
        {
            current = _currentProviderName,
            anyAvailable = AnyAvailable,
            providers = providers
        };

        return JsonSerializer.Serialize(status, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });
    }

    // ================================================================
    // 测试连接：调用一次最简 chat，返回人类可读结果
    // ================================================================
    public static async Task<string> TestAsync(
        string? providerName = null, CancellationToken ct = default)
    {
        var p = string.IsNullOrWhiteSpace(providerName)
            ? Current
            : Get(providerName!);

        if (!p.IsAvailable)
            return $"未配置 API Key — 请在下方填写 [{p.DisplayName}] 的 Key，或设置环境变量 API_KEY";

        try
        {
            var reply = await p.ChatAsync(
                "你是连接测试助手，只回复 OK 两个字。",
                "ping",
                model: GetModel(p.Name),
                ct: ct);

            var trimmed = (reply ?? "").Trim();
            if (trimmed.Length > 80) trimmed = trimmed.Substring(0, 80) + "...";
            return $"连接成功（{p.DisplayName} / {GetModel(p.Name)}）：{trimmed}";
        }
        catch (OperationCanceledException)
        {
            return "测试已取消";
        }
        catch (Exception ex)
        {
            var msg = ex.Message;
            if (msg.Length > 200) msg = msg.Substring(0, 200) + "...";
            return $"连接失败（{p.DisplayName}）：{msg}";
        }
    }

    // ================================================================
    // 持久化
    // ================================================================
    private static void LoadConfig()
    {
        try
        {
            if (!File.Exists(_configFile)) return;
            var json = File.ReadAllText(_configFile);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (root.TryGetProperty("currentProvider", out var cp) &&
                cp.ValueKind == JsonValueKind.String)
            {
                var name = cp.GetString();
                if (!string.IsNullOrEmpty(name) && _providers.ContainsKey(name))
                    _currentProviderName = name;
            }

            if (root.TryGetProperty("models", out var modelsEl) &&
                modelsEl.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in modelsEl.EnumerateObject())
                {
                    if (_providers.ContainsKey(prop.Name) &&
                        prop.Value.ValueKind == JsonValueKind.String)
                        _models[prop.Name] = prop.Value.GetString() ?? "";
                }
            }

            if (root.TryGetProperty("keys", out var keysEl) &&
                keysEl.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in keysEl.EnumerateObject())
                {
                    if (_providers.TryGetValue(prop.Name, out var p) &&
                        prop.Value.ValueKind == JsonValueKind.String)
                    {
                        // 手动 key 覆盖 env 变量（手动设置优先）
                        var key = prop.Value.GetString();
                        p.SetApiKey(key);
                        _keys[prop.Name] = key;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[AiProviderManager] 加载配置失败: {ex.Message}");
        }
    }

    private static void SaveConfig()
    {
        try
        {
            Directory.CreateDirectory(_configDir);

            var data = new
            {
                currentProvider = _currentProviderName,
                models = _models,
                keys = _keys
            };

            var json = JsonSerializer.Serialize(data, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = false
            });

            File.WriteAllText(_configFile, json);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[AiProviderManager] 保存配置失败: {ex.Message}");
        }
    }
}

// ================================================================
// 文生图提供商接口（用于 PDF→PPT 的配图生成）
// 与 IAiProvider（chat）分离，独立配置与持久化。
// ================================================================
public interface IImageGenProvider
{
    string Name { get; }
    string DisplayName { get; }
    string BaseUrl { get; }
    IReadOnlyList<string> AvailableModels { get; }
    string DefaultModel { get; }
    bool IsAvailable { get; }
    void SetApiKey(string? key);

    /// <summary>
    /// 调用文生图 API，返回 PNG/JPG 字节数据。
    /// </summary>
    Task<byte[]> GenerateImageAsync(string prompt, string? model = null,
        CancellationToken ct = default);
}

// ================================================================
// OpenAI 兼容文生图基类
// 调用 POST {BaseUrl}/images/generations，body：{model, prompt, n, size, response_format}
// 支持响应中的 url 或 b64_json 两种格式。
// ================================================================
public abstract class OpenAiCompatibleImageProvider : IImageGenProvider
{
    private static readonly HttpClient HttpClient = new()
    {
        Timeout = TimeSpan.FromMinutes(5) // 单张图最长 5 分钟
    };

    private string? _apiKey;

    protected OpenAiCompatibleImageProvider(string? initialKey = null)
    {
        _apiKey = initialKey;
    }

    public abstract string Name { get; }
    public abstract string DisplayName { get; }
    public abstract string BaseUrl { get; }
    public abstract IReadOnlyList<string> AvailableModels { get; }
    public abstract string DefaultModel { get; }

    /// <summary>请求图片尺寸（如 "1440x720"），各 provider 自定。</summary>
    protected abstract string RequestSize { get; }

    /// <summary>响应中是否优先 b64_json；否则下载 url。</summary>
    protected virtual bool PreferB64Json => false;

    public bool IsAvailable => !string.IsNullOrWhiteSpace(_apiKey);

    public void SetApiKey(string? key)
    {
        _apiKey = string.IsNullOrWhiteSpace(key) ? null : key.Trim();
    }

    public async Task<byte[]> GenerateImageAsync(string prompt, string? model = null,
        CancellationToken ct = default)
    {
        if (!IsAvailable)
            throw new InvalidOperationException(
                $"[{Name}] 图模型 API Key 未配置");

        var useModel = string.IsNullOrWhiteSpace(model) ? DefaultModel : model!;
        var url = $"{BaseUrl.TrimEnd('/')}/images/generations";

        var reqObj = new Dictionary<string, object>
        {
            ["model"] = useModel,
            ["prompt"] = prompt,
            ["n"] = 1,
            ["size"] = RequestSize
        };

        if (PreferB64Json) reqObj["response_format"] = "b64_json";

        var json = JsonSerializer.Serialize(reqObj);

        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Content = new StringContent(json, Encoding.UTF8, "application/json");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
        req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        using var resp = await HttpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead, ct);
        var respText = await resp.Content.ReadAsStringAsync(ct);

        if (!resp.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(
                $"[{Name}] 文生图失败 HTTP {(int)resp.StatusCode}: " +
                Truncate(respText, 500));
        }

        // 解析 data[0].url 或 data[0].b64_json
        string? imageUrl = null;
        string? b64 = null;
        try
        {
            using var doc = JsonDocument.Parse(respText);
            if (doc.RootElement.TryGetProperty("data", out var dataEl) &&
                dataEl.ValueKind == JsonValueKind.Array && dataEl.GetArrayLength() > 0)
            {
                var first = dataEl[0];
                if (first.TryGetProperty("b64_json", out var b64El) &&
                    b64El.ValueKind == JsonValueKind.String)
                {
                    b64 = b64El.GetString();
                }
                if (string.IsNullOrEmpty(b64) &&
                    first.TryGetProperty("url", out var urlEl) &&
                    urlEl.ValueKind == JsonValueKind.String)
                {
                    imageUrl = urlEl.GetString();
                }
            }
        }
        catch (JsonException ex)
        {
            throw new InvalidOperationException(
                $"[{Name}] 文生图响应解析失败: {ex.Message}。原始: " + Truncate(respText, 500));
        }

        if (!string.IsNullOrEmpty(b64))
        {
            return Convert.FromBase64String(b64);
        }

        if (string.IsNullOrEmpty(imageUrl))
        {
            throw new InvalidOperationException(
                $"[{Name}] 文生图响应中未找到图片 URL 或 b64_json");
        }

        // 下载图片
        using var imgResp = await HttpClient.GetAsync(imageUrl, ct);
        if (!imgResp.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(
                $"[{Name}] 下载生成图片失败 HTTP {(int)imgResp.StatusCode}");
        }
        return await imgResp.Content.ReadAsByteArrayAsync(ct);
    }

    private static string Truncate(string s, int max) =>
        s.Length > max ? s.Substring(0, max) + "..." : s;
}

// ================================================================
// 智谱 CogView 文生图
// 文档：https://open.bigmodel.cn/dev/api/cogview/cogview-3-plus
// ================================================================
public sealed class ZhipuCogViewProvider : OpenAiCompatibleImageProvider
{
    public static readonly string[] Models = ["cogview-3-plus", "cogview-4"];

    public ZhipuCogViewProvider(string? apiKey = null) : base(apiKey) { }

    public override string Name => "zhipu";
    public override string DisplayName => "智谱 CogView";
    public override string BaseUrl => "https://open.bigmodel.cn/api/paas/v4";
    public override IReadOnlyList<string> AvailableModels => Models;
    public override string DefaultModel => "cogview-3-plus";
    protected override string RequestSize => "1440x720"; // 2:1 宽屏，适合 PPT 配图
}

// ================================================================
// OpenAI DALL-E 文生图
// ================================================================
public sealed class OpenAiDalleProvider : OpenAiCompatibleImageProvider
{
    public static readonly string[] Models = ["dall-e-3", "dall-e-2"];

    public OpenAiDalleProvider(string? apiKey = null) : base(apiKey) { }

    public override string Name => "openai_dalle";
    public override string DisplayName => "OpenAI DALL·E";
    public override string BaseUrl => "https://api.openai.com/v1";
    public override IReadOnlyList<string> AvailableModels => Models;
    public override string DefaultModel => "dall-e-3";
    protected override string RequestSize => "1792x1024"; // 16:9 宽屏
    protected override bool PreferB64Json => true; // DALL·E 支持 b64，避免二次下载
}

// ================================================================
// 文生图提供商管理器
// 独立配置文件：%LocalAppData%\ConvertPro\ai_image_config.json
// 与 AiProviderManager 风格一致，但完全解耦。
// ================================================================
public static class ImageGenManager
{
    private static readonly Dictionary<string, IImageGenProvider> _providers =
        new(StringComparer.OrdinalIgnoreCase);
    private static string _currentProviderName = "zhipu";
    private static readonly Dictionary<string, string> _models =
        new(StringComparer.OrdinalIgnoreCase);
    private static readonly Dictionary<string, string?> _keys =
        new(StringComparer.OrdinalIgnoreCase);

    private static readonly string _configDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ConvertPro");
    private static readonly string _configFile =
        Path.Combine(_configDir, "ai_image_config.json");

    static ImageGenManager()
    {
        // 实例化两个 provider（暂无 env 变量支持，UI 配置为主）
        var zhipu = new ZhipuCogViewProvider();
        var openai = new OpenAiDalleProvider();

        _providers[zhipu.Name] = zhipu;
        _providers[openai.Name] = openai;

        LoadConfig();
    }

    public static IImageGenProvider Current => _providers[_currentProviderName];

    public static string CurrentProviderName
    {
        get => _currentProviderName;
        set
        {
            if (!_providers.ContainsKey(value))
                throw new ArgumentException($"未知的图模型提供商: {value}");
            if (!string.Equals(_currentProviderName, value, StringComparison.OrdinalIgnoreCase))
            {
                _currentProviderName = value;
                SaveConfig();
            }
        }
    }

    public static IReadOnlyCollection<IImageGenProvider> All => _providers.Values;

    public static IImageGenProvider Get(string name)
    {
        if (!_providers.TryGetValue(name, out var p))
            throw new ArgumentException($"未知的图模型提供商: {name}");
        return p;
    }

    public static void SetApiKey(string providerName, string? apiKey)
    {
        if (!_providers.TryGetValue(providerName, out var p))
            throw new ArgumentException($"未知的图模型提供商: {providerName}");
        p.SetApiKey(apiKey);
        _keys[providerName] = apiKey;
        SaveConfig();
    }

    public static void SetModel(string providerName, string model)
    {
        if (!_providers.ContainsKey(providerName))
            throw new ArgumentException($"未知的图模型提供商: {providerName}");
        _models[providerName] = model;
        SaveConfig();
    }

    public static string GetCurrentModel() => GetModel(_currentProviderName);

    public static string GetModel(string providerName) =>
        _models.TryGetValue(providerName, out var m) && !string.IsNullOrWhiteSpace(m)
            ? m
            : _providers[providerName].DefaultModel;

    public static bool AnyAvailable
    {
        get
        {
            foreach (var p in _providers.Values)
                if (p.IsAvailable) return true;
            return false;
        }
    }

    public static string GetStatusJson()
    {
        var providers = _providers.Values.Select(p => new
        {
            name = p.Name,
            displayName = p.DisplayName,
            baseUrl = p.BaseUrl,
            defaultModel = p.DefaultModel,
            models = p.AvailableModels,
            currentModel = GetModel(p.Name),
            isAvailable = p.IsAvailable,
            hasManualKey = _keys.TryGetValue(p.Name, out var k) && !string.IsNullOrWhiteSpace(k)
        }).ToList();

        var status = new
        {
            current = _currentProviderName,
            anyAvailable = AnyAvailable,
            providers = providers
        };

        return JsonSerializer.Serialize(status, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });
    }

    public static async Task<string> TestAsync(string? providerName = null,
        CancellationToken ct = default)
    {
        var p = string.IsNullOrWhiteSpace(providerName) ? Current : Get(providerName!);

        if (!p.IsAvailable)
            return $"未配置 API Key — 请填写 [{p.DisplayName}] 的 Key";

        try
        {
            var bytes = await p.GenerateImageAsync(
                "A simple blue circle on white background, minimal",
                model: GetModel(p.Name), ct: ct);

            return $"连接成功（{p.DisplayName} / {GetModel(p.Name)}）：生成图片 {bytes.Length / 1024} KB";
        }
        catch (OperationCanceledException)
        {
            return "测试已取消";
        }
        catch (Exception ex)
        {
            var msg = ex.Message;
            if (msg.Length > 200) msg = msg.Substring(0, 200) + "...";
            return $"连接失败（{p.DisplayName}）：{msg}";
        }
    }

    private static void LoadConfig()
    {
        try
        {
            if (!File.Exists(_configFile)) return;
            var json = File.ReadAllText(_configFile);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (root.TryGetProperty("currentProvider", out var cp) &&
                cp.ValueKind == JsonValueKind.String)
            {
                var name = cp.GetString();
                if (!string.IsNullOrEmpty(name) && _providers.ContainsKey(name))
                    _currentProviderName = name;
            }

            if (root.TryGetProperty("models", out var modelsEl) &&
                modelsEl.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in modelsEl.EnumerateObject())
                {
                    if (_providers.ContainsKey(prop.Name) &&
                        prop.Value.ValueKind == JsonValueKind.String)
                        _models[prop.Name] = prop.Value.GetString() ?? "";
                }
            }

            if (root.TryGetProperty("keys", out var keysEl) &&
                keysEl.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in keysEl.EnumerateObject())
                {
                    if (_providers.TryGetValue(prop.Name, out var p) &&
                        prop.Value.ValueKind == JsonValueKind.String)
                    {
                        var key = prop.Value.GetString();
                        p.SetApiKey(key);
                        _keys[prop.Name] = key;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[ImageGenManager] 加载配置失败: {ex.Message}");
        }
    }

    private static void SaveConfig()
    {
        try
        {
            Directory.CreateDirectory(_configDir);

            var data = new
            {
                currentProvider = _currentProviderName,
                models = _models,
                keys = _keys
            };

            var json = JsonSerializer.Serialize(data, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = false
            });

            File.WriteAllText(_configFile, json);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[ImageGenManager] 保存配置失败: {ex.Message}");
        }
    }
}
