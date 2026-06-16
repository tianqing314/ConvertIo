# ConvertPro 离线环境 WebView2 Runtime 修复指南

## 问题描述

ConvertPro 启动时报错：

```
未找到 WebView2 Runtime，程序无法运行。
请从以下地址下载并安装 WebView2 Runtime：
https://developer.microsoft.com/en-us/microsoft-edge/webview2/
```

**根本原因**：WebView2 SDK 默认通过 Windows 注册表查找 Runtime，但某些环境下（手动复制安装、注册表损坏、企业镜像等）注册表键缺失，导致查找失败。即使 Runtime 文件已存在于磁盘，程序仍无法识别。

---

## 解决方案

### 方案一：应用代码修复（推荐）

修改 `MainWindow.xaml.cs` 中的 WebView2 初始化逻辑，直接扫描已知安装目录，绕过注册表查找。

#### 步骤 1：定位代码文件

打开项目目录：
```
ConvertPro/MainWindow.xaml.cs
```

#### 步骤 2：替换 WebView2 初始化方法

找到 `InitializeWebViewAsync` 方法，将原始代码：

```csharp
private async Task InitializeWebViewAsync()
{
    var userDataFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ConvertPro", "WebView2");

    // ⚠️ 使用默认 CreateAsync — WebView2 会自动查找已安装的运行时
    var env = await CoreWebView2Environment.CreateAsync(userDataFolder);
    await webView.EnsureCoreWebView2Async(env);
    // ... 后续代码不变
```

替换为：

```csharp
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

    // ... 后续代码不变
```

#### 步骤 3：编译验证

```powershell
cd 项目根目录
dotnet build
```

确认输出 `已成功生成。0 个错误`。

---

### 方案二：修复注册表（无需改代码）

如果不想修改代码，可以手动添加缺失的注册表键。

#### 步骤 1：确认 Runtime 安装路径

检查以下目录是否存在 WebView2 Runtime：

```
C:\Program Files (x86)\Microsoft\EdgeWebView\Application\
```

或

```
C:\Program Files\Microsoft\EdgeWebView\Application\
```

进入目录，记录版本号文件夹名称（如 `149.0.4022.69`），确认其中包含 `msedgewebview2.exe`。

#### 步骤 2：创建注册表文件

新建文件 `fix-webview2.reg`，内容如下：

```reg
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BEE-13A6279B0908}]
"pv"="149.0.4022.69"
```

> **注意**：将 `149.0.4022.69` 替换为你系统中实际的版本号。

#### 步骤 3：导入注册表

双击 `fix-webview2.reg` 文件，确认导入。

或以管理员身份运行命令行：

```cmd
reg import fix-webview2.reg
```

#### 步骤 4：验证

```cmd
reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BEE-13A6279B0908}"
```

应输出 `pv` 值为对应版本号。

---

## 离线环境下获取 WebView2 Runtime

如果目标机器**完全没有安装** WebView2 Runtime，需要在有网络的机器上下载后拷贝过去。

### 方法一：下载离线安装包

1. **在有网络的电脑上**，访问：
   ```
   https://developer.microsoft.com/en-us/microsoft-edge/webview2/
   ```

2. 下载 **Evergreen Standalone Installer (x64)** 或 **(x86)**：
   - 文件名类似：`MicrosoftEdgeWebview2Setup.exe`
   - 大小约 1-2 MB（在线引导安装器）

   如需完全离线安装包，使用 **Evergreen Standalone (x64) Offline** 下载链接，文件约 150 MB+。

3. 将安装包拷贝到离线电脑，双击运行即可。

### 方法二：直接复制 Runtime 文件

如果无法运行安装程序，可以手动复制：

#### 在有网络的电脑上：

1. 安装 WebView2 Runtime 后，复制整个目录：
   ```
   C:\Program Files (x86)\Microsoft\EdgeWebView\Application\<版本号>\
   ```

2. 将此目录打包（约 300-500 MB），通过 U 盘拷贝。

#### 在离线电脑上：

1. 创建目录结构：
   ```
   C:\Program Files (x86)\Microsoft\EdgeWebView\Application\<版本号>\
   ```

2. 将文件解压到该目录。

3. **配合方案一（代码修复）使用**，无需注册表即可正常运行。

4. **或配合方案二（注册表修复）使用**，手动添加注册表键。

---

## 验证修复结果

### 检查 Runtime 是否存在

```powershell
# PowerShell
Test-Path "C:\Program Files (x86)\Microsoft\EdgeWebView\Application"
Get-ChildItem "C:\Program Files (x86)\Microsoft\EdgeWebView\Application"
```

应输出类似：

```
目录: C:\Program Files (x86)\Microsoft\EdgeWebView\Application

Mode                 LastWriteTime         Length Name
----                 -------------         ------ ----
d-----        2026/6/16   9:18:52                149.0.4022.69
```

### 检查关键文件

确认版本目录下存在 `msedgewebview2.exe`：

```powershell
Test-Path "C:\Program Files (x86)\Microsoft\EdgeWebView\Application\149.0.4022.69\msedgewebview2.exe"
```

输出 `True` 即为正常。

### 运行程序

```powershell
cd ConvertPro目录
dotnet run
```

程序正常启动且不弹出错误提示即为修复成功。

---

## 常见问题

### Q1：修改代码后仍然报错？

检查版本目录下是否存在 `msedgewebview2.exe`。有些目录（如 `SetupMetrics`）不是 Runtime 目录，代码会自动跳过。

### Q2：x64 系统上 Runtime 安装在 x86 路径正常吗？

正常。WebView2 Runtime 默认安装在 `Program Files (x86)` 目录，即使是 64 位系统。代码已同时检查两个路径。

### Q3：企业环境无法修改注册表？

优先使用方案一（代码修复），完全不依赖注册表。Runtime 文件只需存在于磁盘即可。

### Q4：目标机器有 Edge 浏览器但没有 WebView2 Runtime？

Edge 浏览器和 WebView2 Runtime 是独立安装的。有 Edge 不代表有 WebView2 Runtime。需要单独安装。

---

## 文件清单

| 文件 | 说明 |
|------|------|
| `MainWindow.xaml.cs` | 主窗口代码，WebView2 初始化逻辑在此 |
| `fix-webview2.reg` | 注册表修复文件（方案二使用） |
| `MicrosoftEdgeWebview2Setup.exe` | WebView2 在线安装引导器 |

---

## 版本记录

| 日期 | 操作 |
|------|------|
| 2026-06-16 | 初始版本，修复 WebView2 Runtime 注册表查找失败问题 |
