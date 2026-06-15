# OneNoteHyperlinkRemover

OneNote COM 加载项：移除 OneNote 自动将 URL 转换为超链接的行为。

## 功能说明

OneNote 会自动将粘贴或输入的 URL 文本转换为可点击的超链接，且没有内置选项可以禁用此行为。本插件提供：

- **手动移除**：一键扫描当前页面，将自动转换的超链接恢复为纯文本
- **自动监控**：开启后自动检测并移除新出现的自动超链接
- **智能识别**：只移除 OneNote 自动转换的超链接（href 和显示文本相同），保留用户手动创建的有意义的超链接

## 开发环境要求

- Visual Studio 2022（需要"Office/SharePoint 开发"工作负载）
- .NET Framework 4.8
- Microsoft OneNote（Microsoft 365 版本）
- Windows 10/11

## 编译

```powershell
# 使用 Visual Studio 打开 OneNoteHyperlinkRemover.sln 编译
# 或使用命令行：
msbuild OneNoteHyperlinkRemover.sln /p:Configuration=Release
```

编译输出在 `bin\Release\` 目录。

## 注册和安装

### 开发环境（推荐使用 PowerShell 脚本）

```powershell
# 以管理员权限运行 PowerShell
.\Register.ps1 -Configuration Release
```

### 手动注册（使用 bat 脚本）

以管理员权限运行 `Register.bat`。

### 脚本执行的操作

1. 使用 `RegAsm.exe` 注册 COM 组件
2. 在注册表中添加 OneNote 加载项条目：
   `HKCU\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn`

## 使用方法

1. 重启 OneNote
2. 在"开始"选项卡中找到"超链接工具"组
3. 点击"移除超链接"按钮扫描并移除当前页面的自动超链接
4. 点击"自动移除"切换按钮开启/关闭自动监控模式

## 卸载

```powershell
# 以管理员权限运行
.\Unregister.ps1
```

然后重启 OneNote。

## 技术架构

这是一个 **COM Add-in**（不是 VSTO），因为 OneNote 不支持 VSTO 项目模板。

### 核心接口

- `IDTExtensibility2` — COM 加载项生命周期（OnConnection、OnDisconnection 等）
- `IRibbonExtensibility` — Ribbon UI 定义和回调

### 关键文件

| 文件 | 说明 |
|------|------|
| `AddIn.cs` | 入口点，实现 COM 接口和 Ribbon 回调 |
| `OneNoteHelper.cs` | OneNote COM API 封装，管理 IApplication 生命周期 |
| `HyperlinkRemover.cs` | 核心逻辑：解析页面 XML，识别并移除自动超链接 |
| `Ribbon\Ribbon.xml` | Ribbon UI 定义（按钮和切换按钮） |
| `Register.ps1` | 注册脚本（PowerShell） |
| `Unregister.ps1` | 注销脚本（PowerShell） |

### 工作原理

1. 通过 `IApplication.GetPageContent()` 获取当前页面的 XML 内容
2. 使用 LINQ to XML 解析页面结构（Outline → OE → T 元素）
3. 在 T 元素的 CDATA 内容中查找 `<a href="...">text</a>` 标签
4. 识别自动转换的超链接（href 和显示文本相同或显示文本是 URL）
5. 将自动超链接替换为纯文本
6. 通过 `IApplication.UpdatePageContent()` 写回修改后的页面

### COM 对象生命周期

参考 OneMore 项目的设计：**不长期持有 IApplication COM 引用**，每次操作创建新实例并及时释放，避免阻止 OneNote 正常关闭。

## 参考项目

- [OneMore](https://github.com/stevencohn/OneMore) — 功能丰富的 OneNote COM 加载项（C#，.NET Framework 4.8）
- [OneNote JS Add-in](../OneNote_add-ins_manifest/) — 同一功能的 JavaScript Office Add-in 版本

## 日志

插件运行日志位于：
`%LOCALAPPDATA%\OneNoteHyperlinkRemover\addin.log`
