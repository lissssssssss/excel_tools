# WPF Windows 界面（调用现有 Python 核心逻辑）

这个目录下是一个 **WPF (.NET 8)** 桌面程序，用来提供 Windows UI；真正的合并逻辑由 `merge_xls.exe` 执行（它是由仓库根目录的 `merge_xls.py` 打包出来的）。

## 目录结构

```
tmp/
├─ merge_xls.py
├─ requirements.txt
└─ wpf/
   ├─ README_WPF.md
   └─ MergeXlsWpf/
      ├─ MergeXlsWpf.csproj
      ├─ App.xaml
      ├─ MainWindow.xaml
      └─ ...
```

## 目标交付物（给最终用户）

**单一可执行文件**：只需要一个 `MergeXlsWpf.exe`。

- `merge_xls.exe` 会被打包进 `MergeXlsWpf.exe`（作为内嵌资源）
- 运行时自动释放到 `%LocalAppData%\MergeXlsWpf\core\<sha256>\merge_xls.exe` 再执行
- 最终用户机器上 **不需要安装 Python**

## 构建（在你的 Windows 打包机上做一次）

### 1) 打包 Python 核心为 `merge_xls.exe`

在 Windows 机器上，打开 PowerShell：

```powershell
cd <你的项目目录>
py -3.12 -m pip install --upgrade pip
py -3.12 -m pip install -r requirements.txt
py -3.12 -m pip install pyinstaller

# 生成 dist\merge_xls.exe
py -3.12 -m PyInstaller --clean -y --onefile --name merge_xls merge_xls.py
```

生成后把 `dist\merge_xls.exe` 复制到 `wpf\merge_xls.exe`（或直接复制到发布目录，见下方）。

### 2) 打开 / 编译 WPF

二选一：

- **Visual Studio 2022**：打开 `wpf/MergeXlsWpf/MergeXlsWpf.csproj`，直接运行
- **dotnet CLI**：

```powershell
cd wpf/MergeXlsWpf
dotnet restore
dotnet publish -c Release
```

发布产物默认在（self-contained + single-file）：

`wpf\MergeXlsWpf\bin\Release\net8.0-windows\win-x64\publish\`

> `MergeXlsWpf.csproj` 已配置：如果你把 `wpf\merge_xls.exe` 放到位，会在编译时作为资源内嵌进最终 `MergeXlsWpf.exe`。

### 3) 运行与使用

- **源目录**：`/Users/lishengsheng/Documents/铝合金代发` 对应到 Windows 下你的真实目录
- **输出目录**：选择一个可写目录
- 点击 **开始运行**，下方会实时显示 stdout/stderr 日志

## 常见问题

- 如果弹窗提示“未找到内嵌资源”：说明你发布时没有把 `wpf\merge_xls.exe` 放到位，重新复制后再 `dotnet publish`。

## GitHub Actions 自动打包（推荐）

如果你没有 Windows 机器，可以把项目推到 GitHub，然后用 Windows runner 自动产出最终的单一 exe：

- 工作流文件：`.github/workflows/build-windows-single-exe.yml`
- 触发方式：
  - 手动：GitHub → `Actions` → `Build Windows single-exe (WPF + embedded core)` → `Run workflow`
  - 或 push 到 `merge_xls.py` / `requirements.txt` / `wpf/**` 会自动触发
- 产物下载：
  - 进入对应 run，下载 artifact：`MergeXlsWpf-windows-single-exe`
  - 解压后得到 `MergeXlsWpf.exe`（单文件）和 `README_WPF.md`

