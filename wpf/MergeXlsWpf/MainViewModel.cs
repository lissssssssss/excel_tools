using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Forms = System.Windows.Forms;

namespace MergeXlsWpf;

public sealed class MainViewModel : ObservableObject
{
    private string _sourceDir = "";
    private string _outputDir = "";
    private string _logText = "";
    private string _statusText = "就绪";
    private bool _isRunning;

    private CancellationTokenSource? _cts;
    private Process? _proc;

    public MainViewModel()
    {
        BrowseSourceDirCommand = new RelayCommand(BrowseSourceDir, () => IsNotRunning);
        BrowseOutputDirCommand = new RelayCommand(BrowseOutputDir, () => IsNotRunning);

        RunCommand = new RelayCommand(async () => await RunAsync(), () => IsNotRunning);
        CancelCommand = new RelayCommand(Cancel, () => IsRunning);
        ClearLogCommand = new RelayCommand(() => LogText = "", () => IsNotRunning);
    }

    public string SourceDir { get => _sourceDir; set => Set(ref _sourceDir, value); }
    public string OutputDir { get => _outputDir; set => Set(ref _outputDir, value); }
    public string LogText { get => _logText; set => Set(ref _logText, value); }
    public string StatusText { get => _statusText; set => Set(ref _statusText, value); }

    public bool IsRunning
    {
        get => _isRunning;
        private set
        {
            if (Set(ref _isRunning, value))
            {
                Raise(nameof(IsNotRunning));
                (BrowseSourceDirCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (BrowseOutputDirCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (RunCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (CancelCommand as RelayCommand)?.RaiseCanExecuteChanged();
                (ClearLogCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsNotRunning => !IsRunning;

    public ICommand BrowseSourceDirCommand { get; }
    public ICommand BrowseOutputDirCommand { get; }
    public ICommand RunCommand { get; }
    public ICommand CancelCommand { get; }
    public ICommand ClearLogCommand { get; }

    private void AppendLog(string line)
    {
        if (string.IsNullOrEmpty(line)) return;
        System.Windows.Application.Current.Dispatcher.Invoke(() =>
        {
            LogText += line + Environment.NewLine;
        });
    }

    private void BrowseSourceDir()
    {
        var dlg = new Forms.FolderBrowserDialog
        {
            Description = "选择源目录（包含 .xls 文件）",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = false
        };
        var result = dlg.ShowDialog();
        if (result == Forms.DialogResult.OK)
            SourceDir = dlg.SelectedPath;
    }

    private void BrowseOutputDir()
    {
        var dlg = new Forms.FolderBrowserDialog
        {
            Description = "选择输出目录（写入 2月汇总.xlsx 等）",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true
        };
        var result = dlg.ShowDialog();
        if (result == Forms.DialogResult.OK)
            OutputDir = dlg.SelectedPath;
    }

    private async Task RunAsync()
    {
        string coreExePath;
        try
        {
            coreExePath = CorePayload.EnsureExtracted();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                ex.Message, "内嵌核心程序不可用",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error
            );
            return;
        }

        if (string.IsNullOrWhiteSpace(SourceDir) || !Directory.Exists(SourceDir))
        {
            System.Windows.MessageBox.Show(
                "源目录无效，请选择包含 .xls 的目录。", "错误",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error
            );
            return;
        }
        if (string.IsNullOrWhiteSpace(OutputDir))
        {
            System.Windows.MessageBox.Show(
                "输出目录不能为空。", "错误",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error
            );
            return;
        }

        Directory.CreateDirectory(OutputDir);

        LogText = "";
        StatusText = "启动中...";
        IsRunning = true;
        _cts = new CancellationTokenSource();

        var args = new StringBuilder();
        args.Append('"').Append(SourceDir).Append('"').Append(' ');
        args.Append('"').Append(OutputDir).Append('"');

        var psi = new ProcessStartInfo
        {
            FileName = coreExePath,
            Arguments = args.ToString(),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            CreateNoWindow = true,
            WorkingDirectory = Path.GetDirectoryName(coreExePath) ?? Environment.CurrentDirectory,
        };
        // Force UTF-8 output from the embedded Python core (avoid Chinese garbling).
        psi.Environment["PYTHONUTF8"] = "1";
        psi.Environment["PYTHONIOENCODING"] = "utf-8";

        AppendLog("Command:");
        AppendLog($"  \"{psi.FileName}\" {psi.Arguments}");
        AppendLog("");

        try
        {
            _proc = new Process { StartInfo = psi, EnableRaisingEvents = true };
            _proc.OutputDataReceived += (_, e) => { if (e.Data != null) AppendLog(e.Data); };
            _proc.ErrorDataReceived += (_, e) => { if (e.Data != null) AppendLog(e.Data); };

            _proc.Start();
            _proc.BeginOutputReadLine();
            _proc.BeginErrorReadLine();

            StatusText = "运行中...";

            using (_cts.Token.Register(() =>
                   {
                       try { if (_proc is { HasExited: false }) _proc.Kill(entireProcessTree: true); }
                       catch { /* ignore */ }
                   }))
            {
                await _proc.WaitForExitAsync(_cts.Token);
            }

            var code = _proc.ExitCode;
            StatusText = code == 0 ? "完成" : $"退出码: {code}";
            AppendLog("");
            AppendLog($"[WPF] Process exited with code {code}");
        }
        catch (OperationCanceledException)
        {
            StatusText = "已取消";
            AppendLog("");
            AppendLog("[WPF] Canceled.");
        }
        catch (Exception ex)
        {
            StatusText = "失败";
            AppendLog("");
            AppendLog("[WPF] ERROR: " + ex);
            System.Windows.MessageBox.Show(
                ex.Message, "运行失败",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error
            );
        }
        finally
        {
            IsRunning = false;
            _proc?.Dispose();
            _proc = null;
            _cts?.Dispose();
            _cts = null;
        }
    }

    private void Cancel()
    {
        try
        {
            StatusText = "取消中...";
            _cts?.Cancel();
        }
        catch
        {
            // ignore
        }
    }
}

