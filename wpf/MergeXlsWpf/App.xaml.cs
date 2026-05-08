namespace MergeXlsWpf;

public partial class App : System.Windows.Application
{
    private static string LogPath =>
        System.IO.Path.Combine(
            System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
            "MergeXlsWpf",
            "logs",
            "app.log"
        );

    private static void AppendLog(string msg)
    {
        try
        {
            var dir = System.IO.Path.GetDirectoryName(LogPath);
            if (!string.IsNullOrEmpty(dir))
                System.IO.Directory.CreateDirectory(dir);
            System.IO.File.AppendAllText(
                LogPath,
                $"[{System.DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}{System.Environment.NewLine}"
            );
        }
        catch
        {
            // ignore
        }
    }

    protected override void OnStartup(System.Windows.StartupEventArgs e)
    {
        base.OnStartup(e);

        // Allow GBK/CP936 decoding if needed.
        try
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }
        catch
        {
            // ignore
        }

        // Catch early crashes and write a log.
        this.DispatcherUnhandledException += (_, ex) =>
        {
            AppendLog("DispatcherUnhandledException: " + ex.Exception);
            System.Windows.MessageBox.Show(
                "程序启动失败，错误已写入日志：\n" + LogPath + "\n\n" + ex.Exception.Message,
                "启动失败",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error
            );
            ex.Handled = true;
            Shutdown(1);
        };

        System.AppDomain.CurrentDomain.UnhandledException += (_, ex) =>
        {
            AppendLog("UnhandledException: " + ex.ExceptionObject);
        };

        try
        {
            AppendLog("App starting. Version=" + (System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "unknown"));
            var win = new MainWindow();
            MainWindow = win;
            win.Show();
        }
        catch (System.Exception ex)
        {
            AppendLog("OnStartup exception: " + ex);
            System.Windows.MessageBox.Show(
                "程序启动失败，错误已写入日志：\n" + LogPath + "\n\n" + ex.Message,
                "启动失败",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error
            );
            Shutdown(1);
        }
    }
}

