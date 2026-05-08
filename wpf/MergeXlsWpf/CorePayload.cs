using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;

namespace MergeXlsWpf;

public static class CorePayload
{
    private const string ResourceName = "MergeXlsWpf.Payload.merge_xls.exe";

    public static string EnsureExtracted()
    {
        var asm = Assembly.GetExecutingAssembly();
        using var s = asm.GetManifestResourceStream(ResourceName);
        if (s is null)
            throw new InvalidOperationException(
                $"未找到内嵌资源 {ResourceName}。请先在 wpf/ 目录放置 merge_xls.exe 再发布。"
            );

        using var ms = new MemoryStream();
        s.CopyTo(ms);
        var bytes = ms.ToArray();

        var sha = Convert.ToHexString(SHA256.HashData(bytes));
        var baseDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MergeXlsWpf",
            "core",
            sha
        );
        Directory.CreateDirectory(baseDir);

        var exePath = Path.Combine(baseDir, "merge_xls.exe");
        if (!File.Exists(exePath) || new FileInfo(exePath).Length != bytes.Length)
        {
            // atomic-ish write
            var tmp = exePath + ".tmp";
            File.WriteAllBytes(tmp, bytes);
            if (File.Exists(exePath)) File.Delete(exePath);
            File.Move(tmp, exePath);
        }

        return exePath;
    }
}

