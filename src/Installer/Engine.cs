using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using Microsoft.Win32;
namespace SheetPaceInstaller
{
    internal static class Engine
    {
        internal const string ProgId = "SheetPace.Connect", Clsid = "{D19B1B2A-80CA-41D4-9DFD-3E82C147560B}";
        internal const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\SheetPace";
        internal static string InstallPath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "SheetPace"); } }
        internal static byte[] Payload()
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("SheetPace.dll"))
            using (MemoryStream output = new MemoryStream())
            { if (stream == null) throw new InvalidDataException("安装包缺少插件文件。"); stream.CopyTo(output); return output.ToArray(); }
        }
        internal static string Hash(byte[] bytes) { using (SHA256 hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(); }
        internal static void VerifyPayload() { if (Hash(Payload()) != BuildInfo.DllHash) throw new InvalidDataException("插件文件校验失败，请重新获取安装包。"); }
        internal static void CheckEnvironment()
        {
            if (Process.GetProcessesByName("EXCEL").Length != 0) throw new InvalidOperationException("请保存工作簿并关闭所有 Excel 进程后重试。安装程序不会关闭您的 Excel。");
            bool framework = false, excel = false;
            foreach (RegistryView view in Views())
                using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                {
                    using (RegistryKey key = root.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
                        framework |= key != null && Convert.ToInt32(key.GetValue("Release", 0)) >= 528040;
                    using (RegistryKey key = root.OpenSubKey(@"Software\Classes\Excel.Application\CLSID")) excel |= key != null;
                }
            if (!framework) throw new InvalidOperationException("需要 .NET Framework 4.8。请先安装 Microsoft .NET Framework 4.8。");
            if (!excel) throw new InvalidOperationException("未检测到 Microsoft Excel 桌面版。请先安装 Office Excel。");
        }
        internal static RegistryView[] Views() { return Environment.Is64BitOperatingSystem ? new[] { RegistryView.Registry32, RegistryView.Registry64 } : new[] { RegistryView.Registry32 }; }
        internal static void Register(string directory, string testPrefix)
        {
            foreach (RegistryView view in Views())
                using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                    foreach (RegistrationEntry entry in RegistrationPlan.Create(directory, testPrefix))
                        using (RegistryKey key = root.CreateSubKey(entry.Path)) key.SetValue(entry.Name, entry.Value, entry.Kind);
        }
        internal static void Unregister(string testPrefix)
        {
            foreach (RegistryView view in Views())
                using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                    foreach (string path in RegistrationPlan.RemovalRoots(testPrefix)) root.DeleteSubKeyTree(path, false);
        }

        internal static void Install()
        {
            CheckEnvironment(); VerifyPayload();
            Directory.CreateDirectory(InstallPath);
            string dll = Path.Combine(InstallPath, "SheetPace.dll"), backup = dll + ".previous";
            bool existed = File.Exists(dll);
            if (existed) File.Copy(dll, backup, true);
            try
            {
                File.WriteAllBytes(dll, Payload());
                string uninstall = Path.Combine(InstallPath, "SheetPace.Uninstall.exe");
                if (!String.Equals(Assembly.GetExecutingAssembly().Location, uninstall, StringComparison.OrdinalIgnoreCase)) File.Copy(Assembly.GetExecutingAssembly().Location, uninstall, true);
                Register(InstallPath, "");
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(UninstallKey))
                {
                    key.SetValue("DisplayName", "SheetPace — Excel 筛选计数与行列定位"); key.SetValue("DisplayVersion", "1.0.0");
                    key.SetValue("Publisher", "SheetPace"); key.SetValue("InstallLocation", InstallPath);
                    key.SetValue("DisplayIcon", uninstall + ",0"); key.SetValue("UninstallString", "\"" + uninstall + "\" /uninstall");
                    key.SetValue("NoModify", 1, RegistryValueKind.DWord); key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                    key.SetValue("EstimatedSize", (int)((new FileInfo(uninstall).Length + new FileInfo(dll).Length) / 1024), RegistryValueKind.DWord);
                }
                File.WriteAllText(Path.Combine(InstallPath, "install-marker.txt"), Clsid);
                if (File.Exists(backup)) File.Delete(backup);
            }
            catch
            {
                if (existed) { File.Copy(backup, dll, true); Register(InstallPath, ""); }
                else { Unregister(""); if (File.Exists(dll)) File.Delete(dll); }
                throw;
            }
        }
        internal static void Uninstall()
        {
            if (Process.GetProcessesByName("EXCEL").Length != 0) throw new InvalidOperationException("请关闭所有 Excel 进程后再卸载。");
            Unregister(""); Registry.CurrentUser.DeleteSubKeyTree(UninstallKey, false);
            foreach (string name in new[] { "SheetPace.dll", "SheetPace.dll.previous", "install-marker.txt" })
            { string file = Path.Combine(InstallPath, name); if (File.Exists(file)) File.Delete(file); }
            string exe = Path.Combine(InstallPath, "SheetPace.Uninstall.exe");
            if (!String.Equals(Assembly.GetExecutingAssembly().Location, exe, StringComparison.OrdinalIgnoreCase))
            { if (File.Exists(exe)) File.Delete(exe); if (Directory.Exists(InstallPath)) Directory.Delete(InstallPath, false); }
        }
        internal static void ScheduleSelfCleanup()
        {
            if (!String.Equals(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), InstallPath, StringComparison.OrdinalIgnoreCase)) return;
            string target = Path.Combine(InstallPath, "SheetPace.Uninstall.exe");
            // A hidden native PowerShell cleanup uses literal paths; no recursive deletion or temp executable.
            string command = "Wait-Process -Id " + Process.GetCurrentProcess().Id + " -ErrorAction SilentlyContinue; " +
                "Remove-Item -LiteralPath '" + target.Replace("'", "''") + "' -Force -ErrorAction SilentlyContinue; " +
                "$cleanupDirectory='" + InstallPath.Replace("'", "''") + "'; if ([IO.Directory]::Exists($cleanupDirectory) -and [IO.Directory]::GetFileSystemEntries($cleanupDirectory).Length -eq 0) { [IO.Directory]::Delete($cleanupDirectory, $false) }";
            string encoded = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(command));
            string shell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"WindowsPowerShell\v1.0\powershell.exe");
            Process.Start(new ProcessStartInfo(shell, "-NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand " + encoded) { UseShellExecute = false, CreateNoWindow = true });
        }
        internal static string SelfTest(string output)
        {
            VerifyPayload(); Directory.CreateDirectory(output);
            string dll = Path.Combine(output, "SheetPace.dll"); File.WriteAllBytes(dll, Payload());
            Assembly assembly = Assembly.LoadFile(dll);
            if (assembly.GetType("SheetPace.Connect") == null) throw new Exception("COM 类不存在");
            System.Collections.Generic.List<RegistrationEntry> entries = RegistrationPlan.Create(output, "");
            System.Collections.Generic.List<string> plan = new System.Collections.Generic.List<string>();
            foreach (RegistrationEntry entry in entries)
            {
                if (!entry.Path.StartsWith(@"Software\", StringComparison.Ordinal)) throw new Exception("注册路径错误");
                plan.Add(entry.Path + " | " + entry.Name + " | " + entry.Kind + " | " + entry.Value);
            }
            if (!entries.Exists(e => e.Name == "LoadBehavior" && (int)e.Value == 3)) throw new Exception("加载配置不完整");
            if (!entries.Exists(e => e.Name == "CodeBase" && (string)e.Value == new Uri(dll).AbsoluteUri)) throw new Exception("CodeBase 不匹配");
            foreach (RegistrationEntry entry in entries)
                if (!Array.Exists(RegistrationPlan.RemovalRoots(""), r => entry.Path == r || entry.Path.StartsWith(r + (char)92))) throw new Exception("存在无法卸载的注册项");
            File.WriteAllLines(Path.Combine(output, "registration-plan.txt"), plan);
            string report = "PASS: embedded DLL SHA256, COM class, registration/removal plan (no registry writes)\r\nSHA256=" + BuildInfo.DllHash;
            File.WriteAllText(Path.Combine(output, "installer-test.txt"), report); return report;
        }
    }
}
