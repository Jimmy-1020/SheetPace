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
        internal static byte[] Payload() { return Resource("SheetPace.dll"); }
        internal static byte[] HelperPayload() { return Resource("SheetPace.NativeMenu.exe"); }
        internal static byte[] HostPayload() { return Resource("SheetPace.VisualHost.exe"); }
        private static byte[] Resource(string name)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            using (MemoryStream output = new MemoryStream())
            { if (stream == null) throw new InvalidDataException("安装包缺少插件文件。"); stream.CopyTo(output); return output.ToArray(); }
        }
        internal static string Hash(byte[] bytes) { using (SHA256 hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(); }
        internal static void VerifyPayload() { if (Hash(Payload()) != BuildInfo.DllHash || Hash(HelperPayload()) != BuildInfo.HelperHash || Hash(HostPayload()) != BuildInfo.HostHash) throw new InvalidDataException("插件文件校验失败，请重新获取安装包。"); }
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
            string helper = Path.Combine(InstallPath, "SheetPace.NativeMenu.exe"), helperBackup = helper + ".previous";
            string host = Path.Combine(InstallPath, "SheetPace.VisualHost.exe"), hostBackup = host + ".previous";
            bool hostExisted = File.Exists(host);
            if (hostExisted) File.Copy(host, hostBackup, true);
            bool helperExisted = File.Exists(helper);
            if (helperExisted) File.Copy(helper, helperBackup, true);
            bool existed = File.Exists(dll);
            if (existed) File.Copy(dll, backup, true);
            try
            {
                File.WriteAllBytes(dll, Payload());
                File.WriteAllBytes(helper, HelperPayload());
                File.WriteAllBytes(host, HostPayload());
                string uninstall = Path.Combine(InstallPath, "SheetPace.Uninstall.exe");
                if (!String.Equals(Assembly.GetExecutingAssembly().Location, uninstall, StringComparison.OrdinalIgnoreCase)) File.Copy(Assembly.GetExecutingAssembly().Location, uninstall, true);
                Register(InstallPath, "");
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(UninstallKey))
                {
                    key.SetValue("DisplayName", "SheetPace — Excel 筛选计数与行列定位"); key.SetValue("DisplayVersion", "1.0.1");
                    key.SetValue("Publisher", "SheetPace"); key.SetValue("InstallLocation", InstallPath);
                    key.SetValue("DisplayIcon", uninstall + ",0"); key.SetValue("UninstallString", "\"" + uninstall + "\" /uninstall");
                    key.SetValue("NoModify", 1, RegistryValueKind.DWord); key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                    key.SetValue("EstimatedSize", (int)((new FileInfo(uninstall).Length + new FileInfo(dll).Length + new FileInfo(helper).Length + new FileInfo(host).Length) / 1024), RegistryValueKind.DWord);
                }
                File.WriteAllText(Path.Combine(InstallPath, "install-marker.txt"), Clsid);
                RestoreDisabledAddin();
            }
            catch
            {
                if (hostExisted) File.Copy(hostBackup, host, true);
                else if (File.Exists(host)) File.Delete(host);
                if (helperExisted) File.Copy(helperBackup, helper, true);
                else if (File.Exists(helper)) File.Delete(helper);
                if (existed) { File.Copy(backup, dll, true); Register(InstallPath, ""); }
                else { Unregister(""); if (File.Exists(dll)) File.Delete(dll); }
                throw;
            }
            // Installation is committed; leftover backups must not trigger rollback.
            foreach (string previous in new[] { backup, helperBackup, hostBackup })
            {
                try { if (File.Exists(previous)) File.Delete(previous); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
        internal static void RestoreDisabledAddin()
        {
            // Excel keeps disabled DLLs across updates. Recover only our exact DLL + ProgID.
            string dll = Path.Combine(InstallPath, "SheetPace.dll");
            foreach (RegistryView view in Views())
                using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                    foreach (string version in new[] { "16.0", "15.0" })
                    {
                        string path = @"Software\Microsoft\Office\" + version + @"\Excel\Resiliency\DisabledItems";
                        using (RegistryKey key = root.OpenSubKey(path, true))
                        {
                            if (key == null) continue;
                            foreach (string name in key.GetValueNames())
                            {
                                byte[] data = key.GetValue(name) as byte[];
                                if (!MatchesDisabledAddin(data, dll)) continue;
                                string recovery = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SheetPace", "recovery");
                                Directory.CreateDirectory(recovery);
                                string file = Path.Combine(recovery, "disabled-" + version + "-" + view + "-" + Guid.NewGuid().ToString("N"));
                                File.WriteAllBytes(file + ".bin", data);
                                File.WriteAllText(file + ".txt", path + Environment.NewLine + name);
                                key.DeleteValue(name, false);
                            }
                        }
                    }
        }
        internal static bool MatchesDisabledAddin(byte[] data, string dll)
        {
            if (data == null) return false;
            string text = System.Text.Encoding.Unicode.GetString(data);
            return text.IndexOf(dll + (char)0, StringComparison.OrdinalIgnoreCase) >= 0 &&
                text.IndexOf(ProgId + (char)0, StringComparison.OrdinalIgnoreCase) >= 0;
        }
        internal static void Uninstall()
        {
            if (Process.GetProcessesByName("EXCEL").Length != 0) throw new InvalidOperationException("请关闭所有 Excel 进程后再卸载。");
            Unregister(""); Registry.CurrentUser.DeleteSubKeyTree(UninstallKey, false);
            foreach (string name in new[] { "SheetPace.dll", "SheetPace.dll.previous", "SheetPace.NativeMenu.exe", "SheetPace.NativeMenu.exe.previous", "SheetPace.VisualHost.exe", "SheetPace.VisualHost.exe.previous", "install-marker.txt" })
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
            string testDll = Path.Combine(output, "SheetPace.dll");
            byte[] ours = System.Text.Encoding.Unicode.GetBytes(testDll + (char)0 + ProgId + (char)0);
            byte[] otherPath = System.Text.Encoding.Unicode.GetBytes(testDll + ".other" + (char)0 + ProgId + (char)0);
            byte[] otherId = System.Text.Encoding.Unicode.GetBytes(testDll + (char)0 + ProgId + ".other" + (char)0);
            if (!MatchesDisabledAddin(ours, testDll) || MatchesDisabledAddin(otherPath, testDll) || MatchesDisabledAddin(otherId, testDll) || MatchesDisabledAddin(null, testDll))
                throw new Exception("禁用项目恢复匹配不正确");
            string dll = Path.Combine(output, "SheetPace.dll"); File.WriteAllBytes(dll, Payload());
            File.WriteAllBytes(Path.Combine(output, "SheetPace.NativeMenu.exe"), HelperPayload());
            File.WriteAllBytes(Path.Combine(output, "SheetPace.VisualHost.exe"), HostPayload());
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
            string report = "PASS: embedded DLL SHA256, COM class, registration/removal plan, targeted disabled-item matching (no registry writes)\r\nSHA256=" + BuildInfo.DllHash + Environment.NewLine + "HelperSHA256=" + BuildInfo.HelperHash + Environment.NewLine + "HostSHA256=" + BuildInfo.HostHash;
            File.WriteAllText(Path.Combine(output, "installer-test.txt"), report); return report;
        }
    }
}
