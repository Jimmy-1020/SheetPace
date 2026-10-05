using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
namespace SheetPace
{
    internal sealed class VisualHostBridge : IDisposable
    {
        private Process child;
        public int ProcessId { get { return child == null ? 0 : child.Id; } }
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int length);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wparam, IntPtr lparam);
        internal VisualHostBridge(IntPtr excel, uint process)
        {
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SheetPace.VisualHost.exe");
            child = Process.Start(new ProcessStartInfo(path, excel.ToInt64() + " " + process)
            { UseShellExecute = false, CreateNoWindow = true });
        }
        private IntPtr Window()
        {
            Process target = child; if (target == null || target.HasExited) return IntPtr.Zero;
            IntPtr found = IntPtr.Zero;
            NativeMethods.EnumWindows(delegate(IntPtr hwnd, IntPtr unused)
            {
                uint process; NativeMethods.GetWindowThreadProcessId(hwnd, out process);
                if (process != target.Id) return true;
                StringBuilder title = new StringBuilder(128); GetWindowText(hwnd, title, title.Capacity);
                if (title.ToString() != "SheetPace Visual Host") return true;
                found = hwnd; return false;
            }, IntPtr.Zero);
            return found;
        }
        internal void Send(int command, int value)
        {
            try
            {
                IntPtr window = Window();
                if (window != IntPtr.Zero) PostMessage(window, 0x8000 + (uint)command, IntPtr.Zero, new IntPtr(value));
                else Log.Write("界面进程尚未就绪或已退出", null);
            }
            catch (Exception ex) { Log.Write("界面操作传递失败", ex); }
        }
        public void Dispose()
        {
            Process target = child; if (target == null) return;
            try
            {
                IntPtr window = Window();
                if (window != IntPtr.Zero) PostMessage(window, 0x10, IntPtr.Zero, IntPtr.Zero);
                // The STA join pumps COM so the visual host can release Excel proxies.
                Thread waiter = new Thread(delegate()
                {
                    try { if (!target.WaitForExit(4000)) target.Kill(); } catch { }
                }) { IsBackground = true };
                waiter.Start(); waiter.Join(4500);
            }
            catch (Exception ex) { Log.Write("界面进程停止失败", ex); }
            finally { child = null; target.Dispose(); }
        }
    }
}
