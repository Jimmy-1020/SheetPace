using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
namespace SheetPace
{
    public static class VisualHostService
    {
        [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object result);
        [DllImport("user32.dll")] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
        public static int Run(IntPtr excel, uint process)
        {
            SetThreadDpiAwarenessContext(new IntPtr(-4)); Application.EnableVisualStyles();
            Connect runtime = new Connect(); dynamic app = null; Array custom = new object[0];
            using (ApplicationContext context = new ApplicationContext())
            using (HostWindow dispatcher = new HostWindow(runtime, context))
            using (System.Windows.Forms.Timer watchdog = new System.Windows.Forms.Timer { Interval = 700 })
            {
                IntPtr commandWindow = dispatcher.Handle;
                try
                {
                    for (int attempt = 0; attempt < 100 && app == null; attempt++)
                    {
                        if (!NativeMethods.IsWindow(excel)) return 1;
                        IntPtr grid = IntPtr.Zero;
                        NativeMethods.EnumChildWindows(excel, delegate(IntPtr window, IntPtr unused)
                        { if (NativeMethods.ClassName(window) == "EXCEL7") { grid = window; return false; } return true; }, IntPtr.Zero);
                        if (grid != IntPtr.Zero)
                        {
                            Guid iid = new Guid("00020400-0000-0000-C000-000000000046"); object result = null;
                            try
                            {
                                if (AccessibleObjectFromWindow(grid, 0xFFFFFFF0, ref iid, out result) == 0)
                                { dynamic native = result; app = native.Application; }
                            }
                            catch (COMException) { }
                            finally { ExcelContext.Release(result); }
                        }
                        if (app == null) Thread.Sleep(100);
                    }
                    if (app == null) throw new InvalidOperationException("无法连接指定 Excel 窗口");
                    runtime.OnConnection((object)app, ConnectMode.External, null, ref custom);
                    watchdog.Tick += delegate
                    {
                        try { using (Process parent = Process.GetProcessById((int)process)) if (parent.HasExited) dispatcher.Close(); }
                        catch { dispatcher.Close(); }
                    };
                    watchdog.Start(); Application.Run(context);
                    return 0;
                }
                catch (Exception ex) { Log.Write("界面进程启动失败", ex); return 1; }
                finally
                {
                    watchdog.Stop(); runtime.OnDisconnection(DisconnectMode.UserClosed, ref custom);
                    ExcelContext.Release((object)app); app = null; GC.Collect(); GC.WaitForPendingFinalizers();
                }
            }
        }
        private sealed class HostWindow : Form
        {
            private readonly Connect runtime;
            private readonly ApplicationContext context;
            internal HostWindow(Connect instance, ApplicationContext loop)
            { runtime = instance; context = loop; Text = "SheetPace Visual Host"; ShowInTaskbar = false; FormBorderStyle = FormBorderStyle.None; }
            protected override void WndProc(ref Message message)
            {
                if (message.Msg >= 0x8001 && message.Msg <= 0x8004)
                {
                    switch (message.Msg)
                    {
                        case 0x8001: runtime.OnShowFilter(null); break;
                        case 0x8002: runtime.OnSettings(null); break;
                        case 0x8003: runtime.OnHelp(null); break;
                        case 0x8004: runtime.OnToggleHover(null, message.LParam != IntPtr.Zero); break;
                    }
                    return;
                }
                base.WndProc(ref message);
            }
            protected override void OnFormClosed(FormClosedEventArgs args)
            { context.ExitThread(); base.OnFormClosed(args); }
        }
    }
}
