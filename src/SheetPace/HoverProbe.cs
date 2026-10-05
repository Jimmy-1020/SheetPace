using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
namespace SheetPace
{
    // A native message-only timer runs on Excel's STA. No forms or UI Automation.
    internal sealed class HoverProbe : IDisposable
    {
        [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate IntPtr WindowProc(IntPtr hwnd, uint message, IntPtr wparam, IntPtr lparam);
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] private struct WindowClass
        {
            public uint Style; public WindowProc Procedure; public int ClassExtra, WindowExtra;
            public IntPtr Instance, Icon, Cursor, Background;
            public string MenuName, ClassName;
        }
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern ushort RegisterClassW(ref WindowClass value);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool UnregisterClassW(string name, IntPtr instance);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr CreateWindowExW(uint ex, string cls, string title, uint style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);
        [DllImport("user32.dll")] private static extern IntPtr DefWindowProcW(IntPtr hwnd, uint msg, IntPtr wp, IntPtr lp);
        [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern UIntPtr SetTimer(IntPtr hwnd, UIntPtr id, uint interval, IntPtr callback);
        [DllImport("user32.dll")] private static extern bool KillTimer(IntPtr hwnd, UIntPtr id);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandleW(string name);
        private readonly WindowProc procedure;
        private readonly string className = "SheetPace.HoverProbe." + Guid.NewGuid().ToString("N");
        private readonly IntPtr instance = GetModuleHandleW(null);
        private readonly Func<bool> enabled;
        private readonly Action refreshSettings;
        private int settingsStamp;
        private readonly HoverPacket packet;
        private dynamic app;
        private IntPtr window;
        private bool registered, stopping, busy;
        private string geometryKey, headerKey;
        private Rectangle rectangle;
        private int headerStamp;
        private readonly List<int[]> headers = new List<int[]>();
        private readonly HoverFrame latest = new HoverFrame();
        internal HoverProbe(object application, uint process, Func<bool> active, Action refresh)
        {
            app = application; enabled = active; refreshSettings = refresh; procedure = Dispatch;
            packet = new HoverPacket(process, true);
            try
            {
                WindowClass cls = new WindowClass { Procedure = procedure, Instance = instance, ClassName = className };
                if (RegisterClassW(ref cls) == 0) throw new System.ComponentModel.Win32Exception(); registered = true;
                window = CreateWindowExW(0, className, "", 0, 0, 0, 0, 0, new IntPtr(-3), IntPtr.Zero, instance, IntPtr.Zero);
                if (window == IntPtr.Zero || SetTimer(window, new UIntPtr(1), 33, IntPtr.Zero) == UIntPtr.Zero) throw new System.ComponentModel.Win32Exception();
            }
            catch { Dispose(); throw; }
        }
        private IntPtr Dispatch(IntPtr hwnd, uint message, IntPtr wp, IntPtr lp)
        {
            if (message == 0x113) { if (!stopping && !busy) Sample(); return IntPtr.Zero; }
            return DefWindowProcW(hwnd, message, wp, lp);
        }
        private void Sample()
        {
            busy = true; Stopwatch watch = Stopwatch.StartNew(); dynamic view = null, cell = null, sheet = null;
            try
            {
                latest.Visible = false;
                if (unchecked(Environment.TickCount - settingsStamp) >= 1000) { settingsStamp = Environment.TickCount; refreshSettings(); }
                if (!enabled()) return;
                IntPtr main = new IntPtr((int)app.Hwnd);
                if (NativeMethods.GetForegroundWindow() != main) { geometryKey = null; return; }
                NativeMethods.POINT cursor; NativeMethods.GetCursorPos(out cursor); Point pointer = new Point(cursor.X, cursor.Y);
                if ((NativeMethods.GetAsyncKeyState(1) & 0x8000) != 0 || (NativeMethods.GetAsyncKeyState(2) & 0x8000) != 0) { geometryKey = null; return; }
                Rectangle grid = GridGeometry.FindGrid(main, pointer); if (grid.IsEmpty) return;
                view = app.ActiveWindow; cell = view.RangeFromPoint(pointer.X, pointer.Y); if (cell == null) return;
                sheet = app.ActiveSheet; string name = (string)sheet.Name;
                int row = (int)cell.Row, column = (int)cell.Column;
                string layout = main + ":" + name + ":" + row + ":" + column + ":" + grid + ":" + view.Zoom + ":" + view.ScrollRow + ":" + view.ScrollColumn;
                if (geometryKey != layout || !rectangle.Contains(pointer))
                { rectangle = GridGeometry.CellRectangle(view, cell, grid, pointer); geometryKey = layout; }
                string key = main + ":" + name;
                if (headerKey != key || unchecked(Environment.TickCount - headerStamp) > 1000) RefreshHeaders(sheet, key);
                bool header = false;
                foreach (int[] range in headers) if (row == range[0] && column >= range[1] && column <= range[2]) { header = true; break; }
                latest.Visible = true; latest.Window = main; latest.Pointer = pointer; latest.Row = row; latest.Column = column;
                latest.SheetName = name; latest.Grid = grid; latest.Cell = rectangle; latest.FilterHeader = header;
            }
            catch { geometryKey = null; latest.Visible = false; }
            finally
            {
                ExcelContext.Release(sheet); ExcelContext.Release(cell); ExcelContext.Release(view);
                watch.Stop(); latest.ElapsedMicroseconds = (int)Math.Min(Int32.MaxValue, watch.ElapsedTicks * 1000000L / Stopwatch.Frequency);
                latest.Timestamp = Environment.TickCount;
                try { packet.Publish(latest); } catch { }
                busy = false;
            }
        }
        private void AddHeader(dynamic source)
        {
            dynamic columns = null;
            try { columns = source.Columns; int column = (int)source.Column; headers.Add(new[] { (int)source.Row, column, column + (int)columns.Count - 1 }); }
            finally { ExcelContext.Release(columns); }
        }
        private void RefreshHeaders(dynamic sheet, string key)
        {
            headers.Clear(); dynamic filter = null, source = null, tables = null;
            try
            {
                if ((bool)sheet.AutoFilterMode)
                { filter = sheet.AutoFilter; source = filter.Range; AddHeader(source); ExcelContext.Release(source); source = null; }
                tables = sheet.ListObjects;
                for (int i = 1; i <= (int)tables.Count; i++)
                {
                    dynamic table = null;
                    try { table = tables.Item(i); source = table.Range; AddHeader(source); }
                    finally { ExcelContext.Release(source); source = null; ExcelContext.Release(table); }
                }
                headerKey = key; headerStamp = Environment.TickCount;
            }
            finally { ExcelContext.Release(source); ExcelContext.Release(filter); ExcelContext.Release(tables); }
        }
        public void Dispose()
        {
            if (stopping) return; stopping = true;
            if (window != IntPtr.Zero) { KillTimer(window, new UIntPtr(1)); DestroyWindow(window); window = IntPtr.Zero; }
            if (registered) { UnregisterClassW(className, instance); registered = false; }
            packet.Dispose(); app = null; GC.KeepAlive(procedure);
        }
    }
}
