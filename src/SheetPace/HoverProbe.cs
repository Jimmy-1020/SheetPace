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
        private readonly Func<Settings> currentSettings;
        private readonly Action refreshSettings;
        private int settingsStamp;
        private readonly HoverPacket packet;
        private dynamic app;
        private IntPtr window;
        private bool registered, stopping, busy, failureLogged;
        private string geometryKey, headerKey;
        private Rectangle rectangle, targetGrid;
        private HighlightBands bands = new HighlightBands();
        private SelectionSnapshot selectionSnapshot;
        private string selectionAddress;
        private int geometryStamp;
        private int headerStamp;
        private readonly List<int[]> headers = new List<int[]>();
        private readonly HoverFrame latest = new HoverFrame();
        internal HoverProbe(object application, uint process, Func<Settings> active, Action refresh)
        {
            app = application; currentSettings = active; refreshSettings = refresh; procedure = Dispatch;
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
            busy = true; Stopwatch watch = Stopwatch.StartNew(); dynamic view = null, cell = null, selected = null, active = null, targetArea = null, sheet = null;
            try
            {
                latest.Visible = latest.PointerValid = latest.FilterHeader = false;
                latest.RowBands = latest.ColumnBands = new Rectangle[0];
                latest.Row = latest.Column = latest.TargetRow = latest.TargetColumn = 0;
                if (unchecked(Environment.TickCount - settingsStamp) >= 1000) { settingsStamp = Environment.TickCount; refreshSettings(); }
                Settings settings = currentSettings();
                if (settings == null || (!settings.HoverEnabled && !settings.NativeCountsEnabled)) return;
                latest.Mode = settings.Mode;
                IntPtr main = new IntPtr((int)app.Hwnd);
                if (NativeMethods.GetForegroundWindow() != main) { geometryKey = null; return; }
                NativeMethods.POINT cursor; NativeMethods.GetCursorPos(out cursor); Point pointer = new Point(cursor.X, cursor.Y);
                if ((NativeMethods.GetAsyncKeyState(1) & 0x8000) != 0 || (NativeMethods.GetAsyncKeyState(2) & 0x8000) != 0) { geometryKey = null; return; }
                view = app.ActiveWindow; sheet = app.ActiveSheet; string name = (string)sheet.Name;
                latest.Window = main; latest.Pointer = pointer; latest.SheetName = name;
                Rectangle pointerGrid = GridGeometry.FindGrid(main, pointer);
                if (!pointerGrid.IsEmpty) cell = view.RangeFromPoint(pointer.X, pointer.Y);
                if (cell != null)
                {
                    latest.PointerValid = true; latest.Row = (int)cell.Row; latest.Column = (int)cell.Column;
                    if (settings.NativeCountsEnabled)
                    {
                        string key = main + ":" + view.Caption + ":" + name;
                        if (headerKey != key || unchecked(Environment.TickCount - headerStamp) > 1000) RefreshHeaders(sheet, key);
                        foreach (int[] range in headers)
                            if (latest.Row == range[0] && latest.Column >= range[1] && latest.Column <= range[2]) { latest.FilterHeader = true; break; }
                    }
                }
                if (!settings.HoverEnabled) return;
                dynamic target = cell; string address = "";
                if (settings.Mode == HighlightMode.ClickSelection)
                { selected = app.Selection; active = app.ActiveCell; target = active; address = (string)selected.Address; }
                if (target == null) return;
                targetArea = target.MergeArea; latest.TargetRow = (int)targetArea.Row; latest.TargetColumn = (int)targetArea.Column;
                List<Rectangle> grids = settings.Mode == HighlightMode.ClickSelection ? GridGeometry.FindGrids(main) : new List<Rectangle> { pointerGrid };
                string selectionIdentity = main + ":" + view.Caption + ":" + name + ":" + address + ":" + targetArea.Address;
                string layout = selectionIdentity + ":" + settings.Mode + ":" + latest.TargetRow + ":" + latest.TargetColumn + ":" + String.Join(";", grids)
                    + ":" + view.Zoom + ":" + view.ScrollRow + ":" + view.ScrollColumn + ":" + view.SplitRow + ":" + view.SplitColumn;
                bool refresh = settings.Mode == HighlightMode.ClickSelection && unchecked(Environment.TickCount - geometryStamp) >= 1000;
                if (geometryKey != layout || refresh || (settings.Mode == HighlightMode.FollowMouse && !rectangle.Contains(pointer)))
                {
                    HighlightBands next;
                    if (settings.Mode == HighlightMode.ClickSelection)
                    {
                        if (selectionSnapshot == null || selectionAddress != selectionIdentity)
                        { selectionSnapshot = SelectionSnapshot.Capture(selected); selectionAddress = selectionIdentity; }
                        next = SelectionGeometry.Project(view, grids, selectionSnapshot); targetGrid = next.Grid;
                        if (next.Rows.Length == 1 && next.Columns.Length == 1) rectangle = Rectangle.Intersect(next.Rows[0], next.Columns[0]);
                        else rectangle = GridGeometry.SelectionRectangle(view, target, grids, pointer, out pointerGrid);
                    }
                    else
                    {
                        targetGrid = pointerGrid; rectangle = GridGeometry.CellRectangle(view, target, pointerGrid, pointer);
                        Rectangle clipped = Rectangle.Intersect(targetGrid, rectangle);
                        next = new HighlightBands { Grid = targetGrid,
                            Rows = new[] { new Rectangle(targetGrid.Left, clipped.Top, targetGrid.Width, clipped.Height) },
                            Columns = new[] { new Rectangle(clipped.Left, targetGrid.Top, clipped.Width, targetGrid.Height) } };
                    }
                    if (next.Grid != bands.Grid || !Same(next.Rows, bands.Rows) || !Same(next.Columns, bands.Columns)) latest.ShapeRevision = unchecked(latest.ShapeRevision + 1);
                    bands = next; geometryKey = layout; geometryStamp = Environment.TickCount;
                }
                latest.Grid = targetGrid; latest.Cell = rectangle; latest.RowBands = bands.Rows; latest.ColumnBands = bands.Columns; latest.Visible = bands.Visible; failureLogged = false;
            }
            catch (Exception ex) { geometryKey = null; latest.Visible = false; if (!failureLogged) { Log.Write("选区坐标采样失败", ex); failureLogged = true; } }
            finally
            {
                ExcelContext.Release(targetArea); ExcelContext.Release(active); ExcelContext.Release(selected); ExcelContext.Release(sheet); ExcelContext.Release(cell); ExcelContext.Release(view);
                watch.Stop(); latest.ElapsedMicroseconds = (int)Math.Min(Int32.MaxValue, watch.ElapsedTicks * 1000000L / Stopwatch.Frequency);
                latest.Timestamp = Environment.TickCount;
                try { packet.Publish(latest); } catch { }
                busy = false;
            }
        }
        private static bool Same(Rectangle[] first, Rectangle[] second)
        {
            if (first.Length != second.Length) return false;
            for (int i = 0; i < first.Length; i++) if (first[i] != second[i]) return false;
            return true;
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
