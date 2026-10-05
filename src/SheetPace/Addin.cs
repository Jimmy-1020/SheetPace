using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
namespace SheetPace
{
    [ComVisible(true), Guid("D19B1B2A-80CA-41D4-9DFD-3E82C147560B"), ProgId("SheetPace.Connect"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private dynamic app, ribbon;
        private Settings settings;
        private Timer timer;
        private OverlayWindow hover, counts;
        private NativeFilterWatcher watcher;
        private uint process;
        private bool inTick, dialogOpen, wasPressed, nativeOpen;
        private int tick, pendingRow, pendingColumn, lastRow, lastColumn, steadyTicks;
        private string lastSheet, hoverKey, geometryKey;
        private Rectangle lastRectangle;
        private ExcelContext countContext;
        private Point lastPoint, clickPoint;
        private bool clickPending;
        private IntPtr mouseHook;
        private NativeMethods.MouseHookProc mouseCallback;
        public string GetCustomUI(string ribbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
+<ribbon><tabs><tab id='sheetpaceTab' label='SheetPace'>
+<group id='sheetpaceFilter' label='计数筛选'><button id='showCountFilter' label='计数筛选' size='large' imageMso='Filter' onAction='OnShowFilter' screentip='显示本列各值的数量' supertip='选中目标列中的一个单元格，再打开搜索、多选和计数窗口。'/></group>
+<group id='sheetpacePointer' label='行列定位'><toggleButton id='hoverToggle' label='鼠标光影' size='large' imageMso='ConditionalFormattingHighlightCellsRules' onAction='OnToggleHover' getPressed='GetHoverPressed'/><button id='settings' label='颜色与透明度' size='large' imageMso='ColorPicker' onAction='OnSettings'/></group>
+<group id='sheetpaceHelp' label='使用帮助'><button id='help' label='使用说明' imageMso='Help' onAction='OnHelp'/></group>
+</tab></tabs></ribbon></customUI>".Replace("\n+", "\n");
        }
        public void OnConnection(object application, ConnectMode mode, object addInInst, ref Array custom)
        {
            app = application;
            try
            {
                settings = Settings.Load(Settings.DefaultPath); hover = new OverlayWindow(); counts = new OverlayWindow();
                IntPtr hwnd = new IntPtr((int)app.Hwnd); NativeMethods.GetWindowThreadProcessId(hwnd, out process);
                if (process == 0) process = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
                watcher = new NativeFilterWatcher(process);
                mouseCallback = MouseEvent; mouseHook = NativeMethods.SetWindowsHookEx(14, mouseCallback, NativeMethods.GetModuleHandle(null), 0);
                timer = new Timer { Interval = 90 }; timer.Tick += Tick; timer.Start();
                Log.Write("SheetPace 已连接 Excel " + Convert.ToString(app.Version), null);
            }
            catch (Exception ex) { Log.Write("加载项初始化失败", ex); Stop(); }
        }
        public void OnRibbonLoad(object ui) { ribbon = ui; }
        public bool GetHoverPressed(object control) { return settings == null || settings.HoverEnabled; }
        public void OnToggleHover(object control, bool pressed)
        {
            if (settings == null) return;
            settings.HoverEnabled = pressed; SaveSettings(); hoverKey = null;
            if (!pressed && hover != null) hover.Hide();
        }
        public void OnSettings(object control)
        {
            if (app == null || settings == null) return;
            dialogOpen = true; HideOverlays();
            try
            {
                using (SettingsForm form = new SettingsForm(settings))
                    if (form.ShowDialog(new WindowOwner(new IntPtr((int)app.Hwnd))) == DialogResult.OK)
                    {
                        settings = form.Result; SaveSettings(); hoverKey = null;
                        if (ribbon != null) ribbon.Invalidate();
                    }
            }
            catch (Exception ex) { Log.Write("设置窗口异常", ex); }
            finally { dialogOpen = false; }
        }
        private void SaveSettings()
        {
            try { settings.Save(Settings.DefaultPath); }
            catch (Exception ex) { MessageBox.Show("设置保存失败：" + ex.Message, "SheetPace"); }
        }
        public void OnShowFilter(object control)
        {
            if (app == null) return;
            dialogOpen = true; HideOverlays(); dynamic cell = null;
            try
            {
                cell = app.ActiveCell;
                using (ExcelContext context = ExcelContext.Capture((object)app, (int)cell.Row, (int)cell.Column))
                {
                    context.ReadSnapshot();
                    using (FilterForm form = new FilterForm(context)) form.ShowDialog(new WindowOwner(new IntPtr((int)app.Hwnd)));
                }
                DisposeContext();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "SheetPace · 计数筛选", MessageBoxButtons.OK, MessageBoxIcon.Information); }
            finally { ExcelContext.Release(cell); dialogOpen = false; hoverKey = null; }
        }
        public void OnHelp(object control)
        {
            MessageBox.Show("1. 打开列标题的筛选按钮，支持的原生列表会显示数量标签。\n2. 选中目标列中的一个格子，点击“计数筛选”，可搜索、多选并按数量筛选。\n3. 鼠标移动到单元格即可显示行列光影。\n4. 点击“颜色与透明度”调整外观。\n\n计数排除标题，包含筛选区域内所有隐藏行；空白单独计数。\n原生菜单增强依赖 Office 的可访问性支持；不能识别时请使用计数筛选窗口。\n\n版本 1.0.0 · Windows Excel 2016 及以上\n诊断日志：" + Settings.DirectoryPath, "SheetPace 使用说明");
        }
        private void Tick(object sender, EventArgs args)
        {
            if (inTick || app == null || dialogOpen) return;
            inTick = true; dynamic window = null, cell = null;
            try
            {
                if (!NativeMethods.IsProcessForeground(process)) { HideOverlays(); nativeOpen = false; return; }
                NativeMethods.POINT p; NativeMethods.GetCursorPos(out p); Point pointer = new Point(p.X, p.Y);
                bool pressed = (NativeMethods.GetAsyncKeyState(1) & 0x8000) != 0;
                if (pressed && !wasPressed && lastRow > 0) { pendingRow = lastRow; pendingColumn = lastColumn; }
                wasPressed = pressed;
                if (++tick % 3 == 0 && settings.NativeCountsEnabled) watcher.Request();
                FilterMenu menu = watcher.Latest;
                bool isMenu = menu != null && NativeMethods.IsWindow(menu.Window) && NativeMethods.IsWindowVisible(menu.Window) && unchecked(Environment.TickCount - (int)menu.Timestamp) < 1500;
                if (isMenu && settings.NativeCountsEnabled)
                {
                    hover.Hide(); hoverKey = null;
                    if (!nativeOpen) PrepareNativeCount();
                    nativeOpen = true; DrawCounts(menu); return;
                }
                nativeOpen = false; counts.Hide();
                IntPtr main = new IntPtr((int)app.Hwnd);
                // Do not draw over dialogs, ribbon, a different workbook window, or outside the cell grid.
                if (NativeMethods.GetForegroundWindow() != main) { hover.Hide(); hoverKey = null; return; }
                Rectangle grid = GridGeometry.FindGrid(main, pointer);
                if (grid.IsEmpty || pressed || (NativeMethods.GetAsyncKeyState(0x02) & 0x8000) != 0) { hover.Hide(); hoverKey = null; return; }
                window = app.ActiveWindow; cell = window.RangeFromPoint(pointer.X, pointer.Y);
                if (cell == null) { hover.Hide(); hoverKey = null; lastRow = 0; return; }
                int row = (int)cell.Row, column = (int)cell.Column;
                string sheet = (string)app.ActiveSheet.Name;
                bool moved = pointer != lastPoint || row != lastRow || column != lastColumn || sheet != lastSheet;
                steadyTicks = moved ? 0 : steadyTicks + 1;
                lastRow = row; lastColumn = column; lastSheet = sheet; lastPoint = pointer;
                pendingRow = row; pendingColumn = column;
                // Precompute while hovering a filter header, before Excel enters its popup loop.
                if (settings.NativeCountsEnabled && steadyTicks == 3) TryCountContext(row, column, false);
                if (!settings.HoverEnabled) { hover.Hide(); return; }
                string layout = main + ":" + sheet + ":" + row + ":" + column + ":" + grid + ":" + window.Zoom + ":" + window.ScrollRow + ":" + window.ScrollColumn;
                if (hoverKey == null || geometryKey != layout || !lastRectangle.Contains(pointer))
                { lastRectangle = GridGeometry.CellRectangle(window, cell, grid, pointer); geometryKey = layout; }
                Rectangle rect = lastRectangle;
                string key = sheet + ":" + row + ":" + column + ":" + grid + ":" + rect;
                if (hoverKey != key || !hover.Visible) { hover.DrawCross(grid, rect, settings); hoverKey = key; }
            }
            catch (COMException) { HideOverlays(); hoverKey = null; }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { HideOverlays(); hoverKey = null; }
            catch (Exception ex) { HideOverlays(); Log.Write("光影更新失败", ex); }
            finally { ExcelContext.Release(cell); ExcelContext.Release(window); inTick = false; }
        }
        private void TryCountContext(int row, int column, bool refresh)
        {
            ExcelContext next = null;
            try
            {
                if (!refresh && countContext != null && countContext.HeaderRow == row && countContext.Column == column && countContext.SheetName == (string)app.ActiveSheet.Name) return;
                next = ExcelContext.Capture((object)app, row, column);
                if (!refresh && row != next.HeaderRow) return;
                next.ReadSnapshot(); DisposeContext(); countContext = next; next = null;
            }
            catch { /* Excel may be editing or in a native menu. Keep a prewarmed snapshot. */ }
            finally { if (next != null) next.Dispose(); }
        }
        private IntPtr MouseEvent(int code, IntPtr message, IntPtr data)
        {
            if (code >= 0 && message.ToInt32() == 0x201 && NativeMethods.IsProcessForeground(process))
            {
                NativeMethods.MOUSEHOOK info = (NativeMethods.MOUSEHOOK)Marshal.PtrToStructure(data, typeof(NativeMethods.MOUSEHOOK));
                clickPoint = new Point(info.Point.X, info.Point.Y); clickPending = true;
            }
            return NativeMethods.CallNextHookEx(mouseHook, code, message, data);
        }
        private void PrepareNativeCount()
        {
            dynamic cell = null, window = null;
            try
            {
                if (clickPending) { window = app.ActiveWindow; cell = window.RangeFromPoint(clickPoint.X, clickPoint.Y); }
                if (cell == null) cell = app.ActiveCell;
                if (cell != null) TryCountContext((int)cell.Row, (int)cell.Column, true);
                if (countContext != null && !countContext.MatchesActive()) DisposeContext();
            }
            catch { if (countContext != null && (countContext.Column != pendingColumn || countContext.SheetName != lastSheet)) DisposeContext(); }
            finally { clickPending = false; ExcelContext.Release(cell); ExcelContext.Release(window); }
        }
        private void DrawCounts(FilterMenu menu)
        {
            if (countContext == null || countContext.Snapshot == null) { counts.Hide(); return; }
            counts.Render(menu.Bounds, delegate(Graphics graphics)
            {
                using (Font font = new Font("Microsoft YaHei UI", 8.5F))
                using (Brush background = new SolidBrush(Color.FromArgb(245, 242, 247, 253)))
                using (Brush ink = new SolidBrush(Color.FromArgb(255, 36, 87, 155)))
                using (StringFormat format = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Center })
                {
                    foreach (FilterRow row in menu.Rows)
                    {
                        int value;
                        if (!countContext.Snapshot.TryGetCount(row.Name, row.Year, row.Month, out value)) continue;
                        Rectangle tag = new Rectangle(menu.Bounds.Width - 70, row.Bounds.Top - menu.Bounds.Top + 1, 52, Math.Max(15, row.Bounds.Height - 2));
                        if (tag.Top < 0 || tag.Bottom > menu.Bounds.Height) continue;
                        graphics.FillRectangle(background, tag);
                        graphics.DrawString("(" + value.ToString("N0") + ")", font, ink, tag, format);
                    }
                }
            });
        }
        private void HideOverlays() { if (hover != null) hover.Hide(); if (counts != null) counts.Hide(); }
        private void DisposeContext() { if (countContext != null) { countContext.Dispose(); countContext = null; } }
        private void Stop()
        {
            if (mouseHook != IntPtr.Zero) { NativeMethods.UnhookWindowsHookEx(mouseHook); mouseHook = IntPtr.Zero; }
            if (timer != null) { timer.Stop(); timer.Dispose(); timer = null; }
            if (watcher != null) { watcher.Dispose(); watcher = null; }
            if (hover != null) { hover.Dispose(); hover = null; }
            if (counts != null) { counts.Dispose(); counts = null; }
            DisposeContext(); ribbon = null; app = null;
        }
        public void OnDisconnection(DisconnectMode mode, ref Array custom) { Stop(); }
        public void OnBeginShutdown(ref Array custom) { Stop(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
    }
}
