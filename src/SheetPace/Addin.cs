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
        private CrossOverlay hover;
        private OverlayWindow counts;
        private HoverProbe hoverProbe;
        private HoverPacket hoverPacket;
        private string lastCountKey;
        private NativeFilterWatcher watcher;
        private uint process;
        private bool inTick, dialogOpen, wasPressed, nativeOpen;
        private int tick, pendingRow, pendingColumn, lastRow, lastColumn, steadyTicks, clickStamp;
        private string lastSheet, hoverKey, geometryKey;
        private Rectangle lastRectangle;
        private ExcelContext countContext;
        private Point lastPoint, clickPoint;
        private bool clickPending, bridge;
        private VisualHostBridge visualHost;
        private DateTime settingsStamp;
        private RibbonImages images;
        public int VisualHostProcessId { get { return visualHost == null ? 0 : visualHost.ProcessId; } }
        public string GetCustomUI(string ribbonID)
        {
            Log.Write("PID=" + System.Diagnostics.Process.GetCurrentProcess().Id + " Ribbon GetCustomUI", null);
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
+<ribbon><tabs><tab id='sheetpaceTab' label='SheetPace'>
+<group id='sheetpaceFilter' label='计数筛选'><button id='showCountFilter' label='计数筛选' size='large' imageMso='Filter' onAction='OnShowFilter' screentip='显示本列各值的数量' supertip='选中目标列中的一个单元格，再打开搜索、多选和计数窗口。'/></group>
+<group id='sheetpacePointer' label='行列定位'><toggleButton id='hoverToggle' label='鼠标光影' size='large' getImage='GetHoverImage' screentip='显示选中单元格或鼠标所在格的行列光影' onAction='OnToggleHover' getPressed='GetHoverPressed'/><button id='settings' label='光影设置' size='large' getImage='GetSettingsImage' screentip='设置光影模式、颜色与透明度' onAction='OnSettings'/></group>
+<group id='sheetpaceHelp' label='使用帮助'><button id='help' label='使用说明' imageMso='Help' onAction='OnHelp'/></group>
+</tab></tabs></ribbon></customUI>".Replace("\n+", "\n");
        }
        public void OnConnection(object application, ConnectMode mode, object addInInst, ref Array custom)
        {
            Log.Write("PID=" + System.Diagnostics.Process.GetCurrentProcess().Id + " OnConnection " + mode, null);
            app = application;
            try
            {
                settings = Settings.Load(Settings.DefaultPath);
                IntPtr hwnd = new IntPtr((int)app.Hwnd); NativeMethods.GetWindowThreadProcessId(hwnd, out process);
                if (process == 0) throw new InvalidOperationException("Excel 窗口不可用");
                bridge = process == (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
                if (bridge)
                {
                    if (addInInst != null) { dynamic instance = addInInst; instance.Object = this; }
                    hoverProbe = new HoverProbe(application, process, delegate { return settings; }, delegate { RefreshRibbonSettings(null); });
                    visualHost = new VisualHostBridge(hwnd, process);
                    settingsStamp = System.IO.File.Exists(Settings.DefaultPath) ? System.IO.File.GetLastWriteTimeUtc(Settings.DefaultPath) : DateTime.MinValue;
                    Log.Write("PID=" + process + " VisualHost=" + visualHost.ProcessId, null);
                    return;
                }
                // All windows and UI Automation live in the standalone STA visual host.
                hover = new CrossOverlay { Text = "SheetPace Hover Overlay" };
                try { hoverPacket = new HoverPacket(process, false); }
                catch (System.IO.FileNotFoundException) { /* Direct test connections can use the legacy geometry path. */ }
                counts = new OverlayWindow { Text = "SheetPace Count Overlay" };
                watcher = new NativeFilterWatcher(process);
                timer = new Timer { Interval = hoverPacket == null ? 90 : 33 }; timer.Tick += Tick; timer.Start();
                Log.Write("SheetPace 已连接 Excel " + Convert.ToString(app.Version), null);
            }
            catch (Exception ex) { Log.Write("加载项初始化失败", ex); Stop(); }
        }
        private void RefreshRibbonSettings(object unused)
        {
            try
            {
                DateTime stamp = System.IO.File.Exists(Settings.DefaultPath) ? System.IO.File.GetLastWriteTimeUtc(Settings.DefaultPath) : DateTime.MinValue;
                if (stamp == settingsStamp) return;
                settingsStamp = stamp; settings = Settings.Load(Settings.DefaultPath);
                if (ribbon != null) ribbon.Invalidate();
            }
            catch { /* Office can be busy or closing. */ }
        }
        [return: MarshalAs(UnmanagedType.IDispatch)]
        public object GetHoverImage(object control) { if (images == null) images = new RibbonImages(); return images.Hover; }
        [return: MarshalAs(UnmanagedType.IDispatch)]
        public object GetSettingsImage(object control) { if (images == null) images = new RibbonImages(); return images.Settings; }
        public int GetHighlightMode(object control) { if (bridge) settings = Settings.Load(Settings.DefaultPath); return (int)(settings == null ? HighlightMode.ClickSelection : settings.Mode); }
        public void OnSetMode(object control, int mode)
        {
            if (!Enum.IsDefined(typeof(HighlightMode), mode)) return;
            if (bridge) { if (visualHost != null) visualHost.Send(5, mode); return; }
            if (settings == null) return;
            settings.Mode = (HighlightMode)mode; SaveSettings(); hoverKey = null;
        }
        public void OnRibbonLoad(object ui) { ribbon = ui; }
        public bool GetHoverPressed(object control) { if (bridge) settings = Settings.Load(Settings.DefaultPath); return settings == null || settings.HoverEnabled; }
        public void OnToggleHover(object control, bool pressed)
        {
            if (bridge) { if (visualHost != null) visualHost.Send(4, pressed ? 1 : 0); return; }
            if (settings == null) return;
            settings.HoverEnabled = pressed; SaveSettings(); hoverKey = null;
            if (!pressed && hover != null) hover.Hide();
        }
        public void OnSettings(object control)
        {
            if (bridge) { if (visualHost != null) visualHost.Send(2, 0); return; }
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
            if (bridge) { if (visualHost != null) visualHost.Send(1, 0); return; }
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
            if (bridge) { if (visualHost != null) visualHost.Send(3, 0); return; }
            MessageBox.Show("1. 打开列标题的筛选按钮，支持的原生列表会显示数量标签。\n2. 选中目标列中的一个格子，点击“计数筛选”，可搜索、多选并按数量筛选。\n3. 默认点击选中单元格显示十字光影。\n4. 点击“光影设置”切换鼠标跟随，或调整颜色与透明度。\n\n计数排除标题，包含筛选区域内所有隐藏行；空白单独计数。\n原生菜单增强依赖 Office 的可访问性支持；不能识别时请使用计数筛选窗口。\n\n版本 1.0.3 · Windows Excel 2016 及以上\n诊断日志：" + Settings.DirectoryPath, "SheetPace 使用说明");
        }
        private void Tick(object sender, EventArgs args)
        {
            if (inTick || app == null || dialogOpen) return;
            inTick = true; dynamic window = null, cell = null;
            try
            {
                if (!NativeMethods.IsProcessForeground(process)) { HideOverlays(); nativeOpen = false; return; }
                NativeMethods.POINT p; NativeMethods.GetCursorPos(out p); Point pointer = new Point(p.X, p.Y);
                short buttonState = NativeMethods.GetAsyncKeyState(1);
                bool pressed = (buttonState & 0x8000) != 0;
                // Poll input on Excel's timer; native callback pointers cannot outlive shutdown.
                if ((buttonState & 1) != 0 || (pressed && !wasPressed)) { clickPoint = pointer; clickPending = true; clickStamp = Environment.TickCount; }
                if (pressed && !wasPressed && lastRow > 0) { pendingRow = lastRow; pendingColumn = lastColumn; }
                wasPressed = pressed;
                if (++tick % (hoverPacket == null ? 3 : 8) == 0 && settings.NativeCountsEnabled && watcher != null) watcher.Request();
                FilterMenu menu = watcher == null ? null : watcher.Latest;
                bool isMenu = menu != null && NativeMethods.IsWindow(menu.Window) && NativeMethods.IsWindowVisible(menu.Window) && unchecked(Environment.TickCount - (int)menu.Timestamp) < 1500;
                if (isMenu && settings.NativeCountsEnabled)
                {
                    hover.Hide(); hoverKey = null;
                    if (!nativeOpen) { PrepareNativeCount(menu); TraceNativeCount(menu); }
                    nativeOpen = true; DrawCounts(menu); return;
                }
                nativeOpen = false; counts.Hide(); lastCountKey = null;
                if (hoverPacket != null)
                {
                    HoverFrame frame = hoverPacket.Read();
                    if (frame == null || unchecked(Environment.TickCount - frame.Timestamp) > 250)
                    { hover.Hide(); hoverKey = null; return; }
                    if (frame.PointerValid)
                    {
                        bool movedFrame = frame.Pointer != lastPoint || frame.Row != lastRow || frame.Column != lastColumn || frame.SheetName != lastSheet;
                        steadyTicks = movedFrame ? 0 : steadyTicks + 1;
                        lastRow = frame.Row; lastColumn = frame.Column; lastSheet = frame.SheetName; lastPoint = frame.Pointer;
                        pendingRow = frame.Row; pendingColumn = frame.Column;
                        if (settings.NativeCountsEnabled && frame.FilterHeader && (steadyTicks == 0 || steadyTicks == 3)) TryCountContext(frame.Row, frame.Column, false);
                    }
                    if (!settings.HoverEnabled || !frame.Visible) { hover.Hide(); hoverKey = null; return; }
                    string packetKey = frame.Window + ":" + frame.SheetName + ":" + frame.Grid + ":" + frame.Cell + ":" + settings.Alpha + ":" + settings.HighlightColor.ToArgb();
                    if (hoverKey != packetKey || !hover.Visible) { hover.DrawCross(frame.Grid, frame.Cell, settings); hoverKey = packetKey; }
                    return;
                }
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
                if (settings.NativeCountsEnabled && (steadyTicks == 0 || steadyTicks == 3)) TryCountContext(row, column, false);
                if (!settings.HoverEnabled) { hover.Hide(); return; }
                if (settings.Mode == HighlightMode.ClickSelection)
                {
                    dynamic selected = null;
                    try
                    {
                        selected = app.ActiveCell; Rectangle selectedGrid;
                        Rectangle selectedRect = GridGeometry.SelectionRectangle(window, selected, GridGeometry.FindGrids(main), pointer, out selectedGrid);
                        hover.DrawCross(selectedGrid, selectedRect, settings);
                    }
                    finally { ExcelContext.Release(selected); }
                    return;
                }

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
        private void PrepareNativeCount(FilterMenu menu)
        {
            dynamic cell = null, window = null;
            try
            {
                // Opening a filter arrow does not move ActiveCell. Keep the hovered
                // header when the popup is anchored there, even if hit testing is blocked
                // by Excel's popup loop. Keyboard menus at another header use ActiveCell.
                bool atHeader = countContext != null && countContext.HeaderRow == lastRow &&
                    countContext.Column == lastColumn && countContext.SheetName == lastSheet;
                bool besidePointer = Math.Min(Math.Abs(menu.Bounds.Left - lastPoint.X), Math.Abs(menu.Bounds.Right - lastPoint.X)) <= 32;
                bool belowPointer = menu.Bounds.Top >= lastPoint.Y - 8 && menu.Bounds.Top <= lastPoint.Y + 48;
                bool abovePointer = menu.Bounds.Bottom >= lastPoint.Y - 48 && menu.Bounds.Bottom <= lastPoint.Y + 8;
                if (atHeader && besidePointer && (belowPointer || abovePointer) && countContext.MatchesActive())
                { TryCountContext(lastRow, lastColumn, true); return; }
                bool clickBeside = Math.Min(Math.Abs(menu.Bounds.Left - clickPoint.X), Math.Abs(menu.Bounds.Right - clickPoint.X)) <= 32;
                bool clickBelow = menu.Bounds.Top >= clickPoint.Y - 8 && menu.Bounds.Top <= clickPoint.Y + 48;
                bool clickAbove = menu.Bounds.Bottom >= clickPoint.Y - 48 && menu.Bounds.Bottom <= clickPoint.Y + 8;
                // A click on a different cell or a dismissed dialog must not override a keyboard menu.
                if (clickPending && unchecked(Environment.TickCount - clickStamp) <= 1500 && clickBeside && (clickBelow || clickAbove))
                { window = app.ActiveWindow; cell = window.RangeFromPoint(clickPoint.X, clickPoint.Y); }
                if (cell == null) cell = app.ActiveCell;
                if (cell != null) TryCountContext((int)cell.Row, (int)cell.Column, true);
                if (countContext != null && !countContext.MatchesActive()) DisposeContext();
            }
            catch { if (countContext != null && (countContext.Column != pendingColumn || countContext.SheetName != lastSheet)) DisposeContext(); }
            finally { clickPending = false; ExcelContext.Release(cell); ExcelContext.Release(window); }
        }
        private void TraceNativeCount(FilterMenu menu)
        {
            string output = Environment.GetEnvironmentVariable("SHEETPACE_NATIVE_DIAGNOSTICS");
            if (String.IsNullOrEmpty(output)) return;
            try
            {
                System.Text.StringBuilder text = new System.Text.StringBuilder();
                text.AppendLine("PID=" + process + " column=" + (countContext == null ? 0 : countContext.Column) + " last=" + lastRow + "," + lastColumn + " pending=" + pendingRow + "," + pendingColumn + " menu=" + menu.Bounds);
                if (countContext != null && countContext.Snapshot != null)
                {
                    foreach (CountItem item in countContext.Snapshot.Items) text.AppendLine("SNAPSHOT " + item.Label + "=" + item.Count);
                    foreach (FilterRow row in menu.Rows)
                    {
                        int value; bool found = countContext.Snapshot.TryGetCount(row.Name, row.Year, row.Month, out value);
                        text.AppendLine("ROW " + row.Name + " found=" + found + " count=" + value + " bounds=" + row.Bounds);
                    }
                }
                System.IO.File.AppendAllText(output, text.ToString());
            }
            catch { }
        }
        private void DrawCounts(FilterMenu menu)
        {
            if (countContext == null || countContext.Snapshot == null) { counts.Hide(); return; }
            System.Text.StringBuilder key = new System.Text.StringBuilder(menu.Window + ":" + menu.Bounds);
            foreach (FilterRow row in menu.Rows)
            {
                int value; countContext.Snapshot.TryGetCount(row.Name, row.Year, row.Month, out value);
                key.Append("|").Append(row.Name).Append(row.Bounds).Append(row.Year).Append(row.Month).Append(value);
            }
            string renderedKey = key.ToString(); if (lastCountKey == renderedKey && counts.Visible) return;
            lastCountKey = renderedKey;
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
            if (hoverProbe != null) { hoverProbe.Dispose(); hoverProbe = null; }
            if (visualHost != null) { visualHost.Dispose(); visualHost = null; }
            if (timer != null) { timer.Stop(); timer.Dispose(); timer = null; }
            if (watcher != null) { watcher.Dispose(); watcher = null; }
            if (hoverPacket != null) { hoverPacket.Dispose(); hoverPacket = null; }
            if (hover != null) { hover.Dispose(); hover = null; }
            if (counts != null) { counts.Dispose(); counts = null; }
            if (images != null) { images.Dispose(); images = null; }
            DisposeContext(); ribbon = null; app = null;
        }
        public void OnDisconnection(DisconnectMode mode, ref Array custom) { Log.Write("PID=" + System.Diagnostics.Process.GetCurrentProcess().Id + " OnDisconnection " + mode, null); Stop(); }
        public void OnBeginShutdown(ref Array custom) { Log.Write("PID=" + System.Diagnostics.Process.GetCurrentProcess().Id + " OnBeginShutdown", null); Stop(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { Log.Write("PID=" + System.Diagnostics.Process.GetCurrentProcess().Id + " OnStartupComplete", null); }
    }
}
