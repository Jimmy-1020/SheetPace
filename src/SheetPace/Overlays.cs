using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Runtime.InteropServices;
namespace SheetPace
{
    internal sealed class OverlayWindow : Form
    {
        public OverlayWindow() { FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; }
        public new bool Visible { get { return IsHandleCreated && NativeMethods.IsWindowVisible(Handle); } }
        public new void Hide() { if (IsHandleCreated) NativeMethods.ShowWindow(Handle, 0); }
        protected override bool ShowWithoutActivation { get { return true; } }
        protected override CreateParams CreateParams
        {
            get { CreateParams p = base.CreateParams; p.ExStyle |= NativeMethods.WS_EX_LAYERED | NativeMethods.WS_EX_TRANSPARENT | NativeMethods.WS_EX_NOACTIVATE | NativeMethods.WS_EX_TOOLWINDOW; return p; }
        }
        protected override void WndProc(ref Message message)
        {
            // Decorative, click-through windows expose no interactive accessibility provider.
            // This also prevents UI Automation from retaining managed providers at Office shutdown.
            if (message.Msg == 0x3D) { message.Result = IntPtr.Zero; return; }
            if (message.Msg == 0x84) { message.Result = new IntPtr(-1); return; }
            if (message.Msg == 0x21) { message.Result = new IntPtr(3); return; }
            base.WndProc(ref message);
        }
        public void Render(Rectangle screen, Action<Graphics> paint)
        {
            if (screen.Width <= 0 || screen.Height <= 0) { Hide(); return; }
            using (Bitmap bitmap = new Bitmap(screen.Width, screen.Height, PixelFormat.Format32bppPArgb))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap)) { graphics.Clear(Color.Transparent); paint(graphics); }
                IntPtr dc = NativeMethods.GetDC(IntPtr.Zero), mem = NativeMethods.CreateCompatibleDC(dc);
                NativeMethods.BITMAPINFO info = new NativeMethods.BITMAPINFO { Size = 40, Width = bitmap.Width, Height = -bitmap.Height, Planes = 1, BitCount = 32 };
                IntPtr bits, image = NativeMethods.CreateDIBSection(dc, ref info, 0, out bits, IntPtr.Zero, 0);
                if (image == IntPtr.Zero) { NativeMethods.DeleteDC(mem); NativeMethods.ReleaseDC(IntPtr.Zero, dc); throw new System.ComponentModel.Win32Exception(); }
                BitmapData locked = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppPArgb);
                try { byte[] pixels = new byte[bitmap.Width * bitmap.Height * 4]; Marshal.Copy(locked.Scan0, pixels, 0, pixels.Length); Marshal.Copy(pixels, 0, bits, pixels.Length); }
                finally { bitmap.UnlockBits(locked); }
                IntPtr old = NativeMethods.SelectObject(mem, image);
                try
                {
                    NativeMethods.POINT dest = new NativeMethods.POINT(screen.X, screen.Y), source = new NativeMethods.POINT(0, 0);
                    NativeMethods.SIZE size = new NativeMethods.SIZE(screen.Width, screen.Height);
                    NativeMethods.BLENDFUNCTION blend = new NativeMethods.BLENDFUNCTION { SourceConstantAlpha = 255, AlphaFormat = 1 };
                    if (!NativeMethods.UpdateLayeredWindow(Handle, dc, ref dest, ref size, mem, ref source, 0, ref blend, 2)) throw new System.ComponentModel.Win32Exception();
                    NativeMethods.SetWindowPos(Handle, NativeMethods.HWND_TOPMOST, screen.X, screen.Y, screen.Width, screen.Height, NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_SHOWWINDOW);
                }
                finally { NativeMethods.SelectObject(mem, old); NativeMethods.DeleteObject(image); NativeMethods.DeleteDC(mem); NativeMethods.ReleaseDC(IntPtr.Zero, dc); }
            }
        }
        public void Fill(Rectangle screen, Color color, int alpha)
        {
            if (screen.Width <= 0 || screen.Height <= 0 || alpha == 0) { Hide(); return; }
            Color rgb = Color.FromArgb(color.R, color.G, color.B);
            if (BackColor != rgb) BackColor = rgb;
            NativeMethods.SetLayeredWindowAttributes(Handle, 0, (byte)alpha, 2);
            NativeMethods.SetWindowPos(Handle, NativeMethods.HWND_TOPMOST, screen.X, screen.Y, screen.Width, screen.Height, NativeMethods.SWP_NOACTIVATE | NativeMethods.SWP_SHOWWINDOW);
        }
        public void DrawCross(Rectangle grid, Rectangle cell, Settings settings)
        {
            if (settings.Alpha == 0) { Hide(); return; }
            Rectangle local = new Rectangle(cell.X - grid.X, cell.Y - grid.Y, cell.Width, cell.Height);
            Render(grid, delegate(Graphics graphics)
            {
                using (Region region = new Region(new Rectangle(0, local.Y, grid.Width, local.Height)))
                using (Brush brush = new SolidBrush(Color.FromArgb(settings.Alpha, settings.HighlightColor)))
                {
                    region.Union(new Rectangle(local.X, 0, local.Width, grid.Height)); graphics.FillRegion(brush, region);
                }
            });
        }
    }
    internal sealed class CrossOverlay : IDisposable
    {
        private readonly OverlayWindow row = new OverlayWindow();
        private readonly OverlayWindow above = new OverlayWindow();
        private readonly OverlayWindow below = new OverlayWindow();
        public string Text { set { row.Text = value; above.Text = value + " Column Above"; below.Text = value + " Column Below"; } }
        public System.Windows.Forms.Form PrimaryWindow { get { return row; } }
        public bool Visible { get { return row.Visible || above.Visible || below.Visible; } }
        public void Hide() { row.Hide(); above.Hide(); below.Hide(); }
        public void DrawCross(Rectangle grid, Rectangle cell, Settings settings)
        {
            Rectangle clipped = Rectangle.Intersect(grid, cell);
            if (clipped.IsEmpty || settings.Alpha == 0) { Hide(); return; }
            // Three disjoint strips avoid a full-grid bitmap and double opacity at the intersection.
            row.Fill(new Rectangle(grid.Left, clipped.Top, grid.Width, clipped.Height), settings.HighlightColor, settings.Alpha);
            above.Fill(new Rectangle(clipped.Left, grid.Top, clipped.Width, clipped.Top - grid.Top), settings.HighlightColor, settings.Alpha);
            below.Fill(new Rectangle(clipped.Left, clipped.Bottom, clipped.Width, grid.Bottom - clipped.Bottom), settings.HighlightColor, settings.Alpha);
        }
        public void Dispose() { row.Dispose(); above.Dispose(); below.Dispose(); }
    }
    internal static class GridGeometry
    {
        public static Rectangle FindGrid(IntPtr parent, Point pointer)
        {
            Rectangle found = Rectangle.Empty;
            NativeMethods.EnumChildWindows(parent, delegate(IntPtr hwnd, IntPtr unused)
            {
                if (NativeMethods.ClassName(hwnd) == "EXCEL7" && NativeMethods.IsWindowVisible(hwnd))
                {
                    Rectangle rect = NativeMethods.ClientRectangleOnScreen(hwnd);
                    if (rect.Contains(pointer)) { found = rect; return false; }
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }
        public static Rectangle CellRectangle(dynamic window, dynamic cell, Rectangle grid, Point pointer)
        {
            dynamic merged = null;
            try
            {
                merged = cell.MergeArea;
                // Hit testing uses the same physical screen coordinates as the mouse,
                // avoiding Excel's point conversion differences at display scaling.
                int row = (int)merged.Row, column = (int)merged.Column;
                int left = Edge(window, grid.Left, pointer.X, pointer, true, column, true);
                int right = Edge(window, pointer.X, grid.Right - 1, pointer, true, column, false) + 1;
                int top = Edge(window, grid.Top, pointer.Y, pointer, false, row, true);
                int bottom = Edge(window, pointer.Y, grid.Bottom - 1, pointer, false, row, false) + 1;
                return Rectangle.FromLTRB(left, top, right, bottom);
            }
            finally { ExcelContext.Release(merged); }
        }
        private static int Edge(dynamic window, int lower, int upper, Point pointer, bool horizontal, int expected, bool leading)
        {
            while (lower < upper)
            {
                int midpoint = leading ? (lower + upper) / 2 : (lower + upper + 1) / 2;
                dynamic probe = null, area = null; bool match = false;
                try
                {
                    probe = window.RangeFromPoint(horizontal ? midpoint : pointer.X, horizontal ? pointer.Y : midpoint);
                    if (probe != null) { area = probe.MergeArea; match = (int)(horizontal ? area.Column : area.Row) == expected; }
                }
                catch { }
                finally { ExcelContext.Release(area); ExcelContext.Release(probe); }
                if (leading) { if (match) upper = midpoint; else lower = midpoint + 1; }
                else { if (match) lower = midpoint; else upper = midpoint - 1; }
            }
            return lower;
        }
    }
}
