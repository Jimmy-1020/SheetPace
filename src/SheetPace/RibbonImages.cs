using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
namespace SheetPace
{
    // OLE pictures are created without an AxHost or any WinForms window in Excel.
    internal sealed class RibbonImages : IDisposable
    {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void OleCreatePictureIndirect(IntPtr description, ref Guid iid, [MarshalAs(UnmanagedType.Bool)] bool own, [MarshalAs(UnmanagedType.Interface)] out object picture);
        private object hover, settings;
        internal object Hover { get { return hover ?? (hover = Picture(false)); } }
        internal object Settings { get { return settings ?? (settings = Picture(true)); } }
        internal static Bitmap Bitmap(bool gear)
        {
            Bitmap image = new Bitmap(32, 32, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(image))
            using (Pen grid = new Pen(Color.FromArgb(116, 139, 167), 1))
            using (Brush blue = new SolidBrush(Color.FromArgb(55, 133, 233)))
            using (Brush pale = new SolidBrush(Color.FromArgb(225, 238, 253)))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias; g.Clear(Color.Transparent);
                g.FillRectangle(pale, 2, 3, 27, 25);
                g.FillRectangle(blue, 11, 3, 9, 25); g.FillRectangle(blue, 2, 11, 27, 8);
                for (int x = 2; x <= 29; x += 9) g.DrawLine(grid, x, 3, x, 28);
                for (int y = 3; y <= 28; y += 8) g.DrawLine(grid, 2, y, 29, y);
                if (!gear)
                {
                    Point[] cursor = { new Point(18, 15), new Point(18, 30), new Point(22, 26), new Point(25, 31), new Point(28, 29), new Point(25, 24), new Point(31, 24) };
                    using (Brush ink = new SolidBrush(Color.FromArgb(29, 56, 92)))
                    using (Pen edge = new Pen(Color.White, 1.4F)) { g.FillPolygon(ink, cursor); g.DrawPolygon(edge, cursor); }
                }
                else
                {
                    g.FillEllipse(Brushes.White, 14, 14, 17, 17);
                    using (Pen cog = new Pen(Color.FromArgb(49, 70, 98), 4))
                    using (Pen tooth = new Pen(Color.FromArgb(49, 70, 98), 3))
                    {
                        g.DrawEllipse(cog, 18, 18, 8, 8);
                        for (int i = 0; i < 8; i++)
                        {
                            double angle = i * Math.PI / 4;
                            g.DrawLine(tooth, 22 + (float)Math.Cos(angle) * 5, 22 + (float)Math.Sin(angle) * 5, 22 + (float)Math.Cos(angle) * 7, 22 + (float)Math.Sin(angle) * 7);
                        }
                    }
                    g.FillEllipse(Brushes.White, 20, 20, 4, 4);
                }
            }
            return image;
        }
        private static object Picture(bool gear)
        {
            using (Bitmap bitmap = Bitmap(gear))
            {
                IntPtr handle = bitmap.GetHbitmap(Color.Transparent);
                int size = IntPtr.Size == 8 ? 24 : 20;
                IntPtr description = Marshal.AllocCoTaskMem(size);
                try
                {
                    for (int i = 0; i < size; i++) Marshal.WriteByte(description, i, 0);
                    Marshal.WriteInt32(description, 0, size); Marshal.WriteInt32(description, 4, 1);
                    Marshal.WriteIntPtr(description, 8, handle);
                    Guid iid = new Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB"); object picture;
                    OleCreatePictureIndirect(description, ref iid, true, out picture); handle = IntPtr.Zero; return picture;
                }
                finally { Marshal.FreeCoTaskMem(description); if (handle != IntPtr.Zero) NativeMethods.DeleteObject(handle); }
            }
        }
        public void Dispose()
        {
            ExcelContext.Release(hover); ExcelContext.Release(settings); hover = null; settings = null;
        }
    }
}
