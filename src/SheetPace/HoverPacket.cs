using System;
using System.Drawing;
using System.IO.MemoryMappedFiles;
using System.Text;
using System.Threading;
namespace SheetPace
{
    internal sealed class HoverFrame
    {
        internal int Sequence, Row, Column, ElapsedMicroseconds, Timestamp;
        internal bool Visible, FilterHeader;
        internal IntPtr Window;
        internal Point Pointer;
        internal Rectangle Grid, Cell;
        internal string SheetName;
    }
    internal sealed class HoverPacket : IDisposable
    {
        private MemoryMappedFile file;
        private MemoryMappedViewAccessor view;
        private readonly bool writer;
        private int sequence;
        internal static string Name(uint process) { return "Local" + (char)92 + "SheetPace.Hover." + process; }
        internal HoverPacket(uint process, bool write)
        {
            writer = write;
            file = write ? MemoryMappedFile.CreateNew(Name(process), 256) : MemoryMappedFile.OpenExisting(Name(process), MemoryMappedFileRights.Read);
            view = file.CreateViewAccessor(0, 256, write ? MemoryMappedFileAccess.ReadWrite : MemoryMappedFileAccess.Read);
        }
        internal void Publish(HoverFrame frame)
        {
            if (!writer || view == null) return;
            sequence = unchecked(sequence + 2);
            view.Write(0, sequence - 1); Thread.MemoryBarrier();
            view.Write(4, frame.Visible ? 1 : 0); view.Write(8, frame.Window.ToInt64());
            view.Write(16, frame.Row); view.Write(20, frame.Column);
            view.Write(24, frame.Pointer.X); view.Write(28, frame.Pointer.Y);
            WriteRectangle(32, frame.Grid); WriteRectangle(48, frame.Cell);
            view.Write(64, frame.FilterHeader ? 1 : 0); view.Write(68, frame.ElapsedMicroseconds); view.Write(72, frame.Timestamp);
            byte[] name = new byte[64]; byte[] text = Encoding.Unicode.GetBytes(frame.SheetName ?? "");
            Buffer.BlockCopy(text, 0, name, 0, Math.Min(text.Length, 62)); view.WriteArray(80, name, 0, name.Length);
            Thread.MemoryBarrier(); view.Write(0, sequence);
        }
        private void WriteRectangle(int offset, Rectangle rect)
        { view.Write(offset, rect.X); view.Write(offset + 4, rect.Y); view.Write(offset + 8, rect.Width); view.Write(offset + 12, rect.Height); }
        private Rectangle ReadRectangle(int offset)
        { return new Rectangle(view.ReadInt32(offset), view.ReadInt32(offset + 4), view.ReadInt32(offset + 8), view.ReadInt32(offset + 12)); }
        internal HoverFrame Read()
        {
            if (view == null) return null;
            for (int attempt = 0; attempt < 3; attempt++)
            {
                int before = view.ReadInt32(0); if (before == 0 || (before & 1) != 0) continue;
                Thread.MemoryBarrier();
                HoverFrame frame = new HoverFrame { Sequence = before, Visible = view.ReadInt32(4) != 0, Window = new IntPtr(view.ReadInt64(8)),
                    Row = view.ReadInt32(16), Column = view.ReadInt32(20), Pointer = new Point(view.ReadInt32(24), view.ReadInt32(28)),
                    Grid = ReadRectangle(32), Cell = ReadRectangle(48), FilterHeader = view.ReadInt32(64) != 0,
                    ElapsedMicroseconds = view.ReadInt32(68), Timestamp = view.ReadInt32(72) };
                byte[] text = new byte[64]; view.ReadArray(80, text, 0, text.Length); frame.SheetName = Encoding.Unicode.GetString(text).TrimEnd((char)0);
                Thread.MemoryBarrier(); if (before == view.ReadInt32(0)) return frame;
            }
            return null;
        }
        public void Dispose()
        { if (view != null) { view.Dispose(); view = null; } if (file != null) { file.Dispose(); file = null; } }
    }
}
