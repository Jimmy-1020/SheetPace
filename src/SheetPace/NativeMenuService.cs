using System;
using System.Drawing;
using System.IO;
namespace SheetPace
{
    internal static class NativeMenuProtocol
    {
        private const int Signature = 0x53504D31;
        private static void WriteRect(BinaryWriter writer, Rectangle rect)
        { writer.Write(rect.X); writer.Write(rect.Y); writer.Write(rect.Width); writer.Write(rect.Height); }
        private static Rectangle ReadRect(BinaryReader reader)
        { return new Rectangle(reader.ReadInt32(), reader.ReadInt32(), reader.ReadInt32(), reader.ReadInt32()); }
        internal static void Write(BinaryWriter writer, FilterMenu menu)
        {
            writer.Write(Signature); writer.Write(menu != null);
            if (menu != null)
            {
                writer.Write(menu.Window.ToInt64()); writer.Write(menu.Timestamp); WriteRect(writer, menu.Bounds); writer.Write(menu.Rows.Count);
                foreach (FilterRow row in menu.Rows)
                { writer.Write(row.Name ?? ""); writer.Write(row.Year ?? ""); writer.Write(row.Month ?? ""); WriteRect(writer, row.Bounds); }
            }
            writer.Flush();
        }
        internal static FilterMenu Read(BinaryReader reader)
        {
            if (reader.ReadInt32() != Signature) throw new InvalidDataException("菜单通信版本不匹配");
            if (!reader.ReadBoolean()) return null;
            FilterMenu menu = new FilterMenu { Window = new IntPtr(reader.ReadInt64()), Timestamp = reader.ReadInt64(), Bounds = ReadRect(reader) };
            int count = reader.ReadInt32(); if (count < 0 || count > 2000) throw new InvalidDataException("菜单项目数量无效");
            for (int i = 0; i < count; i++) menu.Rows.Add(new FilterRow { Name = reader.ReadString(), Year = reader.ReadString(), Month = reader.ReadString(), Bounds = ReadRect(reader) });
            return menu;
        }
    }
    public static class NativeMenuService
    {
        // Called by SheetPace.NativeMenu.exe. EOF also releases the helper if Excel terminates.
        public static int Run(uint process)
        {
            NativeFilterScanner scanner = new NativeFilterScanner(process);
            try
            {
                using (BinaryReader reader = new BinaryReader(Console.OpenStandardInput()))
                using (BinaryWriter writer = new BinaryWriter(Console.OpenStandardOutput()))
                    while (reader.ReadByte() == 1)
                    {
                        FilterMenu menu = null;
                        try { if (NativeMethods.IsProcessForeground(process)) menu = scanner.Scan(); }
                        catch (Exception ex) { Log.Write("原生筛选列表不可访问", ex); }
                        NativeMenuProtocol.Write(writer, menu);
                    }
                return 0;
            }
            catch (EndOfStreamException) { return 0; }
            catch (IOException) { return 0; }
            catch (Exception ex) { Log.Write("原生筛选辅助进程失败", ex); return 1; }
        }
    }
}
