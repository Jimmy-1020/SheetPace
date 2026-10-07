using System;
using System.Drawing;
using System.IO;
using SheetPace;
class CoreTests
{
    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)] struct Message
    { public IntPtr Window; public uint Id; public UIntPtr WParam; public IntPtr LParam; public uint Time; public Point Pointer; public uint Private; }
    [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool PeekMessageW(out Message message,IntPtr window,uint first,uint last,uint flags);
    [System.Runtime.InteropServices.DllImport("user32.dll")] static extern IntPtr DispatchMessageW(ref Message message);
    static void Pump() { Message message;while(PeekMessageW(out message,IntPtr.Zero,0,0,1))DispatchMessageW(ref message); }
    static int checks;
    static void Check(bool condition, string name) { checks++; if (!condition) throw new Exception("FAIL: " + name); Console.WriteLine("PASS " + name); }
    static object[,] Column(params object[] values) { object[,] data = new object[values.Length, 1]; for (int i = 0; i < values.Length; i++) data[i, 0] = values[i]; return data; }
    static CountSnapshot Count(object data, string format) { return CountSnapshot.Build(data, format, false, ValueFormatter.General); }
    static int Get(CountSnapshot data, string label) { int count; if (!data.TryGetCount(label, null, null, out count)) return -1; return count; }
    static int Main(string[] args)
    {
        try
        {
            CountSnapshot data = Count(Column("良好", "优秀", "良好", "中等", null, "", "Foo", "foo", " Foo"), "General");
            Check(data.Total == 9, "all data rows included");
            Check(Get(data, "良好") == 2 && Get(data, "优秀") == 1, "repeated text counts");
            Check(Get(data, "(空白)") == 2 && Get(data, "(Blanks)") == 2, "blank aliases and empty formulas");
            Check(Get(data, "FOO") == 2 && Get(data, " Foo") == 1, "Excel case-insensitive grouping, spaces preserved");
            Check(Get(data, "(全选)") == 9, "select-all count");
            Check(Get(Count(42.0, "General"), "42") == 1, "single cell scalar");
            Array oneBased = Array.CreateInstance(typeof(object), new[] { 3, 1 }, new[] { 1, 1 });
            oneBased.SetValue("x", 1, 1); oneBased.SetValue("x", 2, 1); oneBased.SetValue("y", 3, 1);
            Check(Get(Count(oneBased, "General"), "x") == 2, "one-based COM array");
            Check(ValueFormatter.EscapeCriterion("a*b?~") == "a~*b~?~~", "wildcards escaped");
            Check(ValueFormatter.EscapeCriterion(null) == "=", "blank criterion");
            Check(ValueFormatter.General(new System.Runtime.InteropServices.ErrorWrapper(2042)) == "#N/A", "error value");
            Check(ValueFormatter.IsDateFormat("yyyy-mm-dd") && ValueFormatter.IsDateFormat("d-mmm"), "date format detection");
            Check(!ValueFormatter.IsDateFormat("General") && !ValueFormatter.IsDateFormat("0.00") && !ValueFormatter.IsDateFormat("0 \"days\""), "literal text is not date");
            double day = new DateTime(2026, 10, 5).ToOADate();
            data = Count(Column(day, day, day + 1), "yyyy-mm-dd"); int n;
            Check(data.TryGetCount("2026", null, null, out n) && n == 3, "year group");
            Check(data.TryGetCount("10月", "2026", null, out n) && n == 3, "Chinese month group");
            Check(data.TryGetCount("5", "2026", "10月", out n) && n == 2, "day group");
            Settings settings = new Settings(); Check(settings.Mode == HighlightMode.ClickSelection, "default highlight mode is click selection"); settings.Transparency = 100; Check(settings.Alpha == 0, "100 percent transparency");
            settings.Transparency = 0; Check(settings.Alpha == 255, "zero percent transparency");
            settings.HighlightColor = Color.FromArgb(21, 45, 89); settings.Transparency = 76; settings.HoverEnabled = false; settings.Mode = HighlightMode.FollowMouse;
            string dir = args.Length > 0 ? args[0] : Path.Combine(Path.GetTempPath(), "SheetPace-tests"); Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "settings-test.ini"); settings.Save(path); settings.Save(path);
            Settings loaded = Settings.Load(path);
            Check(loaded.Transparency == 76 && loaded.HighlightColor.ToArgb() == settings.HighlightColor.ToArgb() && !loaded.HoverEnabled && loaded.Mode == HighlightMode.FollowMouse, "settings atomic save and roundtrip");
            File.WriteAllText(path, "Transparency=150" + Environment.NewLine + "Color=invalid");
            Check(Settings.Load(path).Transparency == 100, "corrupt settings clamped"); File.Delete(path);
            File.WriteAllText(path, "HoverEnabled=True" + Environment.NewLine + "Transparency=61");
            Check(Settings.Load(path).Mode == HighlightMode.ClickSelection && Settings.Load(path).Transparency == 61, "old settings migrate to selection and keep appearance");
            File.WriteAllText(path, "Mode=99"); Check(Settings.Load(path).Mode == HighlightMode.ClickSelection, "invalid highlight mode falls back to selection"); File.Delete(path);
            using (RibbonImages icons = new RibbonImages())
            {
                foreach (object picture in new[] { icons.Hover, icons.Settings })
                {
                    Guid iid = new Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB");
                    IntPtr dispatch = System.Runtime.InteropServices.Marshal.GetIDispatchForObject(picture), result;
                    int hr = System.Runtime.InteropServices.Marshal.QueryInterface(dispatch, ref iid, out result);
                    Check(hr == 0 && result != IntPtr.Zero, "Ribbon icon exposes native IPictureDisp");
                    if (result != IntPtr.Zero) System.Runtime.InteropServices.Marshal.Release(result);
                    System.Runtime.InteropServices.Marshal.Release(dispatch);
                }
                using (Bitmap bitmap = RibbonImages.Bitmap(false)) bitmap.Save(Path.Combine(dir, "hover-icon.png"));
                using (Bitmap bitmap = RibbonImages.Bitmap(true)) bitmap.Save(Path.Combine(dir, "settings-icon.png"));
            }
            int size = 100000; object[,] large = new object[size, 1]; for (int i = 0; i < size; i++) large[i, 0] = i % 23;
            data = Count(large, "General"); Check(data.Total == size && data.Items.Count == 23, "100000 row frequency calculation");
            FilterMenu menu = new FilterMenu { Window = new IntPtr(123456), Timestamp = 7654321, Bounds = new Rectangle(10, 20, 300, 400) };
            menu.Rows.Add(new FilterRow { Name = "5", Year = "2026", Month = "10月", Bounds = new Rectangle(30, 40, 200, 18) });
            using (MemoryStream stream = new MemoryStream())
            {
                NativeMenuProtocol.Write(new BinaryWriter(stream), menu); stream.Position = 0;
                FilterMenu received = NativeMenuProtocol.Read(new BinaryReader(stream));
                Check(received.Window == menu.Window && received.Timestamp == menu.Timestamp && received.Bounds == menu.Bounds && received.Rows.Count == 1 && received.Rows[0].Month == "10月" && received.Rows[0].Year == "2026" && received.Rows[0].Bounds == menu.Rows[0].Bounds, "helper preserves native menu geometry and date hierarchy");
            }
            using (MemoryStream stream = new MemoryStream())
            {
                NativeMenuProtocol.Write(new BinaryWriter(stream), null); stream.Position = 0;
                Check(NativeMenuProtocol.Read(new BinaryReader(stream)) == null, "helper reports closed native menu");
            }
            using (MemoryStream stream = new MemoryStream(new byte[4]))
            {
                bool rejected = false; try { NativeMenuProtocol.Read(new BinaryReader(stream)); } catch (InvalidDataException) { rejected = true; }
                Check(rejected, "helper rejects incompatible protocol");
            }
            IndexSpan[] spans = SelectionSnapshot.Merge(new System.Collections.Generic.List<IndexSpan> {
                new IndexSpan(9, 11), new IndexSpan(3, 3), new IndexSpan(1, 2), new IndexSpan(10, 12), new IndexSpan(6, 6) });
            Check(spans.Length == 3 && spans[0].First == 1 && spans[0].Last == 3 && spans[1].First == 6 && spans[1].Last == 6 && spans[2].First == 9 && spans[2].Last == 12, "selection union merges overlap and adjacency, preserves gaps");
            uint testProcess = 0x40000000u + (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            using (HoverPacket writer = new HoverPacket(testProcess, true))
            using (HoverPacket reader = new HoverPacket(testProcess, false))
            {
                Check(reader.Read() == null, "hover reader ignores uninitialized packet");
                writer.Publish(new HoverFrame { Visible = true, Window = new IntPtr(123456), Row = 7, Column = -7,
                    ShapeRevision = 7, RowBands = new[] { new Rectangle(7, 2, 300, 4), new Rectangle(7, 20, 300, 4) }, ColumnBands = new[] { new Rectangle(8, 2, 30, 400) }, Pointer = new Point(7, -7), Grid = new Rectangle(7, 2, 3, 4), Cell = new Rectangle(5, 7, 8, 9),
                    SheetName = "定位测试", FilterHeader = true, PointerValid = true, TargetRow = 12, TargetColumn = 4, Mode = HighlightMode.ClickSelection, ElapsedMicroseconds = 1200, Timestamp = 9 });
                HoverFrame received = reader.Read();
                Check(received != null && received.Row == 7 && received.Column == -7 && received.Pointer == new Point(7, -7) && received.Grid.X == 7 && received.Cell.Y == 7 && received.SheetName == "定位测试" && received.FilterHeader && received.Window == new IntPtr(123456) && received.ElapsedMicroseconds == 1200 && received.PointerValid && received.TargetRow == 12 && received.TargetColumn == 4 && received.Mode == HighlightMode.ClickSelection, "hover packet geometry and metadata roundtrip");
                Check(received.ShapeRevision == 7 && received.RowBands.Length == 2 && received.ColumnBands.Length == 1 && received.RowBands[1].Y == 20 && received.ColumnBands[0].Height == 400, "variable-length highlight bands roundtrip beyond fixed header");
                bool oversized = false;
                try { writer.Publish(new HoverFrame { RowBands = new Rectangle[HoverPacket.MaxBands + 1] }); } catch (ArgumentOutOfRangeException) { oversized = true; }
                Check(oversized && reader.Read().Row == 7, "excess bands rejected before overwriting valid frame");
                System.Threading.Thread producer = new System.Threading.Thread(delegate()
                {
                    for (int i = 8; i < 5008; i++) {
                        Rectangle[] rowBands = new Rectangle[i % 9 + 1];
                        for (int j = 0; j < rowBands.Length; j++) rowBands[j] = new Rectangle(i, j, 300, i);
                        writer.Publish(new HoverFrame { Visible = true, Row = i, Column = -i, ShapeRevision = i, RowBands = rowBands,
                            ColumnBands = new[] { new Rectangle(-i, 0, 20, i) }, Pointer = new Point(i, -i), Grid = new Rectangle(i, 2, 3, 4), Cell = new Rectangle(5, i, 8, 9) });
                    }
                });
                producer.Start(); bool coherent = true; int samples = 0;
                while (producer.IsAlive)
                {
                    HoverFrame frame = reader.Read(); if (frame == null) continue; samples++;
                    if (frame.Column != -frame.Row || frame.Pointer.X != frame.Row || frame.Pointer.Y != -frame.Row || frame.Grid.X != frame.Row || frame.Cell.Y != frame.Row || frame.ShapeRevision != frame.Row) coherent = false;
                    if (frame.Row >= 8) {
                        if (frame.RowBands.Length != frame.Row % 9 + 1 || frame.ColumnBands.Length != 1 || frame.ColumnBands[0].X != -frame.Row) coherent = false;
                        foreach (Rectangle band in frame.RowBands) if (band.X != frame.Row || band.Height != frame.Row) coherent = false;
                    }
                }
                producer.Join();
                Check(coherent && samples > 0 && reader.Read().Row == 5007, "hover sequence protects concurrent readers from torn frames");
            }
            int refreshes = 0;
            using (HoverProbe probe = new HoverProbe(new object(), testProcess, delegate { return new Settings { HoverEnabled = false, NativeCountsEnabled = false }; }, delegate { refreshes++; }))
            using (HoverPacket reader = new HoverPacket(testProcess, false))
            {
                System.Threading.Thread.Sleep(60); Pump();
                HoverFrame frame = reader.Read();
                Check(frame != null && !frame.Visible && frame.Timestamp != 0 && refreshes > 0, "native STA timer publishes hidden frame without Excel or forms");
            }
            using (HoverProbe probe = new HoverProbe(new object(), testProcess, delegate { return new Settings { HoverEnabled = false, NativeCountsEnabled = false }; }, delegate { }))
            using (HoverPacket reader = new HoverPacket(testProcess, false))
            {
                System.Threading.Thread.Sleep(60); Pump();
                Check(reader.Read() != null, "native probe releases timer, class and mapping before reconnect");
            }
            Pump();
            Console.WriteLine("PASS " + checks + " checks"); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
