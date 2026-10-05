using System;
using System.Drawing;
using System.IO;
using SheetPace;
class CoreTests
{
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
            Settings settings = new Settings(); settings.Transparency = 100; Check(settings.Alpha == 0, "100 percent transparency");
            settings.Transparency = 0; Check(settings.Alpha == 255, "zero percent transparency");
            settings.HighlightColor = Color.FromArgb(21, 45, 89); settings.Transparency = 76; settings.HoverEnabled = false;
            string dir = args.Length > 0 ? args[0] : Path.Combine(Path.GetTempPath(), "SheetPace-tests"); Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "settings-test.ini"); settings.Save(path); settings.Save(path);
            Settings loaded = Settings.Load(path);
            Check(loaded.Transparency == 76 && loaded.HighlightColor.ToArgb() == settings.HighlightColor.ToArgb() && !loaded.HoverEnabled, "settings atomic save and roundtrip");
            File.WriteAllText(path, "Transparency=150" + Environment.NewLine + "Color=invalid");
            Check(Settings.Load(path).Transparency == 100, "corrupt settings clamped"); File.Delete(path);
            int size = 100000; object[,] large = new object[size, 1]; for (int i = 0; i < size; i++) large[i, 0] = i % 23;
            data = Count(large, "General"); Check(data.Total == size && data.Items.Count == 23, "100000 row frequency calculation");
            Console.WriteLine("PASS " + checks + " checks"); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
