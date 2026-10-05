using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using SheetPace;
class ExcelIntegration
{
    static int checks;
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    static void Check(bool condition, string name) { checks++; if (!condition) throw new Exception("FAIL: " + name); Console.WriteLine("PASS " + name); }
    static HashSet<string> Select(params string[] values) { return new HashSet<string>(values, StringComparer.OrdinalIgnoreCase); }
    static int Get(CountSnapshot data, string label) { int n; return data.TryGetCount(label, null, null, out n) ? n : -1; }
    [STAThread] static int Main(string[] args)
    {
        dynamic app = null, book = null, sheet = null; object extraBook = null; Connect addin = null;
        try
        {
            string output = Path.GetFullPath(args.Length > 0 ? args[0] : "."); Directory.CreateDirectory(output);
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            uint pid; GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out pid); File.WriteAllText(Path.Combine(output, "owned-excel-pid.txt"), pid.ToString());
            app.Visible = false; app.DisplayAlerts = false; book = app.Workbooks.Add(); sheet = book.Worksheets[1];
            object[,] data = new object[,] { { "评级", "组别" }, { "优秀", "A" }, { "良好", "A" }, { "良好", "B" }, { "中等", "A" }, { null, "B" }, { "a*b", "A" }, { "a?b", "B" } };
            sheet.Range["A1:B8"].Value2 = data; sheet.Range["A1:B8"].AutoFilter(1);
            using (ExcelContext ctx = ExcelContext.Capture((object)app, 3, 1))
            {
                CountSnapshot counts = ctx.ReadSnapshot();
                Check(counts.Total == 7 && Get(counts, "良好") == 2 && Get(counts, "(空白)") == 1, "Excel full-column counts");
                sheet.Range["A1:B8"].AutoFilter(2, "A");
                Check(Get(ctx.ReadSnapshot(), "良好") == 2, "hidden rows counted");
                ctx.Apply(Select("良好"));
                Check(!(bool)sheet.Rows[3].Hidden && (bool)sheet.Rows[4].Hidden && (bool)sheet.Rows[2].Hidden, "selected values and other-column filter preserved");
                HashSet<string> all = new HashSet<string>(); foreach (CountItem item in ctx.Snapshot.Items) all.Add(item.Label);
                ctx.Apply(all); Check(!(bool)sheet.Rows[2].Hidden && (bool)sheet.Rows[4].Hidden, "clear one column only");
                sheet.ShowAllData(); ctx.Apply(Select("a*b"));
                Check(!(bool)sheet.Rows[7].Hidden && (bool)sheet.Rows[8].Hidden, "literal asterisk criterion");
                ctx.Apply(Select("(空白)")); Check(!(bool)sheet.Rows[6].Hidden && (bool)sheet.Rows[2].Hidden, "blank criterion");
                bool rejected = false; try { ctx.Apply(Select()); } catch (InvalidOperationException) { rejected = true; }
                Check(rejected, "empty selection rejected"); sheet.ShowAllData();
            }
            sheet.AutoFilterMode = false;
            sheet.Range["D1:D4"].Value2 = new object[,] { { "数字" }, { 1.234 }, { 1.235 }, { 1.234 } };
            sheet.Range["D2:D4"].NumberFormat = "0.00";
            using (ExcelContext ctx = ExcelContext.Capture((object)app, 2, 4))
            {
                CountSnapshot counts = ctx.ReadSnapshot(); Check(Get(counts, "1.23") == 2 && Get(counts, "1.24") == 1, "formatted numeric counts via Excel TEXT array");
                ctx.Apply(Select("1.23")); Check(!(bool)sheet.Rows[2].Hidden && (bool)sheet.Rows[3].Hidden && !(bool)sheet.Rows[4].Hidden, "formatted numeric group filters original values");
            }
            if ((bool)sheet.FilterMode) sheet.ShowAllData(); sheet.AutoFilterMode = false;
            double date = new DateTime(2026, 10, 5).ToOADate();
            sheet.Range["G1:G4"].Value2 = new object[,] { { "日期" }, { date }, { date + 1 }, { date } }; sheet.Range["G2:G4"].NumberFormat = "yyyy-mm-dd";
            using (ExcelContext ctx = ExcelContext.Capture((object)app, 2, 7))
            {
                Check(Get(ctx.ReadSnapshot(), "2026-10-05") == 2, "formatted date counts");
                ctx.Apply(Select("2026-10-05")); Check(!(bool)sheet.Rows[2].Hidden && (bool)sheet.Rows[3].Hidden && !(bool)sheet.Rows[4].Hidden, "date value filter");
            }
            dynamic tableSheet = book.Worksheets.Add(); tableSheet.Name = "Tables";
            tableSheet.Range["A1:B4"].Value2 = new object[,] { { "值", "组" }, { "x", "A" }, { "y", "B" }, { "x", "B" } };
            tableSheet.Range["E1:F4"].Value2 = new object[,] { { "值", "组" }, { "red", "C" }, { "blue", "D" }, { "red", "D" } };
            dynamic table1 = tableSheet.ListObjects.Add(1, tableSheet.Range["A1:B4"], Type.Missing, 1);
            dynamic table2 = tableSheet.ListObjects.Add(1, tableSheet.Range["E1:F4"], Type.Missing, 1);
            table1.Range.AutoFilter(1, "y"); table2.Range.AutoFilter(1, "red");
            using (ExcelContext ctx = ExcelContext.Capture((object)app, 2, 5))
            {
                Check(Get(ctx.ReadSnapshot(), "red") == 2, "second table counts use correct range");
                HashSet<string> current = ctx.CurrentSelection();
                Check(current.Count == 1 && current.Contains("red"), "second table reads its own criteria");
                ctx.Apply(Select("blue")); Check(table2.AutoFilter.Filters.Item(1).On, "table filter applies");
                dynamic second = app.Workbooks.Add(); extraBook = (object)second; second.Worksheets[1].Name = "Tables";
                bool rejected = false; try { ctx.Apply(Select("red")); } catch (InvalidOperationException) { rejected = true; }
                Check(rejected, "same-name worksheet in another workbook rejected");
                book.Activate(); tableSheet.Activate();
            }
            Marshal.ReleaseComObject(table1); Marshal.ReleaseComObject(table2); Marshal.ReleaseComObject(tableSheet);
            Array custom = new object[0]; addin = new Connect(); addin.OnConnection((object)app, ConnectMode.External, null, ref custom);
            Check(addin.GetCustomUI("Microsoft.Excel.Workbook").Contains("计数筛选"), "ribbon XML and add-in lifecycle");
            addin.OnDisconnection(DisconnectMode.UserClosed, ref custom); addin = null;
            book.SaveAs(Path.Combine(output, "SheetPace-test.xlsx"), 51);
            Console.WriteLine("PASS " + checks + " Excel checks · Excel " + app.Version); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            if (addin != null) { Array a = new object[0]; addin.OnDisconnection(DisconnectMode.UserClosed, ref a); }
            try { if (book != null) book.Close(false); } catch { }
            try { if (app != null) app.Quit(); } catch { }
            foreach (object item in new object[] { extraBook, (object)sheet, (object)book, (object)app }) if (item != null && Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item);
            GC.Collect(); GC.WaitForPendingFinalizers();
        }
    }
}
