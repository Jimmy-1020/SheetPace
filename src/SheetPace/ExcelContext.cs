using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
namespace SheetPace
{
    public sealed class ExcelContext : IDisposable
    {
        private readonly dynamic app;
        private dynamic sheet, range, workbook, sourceFilter;
        private bool date1904;
        public bool SelectionKnown = true;
        public bool MatchesActive()
        {
            dynamic active = null;
            try { active = app.ActiveSheet; return SameComObject((object)active, (object)sheet); }
            finally { Release(active); }
        }
        private static bool SameComObject(object first, object second)
        {
            IntPtr a = IntPtr.Zero, b = IntPtr.Zero;
            try { a = Marshal.GetIUnknownForObject(first); b = Marshal.GetIUnknownForObject(second); return a == b; }
            finally { if (a != IntPtr.Zero) Marshal.Release(a); if (b != IntPtr.Zero) Marshal.Release(b); }
        }
        public readonly int Field, HeaderRow, Column;
        public readonly string Title, Address, SheetName;
        public CountSnapshot Snapshot { get; private set; }
        public bool MixedFormats;
        private ExcelContext(object application, object worksheet, object source, object filter, int column)
        {
            app = application; sheet = worksheet; range = source; sourceFilter = filter;
            workbook = sheet.Parent; date1904 = (bool)workbook.Date1904;
            Column = column; HeaderRow = (int)range.Row; Field = column - (int)range.Column + 1;
            if (Field < 1 || Field > (int)range.Columns.Count) throw new InvalidOperationException("请选择数据区域内的单元格。");
            dynamic header = range.Cells[1, Field];
            try { Title = Convert.ToString(header.Text); } finally { Release(header); }
            SheetName = (string)sheet.Name;
            Address = (string)range.Address[false, false, 1];
        }
        public static ExcelContext Capture(object application, int row, int column)
        {
            dynamic app = application, sheet = app.ActiveSheet, cell = null, source = null, capturedFilter = null;
            try
            {
                if ((int)sheet.Type != -4167) throw new InvalidOperationException("此功能用于普通工作表。");
                cell = sheet.Cells[row, column];
                dynamic tables = sheet.ListObjects;
                try
                {
                    for (int i = 1; i <= (int)tables.Count; i++)
                    {
                        dynamic table = tables.Item(i), tr = table.Range;
                        bool inside = row >= (int)tr.Row && row < (int)tr.Row + (int)tr.Rows.Count && column >= (int)tr.Column && column < (int)tr.Column + (int)tr.Columns.Count;
                        if (inside)
                        {
                            capturedFilter = table.AutoFilter; source = capturedFilter.Range;
                             Release(tr); Release(table); break;
                        }
                        Release(tr); Release(table);
                    }
                }
                finally { Release(tables); }
                if (source == null && (bool)sheet.AutoFilterMode)
                {
                    dynamic filter = sheet.AutoFilter, candidate = filter.Range;
                    if (column >= (int)candidate.Column && column < (int)candidate.Column + (int)candidate.Columns.Count && row >= (int)candidate.Row && row < (int)candidate.Row + (int)candidate.Rows.Count) { source = candidate; capturedFilter = filter; }
                    else { Release(candidate); Release(filter); throw new InvalidOperationException("当前工作表的自动筛选位于其他区域，请选中该区域的列，或先清除原筛选。"); }
                }
                if (source == null) source = cell.CurrentRegion;
                if ((int)source.Rows.Count < 2) throw new InvalidOperationException("该区域需要一行标题及至少一行数据。");
                ExcelContext context = new ExcelContext(application, sheet, source, capturedFilter, column);
                sheet = null; source = null; capturedFilter = null; return context;
            }
            finally { Release(cell); Release(source); Release(sheet); Release(capturedFilter); }
        }
        public CountSnapshot ReadSnapshot()
        {
            dynamic data = null, first = null, last = null;
            try
            {
                int rows = (int)range.Rows.Count - 1;
                first = sheet.Cells[HeaderRow + 1, Column]; last = sheet.Cells[HeaderRow + rows, Column];
                data = sheet.Range[first, last];
                object values = data.Value2;
                object nf = data.NumberFormat;
                MixedFormats = nf == null || nf == DBNull.Value;
                string format = MixedFormats ? "General" : Convert.ToString(nf);
                // Date system belongs to the captured workbook.
                Dictionary<string, string> formatted = new Dictionary<string, string>();
                // One array evaluation avoids a COM call per cell/unique value.
                if (!MixedFormats && format != "General" && format != "@")
                {
                    try
                    {
                        string address = (string)data.Address[false, false, 1];
                        object textValues = sheet.Evaluate("IF(ROW(" + address + "),TEXT(" + address + ",\"" + format.Replace("\"", "\"\"") + "\"))");
                        Array raw = values as Array, texts = textValues as Array;
                        if (raw != null && texts != null && raw.GetLength(0) == texts.GetLength(0))
                            for (int i = 0; i < raw.GetLength(0); i++)
                                formatted[ValueFormatter.RawKey(raw.GetValue(raw.GetLowerBound(0) + i, raw.GetLowerBound(1)))] = Convert.ToString(texts.GetValue(texts.GetLowerBound(0) + i, texts.GetLowerBound(1)));
                        else if (raw == null) formatted[ValueFormatter.RawKey(values)] = Convert.ToString(app.WorksheetFunction.Text(values, format));
                    }
                    catch (Exception ex) { Log.Write("列显示格式解析回退", ex); }
                }
                Snapshot = CountSnapshot.Build(values, format, date1904, delegate(object value)
                {
                    string label; return formatted.TryGetValue(ValueFormatter.RawKey(value), out label) ? label : ValueFormatter.General(value);
                });
                return Snapshot;
            }
            finally { Release(data); Release(first); Release(last); }
        }
        public HashSet<string> CurrentSelection()
        {
            HashSet<string> selected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (CountItem item in Snapshot.Items) selected.Add(item.Label);
            dynamic filter = null, filters = null, current = null;
            try
            {
                if (sourceFilter == null) return selected;
                filters = sourceFilter.Filters; current = filters.Item(Field);
                if (!(bool)current.On) return selected;
                object criteria = current.Criteria1;
                Array array = criteria as Array;
                HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (array == null) keys.Add(Convert.ToString(criteria)); else foreach (object value in array) keys.Add(Convert.ToString(value));
                HashSet<string> matching = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (CountItem item in Snapshot.Items)
                    if (keys.Contains(item.Label) || keys.Contains("=" + item.Label) || item.Values.Any(v => keys.Contains(ValueFormatter.EscapeCriterion(v)) || keys.Contains("=" + ValueFormatter.EscapeCriterion(v)) || keys.Contains(ValueFormatter.General(v)))) matching.Add(item.Label);
                if (matching.Count > 0) return matching;
                SelectionKnown = false;
            }
            catch { SelectionKnown = false; }
            finally { Release(current); Release(filters); Release(filter); }
            return selected;
        }
        public void Apply(HashSet<string> selected)
        {
            if (!MatchesActive()) throw new InvalidOperationException("工作表已切换，请重新打开计数筛选。");
            if ((bool)sheet.ProtectContents) throw new InvalidOperationException("工作表已保护，请先解除保护后筛选。");
            if (selected.Count == 0) throw new InvalidOperationException("请至少勾选一个值。");
            if (sourceFilter == null && !(bool)sheet.AutoFilterMode) range.AutoFilter(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            if (selected.Count == Snapshot.Items.Count) { range.AutoFilter(Field, Type.Missing, Type.Missing, Type.Missing, Type.Missing); return; }
            List<object> criteria = Snapshot.Items.Where(i => selected.Contains(i.Label))
                .Select(i => (object)(i.Label == "(空白)" ? "=" : ValueFormatter.EscapeCriterion(i.Label))).ToList();
            if (criteria.Count > 10000) throw new InvalidOperationException("Excel 单列值筛选最多支持 10000 个筛选值，请缩小勾选范围。");
            range.AutoFilter(Field, criteria.ToArray(), 7, Type.Missing, Type.Missing);
        }
        internal static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        public void Dispose() { Snapshot = null; Release(range); Release(sheet); Release(workbook); Release(sourceFilter); range = null; sheet = null; workbook = null; sourceFilter = null; }
    }
}
