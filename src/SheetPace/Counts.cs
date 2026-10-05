using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
namespace SheetPace
{
    public sealed class CountItem
    {
        public string Label;
        public int Count;
        public readonly List<object> Values = new List<object>();
    }
    public sealed class CountSnapshot
    {
        public readonly List<CountItem> Items = new List<CountItem>();
        public int Total;
        public bool IsDate;
        private readonly Dictionary<string, int> lookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> dates = new Dictionary<string, int>();
        public bool TryGetCount(string label, string year, string month, out int count)
        {
            if (label == "(全选)" || label == "(Select All)" || label == "(全選)") { count = Total; return true; }
            if (label == "(空白)" || label == "(Blanks)" || label == "(空白单元格)") label = "(空白)";
            if (IsDate)
            {
                string key = label;
                if (!String.IsNullOrEmpty(year)) key = year + "/" + label;
                if (!String.IsNullOrEmpty(month)) key = year + "/" + month + "/" + label;
                if (dates.TryGetValue(key, out count)) return true;
            }
            return lookup.TryGetValue(label, out count);
        }
        private void AddDate(string key) { int n; dates.TryGetValue(key, out n); dates[key] = n + 1; }
        public static CountSnapshot Build(object values, string numberFormat, bool date1904, Func<object, string> display)
        {
            CountSnapshot result = new CountSnapshot(); result.IsDate = ValueFormatter.IsDateFormat(numberFormat);
            Dictionary<string, CountItem> groups = new Dictionary<string, CountItem>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> labels = new Dictionary<string, string>();
            HashSet<string> rawValues = new HashSet<string>();
            Array array = values as Array;
            int length = array == null ? 1 : array.GetLength(0);
            int start = array == null ? 0 : array.GetLowerBound(0);
            for (int i = 0; i < length; i++)
            {
                object value = array == null ? values : array.GetValue(start + i, array.GetLowerBound(1));
                string rawKey = ValueFormatter.RawKey(value), label;
                if (!labels.TryGetValue(rawKey, out label))
                {
                    label = value == null || (value is string && (string)value == "") ? "(空白)" : display(value);
                    labels[rawKey] = label;
                }
                CountItem item;
                if (!groups.TryGetValue(label, out item)) { item = new CountItem { Label = label }; groups.Add(label, item); }
                item.Count++; result.Total++;
                if (rawValues.Add(rawKey)) item.Values.Add(value);
                if (result.IsDate && value is double)
                {
                    try
                    {
                        DateTime date = DateTime.FromOADate((double)value + (date1904 ? 1462 : 0));
                        string y = date.Year.ToString(), m = date.Month.ToString(), d = date.Day.ToString();
                        result.AddDate(y); result.AddDate(y + "/" + m); result.AddDate(y + "/" + m + "/" + d);
                        result.AddDate(y + "/" + m + "月"); result.AddDate(y + "/" + m + "月/" + d);
                        string en = date.ToString("MMMM", CultureInfo.GetCultureInfo("en-US"));
                        result.AddDate(y + "/" + en); result.AddDate(y + "/" + en + "/" + d);
                    }
                    catch (ArgumentException) { }
                }
            }
            result.Items.AddRange(groups.Values.OrderBy(x => x.Label == "(空白)" ? 1 : 0).ThenBy(x => x.Label, StringComparer.CurrentCultureIgnoreCase));
            foreach (CountItem item in result.Items) result.lookup[item.Label] = item.Count;
            return result;
        }
    }
    public static class ValueFormatter
    {
        public static bool IsDateFormat(string format)
        {
            if (String.IsNullOrEmpty(format)) return false;
            if (format.Equals("General", StringComparison.OrdinalIgnoreCase) || format == "@") return false;
            System.Text.StringBuilder cleaned = new System.Text.StringBuilder();
            bool quoted = false, bracket = false;
            for (int i = 0; i < format.Length; i++)
            {
                char c = format[i];
                if (c == (char)34) { quoted = !quoted; continue; }
                if (quoted) continue;
                if (c == (char)92 || c == '_' || c == '*') { i++; continue; }
                if (c == '[') { bracket = true; continue; }
                if (c == ']') { bracket = false; continue; }
                if (!bracket) cleaned.Append(Char.ToLowerInvariant(c));
            }
            string clean = cleaned.ToString();
            return clean.Contains("y") || clean.Contains("d") || clean.Contains("m") || clean.Contains("h:");
        }
        public static string RawKey(object value)
        {
            if (value == null || (value is string && (string)value == "")) return "blank";
            ErrorWrapper error = value as ErrorWrapper;
            if (error != null) return "error:" + error.ErrorCode;
            return value.GetType().FullName + ":" + Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        public static string General(object value)
        {
            if (value == null) return "(空白)";
            ErrorWrapper error = value as ErrorWrapper;
            if (error != null)
            {
                switch (error.ErrorCode & 0xFFFF) { case 2007: return "#DIV/0!"; case 2015: return "#VALUE!"; case 2023: return "#REF!"; case 2029: return "#NAME?"; case 2036: return "#NUM!"; case 2042: return "#N/A"; default: return "#ERROR!"; }
            }
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            return Convert.ToString(value, CultureInfo.CurrentCulture);
        }
        public static string EscapeCriterion(object value)
        {
            if (value == null || (value is string && (string)value == "")) return "=";
            return General(value).Replace("~", "~~").Replace("*", "~*").Replace("?", "~?");
        }
    }
}
