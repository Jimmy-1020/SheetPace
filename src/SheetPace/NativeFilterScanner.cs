using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Automation;
namespace SheetPace
{
    // UI Automation runs only in the helper executable, never inside Excel.
    internal sealed class NativeFilterScanner
    {
        private readonly uint process;
        internal NativeFilterScanner(uint processID) { process = processID; }
        internal FilterMenu Scan()
        {
            List<IntPtr> candidates = new List<IntPtr>();
            NativeMethods.EnumWindows(delegate(IntPtr hwnd, IntPtr unused)
            {
                uint owner; NativeMethods.GetWindowThreadProcessId(hwnd, out owner);
                if (owner != process || !NativeMethods.IsWindowVisible(hwnd)) return true;
                string cls = NativeMethods.ClassName(hwnd);
                if (cls != "XLMAIN" && !cls.StartsWith("WindowsForms", StringComparison.Ordinal)) candidates.Add(hwnd);
                if (cls == "XLMAIN") NativeMethods.EnumChildWindows(hwnd, delegate(IntPtr child, IntPtr ignored)
                {
                    string childClass = NativeMethods.ClassName(child);
                    if (NativeMethods.IsWindowVisible(child) && (childClass == "NetUIHWND" || childClass == "EXCEL6")) candidates.Add(child);
                    return true;
                }, IntPtr.Zero);
                return true;
            }, IntPtr.Zero);
            foreach (IntPtr candidate in candidates)
            {
                NativeMethods.RECT winRect; NativeMethods.GetWindowRect(candidate, out winRect);
                Rectangle bounds = winRect.ToRectangle();
                if (bounds.Width < 100 || bounds.Width > 950 || bounds.Height < 100 || bounds.Height > 1600) continue;
                AutomationElement root = AutomationElement.FromHandle(candidate);
                if (root == null) continue;
                Condition types = new OrCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                AutomationElementCollection controls = root.FindAll(TreeScope.Descendants, types);
                bool all = false;
                for (int i = 0; i < Math.Min(controls.Count, 2000); i++) if (IsAll(controls[i].Current.Name)) { all = true; break; }
                if (!all) continue;
                FilterMenu menu = new FilterMenu { Window = candidate, Bounds = bounds, Timestamp = Environment.TickCount };
                for (int i = 0; i < Math.Min(controls.Count, 2000); i++)
                {
                    AutomationElement element = controls[i];
                    if (element.Current.IsOffscreen) continue;
                    System.Windows.Rect r = element.Current.BoundingRectangle;
                    if (r.IsEmpty || r.Width < 1 || r.Height < 1) continue;
                    FilterRow row = new FilterRow { Name = element.Current.Name, Bounds = Rectangle.Round(new RectangleF((float)r.X, (float)r.Y, (float)r.Width, (float)r.Height)) };
                    if (IsAll(row.Name)) row.Name = "(全选)";
                    AutomationElement parent = element;
                    for (int depth = 0; depth < 7; depth++)
                    {
                        parent = TreeWalker.ControlViewWalker.GetParent(parent); if (parent == null || parent == root) break;
                        if (parent.Current.ControlType != ControlType.TreeItem) continue;
                        string name = parent.Current.Name; int year;
                        if (int.TryParse(name, out year) && year >= 1900 && year <= 9999) row.Year = name;
                        else if (String.IsNullOrEmpty(row.Month)) row.Month = name;
                    }
                    menu.Rows.Add(row);
                }
                if (menu.Rows.Count > 0) return menu;
            }
            return null;
        }
        private static bool IsAll(string name) { return name == "(全选)" || name == "全选" || name == "(全選)" || name == "(Select All)" || name == "Select All"; }
    }
}
