using System;
using System.Collections.Generic;
using System.Drawing;
namespace SheetPace
{
    internal struct IndexSpan
    {
        internal int First, Last;
        internal IndexSpan(int first, int last) { First = first; Last = last; }
    }
    internal sealed class SelectionSnapshot
    {
        internal IndexSpan[] Rows, Columns;
        internal static IndexSpan[] Merge(List<IndexSpan> source)
        {
            source.Sort(delegate(IndexSpan a, IndexSpan b) { return a.First.CompareTo(b.First); });
            List<IndexSpan> result = new List<IndexSpan>();
            foreach (IndexSpan span in source)
            {
                if (result.Count == 0 || span.First > result[result.Count - 1].Last + 1) result.Add(span);
                else { IndexSpan previous = result[result.Count - 1]; previous.Last = Math.Max(previous.Last, span.Last); result[result.Count - 1] = previous; }
            }
            return result.ToArray();
        }
        internal static SelectionSnapshot Capture(dynamic selection)
        {
            List<IndexSpan> rows = new List<IndexSpan>(), columns = new List<IndexSpan>(); dynamic areas = null;
            try
            {
                areas = selection.Areas;
                int count = (int)areas.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic area = null, rowRange = null, columnRange = null, merged = null;
                    try
                    {
                        area = areas.Item(i); rowRange = area.Rows; columnRange = area.Columns;
                        int row = (int)area.Row, column = (int)area.Column, height = (int)rowRange.Count, width = (int)columnRange.Count;
                        if (height == 1 && width == 1)
                        {
                            merged = area.MergeArea; ExcelContext.Release(rowRange); ExcelContext.Release(columnRange);
                            rowRange = null; columnRange = null; rowRange = merged.Rows; columnRange = merged.Columns;
                            row = (int)merged.Row; column = (int)merged.Column; height = (int)rowRange.Count; width = (int)columnRange.Count;
                        }
                        rows.Add(new IndexSpan(row, row + height - 1)); columns.Add(new IndexSpan(column, column + width - 1));
                    }
                    finally { ExcelContext.Release(merged); ExcelContext.Release(columnRange); ExcelContext.Release(rowRange); ExcelContext.Release(area); }
                }
            }
            finally { ExcelContext.Release(areas); }
            return new SelectionSnapshot { Rows = Merge(rows), Columns = Merge(columns) };
        }
    }
    internal sealed class HighlightBands
    {
        internal Rectangle Grid;
        internal Rectangle[] Rows = new Rectangle[0], Columns = new Rectangle[0];
        internal bool Visible { get { return Rows.Length != 0 || Columns.Length != 0; } }
    }
    internal static class SelectionGeometry
    {
        private sealed class Axis
        {
            internal bool Horizontal;
            internal int Fixed, First, Last, MinIndex, MaxIndex;
        }
        private static int Index(dynamic window, bool horizontal, int variable, int fixedCoordinate)
        {
            dynamic hit = null;
            try
            {
                hit = window.RangeFromPoint(horizontal ? variable : fixedCoordinate, horizontal ? fixedCoordinate : variable);
                return hit == null ? 0 : (int)(horizontal ? hit.Column : hit.Row);
            }
            catch { return 0; }
            finally { ExcelContext.Release(hit); }
        }
        private static Axis VisibleAxis(dynamic window, Rectangle grid, bool horizontal)
        {
            int first = horizontal ? grid.Left : grid.Top, last = (horizontal ? grid.Right : grid.Bottom) - 1;
            int[] references = horizontal ? new[] { grid.Bottom - 48, grid.Top + grid.Height / 2, grid.Top + 80 }
                : new[] { grid.Right - 40, grid.Left + grid.Width / 2, grid.Left + 80 };
            foreach (int reference in references)
            {
                int seed = first + (last - first) / 2;
                if (Index(window, horizontal, seed, reference) == 0) continue;
                int low = first, high = seed;
                while (low < high) { int mid = low + (high - low) / 2; if (Index(window, horizontal, mid, reference) != 0) high = mid; else low = mid + 1; }
                int start = low; low = seed; high = last;
                while (low < high) { int mid = low + (high - low + 1) / 2; if (Index(window, horizontal, mid, reference) != 0) low = mid; else high = mid - 1; }
                int end = low;
                return new Axis { Horizontal = horizontal, Fixed = reference, First = start, Last = end,
                    MinIndex = Index(window, horizontal, start, reference), MaxIndex = Index(window, horizontal, end, reference) };
            }
            return null;
        }
        private static int LowerBound(dynamic window, Axis axis, int index)
        {
            int low = axis.First, high = axis.Last + 1;
            while (low < high)
            {
                int mid = low + (high - low) / 2;
                if (Index(window, axis.Horizontal, mid, axis.Fixed) < index) low = mid + 1; else high = mid;
            }
            return low;
        }
        private static void Project(dynamic window, Axis axis, IndexSpan[] spans, Rectangle grid, List<Rectangle> output)
        {
            foreach (IndexSpan span in spans)
            {
                if (span.Last < axis.MinIndex || span.First > axis.MaxIndex) continue;
                int first = LowerBound(window, axis, span.First), end = LowerBound(window, axis, span.Last + 1);
                if (end <= first) continue;
                output.Add(axis.Horizontal ? new Rectangle(first, grid.Top, end - first, grid.Height)
                    : new Rectangle(grid.Left, first, grid.Width, end - first));
            }
        }
        internal static HighlightBands Project(dynamic window, IList<Rectangle> grids, SelectionSnapshot selection)
        {
            List<Rectangle> rows = new List<Rectangle>(), columns = new List<Rectangle>(); Rectangle bounds = Rectangle.Empty;
            foreach (Rectangle candidate in grids)
            {
                Axis x = VisibleAxis(window, candidate, true), y = VisibleAxis(window, candidate, false);
                if (x == null || y == null) continue;
                // Keep row/column headers, exclude scrollbars and worksheet tabs.
                Rectangle grid = Rectangle.FromLTRB(candidate.Left, candidate.Top, x.Last + 1, y.Last + 1);
                Project(window, y, selection.Rows, grid, rows); Project(window, x, selection.Columns, grid, columns);
                bounds = bounds.IsEmpty ? grid : Rectangle.Union(bounds, grid);
            }
            return new HighlightBands { Grid = bounds, Rows = rows.ToArray(), Columns = columns.ToArray() };
        }
    }
}
