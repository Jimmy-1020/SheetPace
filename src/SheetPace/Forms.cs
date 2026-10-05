using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
namespace SheetPace
{
    internal static class Theme
    {
        public static readonly Color Blue = Color.FromArgb(35, 105, 220);
        public static void Apply(Form form)
        {
            form.Font = new Font("Microsoft YaHei UI", 9F); form.BackColor = Color.White;
            form.StartPosition = FormStartPosition.CenterParent; form.AutoScaleMode = AutoScaleMode.Dpi;
            form.ShowInTaskbar = false; form.MinimizeBox = false; form.MaximizeBox = false;
        }
        public static Button Button(string text, int x, int y, int width, bool primary)
        {
            Button b = new Button { Text = text, Location = new Point(x, y), Size = new Size(width, 34), FlatStyle = FlatStyle.Flat };
            b.FlatAppearance.BorderColor = primary ? Blue : Color.FromArgb(210, 215, 225);
            if (primary) { b.BackColor = Blue; b.ForeColor = Color.White; }
            return b;
        }
        public static Label Label(string text, int x, int y, int width, int height)
        { return new Label { Text = text, Location = new Point(x, y), Size = new Size(width, height) }; }
    }
    internal sealed class FilterForm : Form
    {
        private readonly ExcelContext context;
        private readonly HashSet<string> selected;
        private readonly ListView list;
        private readonly TextBox search;
        private readonly Label summary;
        private bool populating;
        public FilterForm(ExcelContext source)
        {
            context = source; selected = context.CurrentSelection(); Theme.Apply(this);
            Text = "SheetPace · 计数筛选"; ClientSize = new Size(466, 606); FormBorderStyle = FormBorderStyle.FixedDialog;
            Label title = Theme.Label(String.IsNullOrEmpty(context.Title) ? "当前列" : context.Title, 22, 20, 420, 30);
            title.Font = new Font(Font.FontFamily, 15, FontStyle.Bold); Controls.Add(title);
            Controls.Add(Theme.Label(context.SheetName + " · " + context.Address + " · 排除标题，包含隐藏行", 22, 58, 426, 24));
            search = new TextBox { Location = new Point(22, 90), Size = new Size(422, 28) }; Controls.Add(search);
            Controls.Add(Theme.Label("搜索筛选值", 24, 122, 400, 24)); search.TextChanged += delegate { Populate(); };
            Button all = Theme.Button("全选", 22, 150, 80, false), none = Theme.Button("清空", 112, 150, 80, false);
            Button matches = Theme.Button("只选搜索结果", 202, 150, 130, false);
            all.Click += delegate { foreach (CountItem item in context.Snapshot.Items) selected.Add(item.Label); Populate(); };
            none.Click += delegate { selected.Clear(); Populate(); };
            matches.Click += delegate { selected.Clear(); foreach (CountItem item in Matches()) selected.Add(item.Label); Populate(); };
            Controls.Add(all); Controls.Add(none); Controls.Add(matches);
            list = new ListView { Location = new Point(22, 196), Size = new Size(422, 292), View = View.Details, CheckBoxes = true, FullRowSelect = true, HideSelection = false };
            list.Columns.Add("筛选值", 310); list.Columns.Add("数量", 86, HorizontalAlignment.Right);
            list.ItemChecked += delegate(object sender, ItemCheckedEventArgs e)
            {
                if (populating) return;
                string key = (string)e.Item.Tag;
                if (e.Item.Checked) selected.Add(key); else selected.Remove(key);
                RefreshSummary();
            };
            Controls.Add(list); summary = Theme.Label("", 22, 502, 422, 38); Controls.Add(summary);
            Button cancel = Theme.Button("取消", 228, 555, 100, false), apply = Theme.Button("应用筛选", 340, 555, 104, true);
            cancel.DialogResult = DialogResult.Cancel; CancelButton = cancel; AcceptButton = apply;
            apply.Click += delegate
            {
                try { context.Apply(selected); DialogResult = DialogResult.OK; Close(); }
                catch (Exception ex) { MessageBox.Show(this, ex.Message, "无法应用筛选", MessageBoxButtons.OK, MessageBoxIcon.Information); }
            };
            Controls.Add(cancel); Controls.Add(apply); Populate();
        }
        private IEnumerable<CountItem> Matches() { return context.Snapshot.Items.Where(i => i.Label.IndexOf(search.Text, StringComparison.CurrentCultureIgnoreCase) >= 0); }
        private void Populate()
        {
            populating = true; list.BeginUpdate();
            try
            {
                list.Items.Clear();
                foreach (CountItem item in Matches())
                {
                    ListViewItem row = new ListViewItem(item.Label); row.Tag = item.Label;
                    row.SubItems.Add(item.Count.ToString("N0")); row.Checked = selected.Contains(item.Label); list.Items.Add(row);
                }
            }
            finally { list.EndUpdate(); populating = false; }
            RefreshSummary();
        }
        private void RefreshSummary()
        {
            summary.Text = "共 " + context.Snapshot.Total.ToString("N0") + " 行 · " + context.Snapshot.Items.Count.ToString("N0") + " 个值 · 已选 " + selected.Count + " 个";
            if (!context.SelectionKnown) summary.Text += Environment.NewLine + "现有条件无法显示；应用将替换本列条件。";
            if (context.MixedFormats) summary.Text += "\n本列格式混合，按原始值统计。";
        }
    }
    internal sealed class SettingsForm : Form
    {
        public readonly Settings Result;
        private readonly Panel preview;
        public SettingsForm(Settings settings)
        {
            Result = settings.Copy(); Theme.Apply(this); Text = "SheetPace · 光影设置";
            ClientSize = new Size(454, 600); FormBorderStyle = FormBorderStyle.FixedDialog;
            Label title = Theme.Label("光影设置", 24, 20, 400, 32);
            title.Font = new Font(Font.FontFamily, 15, FontStyle.Bold); Controls.Add(title);
            CheckBox hover = new CheckBox { Text = "启用鼠标光影", Checked = Result.HoverEnabled, Location = new Point(24, 68), AutoSize = true };
            Controls.Add(hover);
            Controls.Add(Theme.Label("光影模式", 24, 110, 100, 26));
            RadioButton selected = new RadioButton { Name = "selectionMode", Text = "点击选中（默认）", Checked = Result.Mode == HighlightMode.ClickSelection, Location = new Point(24, 142), AutoSize = true };
            RadioButton follow = new RadioButton { Name = "followMode", Text = "鼠标跟随", Checked = Result.Mode == HighlightMode.FollowMouse, Location = new Point(254, 142), AutoSize = true };
            Label hint = Theme.Label("", 24, 178, 406, 40);
            Action updateMode = delegate
            {
                Result.Mode = follow.Checked ? HighlightMode.FollowMouse : HighlightMode.ClickSelection;
                hint.Text = follow.Checked ? "鼠标移到哪个单元格，十字光影就跟随到该格。" : "十字光影定位当前选中的单元格，移动鼠标不改变位置。";
            };
            selected.CheckedChanged += delegate { updateMode(); }; follow.CheckedChanged += delegate { updateMode(); };
            Controls.Add(selected); Controls.Add(follow); Controls.Add(hint); updateMode();
            CheckBox counts = new CheckBox { Text = "启用原生筛选数量标签", Checked = Result.NativeCountsEnabled, Location = new Point(24, 224), AutoSize = true };
            hover.CheckedChanged += delegate { Result.HoverEnabled = hover.Checked; preview.Invalidate(); };
            counts.CheckedChanged += delegate { Result.NativeCountsEnabled = counts.Checked; };
            Controls.Add(counts);
            Controls.Add(Theme.Label("光影颜色", 24, 266, 100, 26));
            Button color = Theme.Button("选择颜色", 320, 260, 110, false); color.BackColor = Result.HighlightColor;
            color.Click += delegate
            {
                using (ColorDialog picker = new ColorDialog { Color = Result.HighlightColor, FullOpen = true })
                    if (picker.ShowDialog(this) == DialogResult.OK) { Result.HighlightColor = picker.Color; color.BackColor = picker.Color; preview.Invalidate(); }
            };
            Controls.Add(color);
            Label opacity = Theme.Label("透明度 " + Result.Transparency + "%", 24, 308, 200, 25); Controls.Add(opacity);
            TrackBar slider = new TrackBar { Minimum = 0, Maximum = 100, Value = Result.Transparency, TickFrequency = 10, Location = new Point(18, 338), Size = new Size(418, 44) };
            slider.ValueChanged += delegate { Result.Transparency = slider.Value; opacity.Text = "透明度 " + Result.Transparency + "%"; preview.Invalidate(); }; Controls.Add(slider);
            preview = new Panel { Location = new Point(24, 398), Size = new Size(406, 118), BorderStyle = BorderStyle.FixedSingle };
            preview.Paint += delegate(object sender, PaintEventArgs e)
            {
                int width = preview.ClientSize.Width, height = preview.ClientSize.Height;
                using (Pen pen = new Pen(Color.FromArgb(215, 224, 238)))
                {
                    for (int x = 0; x < width; x += 81) e.Graphics.DrawLine(pen, x, 0, x, height);
                    for (int y = 0; y < height; y += 29) e.Graphics.DrawLine(pen, 0, y, width, y);
                }
                if (Result.HoverEnabled)
                    using (Brush brush = new SolidBrush(Color.FromArgb(Result.Alpha, Result.HighlightColor)))
                    using (Region region = new Region(new Rectangle(0, 29, width, 29)))
                    { region.Union(new Rectangle(162, 0, 81, height)); e.Graphics.FillRegion(brush, region); }
                e.Graphics.DrawString("单元格内容", Font, Brushes.DimGray, 171, 34);
            };
            Controls.Add(preview);
            Button cancel = Theme.Button("取消", 214, 544, 100, false), save = Theme.Button("保存设置", 326, 544, 104, true);
            cancel.DialogResult = DialogResult.Cancel; save.DialogResult = DialogResult.OK;
            Controls.Add(cancel); Controls.Add(save); CancelButton = cancel; AcceptButton = save;
        }
    }
    internal sealed class WindowOwner : IWin32Window
    {
        public IntPtr Handle { get; private set; }
        public WindowOwner(IntPtr hwnd) { Handle = hwnd; }
    }
}
