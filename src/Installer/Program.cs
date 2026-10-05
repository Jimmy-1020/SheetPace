using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
namespace SheetPaceInstaller
{
    internal static class Program
    {
        [STAThread] private static int Main(string[] args)
        {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                if (args.Length > 0 && args[0] == "/selftest")
                { if (args.Length != 2) return 2; Engine.SelfTest(Path.GetFullPath(args[1])); return 0; }
                if (args.Length > 0 && args[0] == "/extract")
                { Engine.VerifyPayload(); Directory.CreateDirectory(args[1]); File.WriteAllBytes(Path.Combine(args[1], "SheetPace.dll"), Engine.Payload()); return 0; }
                if (args.Length > 0 && args[0] == "/install-quiet") { Engine.Install(); return 0; }
                bool uninstall = args.Length > 0 && args[0] == "/uninstall";
                using (SetupForm form = new SetupForm(uninstall)) Application.Run(form);
                return 0;
            }
            catch (Exception ex)
            {
                if (args.Length > 1 && (args[0] == "/selftest" || args[0] == "/extract"))
                { Directory.CreateDirectory(args[1]); File.WriteAllText(Path.Combine(args[1], "error.txt"), ex.ToString()); }
                else MessageBox.Show(ex.Message, "SheetPace", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }
    }
    internal sealed class SetupForm : Form
    {
        private readonly bool uninstall;
        private readonly Button action;
        private readonly Label status;
        public SetupForm(bool remove)
        {
            uninstall = remove; Text = remove ? "卸载 SheetPace" : "安装 SheetPace";
            Font = new Font("Microsoft YaHei UI", 10); BackColor = Color.White; ClientSize = new Size(586, 454);
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen; AutoScaleMode = AutoScaleMode.Dpi;
            Label title = new Label { Text = "SheetPace", Font = new Font(Font.FontFamily, 27, FontStyle.Bold), ForeColor = Color.FromArgb(35, 105, 220), Location = new Point(28, 24), Size = new Size(520, 56) }; Controls.Add(title);
            Controls.Add(new Label { Text = "Excel 筛选计数与行列定位", Font = new Font(Font.FontFamily, 13), Location = new Point(30, 88), Size = new Size(520, 36) });
            Controls.Add(new Label { Text = remove ? "卸载将移除 Excel 加载项和程序文件。\n您的工作簿不会被修改，个人光影设置将保留。" : "• 筛选值数量：搜索、多选与次数统计\n• 鼠标行列光影：移动鼠标即可定位\n• 自定义颜色与透明度，实时预览\n\nWindows Excel 2016 及以上 · 32 / 64 位\n当前用户安装，无需管理员权限", Location = new Point(30, 140), Size = new Size(526, 170) });
            status = new Label { Text = "请先保存工作簿，并关闭所有 Excel 窗口。\n安装位置：" + Engine.InstallPath, ForeColor = Color.DimGray, Location = new Point(30, 310), Size = new Size(526, 74) }; Controls.Add(status);
            Button cancel = new Button { Text = "取消", Location = new Point(326, 396), Size = new Size(94, 36), FlatStyle = FlatStyle.Flat };
            cancel.Click += delegate { Close(); }; Controls.Add(cancel);
            action = new Button { Text = remove ? "确认卸载" : "安装 / 更新", Location = new Point(432, 396), Size = new Size(124, 36), FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(35, 105, 220), ForeColor = Color.White };
            action.Click += Act; Controls.Add(action); AcceptButton = action; CancelButton = cancel;
        }
        private void Act(object sender, EventArgs e)
        {
            action.Enabled = false;
            try
            {
                if (uninstall) Engine.Uninstall(); else Engine.Install();
                status.ForeColor = Color.FromArgb(30, 130, 75);
                status.Text = uninstall ? "卸载完成。" : "安装完成。重新打开 Excel，即可在功能区看到 SheetPace。";
                action.Text = "完成"; action.Click -= Act;
                action.Click += delegate { Close(); }; action.Enabled = true;
                // Cleanup also works when the window's close button is used.
                if (uninstall) FormClosed += delegate { Engine.ScheduleSelfCleanup(); };
            }
            catch (Exception ex) { status.ForeColor = Color.Firebrick; status.Text = ex.Message; action.Enabled = true; }
        }
    }
}
