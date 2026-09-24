using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OrthoLink.Setup
{
    internal static class Program
    {
        [DllImport("user32.dll")] private static extern bool SetProcessDPIAware();

        /// <summary>
        /// No arguments: show the window. "/install" or "/uninstall" run directly;
        /// add "/quiet" to skip every dialog (exit code 0 = ok, 1 = failed, 2 = PowerPoint is running).
        /// </summary>
        [STAThread]
        private static int Main(string[] args)
        {
            try { SetProcessDPIAware(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool quiet = Has(args, "/quiet") || Has(args, "/silent") || Has(args, "/s");
            string action = Has(args, "/uninstall") ? "uninstall" : Has(args, "/install") ? "install" : null;

            if (action == null)
            {
                Application.Run(new SetupForm());
                return 0;
            }

            if (Installer.PowerPointRunning())
            {
                if (!quiet) MessageBox.Show("请先关闭 PowerPoint，再重新运行。", "OrthoLink 安装程序", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 2;
            }
            try
            {
                if (action == "install")
                {
                    Installer.Install(s => { });
                    if (!quiet) MessageBox.Show("安装完成。重新打开 PowerPoint 后，功能区会出现 OrthoLink 选项卡。", "OrthoLink 安装程序", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    if (!quiet && MessageBox.Show("确定要卸载 OrthoLink 吗？", "OrthoLink 安装程序", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return 0;
                    Installer.Uninstall(s => { });
                    if (!quiet) MessageBox.Show("OrthoLink 已卸载。", "OrthoLink 安装程序", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return 0;
            }
            catch (Exception ex)
            {
                Installer.Log(action + " FAILED: " + ex);
                if (!quiet) MessageBox.Show("操作失败：" + ex.Message + "\n\n详细记录在：" + Installer.LogPath, "OrthoLink 安装程序", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        private static bool Has(string[] args, string flag)
        {
            foreach (var a in args) if (string.Equals(a, flag, StringComparison.OrdinalIgnoreCase) || string.Equals(a, "-" + flag.Substring(1), StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }
    }

    /// <summary>One small window: what is installed, and Install / Uninstall / Close.</summary>
    internal sealed class SetupForm : Form
    {
        private readonly Label _status = new Label();
        private readonly Button _install = new Button();
        private readonly Button _uninstall = new Button();
        private readonly Button _close = new Button();

        public SetupForm()
        {
            Text = "OrthoLink 安装程序";
            Font = new Font("Microsoft YaHei UI", 9f);
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(460, 236);
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            var pic = new PictureBox { Bounds = new Rectangle(20, 20, 56, 56), SizeMode = PictureBoxSizeMode.Zoom };
            try { pic.Image = Icon.ExtractAssociatedIcon(Application.ExecutablePath).ToBitmap(); } catch { }
            var title = new Label { Text = "OrthoLink", Font = new Font("Microsoft YaHei UI", 16f, FontStyle.Bold), AutoSize = true, Location = new Point(88, 18) };
            var sub = new Label { Text = "PowerPoint 圆角直角连接线插件    版本 " + Ver.Display, AutoSize = true, Location = new Point(90, 54), ForeColor = Color.FromArgb(90, 90, 90) };
            _status.Bounds = new Rectangle(20, 96, 420, 40);
            var hint = new Label
            {
                Text = "安装到当前用户，不需要管理员权限。安装或卸载前请先关闭 PowerPoint。",
                Bounds = new Rectangle(20, 140, 420, 40), ForeColor = Color.FromArgb(90, 90, 90),
            };

            _install.Bounds = new Rectangle(160, 192, 88, 30); _install.Click += (s, e) => Run(true);
            _uninstall.Text = "卸载"; _uninstall.Bounds = new Rectangle(256, 192, 88, 30); _uninstall.Click += (s, e) => Run(false);
            _close.Text = "关闭"; _close.Bounds = new Rectangle(352, 192, 88, 30); _close.Click += (s, e) => Close();
            CancelButton = _close;

            Controls.AddRange(new Control[] { pic, title, sub, _status, hint, _install, _uninstall, _close });
            RefreshState();
        }

        private void RefreshState()
        {
            string v = Installer.InstalledVersion();
            if (v == null)
            {
                _status.Text = "状态：未安装";
                _install.Text = "安装"; _uninstall.Enabled = false;
            }
            else
            {
                _status.Text = "状态：已安装 " + v + (v == Ver.Display ? "" : "，将更新到 " + Ver.Display) + "\n位置：" + Installer.InstallDir;
                _install.Text = v == Ver.Display ? "重新安装" : "更新"; _uninstall.Enabled = true;
            }
            AcceptButton = _install;
        }

        private void Run(bool install)
        {
            while (Installer.PowerPointRunning())
            {
                if (MessageBox.Show(this, "PowerPoint 正在运行，请先关闭它，然后点“重试”。", "OrthoLink 安装程序",
                        MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning) != DialogResult.Retry) return;
            }
            if (!install && MessageBox.Show(this, "确定要卸载 OrthoLink 吗？", "OrthoLink 安装程序", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            _install.Enabled = _uninstall.Enabled = _close.Enabled = false;
            UseWaitCursor = true;
            try
            {
                Action<string> progress = s => { _status.Text = s; _status.Refresh(); };
                if (install)
                {
                    Installer.Install(progress);
                    RefreshState();
                    MessageBox.Show(this, "安装完成。\n\n重新打开 PowerPoint 后，功能区会出现 OrthoLink 选项卡。", "OrthoLink 安装程序", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    Installer.Uninstall(progress);
                    RefreshState();
                    MessageBox.Show(this, "OrthoLink 已卸载。", "OrthoLink 安装程序", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Close();
                }
            }
            catch (Exception ex)
            {
                Installer.Log((install ? "install" : "uninstall") + " FAILED: " + ex);
                RefreshState();
                MessageBox.Show(this, "操作失败：" + ex.Message + "\n\n详细记录在：" + Installer.LogPath, "OrthoLink 安装程序", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
                _install.Enabled = _close.Enabled = true;
                _uninstall.Enabled = Installer.InstalledVersion() != null;
            }
        }
    }
}
