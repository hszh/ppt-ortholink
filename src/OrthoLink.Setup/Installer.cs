using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OrthoLink.Setup
{
    /// <summary>
    /// Everything the installer does, with no UI. Installs for the current user only:
    /// files go to %LOCALAPPDATA%\OrthoLink, registration goes to HKCU. No admin rights needed.
    /// </summary>
    internal static class Installer
    {
        public const string ProgId = "OrthoLink.AddIn";
        public const string Clsid = "{7C3B6C2E-5A1D-4F0B-9B7E-0D2A6E4F1A11}";
        public const string ClassName = "OrthoLink.Connect";
        public const string DisplayName = "OrthoLink";
        private const string DotNetCategory = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}";
        private const string AddinKey = @"Software\Microsoft\Office\PowerPoint\Addins\" + ProgId;
        private const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\OrthoLink";
        private const string SettingsKey = @"Software\OrthoLink";

        public static string InstallDir
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OrthoLink"); }
        }
        public static string DllPath { get { return Path.Combine(InstallDir, "OrthoLink.AddIn.dll"); } }
        public static string SetupCopyPath { get { return Path.Combine(InstallDir, "OrthoLink-Setup.exe"); } }
        public static string LogPath { get { return Path.Combine(Path.GetTempPath(), "OrthoLink-Setup.log"); } }

        /// <summary>Version currently installed, or null.</summary>
        public static string InstalledVersion()
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(UninstallKey))
                {
                    var v = k == null ? null : k.GetValue("DisplayVersion") as string;
                    if (!string.IsNullOrEmpty(v) && File.Exists(DllPath)) return v;
                }
                if (File.Exists(DllPath)) return AssemblyName.GetAssemblyName(DllPath).Version.ToString(3);
            }
            catch { }
            return null;
        }

        public static bool PowerPointRunning()
        {
            try { return Process.GetProcessesByName("POWERPNT").Length > 0; } catch { return false; }
        }

        public static void Install(Action<string> progress)
        {
            Log("install " + Ver.Display + " -> " + InstallDir);
            progress("正在复制文件…");
            Directory.CreateDirectory(InstallDir);
            Extract("OrthoLink.Setup.OrthoLink.AddIn.dll", DllPath);
            DeleteFile(Path.Combine(InstallDir, "template.pptx"));   // 1.0.0 kept the connector templates there; they are inside the DLL now
            CopySelf();

            progress("正在注册插件…");
            RegisterCom();
            using (var k = Registry.CurrentUser.CreateSubKey(AddinKey))
            {
                k.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                k.SetValue("FriendlyName", DisplayName);
                k.SetValue("Description", "PowerPoint 圆角直角连接线");
            }
            // ask Office not to auto-disable us after a crash elsewhere in the process
            foreach (var ver in new[] { "16.0" })
            {
                try
                {
                    using (var k = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\" + ver + @"\PowerPoint\Resiliency\DoNotDisableAddinList"))
                        k.SetValue(ProgId, 1, RegistryValueKind.DWord);
                }
                catch { }
            }

            progress("正在登记到“应用”列表…");
            using (var k = Registry.CurrentUser.CreateSubKey(UninstallKey))
            {
                k.SetValue("DisplayName", DisplayName);
                k.SetValue("DisplayVersion", Ver.Display);
                k.SetValue("Publisher", "OrthoLink");
                k.SetValue("DisplayIcon", SetupCopyPath + ",0");
                k.SetValue("InstallLocation", InstallDir);
                k.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
                k.SetValue("UninstallString", "\"" + SetupCopyPath + "\" /uninstall");
                k.SetValue("QuietUninstallString", "\"" + SetupCopyPath + "\" /uninstall /quiet");
                k.SetValue("NoModify", 1, RegistryValueKind.DWord);
                k.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                k.SetValue("EstimatedSize", (int)(DirSize(InstallDir) / 1024), RegistryValueKind.DWord);
            }
            Log("install done");
        }

        public static void Uninstall(Action<string> progress)
        {
            Log("uninstall from " + InstallDir);
            progress("正在移除注册信息…");
            UnregisterCom();
            Registry.CurrentUser.DeleteSubKeyTree(AddinKey, false);
            Registry.CurrentUser.DeleteSubKeyTree(UninstallKey, false);
            Registry.CurrentUser.DeleteSubKeyTree(SettingsKey, false);
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\PowerPoint\Resiliency\DoNotDisableAddinList", true))
                    if (k != null) k.DeleteValue(ProgId, false);
            }
            catch { }

            progress("正在删除文件…");
            RemoveInstallDir();
            Log("uninstall done");
        }

        // ------------------------------------------------------------------ COM registration (what RegAsm /codebase would write)

        private static void RegisterCom()
        {
            string asm = "OrthoLink.AddIn, Version=" + Ver.Assembly + ", Culture=neutral, PublicKeyToken=null";
            string codeBase = new Uri(DllPath).AbsoluteUri;
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                using (var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                using (var classes = hkcu.CreateSubKey(@"Software\Classes"))
                {
                    using (var p = classes.CreateSubKey(ProgId))
                    {
                        p.SetValue("", ClassName);
                        using (var c = p.CreateSubKey("CLSID")) c.SetValue("", Clsid);
                    }
                    using (var k = classes.CreateSubKey(@"CLSID\" + Clsid))
                    {
                        k.SetValue("", ClassName);
                        using (var p = k.CreateSubKey("ProgId")) p.SetValue("", ProgId);
                        using (var c = k.CreateSubKey(@"Implemented Categories\" + DotNetCategory)) { }
                        using (var ip = k.CreateSubKey("InprocServer32"))
                        {
                            ip.SetValue("", "mscoree.dll");
                            ip.SetValue("ThreadingModel", "Both");
                            WriteClrValues(ip, asm, codeBase);
                            using (var v = ip.CreateSubKey(Ver.Assembly)) WriteClrValues(v, asm, codeBase);
                        }
                    }
                }
            }
        }

        private static void WriteClrValues(RegistryKey k, string asm, string codeBase)
        {
            k.SetValue("Class", ClassName);
            k.SetValue("Assembly", asm);
            k.SetValue("RuntimeVersion", "v4.0.30319");
            k.SetValue("CodeBase", codeBase);
        }

        private static void UnregisterCom()
        {
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                using (var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                using (var classes = hkcu.OpenSubKey(@"Software\Classes", true))
                {
                    if (classes == null) continue;
                    classes.DeleteSubKeyTree(ProgId, false);
                    classes.DeleteSubKeyTree(@"CLSID\" + Clsid, false);
                }
            }
        }

        // ------------------------------------------------------------------ files

        private static void Extract(string resource, string target)
        {
            using (var s = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource))
            {
                if (s == null) throw new InvalidOperationException("安装包损坏：缺少 " + resource);
                using (var f = new FileStream(target, FileMode.Create, FileAccess.Write)) s.CopyTo(f);
            }
        }

        private static void CopySelf()
        {
            string me = Assembly.GetExecutingAssembly().Location;
            if (string.Equals(Path.GetFullPath(me), Path.GetFullPath(SetupCopyPath), StringComparison.OrdinalIgnoreCase)) return;
            File.Copy(me, SetupCopyPath, true);
            DeleteFile(SetupCopyPath + ":Zone.Identifier");   // drop the "downloaded from the internet" mark on the copy
        }

        /// <summary>Delete our folder. If this very program runs from inside it, hand the job to a detached cmd.exe.</summary>
        private static void RemoveInstallDir()
        {
            string dir = InstallDir;
            if (!Directory.Exists(dir)) return;
            if (!File.Exists(DllPath) && !File.Exists(SetupCopyPath)) return;   // not ours, leave it alone
            string me = Assembly.GetExecutingAssembly().Location;
            bool insideDir = Path.GetFullPath(me).StartsWith(Path.GetFullPath(dir) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
            if (!insideDir)
            {
                Directory.Delete(dir, true);
                return;
            }
            var psi = new ProcessStartInfo("cmd.exe", "/c timeout /t 2 /nobreak >nul & rmdir /s /q \"" + dir + "\"")
            {
                CreateNoWindow = true, UseShellExecute = false, WindowStyle = ProcessWindowStyle.Hidden,
            };
            Process.Start(psi);
        }

        private static long DirSize(string dir)
        {
            long n = 0;
            foreach (var f in Directory.GetFiles(dir)) n += new FileInfo(f).Length;
            return n;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool DeleteFile(string path);

        public static void Log(string line)
        {
            try { File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + line + Environment.NewLine); } catch { }
        }
    }
}
