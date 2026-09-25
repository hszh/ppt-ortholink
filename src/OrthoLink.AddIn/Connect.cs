using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OrthoLink
{
    /// <summary>
    /// The COM add-in entry point. PowerPoint creates this object at startup (IDTExtensibility2),
    /// asks it for ribbon XML (IRibbonExtensibility) and calls the public methods below by name.
    /// </summary>
    [ComVisible(true)]
    [Guid("7C3B6C2E-5A1D-4F0B-9B7E-0D2A6E4F1A11")]
    [ProgId("OrthoLink.AddIn")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Connect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        public const string VersionString = Ver.Display;

        private PowerPoint.Application _app;
        private ConnectorService _svc;
        private Office.IRibbonUI _ribbon;
        private int _busy;                   // > 0 while one of our own operations runs; follow checks wait for it
        private string _lastSelection = "";
        private DateTime _lastCheckError = DateTime.MinValue;

        // follow-up after user input instead of polling: see StartWatching
        private const int CheckDelayMs = 100;
        private Timer _check;                // one-shot, restarted by every mouse / key release
        private HookProc _hookProc;          // kept in a field so it is not garbage collected while Windows calls it
        private IntPtr _hook = IntPtr.Zero;

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetWindowsHookEx(int idHook, HookProc fn, IntPtr module, uint threadId);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int vKey);
        private const int WH_GETMESSAGE = 3, PM_REMOVE = 1, VK_LBUTTON = 0x01;
        private const int WM_KEYUP = 0x0101, WM_SYSKEYUP = 0x0105, WM_LBUTTONUP = 0x0202, WM_RBUTTONUP = 0x0205,
                          WM_MBUTTONUP = 0x0208, WM_POINTERUP = 0x0247;

        private static readonly Side?[] SideItems = { null, Side.Left, Side.Right, Side.Top, Side.Bottom };

        // ================================================================ IDTExtensibility2

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                _app = (PowerPoint.Application)application;
                Settings.Load();
                _svc = new ConnectorService(_app);
                try { ((Office.COMAddIn)addInInst).Object = this; } catch (Exception ex) { Log.Error("expose Object", ex); }
                Log.Write("OnConnection ok, mode=" + connectMode + ", version " + VersionString);
                if (connectMode != ext_ConnectMode.ext_cm_Startup) StartWatching();
            }
            catch (Exception ex) { Log.Error("OnConnection", ex); }
        }

        public void OnStartupComplete(ref Array custom) { StartWatching(); }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                StopWatching();
                if (_svc != null) _svc.Shutdown();
                Log.Write("OnDisconnection " + removeMode);
            }
            catch (Exception ex) { Log.Error("OnDisconnection", ex); }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ================================================================ IRibbonExtensibility

        public string GetCustomUI(string ribbonId)
        {
            using (var s = Assembly.GetExecutingAssembly().GetManifestResourceStream("OrthoLink.Ribbon.xml"))
            using (var r = new StreamReader(s))
                return r.ReadToEnd();
        }

        public void Ribbon_Load(Office.IRibbonUI ribbon) { _ribbon = ribbon; }

        /// <summary>Serves the embedded icon PNGs for controls that use image="..." in Ribbon.xml.</summary>
        public object LoadImage(string imageId)
        {
            try
            {
                using (var s = Assembly.GetExecutingAssembly().GetManifestResourceStream("OrthoLink." + imageId + ".png"))
                {
                    if (s == null) { Log.Write("LoadImage: no resource for " + imageId); return null; }
                    return PictureConverter.ToPictureDisp(System.Drawing.Image.FromStream(s));
                }
            }
            catch (Exception ex) { Log.Error("LoadImage " + imageId, ex); return null; }
        }

        // ================================================================ ribbon callbacks: actions

        public void OnConnect(Office.IRibbonControl control) { Guard("OnConnect", () => ConnectSelected()); }

        public void OnReroute(Office.IRibbonControl control)
        {
            Guard("OnReroute", () =>
            {
                int n = 0;
                foreach (var c in SelectedLinks()) { if (_svc.Update(SlideOf(c), c, true) != null) n++; }
                if (n == 0) Info("请先选中一条由 OrthoLink 生成的连接线。");
                Refresh();
            });
        }

        public void OnUpdateAll(Office.IRibbonControl control)
        {
            Guard("OnUpdateAll", () =>
            {
                var slide = CurrentSlide();
                if (slide != null) _svc.UpdateAll(slide, false);
            });
        }

        public void OnAutoFollow(Office.IRibbonControl control, bool pressed) { Settings.AutoFollow = pressed; Settings.Save(); }

        public void OnSrcSide(Office.IRibbonControl control, string selectedId, int selectedIndex) { Guard("OnSrcSide", () => SetSides(true, selectedIndex)); }
        public void OnDstSide(Office.IRibbonControl control, string selectedId, int selectedIndex) { Guard("OnDstSide", () => SetSides(false, selectedIndex)); }

        public void OnRadiusChange(Office.IRibbonControl control, string text)
        {
            Guard("OnRadiusChange", () =>
            {
                float? r = ParseRadius(text);
                if (!r.HasValue) { Info("请输入一个数字（单位 pt），例如 8。"); Refresh(); return; }
                foreach (var c in SelectedLinks()) _svc.SetRadius(c, r.Value);
                Refresh();
            });
        }

        public void OnDefaultRadiusChange(Office.IRibbonControl control, string text)
        {
            Guard("OnDefaultRadiusChange", () =>
            {
                float? r = ParseRadius(text);
                if (!r.HasValue) { Info("请输入一个数字（单位 pt），例如 8。"); Refresh(); return; }
                Settings.RadiusPt = r.Value; Settings.Save();
                Refresh();
            });
        }

        public void OnGapChange(Office.IRibbonControl control, string text)
        {
            Guard("OnGapChange", () =>
            {
                float? g = ParseRadius(text);
                if (!g.HasValue) { Info("请输入一个数字（单位 pt），例如 4。"); Refresh(); return; }
                foreach (var c in SelectedLinks()) _svc.SetGap(SlideOf(c), c, g.Value);
                Refresh();
            });
        }

        public void OnDefaultGapChange(Office.IRibbonControl control, string text)
        {
            Guard("OnDefaultGapChange", () =>
            {
                float? g = ParseRadius(text);
                if (!g.HasValue) { Info("请输入一个数字（单位 pt），例如 4。"); Refresh(); return; }
                Settings.GapPt = g.Value; Settings.Save();
                Refresh();
            });
        }

        public string GetGapText(Office.IRibbonControl control)
        {
            var c = FirstSelectedLink();
            if (c == null) return "";
            try { return RadiusText(ConnectorService.GetGap(c)); } catch { return ""; }
        }

        public string GetDefaultGapText(Office.IRibbonControl control) { return RadiusText(Settings.GapPt); }

        /// <summary>Accepts "8", "8pt", "8 pt", "直角"; returns null when it is not a number.</summary>
        public static float? ParseRadius(string text)
        {
            if (text == null) return null;
            string t = text.Trim().ToLowerInvariant().Replace("pt", "").Replace("磅", "").Trim();
            if (t == "直角" || t == "无") return 0f;
            float v;
            if (!float.TryParse(t, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out v)) return null;
            if (v < 0) v = 0; if (v > 200) v = 200;
            return v;
        }

        private static string RadiusText(float r)
        {
            return (Math.Abs(r - Math.Round(r)) < 0.01 ? ((int)Math.Round(r)).ToString() : r.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture)) + " pt";
        }

        public void OnArrowEnd(Office.IRibbonControl control, bool pressed)
        {
            Guard("OnArrowEnd", () =>
            {
                foreach (var c in SelectedLinks()) _svc.SetArrowEnd(c, pressed);
                Settings.ArrowEnd = pressed; Settings.Save();
            });
        }

        public void OnArrowStart(Office.IRibbonControl control, bool pressed)
        {
            Guard("OnArrowStart", () =>
            {
                foreach (var c in SelectedLinks()) _svc.SetArrowStart(c, pressed);
                Settings.ArrowStart = pressed; Settings.Save();
            });
        }

        // ================================================================ ribbon callbacks: state (reflect the selected connector)

        public bool GetAutoFollow(Office.IRibbonControl control) { return Settings.AutoFollow; }
        public bool GetHasLink(Office.IRibbonControl control) { return FirstSelectedLink() != null; }

        public string GetRadiusText(Office.IRibbonControl control)
        {
            var c = FirstSelectedLink();
            if (c == null) return "";
            try { return RadiusText(ConnectorService.GetRadius(c)); } catch { return ""; }
        }

        public string GetDefaultRadiusText(Office.IRibbonControl control) { return RadiusText(Settings.RadiusPt); }

        public bool GetArrowEnd(Office.IRibbonControl control)
        {
            var c = FirstSelectedLink();
            if (c != null) { try { return ConnectorService.GetArrowEnd(c); } catch { } }
            return Settings.ArrowEnd;
        }

        public bool GetArrowStart(Office.IRibbonControl control)
        {
            var c = FirstSelectedLink();
            if (c != null) { try { return ConnectorService.GetArrowStart(c); } catch { } }
            return Settings.ArrowStart;
        }

        public int GetSrcSide(Office.IRibbonControl control) { return SideIndex(true); }
        public int GetDstSide(Office.IRibbonControl control) { return SideIndex(false); }

        private int SideIndex(bool source)
        {
            var c = FirstSelectedLink();
            if (c == null) return 0;
            var l = Link.Read(c);
            if (l == null) return 0;
            bool pin = source ? l.SrcPin : l.DstPin;
            if (!pin) return 0;
            var side = source ? l.SrcSide : l.DstSide;
            return Array.IndexOf(SideItems, (Side?)side);
        }

        private void SetSides(bool source, int index)
        {
            var side = SideItems[index];
            var links = SelectedLinks();
            if (links.Count == 0) { Info("请先选中一条连接线。"); return; }
            foreach (var c in links) _svc.SetSide(SlideOf(c), c, source, side);
            Refresh();
        }

        private void Refresh() { try { if (_ribbon != null) _ribbon.Invalidate(); } catch { } }

        // ================================================================ automation API (also used by tests)

        public string Version { get { return VersionString; } }

        /// <summary>Connect the two currently selected shapes (first selected = start). Returns 1 on success.</summary>
        public int ConnectSelected()
        {
            return Busy(() =>
            {
                var sel = _app.ActiveWindow.Selection;
                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count != 2)
                {
                    Info("请先选中两个形状：先点起点形状，再按住 Ctrl 点终点形状，然后点“连接”。");
                    return 0;
                }
                var a = sel.ShapeRange[1]; var b = sel.ShapeRange[2];
                if (Link.IsLink(a) || Link.IsLink(b)) { Info("选中的形状里有连接线，请选两个普通形状。"); return 0; }
                var conn = _svc.Create(SlideOf(a), a, b);
                conn.Select();
                Refresh();
                return 1;
            });
        }

        public object ConnectShapes(object srcShape, object dstShape)
        {
            var a = (PowerPoint.Shape)srcShape; var b = (PowerPoint.Shape)dstShape;
            return Busy(() => (object)_svc.Create(SlideOf(a), a, b));
        }

        /// <summary>side: "" or "Auto" = automatic, else Left/Right/Top/Bottom.</summary>
        public object SetConnectorSide(object connector, bool source, string side)
        {
            var c = (PowerPoint.Shape)connector;
            Side? s = null; Side parsed;
            if (!string.IsNullOrEmpty(side) && !side.Equals("Auto", StringComparison.OrdinalIgnoreCase) && Enum.TryParse(side, true, out parsed)) s = parsed;
            return Busy(() => (object)_svc.SetSide(SlideOf(c), c, source, s));
        }

        public int UpdateSlide(object slide, bool resetMiddle) { return Busy(() => _svc.UpdateAll((PowerPoint.Slide)slide, resetMiddle)); }
        public int FollowSlide(object slide) { return Busy(() => _svc.Follow((PowerPoint.Slide)slide)); }
        public bool AutoFollow { get { return Settings.AutoFollow; } set { Settings.AutoFollow = value; } }
        public float RadiusPt { get { return Settings.RadiusPt; } set { Settings.RadiusPt = value; Settings.Save(); } }
        public float GapPt { get { return Settings.GapPt; } set { Settings.GapPt = value; Settings.Save(); } }
        public object SetConnectorGap(object connector, float gapPt) { var c = (PowerPoint.Shape)connector; return Busy(() => (object)_svc.SetGap(SlideOf(c), c, gapPt)); }
        public float ParseRadiusPt(string text) { var r = ParseRadius(text); return r.HasValue ? r.Value : -1f; }
        public bool ArrowEnd { get { return Settings.ArrowEnd; } set { Settings.ArrowEnd = value; } }
        public bool ArrowStart { get { return Settings.ArrowStart; } set { Settings.ArrowStart = value; } }
        public string LogPath { get { return Log.Path; } }
        public void ActivateTab() { try { if (_ribbon != null) _ribbon.ActivateTab("olTab"); } catch (Exception ex) { Log.Error("ActivateTab", ex); } }

        // ================================================================ auto-follow after user input, ribbon refresh on selection change

        /// <summary>PowerPoint does not tell add-ins that a shape moved. Instead of polling, we watch the messages
        /// PowerPoint's own window thread takes from its queue: releasing a mouse button or a key is when shapes may
        /// have been dragged, resized, nudged, aligned or undone, so a moment later the current slide is checked once.
        /// Nothing runs while the user does nothing.</summary>
        private void StartWatching()
        {
            if (_check != null) return;
            _check = new Timer { Interval = CheckDelayMs };
            _check.Tick += (s, e) => { _check.Stop(); Check(); };
            _hookProc = OnThreadMessage;
            _hook = SetWindowsHookEx(WH_GETMESSAGE, _hookProc, IntPtr.Zero, GetCurrentThreadId());
            if (_hook == IntPtr.Zero) Log.Write("SetWindowsHookEx failed, error " + Marshal.GetLastWin32Error());
            try { ((PowerPoint.EApplication_Event)_app).WindowSelectionChange += OnWindowSelectionChange; }
            catch (Exception ex) { Log.Error("subscribe WindowSelectionChange", ex); }
        }

        private void StopWatching()
        {
            if (_hook != IntPtr.Zero) { UnhookWindowsHookEx(_hook); _hook = IntPtr.Zero; }
            try { if (_app != null) ((PowerPoint.EApplication_Event)_app).WindowSelectionChange -= OnWindowSelectionChange; } catch { }
            if (_check != null) { _check.Stop(); _check.Dispose(); _check = null; }
        }

        private IntPtr OnThreadMessage(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (code >= 0 && wParam.ToInt64() == PM_REMOVE)
                {
                    int msg = Marshal.ReadInt32(lParam, IntPtr.Size);   // MSG.message, right after MSG.hwnd
                    if (msg == WM_LBUTTONUP || msg == WM_RBUTTONUP || msg == WM_MBUTTONUP || msg == WM_POINTERUP ||
                        msg == WM_KEYUP || msg == WM_SYSKEYUP)
                        ScheduleCheck();
                }
            }
            catch { /* never let an exception escape into Windows */ }
            return CallNextHookEx(_hook, code, wParam, lParam);
        }

        /// <summary>(Re)start the countdown to a check; while input keeps coming (typing, key repeat) it keeps moving back.</summary>
        private void ScheduleCheck()
        {
            if (_check == null) return;
            _check.Stop();
            _check.Start();
        }

        private void Check()
        {
            if (_app == null || !Settings.AutoFollow) return;
            if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0) return;   // another drag has begun; its release schedules a check
            if (_busy > 0) { ScheduleCheck(); return; }                   // one of our operations is running; look again after it
            _busy++;
            try
            {
                var slide = CurrentSlide();
                if (slide != null && _svc.Follow(slide) > 0) Refresh();
            }
            catch (Exception ex)
            {
                if ((DateTime.Now - _lastCheckError).TotalSeconds > 10) { Log.Error("Check", ex); _lastCheckError = DateTime.Now; }
            }
            finally { _busy--; }
        }

        private void OnWindowSelectionChange(PowerPoint.Selection sel)
        {
            string sig = SelectionSignature();
            if (sig != _lastSelection) { _lastSelection = sig; Refresh(); }
        }

        private string SelectionSignature()
        {
            try
            {
                var sel = _app.ActiveWindow.Selection;
                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return "";
                var ids = new List<string>();
                foreach (PowerPoint.Shape s in sel.ShapeRange) if (Link.IsLink(s)) ids.Add(s.Id.ToString());
                return string.Join(",", ids);
            }
            catch { return ""; }
        }

        // ================================================================ helpers

        private PowerPoint.Slide CurrentSlide()
        {
            try
            {
                if (_app.Presentations.Count == 0) return null;
                var win = _app.ActiveWindow;
                if (win == null) return null;
                if (win.ViewType != PowerPoint.PpViewType.ppViewNormal && win.ViewType != PowerPoint.PpViewType.ppViewSlide) return null;
                return win.View.Slide as PowerPoint.Slide;
            }
            catch { return null; }
        }

        private static PowerPoint.Slide SlideOf(PowerPoint.Shape s) { return (PowerPoint.Slide)s.Parent; }

        private List<PowerPoint.Shape> SelectedLinks()
        {
            var list = new List<PowerPoint.Shape>();
            try
            {
                var sel = _app.ActiveWindow.Selection;
                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return list;
                foreach (PowerPoint.Shape s in sel.ShapeRange) if (Link.IsLink(s)) list.Add(s);
            }
            catch { }
            return list;
        }

        private PowerPoint.Shape FirstSelectedLink()
        {
            var l = SelectedLinks();
            return l.Count > 0 ? l[0] : null;
        }

        /// <summary>Runs one of our operations. Follow checks that come due meanwhile (PowerPoint can process
        /// messages in the middle of a paste or an export) wait until it is finished.</summary>
        private T Busy<T>(Func<T> f)
        {
            _busy++;
            try { return f(); }
            finally { _busy--; }
        }

        private void Guard(string where, Action a)
        {
            Exception error = null;
            _busy++;
            try { a(); }
            catch (Exception ex) { Log.Error(where, ex); error = ex; }
            finally { _busy--; }
            if (error != null)
                MessageBox.Show("OrthoLink 出错了：" + error.Message + "\n\n详细记录在：" + Log.Path, "OrthoLink", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void Info(string text)
        {
            MessageBox.Show(text, "OrthoLink", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
