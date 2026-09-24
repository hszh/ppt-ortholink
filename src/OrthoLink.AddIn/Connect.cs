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
        private Timer _timer;
        private bool _busy;
        private string _lastSelection = "";
        private DateTime _lastTickError = DateTime.MinValue;

        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int vKey);
        private const int VK_LBUTTON = 0x01;

        private static readonly Side?[] SideItems = { null, Side.Left, Side.Right, Side.Top, Side.Bottom };

        // ================================================================ IDTExtensibility2

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                _app = (PowerPoint.Application)application;
                Settings.Load();
                string dir = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
                _svc = new ConnectorService(_app, Path.Combine(dir, "template.pptx"));
                try { ((Office.COMAddIn)addInInst).Object = this; } catch (Exception ex) { Log.Error("expose Object", ex); }
                Log.Write("OnConnection ok, mode=" + connectMode + ", version " + VersionString);
                if (connectMode != ext_ConnectMode.ext_cm_Startup) StartTimer();
            }
            catch (Exception ex) { Log.Error("OnConnection", ex); }
        }

        public void OnStartupComplete(ref Array custom) { StartTimer(); }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                if (_timer != null) { _timer.Stop(); _timer.Dispose(); _timer = null; }
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
        }

        public object ConnectShapes(object srcShape, object dstShape)
        {
            var a = (PowerPoint.Shape)srcShape; var b = (PowerPoint.Shape)dstShape;
            return _svc.Create(SlideOf(a), a, b);
        }

        /// <summary>side: "" or "Auto" = automatic, else Left/Right/Top/Bottom.</summary>
        public object SetConnectorSide(object connector, bool source, string side)
        {
            var c = (PowerPoint.Shape)connector;
            Side? s = null; Side parsed;
            if (!string.IsNullOrEmpty(side) && !side.Equals("Auto", StringComparison.OrdinalIgnoreCase) && Enum.TryParse(side, true, out parsed)) s = parsed;
            return _svc.SetSide(SlideOf(c), c, source, s);
        }

        public int UpdateSlide(object slide, bool resetMiddle) { return _svc.UpdateAll((PowerPoint.Slide)slide, resetMiddle); }
        public int FollowSlide(object slide) { return _svc.Follow((PowerPoint.Slide)slide); }
        public bool AutoFollow { get { return Settings.AutoFollow; } set { Settings.AutoFollow = value; } }
        public float RadiusPt { get { return Settings.RadiusPt; } set { Settings.RadiusPt = value; Settings.Save(); } }
        public float GapPt { get { return Settings.GapPt; } set { Settings.GapPt = value; Settings.Save(); } }
        public object SetConnectorGap(object connector, float gapPt) { var c = (PowerPoint.Shape)connector; return _svc.SetGap(SlideOf(c), c, gapPt); }
        public float ParseRadiusPt(string text) { var r = ParseRadius(text); return r.HasValue ? r.Value : -1f; }
        public bool ArrowEnd { get { return Settings.ArrowEnd; } set { Settings.ArrowEnd = value; } }
        public bool ArrowStart { get { return Settings.ArrowStart; } set { Settings.ArrowStart = value; } }
        public string LogPath { get { return Log.Path; } }
        public void ActivateTab() { try { if (_ribbon != null) _ribbon.ActivateTab("olTab"); } catch (Exception ex) { Log.Error("ActivateTab", ex); } }

        // ================================================================ timer: auto-follow + ribbon refresh on selection change

        private void StartTimer()
        {
            if (_timer != null) return;
            _timer = new Timer { Interval = 200 };
            _timer.Tick += (s, e) => Tick();
            _timer.Start();
        }

        private void Tick()
        {
            if (_busy || _app == null) return;
            if ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0) return;   // a drag may be in progress
            _busy = true;
            try
            {
                string sig = SelectionSignature();
                if (sig != _lastSelection) { _lastSelection = sig; Refresh(); }
                if (!Settings.AutoFollow) return;
                var slide = CurrentSlide();
                if (slide != null && _svc.Follow(slide) > 0) Refresh();
            }
            catch (Exception ex)
            {
                if ((DateTime.Now - _lastTickError).TotalSeconds > 10) { Log.Error("Tick", ex); _lastTickError = DateTime.Now; }
            }
            finally { _busy = false; }
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

        private static void Guard(string where, Action a)
        {
            try { a(); }
            catch (Exception ex)
            {
                Log.Error(where, ex);
                MessageBox.Show("OrthoLink 出错了：" + ex.Message + "\n\n详细记录在：" + Log.Path, "OrthoLink", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void Info(string text)
        {
            MessageBox.Show(text, "OrthoLink", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
