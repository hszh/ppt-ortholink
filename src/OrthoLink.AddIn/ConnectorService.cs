using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OrthoLink
{
    internal enum Side { Left, Right, Top, Bottom }

    /// <summary>What a connector remembers about itself, stored in the shape's Tags.</summary>
    internal sealed class Link
    {
        public int SrcId, DstId;
        public Side SrcSide = Side.Right, DstSide = Side.Left;
        public float SrcPos = 0.5f, DstPos = 0.5f;
        public bool SrcPin, DstPin;          // true = the user chose this side/position; never auto-changed
        public float Gap;                    // pt between each end and the shape outline
        public string Topo = "HVH";

        private const string TVer = "OL_VER";

        public static bool IsLink(PowerPoint.Shape s)
        {
            try { return s.Tags[TVer] == "1"; } catch { return false; }
        }

        public static Link Read(PowerPoint.Shape s)
        {
            if (!IsLink(s)) return null;
            var t = s.Tags;
            int src, dst;
            if (!int.TryParse(t["OL_SRC"], out src) || !int.TryParse(t["OL_DST"], out dst)) return null;   // tags only half written
            var l = new Link
            {
                SrcId = src,
                DstId = dst,
                SrcSide = ParseSide(t["OL_SRC_SIDE"], Side.Right),
                DstSide = ParseSide(t["OL_DST_SIDE"], Side.Left),
                SrcPos = ParseF(t["OL_SRC_POS"], 0.5f),
                DstPos = ParseF(t["OL_DST_POS"], 0.5f),
                SrcPin = t["OL_SRC_PIN"] == "1",
                DstPin = t["OL_DST_PIN"] == "1",
                Gap = ParseF(t["OL_GAP"], 0f),
                Topo = t["OL_TOPO"],
            };
            if (string.IsNullOrEmpty(l.Topo)) l.Topo = "HVH";
            return l;
        }

        public void Write(PowerPoint.Shape s)
        {
            var t = s.Tags;
            t.Add("OL_SRC", SrcId.ToString());
            t.Add("OL_DST", DstId.ToString());
            t.Add("OL_SRC_SIDE", SrcSide.ToString());
            t.Add("OL_DST_SIDE", DstSide.ToString());
            t.Add("OL_SRC_POS", SrcPos.ToString(System.Globalization.CultureInfo.InvariantCulture));
            t.Add("OL_DST_POS", DstPos.ToString(System.Globalization.CultureInfo.InvariantCulture));
            t.Add("OL_SRC_PIN", SrcPin ? "1" : "0");
            t.Add("OL_DST_PIN", DstPin ? "1" : "0");
            t.Add("OL_GAP", Gap.ToString(System.Globalization.CultureInfo.InvariantCulture));
            t.Add("OL_TOPO", Topo);
            t.Add(TVer, "1");   // last: a shape whose tags are still being written is not taken for a connector
        }

        private static Side ParseSide(string s, Side fallback)
        {
            Side r; return Enum.TryParse(s, true, out r) ? r : fallback;
        }
        private static float ParseF(string s, float fallback)
        {
            float f;
            return float.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out f) ? f : fallback;
        }
    }

    /// <summary>A planned route: segment sequence plus the absolute coordinate of every interior segment
    /// (x for vertical segments, y for horizontal ones).</summary>
    internal sealed class Route
    {
        public string Topo;
        public float[] Params;
        public float Lo = float.NegativeInfinity, Hi = float.PositiveInfinity;   // where the middle of a 3-segment route may go
    }

    /// <summary>Pure geometry helpers, no COM.</summary>
    internal static class Geo
    {
        public const float Margin = 20f;   // how far a route keeps away from a box it leaves "backwards"

        public static bool IsHorizontal(Side s) { return s == Side.Left || s == Side.Right; }

        public static Side Opposite(Side s)
        {
            switch (s)
            {
                case Side.Left: return Side.Right;
                case Side.Right: return Side.Left;
                case Side.Top: return Side.Bottom;
                default: return Side.Top;
            }
        }

        /// <summary>Unit vector pointing out of the box through that side.</summary>
        public static PointF Outward(Side s)
        {
            switch (s)
            {
                case Side.Left: return new PointF(-1, 0);
                case Side.Right: return new PointF(1, 0);
                case Side.Top: return new PointF(0, -1);
                default: return new PointF(0, 1);
            }
        }

        public static PointF Port(RectangleF r, Side side, float pos)
        {
            switch (side)
            {
                case Side.Left: return new PointF(r.Left, r.Top + r.Height * pos);
                case Side.Right: return new PointF(r.Right, r.Top + r.Height * pos);
                case Side.Top: return new PointF(r.Left + r.Width * pos, r.Top);
                default: return new PointF(r.Left + r.Width * pos, r.Bottom);
            }
        }

        public static PointF Center(RectangleF r) { return new PointF(r.Left + r.Width / 2, r.Top + r.Height / 2); }

        /// <summary>Pick both sides from the relative placement of the two boxes (neither end pinned).</summary>
        public static void ChooseSides(RectangleF a, RectangleF b, out Side sa, out Side sb)
        {
            float gapX = Math.Max(b.Left - a.Right, a.Left - b.Right);
            float gapY = Math.Max(b.Top - a.Bottom, a.Top - b.Bottom);
            var ca = Center(a); var cb = Center(b);
            if (gapX >= gapY) sa = cb.X >= ca.X ? Side.Right : Side.Left;
            else sa = cb.Y >= ca.Y ? Side.Bottom : Side.Top;
            sb = Opposite(sa);
        }

        /// <summary>Side of box r that faces point p.</summary>
        public static Side FacingSide(RectangleF r, PointF p)
        {
            var c = Center(r);
            float dx = p.X - c.X, dy = p.Y - c.Y;
            if (Math.Abs(dx) * r.Height >= Math.Abs(dy) * r.Width) return dx >= 0 ? Side.Right : Side.Left;
            return dy >= 0 ? Side.Bottom : Side.Top;
        }

        /// <summary>Route between two ports. Uses the shortest segment sequence that leaves the start box
        /// outward and enters the end box inward; otherwise a longer sequence that goes around.</summary>
        public static Route Plan(Side sa, Side sb, RectangleF ra, RectangleF rb, PointF pa, PointF pb)
        {
            float m = Margin;
            var es = Outward(sa);                 // direction of travel when leaving
            var eo = Outward(sb);
            var ed = new PointF(-eo.X, -eo.Y);    // direction of travel when entering
            bool hs = IsHorizontal(sa), hd = IsHorizontal(sb);

            if (hs && hd)
            {
                // facing boxes too close for the margin: the middle segment runs through the middle of the space between them
                float between = es.X > 0 && ed.X > 0 ? rb.Left - ra.Right : es.X < 0 && ed.X < 0 ? ra.Left - rb.Right : float.PositiveInfinity;
                if (between > 0 && between < 2 * m)
                {
                    float x = es.X > 0 ? ra.Right + between / 2 : ra.Left - between / 2;
                    return new Route { Topo = "HVH", Params = new[] { x }, Lo = x, Hi = x };
                }
                float lo = float.NegativeInfinity, hi = float.PositiveInfinity;
                if (es.X > 0) lo = Math.Max(lo, ra.Right + m); else hi = Math.Min(hi, ra.Left - m);
                if (ed.X > 0) hi = Math.Min(hi, rb.Left - m); else lo = Math.Max(lo, rb.Right + m);
                if (lo <= hi) return new Route { Topo = "HVH", Params = new[] { Clamp((pa.X + pb.X) / 2, lo, hi) }, Lo = lo, Hi = hi };
                float x1 = es.X > 0 ? ra.Right + m : ra.Left - m;
                float x2 = ed.X > 0 ? rb.Left - m : rb.Right + m;
                float yTop = Math.Min(ra.Top, rb.Top) - m, yBot = Math.Max(ra.Bottom, rb.Bottom) + m;
                float y = Math.Abs(pa.Y - yTop) + Math.Abs(pb.Y - yTop) <= Math.Abs(pa.Y - yBot) + Math.Abs(pb.Y - yBot) ? yTop : yBot;
                return new Route { Topo = "HVHVH", Params = new[] { x1, y, x2 } };
            }
            if (!hs && !hd)
            {
                float between = es.Y > 0 && ed.Y > 0 ? rb.Top - ra.Bottom : es.Y < 0 && ed.Y < 0 ? ra.Top - rb.Bottom : float.PositiveInfinity;
                if (between > 0 && between < 2 * m)
                {
                    float y = es.Y > 0 ? ra.Bottom + between / 2 : ra.Top - between / 2;
                    return new Route { Topo = "VHV", Params = new[] { y }, Lo = y, Hi = y };
                }
                float lo = float.NegativeInfinity, hi = float.PositiveInfinity;
                if (es.Y > 0) lo = Math.Max(lo, ra.Bottom + m); else hi = Math.Min(hi, ra.Top - m);
                if (ed.Y > 0) hi = Math.Min(hi, rb.Top - m); else lo = Math.Max(lo, rb.Bottom + m);
                if (lo <= hi) return new Route { Topo = "VHV", Params = new[] { Clamp((pa.Y + pb.Y) / 2, lo, hi) }, Lo = lo, Hi = hi };
                float y1 = es.Y > 0 ? ra.Bottom + m : ra.Top - m;
                float y2 = ed.Y > 0 ? rb.Top - m : rb.Bottom + m;
                float xL = Math.Min(ra.Left, rb.Left) - m, xR = Math.Max(ra.Right, rb.Right) + m;
                float x = Math.Abs(pa.X - xL) + Math.Abs(pb.X - xL) <= Math.Abs(pa.X - xR) + Math.Abs(pb.X - xR) ? xL : xR;
                return new Route { Topo = "VHVHV", Params = new[] { y1, x, y2 } };
            }
            if (hs)
            {
                if ((pb.X - pa.X) * es.X >= m && (pb.Y - pa.Y) * ed.Y >= m) return new Route { Topo = "HV", Params = new float[0] };
                float x1 = es.X > 0 ? ra.Right + m : ra.Left - m;
                float y1 = ed.Y > 0 ? rb.Top - m : rb.Bottom + m;
                return new Route { Topo = "HVHV", Params = new[] { x1, y1 } };
            }
            {
                if ((pb.Y - pa.Y) * es.Y >= m && (pb.X - pa.X) * ed.X >= m) return new Route { Topo = "VH", Params = new float[0] };
                float y1 = es.Y > 0 ? ra.Bottom + m : ra.Top - m;
                float x1 = ed.X > 0 ? rb.Left - m : rb.Right + m;
                return new Route { Topo = "VHVH", Params = new[] { y1, x1 } };
            }
        }

        public static float Clamp(float v, float lo, float hi) { return v < lo ? lo : v > hi ? hi : v; }

        public static float DistanceToRect(PointF p, RectangleF r)
        {
            float dx = Math.Max(Math.Max(r.Left - p.X, 0), p.X - r.Right);
            float dy = Math.Max(Math.Max(r.Top - p.Y, 0), p.Y - r.Bottom);
            return (float)Math.Sqrt(dx * dx + dy * dy);
        }

        /// <summary>Nearest side of r to p, and the position along that side (0..1).</summary>
        public static Side NearestSide(RectangleF r, PointF p, out float pos)
        {
            float dl = Math.Abs(p.X - r.Left), dr = Math.Abs(p.X - r.Right);
            float dt = Math.Abs(p.Y - r.Top), db = Math.Abs(p.Y - r.Bottom);
            float m = Math.Min(Math.Min(dl, dr), Math.Min(dt, db));
            Side s = m == dl ? Side.Left : m == dr ? Side.Right : m == dt ? Side.Top : Side.Bottom;
            pos = IsHorizontal(s) ? (r.Height > 0 ? (p.Y - r.Top) / r.Height : 0.5f) : (r.Width > 0 ? (p.X - r.Left) / r.Width : 0.5f);
            pos = Math.Max(0f, Math.Min(1f, pos));
            return s;
        }

        public static bool Near(RectangleF a, RectangleF b, float eps = 0.05f)
        {
            return Math.Abs(a.X - b.X) < eps && Math.Abs(a.Y - b.Y) < eps &&
                   Math.Abs(a.Width - b.Width) < eps && Math.Abs(a.Height - b.Height) < eps;
        }

        public static float Dist(PointF a, PointF b) { float dx = a.X - b.X, dy = a.Y - b.Y; return (float)Math.Sqrt(dx * dx + dy * dy); }
    }

    /// <summary>Creates and maintains rounded orthogonal connectors on slides.</summary>
    internal sealed class ConnectorService
    {
        public const float AdjPerPt = 0.127f;   // Adjustments value = EMU / 100000; 1 pt = 12700 EMU
        private const int RadiusAdj = 1;        // Adjustments[1] = corner radius; [2..] = interior segments
        private const float SnapDistance = 40f; // how close (pt) a dragged end must land to a box to attach

        private readonly PowerPoint.Application _app;
        private PowerPoint.Presentation _scratch;   // hidden deck where shapes are redrawn to sample their outlines
        private readonly Dictionary<string, byte[]> _templates = new Dictionary<string, byte[]>();

        private sealed class Seen { public RectangleF Src, Dst, Conn; }
        private readonly Dictionary<int, Seen> _seen = new Dictionary<int, Seen>();

        private sealed class CachedOutline { public string Key; public Outline Outline; }
        private readonly Dictionary<int, CachedOutline> _outlines = new Dictionary<int, CachedOutline>();

        public ConnectorService(PowerPoint.Application app)
        {
            _app = app;
        }

        // ------------------------------------------------------------------ public operations

        public PowerPoint.Shape Create(PowerPoint.Slide slide, PowerPoint.Shape src, PowerPoint.Shape dst)
        {
            var link = new Link { SrcId = src.Id, DstId = dst.Id, Gap = Settings.GapPt };
            ResolveSides(link, Rect(src), Rect(dst));
            var conn = Inject(slide, link.Topo);
            conn.Name = UniqueName(slide);
            ApplyStyle(conn, Settings.RadiusPt, Settings.ArrowStart, Settings.ArrowEnd);
            conn = UpdateCore(slide, conn, link, src, dst, true);
            Log.Write("created " + conn.Name + " " + link.Topo + " " + src.Name + " -> " + dst.Name);
            return conn;
        }

        /// <summary>Re-glue one connector to its two boxes. reset also clears user-chosen sides and re-plans.
        /// Returns the (possibly replaced) shape, or null if it is orphaned.</summary>
        public PowerPoint.Shape Update(PowerPoint.Slide slide, PowerPoint.Shape conn, bool reset)
        {
            var link = Link.Read(conn);
            if (link == null) return null;
            var map = IndexShapes(slide);
            PowerPoint.Shape src, dst;
            if (!map.TryGetValue(link.SrcId, out src) || !map.TryGetValue(link.DstId, out dst)) return null;
            if (reset) { link.SrcPin = link.DstPin = false; link.SrcPos = link.DstPos = 0.5f; }
            return UpdateCore(slide, conn, link, src, dst, reset);
        }

        public int UpdateAll(PowerPoint.Slide slide, bool reset)
        {
            int n = 0;
            var map = IndexShapes(slide);
            foreach (var conn in LinksOn(slide))
            {
                var link = Link.Read(conn);
                PowerPoint.Shape src, dst;
                if (link == null || !map.TryGetValue(link.SrcId, out src) || !map.TryGetValue(link.DstId, out dst)) continue;
                UpdateCore(slide, conn, link, src, dst, reset);
                n++;
            }
            return n;
        }

        /// <summary>Pin one end of a connector to a side (null = back to automatic).</summary>
        public PowerPoint.Shape SetSide(PowerPoint.Slide slide, PowerPoint.Shape conn, bool source, Side? side)
        {
            var link = Link.Read(conn);
            if (link == null) return null;
            if (source) { link.SrcPin = side.HasValue; if (side.HasValue) link.SrcSide = side.Value; else link.SrcPos = 0.5f; }
            else { link.DstPin = side.HasValue; if (side.HasValue) link.DstSide = side.Value; else link.DstPos = 0.5f; }
            var map = IndexShapes(slide);
            PowerPoint.Shape src, dst;
            if (!map.TryGetValue(link.SrcId, out src) || !map.TryGetValue(link.DstId, out dst)) return null;
            return UpdateCore(slide, conn, link, src, dst, true);
        }

        /// <summary>Distance between the ends and the shapes they point at.</summary>
        public PowerPoint.Shape SetGap(PowerPoint.Slide slide, PowerPoint.Shape conn, float gapPt)
        {
            var link = Link.Read(conn);
            if (link == null) return null;
            link.Gap = gapPt;
            var map = IndexShapes(slide);
            PowerPoint.Shape src, dst;
            if (!map.TryGetValue(link.SrcId, out src) || !map.TryGetValue(link.DstId, out dst)) { link.Write(conn); return conn; }
            return UpdateCore(slide, conn, link, src, dst, false);
        }

        public static float GetGap(PowerPoint.Shape conn) { var l = Link.Read(conn); return l == null ? 0f : l.Gap; }

        /// <summary>Called after user input. Follows moved boxes, and re-attaches ends the user dragged.</summary>
        public int Follow(PowerPoint.Slide slide)
        {
            int n = 0;
            var map = IndexShapes(slide);
            foreach (var conn in LinksOn(slide))
            {
                try
                {
                    var link = Link.Read(conn);
                    PowerPoint.Shape src, dst;
                    if (link == null || !map.TryGetValue(link.SrcId, out src) || !map.TryGetValue(link.DstId, out dst)) continue;
                    var now = new Seen { Src = Rect(src), Dst = Rect(dst), Conn = Rect(conn) };
                    Seen before;
                    int id = conn.Id;
                    if (!_seen.TryGetValue(id, out before))
                    {
                        _seen[id] = now;   // first sight: learn, don't touch (keeps undo history clean on file open)
                        continue;
                    }
                    bool boxMoved = !Geo.Near(before.Src, now.Src) || !Geo.Near(before.Dst, now.Dst);
                    bool lineMoved = !Geo.Near(before.Conn, now.Conn);
                    if (!boxMoved && !lineMoved) continue;

                    PowerPoint.Shape c2;
                    if (!boxMoved) c2 = ReAnchor(slide, conn, link, src, dst, before.Conn, map);
                    else c2 = UpdateCore(slide, conn, link, src, dst, false);
                    if (c2 != null && c2.Id != id) _seen.Remove(id);
                    n++;
                }
                catch (Exception ex) { Log.Error("Follow", ex); }   // one broken line must not hold up the others
            }
            return n;
        }

        public void ApplyStyle(PowerPoint.Shape conn, float radiusPt, bool arrowStart, bool arrowEnd)
        {
            SetRadius(conn, radiusPt); SetArrowStart(conn, arrowStart); SetArrowEnd(conn, arrowEnd);
        }

        public void SetRadius(PowerPoint.Shape conn, float radiusPt) { conn.Adjustments[RadiusAdj] = radiusPt * AdjPerPt; }
        public static float GetRadius(PowerPoint.Shape conn) { return conn.Adjustments[RadiusAdj] / AdjPerPt; }
        public void SetArrowStart(PowerPoint.Shape conn, bool on) { conn.Line.BeginArrowheadStyle = on ? Office.MsoArrowheadStyle.msoArrowheadTriangle : Office.MsoArrowheadStyle.msoArrowheadNone; }
        public void SetArrowEnd(PowerPoint.Shape conn, bool on) { conn.Line.EndArrowheadStyle = on ? Office.MsoArrowheadStyle.msoArrowheadTriangle : Office.MsoArrowheadStyle.msoArrowheadNone; }
        public static bool GetArrowStart(PowerPoint.Shape conn) { return conn.Line.BeginArrowheadStyle != Office.MsoArrowheadStyle.msoArrowheadNone; }
        public static bool GetArrowEnd(PowerPoint.Shape conn) { return conn.Line.EndArrowheadStyle != Office.MsoArrowheadStyle.msoArrowheadNone; }

        public void Shutdown()
        {
            try { if (_scratch != null) { _scratch.Saved = Office.MsoTriState.msoTrue; _scratch.Close(); } } catch { }
            _scratch = null;
            Clip.Shutdown();
        }

        // ------------------------------------------------------------------ internals

        /// <summary>Decide sides, honouring pinned ends.</summary>
        private static void ResolveSides(Link link, RectangleF ra, RectangleF rb)
        {
            if (!link.SrcPin && !link.DstPin)
            {
                Side sa, sb; Geo.ChooseSides(ra, rb, out sa, out sb);
                link.SrcSide = sa; link.DstSide = sb;
            }
            else if (!link.SrcPin) link.SrcSide = Geo.FacingSide(ra, Geo.Port(rb, link.DstSide, link.DstPos));
            else if (!link.DstPin) link.DstSide = Geo.FacingSide(rb, Geo.Port(ra, link.SrcSide, link.SrcPos));
        }

        private PowerPoint.Shape UpdateCore(PowerPoint.Slide slide, PowerPoint.Shape conn, Link link,
                                            PowerPoint.Shape src, PowerPoint.Shape dst, bool reset)
        {
            var ra = Rect(src); var rb = Rect(dst);

            // remember where the user left the interior segments (relative to the ports they hug)
            float[] oldParams = null; PointF oldPa = PointF.Empty, oldPb = PointF.Empty;
            if (!reset)
            {
                Seen was; RectangleF oa, ob, oc;
                if (_seen.TryGetValue(conn.Id, out was)) { oa = was.Src; ob = was.Dst; oc = was.Conn; } else { oa = ra; ob = rb; oc = Rect(conn); }
                oldPa = PortOf(src, oa, link.SrcSide, link.SrcPos, link.Gap);
                oldPb = PortOf(dst, ob, link.DstSide, link.DstPos, link.Gap);
                // read the segments against the box we last gave the line: when the line was selected and dragged
                // together with the shapes, PowerPoint has already moved it, and counting that move as well as the
                // shapes' move would shift the segments twice
                oldParams = ReadParams(conn, link.Topo, oc);
            }

            ResolveSides(link, ra, rb);
            var pa = PortOf(src, ra, link.SrcSide, link.SrcPos, link.Gap);
            var pb = PortOf(dst, rb, link.DstSide, link.DstPos, link.Gap);
            var plan = Geo.Plan(link.SrcSide, link.DstSide, ra, rb, pa, pb);

            float[] p;
            if (plan.Topo != link.Topo)
            {
                conn = Swap(slide, conn, plan.Topo);
                link.Topo = plan.Topo;
                p = plan.Params;
            }
            else p = oldParams != null ? Carry(oldParams, link.Topo, oldPa, oldPb, pa, pb) : plan.Params;
            if (p.Length == 1) p[0] = Geo.Clamp(p[0], plan.Lo, plan.Hi);   // a kept middle segment must not cut into a box

            link.Write(conn);
            Place(conn, pa, pb);
            WriteParams(conn, link.Topo, p);

            _seen[conn.Id] = new Seen { Src = ra, Dst = rb, Conn = Rect(conn) };
            return conn;
        }

        /// <summary>Move the user's interior segments along with the ports they belong to.
        /// When both ends moved by the same amount (both shapes dragged together) the whole line moves with them.
        /// Otherwise the segment next to the start follows the start, the one next to the end follows the end,
        /// anything else (and the single middle of a 3-segment route) stays where it is.</summary>
        private static float[] Carry(float[] old, string topo, PointF oldPa, PointF oldPb, PointF pa, PointF pb)
        {
            int n = topo.Length;
            var p = (float[])old.Clone();
            float dxa = pa.X - oldPa.X, dya = pa.Y - oldPa.Y, dxb = pb.X - oldPb.X, dyb = pb.Y - oldPb.Y;
            if (Math.Abs(dxa - dxb) < 0.5f && Math.Abs(dya - dyb) < 0.5f)
            {
                for (int k = 1; k <= n - 2; k++) p[k - 1] += topo[k] == 'V' ? dxa : dya;
                return p;
            }
            if (n <= 3) return p;
            for (int k = 1; k <= n - 2; k++)
            {
                bool vertical = topo[k] == 'V';
                if (k == 1) p[k - 1] += vertical ? dxa : dya;
                else if (k == n - 2) p[k - 1] += vertical ? dxb : dyb;
            }
            return p;
        }

        /// <summary>The user dragged the line itself. A moved end gets attached to the box it landed on.</summary>
        private PowerPoint.Shape ReAnchor(PowerPoint.Slide slide, PowerPoint.Shape conn, Link link,
                                          PowerPoint.Shape src, PowerPoint.Shape dst, RectangleF wasRect,
                                          Dictionary<int, PowerPoint.Shape> map)
        {
            bool flipH = conn.HorizontalFlip == Office.MsoTriState.msoTrue, flipV = conn.VerticalFlip == Office.MsoTriState.msoTrue;
            PointF s0, e0, s1, e1;
            Ends(wasRect, flipH, flipV, out s0, out e0);
            Ends(Rect(conn), flipH, flipV, out s1, out e1);
            float ds = Geo.Dist(s0, s1), de = Geo.Dist(e0, e1);
            bool translated = ds > 1 && de > 1 && Math.Abs((s1.X - s0.X) - (e1.X - e0.X)) < 1 && Math.Abs((s1.Y - s0.Y) - (e1.Y - e0.Y)) < 1;

            bool changed = false;
            if (!translated)
            {
                if (ds > 1) changed |= Attach(link, true, s1, src, map);
                if (de > 1) changed |= Attach(link, false, e1, dst, map);
                if (changed)
                {
                    if (!map.TryGetValue(link.SrcId, out src) || !map.TryGetValue(link.DstId, out dst)) return conn;
                    Log.Write("re-anchored " + conn.Name + ": " + link.SrcSide + "@" + link.SrcPos.ToString("0.00") + " -> " + link.DstSide + "@" + link.DstPos.ToString("0.00"));
                }
            }
            // whether attached or not, snap the line back onto its ports (a plain move of the line is undone)
            return UpdateCore(slide, conn, link, src, dst, changed);
        }

        private static bool Attach(Link link, bool source, PointF p, PowerPoint.Shape current, Dictionary<int, PowerPoint.Shape> map)
        {
            PowerPoint.Shape best = null; float bestD = float.MaxValue;
            foreach (var s in map.Values)
            {
                if (Link.IsLink(s)) continue;
                float d = Geo.DistanceToRect(p, Rect(s));
                if (s.Id == current.Id) d -= 5f;   // slight preference for the box it is already on
                if (d < bestD) { bestD = d; best = s; }
            }
            if (best == null || bestD > SnapDistance) return false;
            float pos; var side = Geo.NearestSide(Rect(best), p, out pos);
            if (source) { link.SrcId = best.Id; link.SrcSide = side; link.SrcPos = pos; link.SrcPin = true; }
            else { link.DstId = best.Id; link.DstSide = side; link.DstPos = pos; link.DstPin = true; }
            return true;
        }

        private static void Ends(RectangleF r, bool flipH, bool flipV, out PointF start, out PointF end)
        {
            start = new PointF(flipH ? r.Right : r.Left, flipV ? r.Bottom : r.Top);
            end = new PointF(flipH ? r.Left : r.Right, flipV ? r.Top : r.Bottom);
        }

        /// <summary>Replace a connector by one of another topology, keeping name, formatting and tags.</summary>
        private PowerPoint.Shape Swap(PowerPoint.Slide slide, PowerPoint.Shape old, string topo)
        {
            var fresh = Inject(slide, topo);
            // PickUp/Apply would also copy the adjust values and break the new geometry, so copy the line by hand
            try
            {
                var a = old.Line; var b = fresh.Line;
                b.Visible = a.Visible;
                b.Weight = a.Weight;
                b.DashStyle = a.DashStyle;
                b.ForeColor.RGB = a.ForeColor.RGB;
                b.Transparency = a.Transparency;
                b.BeginArrowheadStyle = a.BeginArrowheadStyle; b.BeginArrowheadLength = a.BeginArrowheadLength; b.BeginArrowheadWidth = a.BeginArrowheadWidth;
                b.EndArrowheadStyle = a.EndArrowheadStyle; b.EndArrowheadLength = a.EndArrowheadLength; b.EndArrowheadWidth = a.EndArrowheadWidth;
            }
            catch (Exception ex) { Log.Error("Swap.CopyLine", ex); }
            fresh.Adjustments[RadiusAdj] = old.Adjustments[RadiusAdj];
            string name = old.Name;
            var tags = old.Tags;
            for (int i = 1; i <= tags.Count; i++) fresh.Tags.Add(tags.Name(i), tags.Value(i));
            _seen.Remove(old.Id);
            old.Delete();
            fresh.Name = name;
            return fresh;
        }

        private static void Place(PowerPoint.Shape c, PointF p, PointF q)
        {
            bool flipH = q.X < p.X, flipV = q.Y < p.Y;
            c.Left = Math.Min(p.X, q.X);
            c.Top = Math.Min(p.Y, q.Y);
            c.Width = Math.Abs(q.X - p.X);
            c.Height = Math.Abs(q.Y - p.Y);
            try
            {
                if ((c.HorizontalFlip == Office.MsoTriState.msoTrue) != flipH) c.Flip(Office.MsoFlipCmd.msoFlipHorizontal);
                if ((c.VerticalFlip == Office.MsoTriState.msoTrue) != flipV) c.Flip(Office.MsoFlipCmd.msoFlipVertical);
            }
            catch (Exception ex) { Log.Error("Place.Flip", ex); }
        }

        /// <summary>Absolute slide coordinates of the interior segments, from the shape's adjust values,
        /// taken as fractions of the given box (where the line was, not necessarily where it is now).</summary>
        private static float[] ReadParams(PowerPoint.Shape c, string topo, RectangleF box)
        {
            int n = topo.Length;
            var p = new float[Math.Max(0, n - 2)];
            bool flipH = c.HorizontalFlip == Office.MsoTriState.msoTrue, flipV = c.VerticalFlip == Office.MsoTriState.msoTrue;
            for (int k = 1; k <= n - 2; k++)
            {
                float adj;
                try { adj = c.Adjustments[k + 1]; } catch { adj = 0.5f; }
                if (topo[k] == 'V') p[k - 1] = flipH ? box.Right - adj * box.Width : box.Left + adj * box.Width;
                else p[k - 1] = flipV ? box.Bottom - adj * box.Height : box.Top + adj * box.Height;
            }
            return p;
        }

        private static void WriteParams(PowerPoint.Shape c, string topo, float[] p)
        {
            int n = topo.Length;
            bool flipH = c.HorizontalFlip == Office.MsoTriState.msoTrue, flipV = c.VerticalFlip == Office.MsoTriState.msoTrue;
            for (int k = 1; k <= n - 2 && k - 1 < p.Length; k++)
            {
                float adj = 0.5f;
                if (topo[k] == 'V' && c.Width > 0.5f)
                    adj = flipH ? (c.Left + c.Width - p[k - 1]) / c.Width : (p[k - 1] - c.Left) / c.Width;
                else if (topo[k] == 'H' && c.Height > 0.5f)
                    adj = flipV ? (c.Top + c.Height - p[k - 1]) / c.Height : (p[k - 1] - c.Top) / c.Height;
                adj = Geo.Clamp(adj, -20f, 21f);
                try { c.Adjustments[k + 1] = adj; } catch (Exception ex) { Log.Error("WriteParams adj" + (k + 1), ex); }
            }
        }

        private static RectangleF Rect(PowerPoint.Shape s) { return new RectangleF(s.Left, s.Top, s.Width, s.Height); }

        // ---- real outlines for non-rectangular shapes

        /// <summary>Where a connector end sits for the given shape: on the real outline, pushed out by gap.</summary>
        private PointF PortOf(PowerPoint.Shape s, RectangleF r, Side side, float pos, float gap)
        {
            PointF p;
            var o = OutlineOf(s);
            p = o != null ? o.Hit(r, side, pos) : Geo.Port(r, side, pos);
            if (gap != 0f) { var d = Geo.Outward(side); p = new PointF(p.X + d.X * gap, p.Y + d.Y * gap); }
            return p;
        }

        /// <summary>Null for rectangles (bounding box is exact) and for things we do not redraw.</summary>
        private Outline OutlineOf(PowerPoint.Shape s)
        {
            try
            {
                if (Math.Abs(s.Rotation) > 0.01f || s.Width < 2 || s.Height < 2) return null;
                string key = OutlineKey(s);
                if (key == null) return null;
                CachedOutline c;
                if (_outlines.TryGetValue(s.Id, out c) && c.Key == key) return c.Outline;
                var o = RenderOutline(s);
                _outlines[s.Id] = new CachedOutline { Key = key, Outline = o };
                return o;
            }
            catch (Exception ex) { Log.Error("OutlineOf", ex); return null; }
        }

        /// <summary>Everything the outline depends on, or null when there is no outline to sample.</summary>
        private static string OutlineKey(PowerPoint.Shape s)
        {
            try
            {
                var sb = new StringBuilder();
                sb.Append(s.Width.ToString("0.0")).Append('|').Append(s.Height.ToString("0.0")).Append('|')
                  .Append(s.HorizontalFlip == Office.MsoTriState.msoTrue ? 'h' : '-').Append(s.VerticalFlip == Office.MsoTriState.msoTrue ? 'v' : '-');
                if (IsPreset(s))
                {
                    sb.Append("|p").Append((int)s.AutoShapeType);
                    var adj = s.Adjustments;
                    for (int i = 1; i <= adj.Count; i++) sb.Append('|').Append(adj[i].ToString("0.000"));
                    return sb.ToString();
                }
                if (s.Type != Office.MsoShapeType.msoFreeform) return null;
                var nodes = ReadNodes(s);
                if (nodes == null) return null;
                foreach (var n in nodes) sb.Append('|').Append(n.X.ToString("0.0")).Append(',').Append(n.Y.ToString("0.0")).Append(n.Curve ? 'c' : 'l');
                return sb.ToString();
            }
            catch { return null; }   // some shapes refuse to report adjust values or points; use their bounding box
        }

        /// <summary>A PowerPoint preset shape (oval, trapezoid, ...) other than a plain rectangle.</summary>
        private static bool IsPreset(PowerPoint.Shape s)
        {
            try
            {
                var type = s.Type;
                if (type != Office.MsoShapeType.msoAutoShape && type != Office.MsoShapeType.msoPlaceholder) return false;
                var a = s.AutoShapeType;
                return a != Office.MsoAutoShapeType.msoShapeRectangle && a != Office.MsoAutoShapeType.msoShapeNotPrimitive &&
                       a != Office.MsoAutoShapeType.msoShapeMixed;
            }
            catch { return false; }
        }

        private struct Node { public float X, Y; public bool Curve; }
        private const int MaxNodes = 4000;   // beyond this, reading the points on every move gets slow

        /// <summary>A freeform's points relative to its bounding box, or null when the path is not one simple
        /// outline (the object model does not say where separate pieces or holes begin).</summary>
        private static List<Node> ReadNodes(PowerPoint.Shape s)
        {
            var nodes = s.Nodes;
            int n = nodes.Count;
            if (n < 3 || n > MaxNodes) return null;
            float left = s.Left, top = s.Top;
            var list = new List<Node>(n);
            for (int i = 1; i <= n; i++)
            {
                var node = nodes[i];
                var pt = (Array)node.Points;
                int r = pt.GetLowerBound(0), c = pt.GetLowerBound(1);
                list.Add(new Node
                {
                    X = Convert.ToSingle(pt.GetValue(r, c)) - left,
                    Y = Convert.ToSingle(pt.GetValue(r, c + 1)) - top,
                    Curve = node.SegmentType == Office.MsoSegmentType.msoSegmentCurve,
                });
            }
            // a curve is three nodes in a row (two control points, then its end); anything else means
            // the path jumps between pieces
            for (int i = 1; i < n; )
            {
                if (!list[i].Curve) { i++; continue; }
                if (i + 2 >= n || !list[i + 1].Curve || !list[i + 2].Curve) return null;
                i += 3;
            }
            return list;
        }

        private const float OutlineMargin = 8f;   // pt of white around the redrawn shape

        /// <summary>Redraw the shape on the hidden scratch slide (same type, size, adjust values and flips, or the
        /// same points for a freeform), black on white, export the slide and sample the picture.
        /// The user's slide, undo list and clipboard are left alone.</summary>
        private Outline RenderOutline(PowerPoint.Shape s)
        {
            float m = OutlineMargin;
            float pageW = Math.Max(72f, s.Width + 2 * m), pageH = Math.Max(72f, s.Height + 2 * m);   // slides are 1 to 56 inches
            if (pageW > 4032f || pageH > 4032f) return null;
            string png = Path.Combine(Path.GetTempPath(), "ortholink_outline_" + Guid.NewGuid().ToString("N") + ".png");
            PowerPoint.Shape dup = null;
            try
            {
                var slide = ScratchSlide();
                _scratch.PageSetup.SlideWidth = pageW;
                _scratch.PageSetup.SlideHeight = pageH;
                dup = Redraw(s, slide, m);
                if (dup == null) return null;
                Blacken(dup);
                float k = Math.Min(3f, 2400f / Math.Max(pageW, pageH));
                slide.Export(png, "PNG", (int)Math.Round(pageW * k), (int)Math.Round(pageH * k));
                var o = Outline.FromSlidePng(png, pageW, pageH, m);
                Log.Write("outline rendered for " + s.Name);
                return o;
            }
            catch (Exception ex) { Log.Error("RenderOutline " + s.Name, ex); return null; }
            finally
            {
                try { if (dup != null) dup.Delete(); } catch { }
                try { if (File.Exists(png)) File.Delete(png); } catch { }
                try { if (_scratch != null) _scratch.Saved = Office.MsoTriState.msoTrue; } catch { }
            }
        }

        /// <summary>A copy of s's geometry with its bounding box at (at, at), or null.</summary>
        private static PowerPoint.Shape Redraw(PowerPoint.Shape s, PowerPoint.Slide slide, float at)
        {
            if (IsPreset(s))
            {
                var d = slide.Shapes.AddShape(s.AutoShapeType, at, at, s.Width, s.Height);
                var from = s.Adjustments; var to = d.Adjustments;
                for (int i = 1; i <= from.Count && i <= to.Count; i++) to[i] = from[i];
                if (s.HorizontalFlip == Office.MsoTriState.msoTrue) d.Flip(Office.MsoFlipCmd.msoFlipHorizontal);
                if (s.VerticalFlip == Office.MsoTriState.msoTrue) d.Flip(Office.MsoFlipCmd.msoFlipVertical);
                return d;
            }
            var nodes = ReadNodes(s);
            if (nodes == null) return null;
            const Office.MsoEditingType corner = Office.MsoEditingType.msoEditingCorner;
            var b = slide.Shapes.BuildFreeform(corner, at + nodes[0].X, at + nodes[0].Y);
            for (int i = 1; i < nodes.Count; )
            {
                var p = nodes[i];
                if (p.Curve)
                {
                    Node q = nodes[i + 1], e = nodes[i + 2];
                    b.AddNodes(Office.MsoSegmentType.msoSegmentCurve, corner, at + p.X, at + p.Y, at + q.X, at + q.Y, at + e.X, at + e.Y);
                    i += 3;
                }
                else { b.AddNodes(Office.MsoSegmentType.msoSegmentLine, corner, at + p.X, at + p.Y); i++; }
            }
            b.AddNodes(Office.MsoSegmentType.msoSegmentLine, corner, at + nodes[0].X, at + nodes[0].Y);
            var f = b.ConvertToShape();
            // the copy must cover exactly the original's box, otherwise we misread the points
            if (Math.Abs(f.Left - at) > 0.5f || Math.Abs(f.Top - at) > 0.5f || Math.Abs(f.Width - s.Width) > 0.5f || Math.Abs(f.Height - s.Height) > 0.5f)
            {
                f.Delete();
                return null;
            }
            return f;
        }

        private static void Blacken(PowerPoint.Shape d)
        {
            try { d.Shadow.Visible = Office.MsoTriState.msoFalse; } catch { }
            try { d.Glow.Radius = 0; } catch { }
            try { d.SoftEdge.Type = Office.MsoSoftEdgeType.msoSoftEdgeTypeNone; } catch { }
            try { d.Reflection.Type = Office.MsoReflectionType.msoReflectionTypeNone; } catch { }
            try { d.ThreeD.Visible = Office.MsoTriState.msoFalse; } catch { }
            d.Fill.Visible = Office.MsoTriState.msoTrue; d.Fill.Solid(); d.Fill.ForeColor.RGB = 0; d.Fill.Transparency = 0;
            d.Line.Visible = Office.MsoTriState.msoTrue; d.Line.Weight = 0.5f; d.Line.ForeColor.RGB = 0; d.Line.Transparency = 0;
        }

        private static Dictionary<int, PowerPoint.Shape> IndexShapes(PowerPoint.Slide slide)
        {
            var map = new Dictionary<int, PowerPoint.Shape>();
            foreach (PowerPoint.Shape s in slide.Shapes) map[s.Id] = s;
            return map;
        }

        private static List<PowerPoint.Shape> LinksOn(PowerPoint.Slide slide)
        {
            var list = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape s in slide.Shapes) if (Link.IsLink(s)) list.Add(s);
            return list;
        }

        private static string UniqueName(PowerPoint.Slide slide)
        {
            var used = new HashSet<string>();
            foreach (PowerPoint.Shape s in slide.Shapes) used.Add(s.Name);
            for (int i = 1; ; i++) { string n = "OrthoLink " + i; if (!used.Contains(n)) return n; }
        }

        // ---- scratch deck + clipboard injection

        /// <summary>The one slide of a hidden, windowless deck: white, no master graphics, emptied.</summary>
        private PowerPoint.Slide ScratchSlide()
        {
            bool alive = false;
            try { alive = _scratch != null && _scratch.Slides.Count > 0; } catch { alive = false; }
            if (!alive)
            {
                _scratch = _app.Presentations.Add(Office.MsoTriState.msoFalse);
                _scratch.Tags.Add("OL_SCRATCH", "1");   // lets scripts tell it apart from user decks
                var s = _scratch.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                s.FollowMasterBackground = Office.MsoTriState.msoFalse;
                s.Background.Fill.Solid();
                s.Background.Fill.ForeColor.RGB = 0xFFFFFF;
                s.DisplayMasterShapes = Office.MsoTriState.msoFalse;
            }
            var slide = _scratch.Slides[1];
            for (int i = slide.Shapes.Count; i >= 1; i--) slide.Shapes[i].Delete();   // leftovers of an earlier failure
            return slide;
        }

        /// <summary>Put a connector of the given topology onto the slide. PowerPoint only lets such a shape in
        /// through the clipboard, so we borrow it: back up what is on it, paste our shape, put it all back.</summary>
        private PowerPoint.Shape Inject(PowerPoint.Slide slide, string topo)
        {
            var saved = Clip.Save();
            try
            {
                Exception last = null;
                for (int attempt = 0; attempt < 3; attempt++)
                {
                    try
                    {
                        Clip.Put(Clip.Gvml, Template(topo));
                        var range = slide.Shapes.Paste();
                        var s = range[1];
                        if (range.Count == 1 && s.Name == "OL_" + topo) return s;
                        Log.Write("paste gave " + s.Name + " instead of OL_" + topo + ", retrying");
                        range.Delete();
                    }
                    catch (Exception ex) { last = ex; }
                    System.Threading.Thread.Sleep(50 * (attempt + 1));
                }
                throw last ?? new InvalidOperationException("没能把连接线贴到幻灯片上。");
            }
            finally
            {
                try { Clip.Restore(saved); } catch (Exception ex) { Log.Error("Clip.Restore", ex); }
            }
        }

        /// <summary>The connector shape for a topology, as generated by tools/build_template.py.</summary>
        private byte[] Template(string topo)
        {
            byte[] data;
            if (_templates.TryGetValue(topo, out data)) return data;
            using (var s = Assembly.GetExecutingAssembly().GetManifestResourceStream("OrthoLink.OL_" + topo + ".gvml"))
            {
                if (s == null) throw new InvalidOperationException("missing connector template OL_" + topo);
                var m = new MemoryStream();
                s.CopyTo(m);
                data = m.ToArray();
            }
            _templates[topo] = data;
            return data;
        }
    }
}
