using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OrthoLink
{
    /// <summary>Edits the original shape in place: its ID, text, style, z-order and animation survive.</summary>
    internal static class PolygonRounder
    {
        private const string Prefix = "OL_POLY_";
        private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

        internal static bool SupportsKind(Office.MsoAutoShapeType kind)
        {
            switch (kind)
            {
                case Office.MsoAutoShapeType.msoShapeRectangle:
                case Office.MsoAutoShapeType.msoShapeIsoscelesTriangle:
                case Office.MsoAutoShapeType.msoShapeRightTriangle:
                case Office.MsoAutoShapeType.msoShapeTrapezoid:
                case Office.MsoAutoShapeType.msoShapeNonIsoscelesTrapezoid:
                case Office.MsoAutoShapeType.msoShapeParallelogram:
                case Office.MsoAutoShapeType.msoShapeDiamond:
                case Office.MsoAutoShapeType.msoShapePentagon:
                case Office.MsoAutoShapeType.msoShapeRegularPentagon:
                case Office.MsoAutoShapeType.msoShapeHexagon:
                case Office.MsoAutoShapeType.msoShapeHeptagon:
                case Office.MsoAutoShapeType.msoShapeOctagon: return true;
                default: return false;
            }
        }

        public static bool IsSupported(PowerPoint.Shape s)
        {
            try
            {
                // Groups need a different coordinate system; require ungrouping first.
                if (s.Child == Office.MsoTriState.msoTrue || !(s.Parent is PowerPoint.Slide) ||
                    s.HasTextFrame != Office.MsoTriState.msoTrue || Link.IsLink(s)) return false;
                return (s.Type == Office.MsoShapeType.msoAutoShape && SupportsKind(s.AutoShapeType)) ||
                       (s.Type == Office.MsoShapeType.msoFreeform && s.Tags[Prefix + "VER"] == "1");
            }
            catch { return false; }
        }

        public static float GetRadius(PowerPoint.Shape s)
        {
            return s.Type == Office.MsoShapeType.msoFreeform ? Number(s.Tags[Prefix + "RADIUS"]) : 0f;
        }

        private sealed class Definition
        {
            public Office.MsoAutoShapeType Kind;
            public float[] Adjustments;
            public List<PointF> Vertices;
            public RectangleF Frame;
        }

        private sealed class Transform
        {
            public RectangleF Frame;
            public float Rotation;
            public bool FlipH, FlipV;
            public Office.MsoTriState Lock;

            public Transform(PowerPoint.Shape s)
            {
                Frame = Bounds(s); Rotation = s.Rotation; Lock = s.LockAspectRatio;
                FlipH = s.HorizontalFlip == Office.MsoTriState.msoTrue;
                FlipV = s.VerticalFlip == Office.MsoTriState.msoTrue;
            }

            public PointF Canonical(PointF p)
            {
                double a = -Rotation * Math.PI / 180, dx = p.X - Center(Frame).X, dy = p.Y - Center(Frame).Y;
                double x = dx * Math.Cos(a) - dy * Math.Sin(a), y = dx * Math.Sin(a) + dy * Math.Cos(a);
                return new PointF((float)(Center(Frame).X + (FlipH ? -x : x)), (float)(Center(Frame).Y + (FlipV ? -y : y)));
            }

            public void Apply(PowerPoint.Shape s)
            {
                RectangleF after = Bounds(s);
                PointF beforeCenter = Center(Frame), afterCenter = Center(after);
                double dx = afterCenter.X - beforeCenter.X, dy = afterCenter.Y - beforeCenter.Y;
                double x = FlipH ? -dx : dx, y = FlipV ? -dy : dy, a = Rotation * Math.PI / 180;
                SetFlips(s, FlipH, FlipV); s.Rotation = Rotation;
                // Node editing tightens the bounding box. Rotate/flip about the original center.
                s.Left = (float)(after.Left - dx + x * Math.Cos(a) - y * Math.Sin(a));
                s.Top = (float)(after.Top - dy + x * Math.Sin(a) + y * Math.Cos(a));
                s.LockAspectRatio = Lock;
            }
        }

        public static bool SetRadius(PowerPoint.Shape s, float radius)
        {
            if (!PolygonGeometry.Finite(radius) || radius < 0 || radius > 200)
                throw new ArgumentOutOfRangeException("radius", "圆角半径须为 0 到 200 pt 的有限数字。");
            if (!IsSupported(s)) return false;
            bool rounded = s.Type == Office.MsoShapeType.msoFreeform;
            if (!rounded && radius == 0) return false;
            string originalName = s.Name;
            var transform = new Transform(s);
            if (transform.Frame.Width < 0.01f || transform.Frame.Height < 0.01f)
                throw new InvalidOperationException("形状过于扁平，无法设置圆角。");
            Definition def = rounded ? ReadDefinition(s) : ReadPreset(s);
            var vertices = new List<PointF>();
            foreach (PointF p in def.Vertices) vertices.Add(new PointF(def.Frame.Left + p.X * def.Frame.Width, def.Frame.Top + p.Y * def.Frame.Height));
            PolygonPath path = PolygonGeometry.Round(vertices, radius); // validate before touching the source
            PolygonPath oldPath = rounded ? ReadPath(s, transform) : null;
            var oldTags = new Dictionary<string, string>();
            for (int i = 1; i <= s.Tags.Count; i++)
            {
                string key = s.Tags.Name(i);
                if (key.StartsWith(Prefix, StringComparison.Ordinal)) oldTags[key] = s.Tags.Value(i);
            }
            try
            {
                Flatten(s);
                if (radius == 0)
                {
                    RestorePreset(s, def);
                    transform.Apply(s);
                    ClearTags(s);
                }
                else
                {
                    WritePath(s, path);
                    RectangleF after = Bounds(s);
                    if (after.Width < 0.01f || after.Height < 0.01f) throw new InvalidOperationException("圆角结果尺寸无效。");
                    transform.Apply(s);
                    SaveDefinition(s, def, after, path.Radius);
                }
                return true;
            }
            catch
            {
                try
                {
                    Flatten(s);
                    if (oldPath != null) WritePath(s, oldPath); else RestorePreset(s, def);
                    transform.Apply(s);
                    ClearTags(s);
                    foreach (var tag in oldTags) s.Tags.Add(tag.Key, tag.Value);
                }
                catch (Exception ex) { Log.Error("PolygonRounder rollback", ex); }
                throw;
            }
            finally { s.Name = originalName; s.LockAspectRatio = transform.Lock; }
        }

        private static Definition ReadPreset(PowerPoint.Shape s)
        {
            var def = new Definition { Kind = s.AutoShapeType, Frame = Bounds(s), Adjustments = new float[s.Adjustments.Count], Vertices = new List<PointF>() };
            for (int i = 0; i < def.Adjustments.Length; i++) def.Adjustments[i] = s.Adjustments[i + 1];
            // Preset Nodes.Points throws in PowerPoint. Sample only our disposable duplicate.
            PowerPoint.Shape copy = null;
            try
            {
                copy = s.Duplicate()[1]; Flatten(copy);
                copy.Left = def.Frame.Left; copy.Top = def.Frame.Top;
                copy.Nodes.SetEditingType(1, Office.MsoEditingType.msoEditingCorner);
                for (int i = 1; i < copy.Nodes.Count; i++)
                {
                    if (copy.Nodes[i].SegmentType != Office.MsoSegmentType.msoSegmentLine) throw new InvalidOperationException("该形状包含曲线，无法作为多边形处理。");
                    PointF p = Point(copy.Nodes[i]);
                    def.Vertices.Add(new PointF((p.X - def.Frame.Left) / def.Frame.Width, (p.Y - def.Frame.Top) / def.Frame.Height));
                }
                return def;
            }
            finally { if (copy != null) copy.Delete(); }
        }

        private static void RestorePreset(PowerPoint.Shape s, Definition def)
        {
            s.AutoShapeType = def.Kind;
            s.Width = def.Frame.Width; s.Height = def.Frame.Height;
            s.Left = def.Frame.Left; s.Top = def.Frame.Top;
            for (int i = 0; i < def.Adjustments.Length; i++) s.Adjustments[i + 1] = def.Adjustments[i];
        }

        private static PolygonPath ReadPath(PowerPoint.Shape s, Transform t)
        {
            var path = new PolygonPath { Start = t.Canonical(Point(s.Nodes[1])) };
            for (int i = 2; i <= s.Nodes.Count; )
            {
                if (s.Nodes[i].SegmentType == Office.MsoSegmentType.msoSegmentCurve)
                {
                    if (i + 2 > s.Nodes.Count) throw new InvalidOperationException("形状曲线节点不完整。");
                    path.Segments.Add(new PolygonSegment { Curve = true, Control1 = t.Canonical(Point(s.Nodes[i])),
                        Control2 = t.Canonical(Point(s.Nodes[i + 1])), End = t.Canonical(Point(s.Nodes[i + 2])) }); i += 3;
                }
                else { path.Segments.Add(new PolygonSegment { End = t.Canonical(Point(s.Nodes[i])) }); i++; }
            }
            return path;
        }

        private static void WritePath(PowerPoint.Shape s, PolygonPath path)
        {
            var nodes = s.Nodes;
            nodes.SetEditingType(1, Office.MsoEditingType.msoEditingCorner);
            while (nodes.Count > 2) nodes.Delete(nodes.Count - 1);
            // PowerPoint does not move the repeated closing vertex when moving the first node.
            nodes.SetPosition(1, path.Start.X, path.Start.Y);
            nodes.SetPosition(2, path.Start.X, path.Start.Y);
            int index = 1;
            foreach (var seg in path.Segments)
            {
                if (seg.Curve)
                {
                    nodes.Insert(index, Office.MsoSegmentType.msoSegmentCurve, Office.MsoEditingType.msoEditingCorner,
                        seg.Control1.X, seg.Control1.Y, seg.Control2.X, seg.Control2.Y, seg.End.X, seg.End.Y); index += 3;
                }
                else
                {
                    nodes.Insert(index, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, seg.End.X, seg.End.Y); index++;
                }
            }
        }

        private static Definition ReadDefinition(PowerPoint.Shape s)
        {
            int kind, count, adjustments;
            if (!int.TryParse(s.Tags[Prefix + "KIND"], out kind) || !SupportsKind((Office.MsoAutoShapeType)kind) ||
                !int.TryParse(s.Tags[Prefix + "COUNT"], out count) || count < 3 || count > 12 ||
                !int.TryParse(s.Tags[Prefix + "ADJS"], out adjustments) || adjustments < 0 || adjustments > 8)
                throw new InvalidOperationException("形状圆角记录不完整，请撤销后重试。");
            float[] frame = Numbers(s.Tags[Prefix + "FRAME"], 4); RectangleF current = Bounds(s);
            var def = new Definition { Kind = (Office.MsoAutoShapeType)kind, Adjustments = new float[adjustments], Vertices = new List<PointF>(),
                Frame = new RectangleF(current.Left + frame[0] * current.Width, current.Top + frame[1] * current.Height, frame[2] * current.Width, frame[3] * current.Height) };
            if (def.Frame.Width < 0.01f || def.Frame.Height < 0.01f) throw new InvalidOperationException("原始多边形尺寸无效。");
            for (int i = 0; i < adjustments; i++) def.Adjustments[i] = Number(s.Tags[Prefix + "A" + i]);
            for (int i = 0; i < count; i++) { float[] p = Numbers(s.Tags[Prefix + "V" + i], 2); def.Vertices.Add(new PointF(p[0], p[1])); }
            return def;
        }

        private static void SaveDefinition(PowerPoint.Shape s, Definition def, RectangleF after, float radius)
        {
            var tags = s.Tags;
            tags.Add(Prefix + "KIND", ((int)def.Kind).ToString(Invariant));
            tags.Add(Prefix + "COUNT", def.Vertices.Count.ToString(Invariant));
            tags.Add(Prefix + "ADJS", def.Adjustments.Length.ToString(Invariant));
            tags.Add(Prefix + "FRAME", Join((def.Frame.Left - after.Left) / after.Width, (def.Frame.Top - after.Top) / after.Height, def.Frame.Width / after.Width, def.Frame.Height / after.Height));
            for (int i = 0; i < def.Adjustments.Length; i++) tags.Add(Prefix + "A" + i, Join(def.Adjustments[i]));
            for (int i = 0; i < def.Vertices.Count; i++) tags.Add(Prefix + "V" + i, Join(def.Vertices[i].X, def.Vertices[i].Y));
            tags.Add(Prefix + "RADIUS", Join(radius));
            tags.Add(Prefix + "VER", "1");
        }

        private static void ClearTags(PowerPoint.Shape s)
        {
            for (int i = s.Tags.Count; i >= 1; i--) if (s.Tags.Name(i).StartsWith(Prefix, StringComparison.Ordinal)) s.Tags.Delete(s.Tags.Name(i));
        }

        private static void Flatten(PowerPoint.Shape s)
        {
            s.LockAspectRatio = Office.MsoTriState.msoFalse; s.Rotation = 0; SetFlips(s, false, false); s.Rotation = 0;
        }
        private static void SetFlips(PowerPoint.Shape s, bool h, bool v)
        {
            if ((s.HorizontalFlip == Office.MsoTriState.msoTrue) != h) s.Flip(Office.MsoFlipCmd.msoFlipHorizontal);
            if ((s.VerticalFlip == Office.MsoTriState.msoTrue) != v) s.Flip(Office.MsoFlipCmd.msoFlipVertical);
        }
        private static RectangleF Bounds(PowerPoint.Shape s) { return new RectangleF(s.Left, s.Top, s.Width, s.Height); }
        private static PointF Center(RectangleF r) { return new PointF(r.Left + r.Width / 2, r.Top + r.Height / 2); }
        private static PointF Point(PowerPoint.ShapeNode node)
        {
            var p = (Array)node.Points; int r = p.GetLowerBound(0), c = p.GetLowerBound(1);
            return new PointF(Convert.ToSingle(p.GetValue(r, c)), Convert.ToSingle(p.GetValue(r, c + 1)));
        }
        private static string Join(params float[] values) { return string.Join(",", Array.ConvertAll(values, v => v.ToString("R", Invariant))); }
        private static float Number(string text)
        {
            float v;
            if (!float.TryParse(text, NumberStyles.Float, Invariant, out v) || !PolygonGeometry.Finite(v)) throw new InvalidOperationException("形状圆角记录包含无效数字。");
            return v;
        }
        private static float[] Numbers(string text, int count)
        {
            string[] parts = text.Split(',');
            if (parts.Length != count) throw new InvalidOperationException("形状圆角记录不完整。");
            return Array.ConvertAll(parts, Number);
        }
    }
}
