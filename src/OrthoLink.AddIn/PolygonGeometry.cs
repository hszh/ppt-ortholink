using System;
using System.Collections.Generic;
using System.Drawing;

namespace OrthoLink
{
    internal sealed class PolygonSegment
    {
        public bool Curve;
        public PointF Control1, Control2, End;
    }

    internal sealed class PolygonPath
    {
        public PointF Start;
        public float Radius;
        public readonly List<PolygonSegment> Segments = new List<PolygonSegment>();
    }

    /// <summary>Convex polygon fillets in points. Independent of PowerPoint and the clipboard.</summary>
    internal static class PolygonGeometry
    {
        private const double Epsilon = 0.0001;

        public static bool Finite(float v) { return !float.IsNaN(v) && !float.IsInfinity(v); }
        private static PointF Add(PointF a, PointF b, double scale) { return new PointF((float)(a.X + b.X * scale), (float)(a.Y + b.Y * scale)); }
        private static PointF Sub(PointF a, PointF b) { return new PointF(a.X - b.X, a.Y - b.Y); }
        private static double Length(PointF a) { return Math.Sqrt((double)a.X * a.X + (double)a.Y * a.Y); }
        private static double Cross(PointF a, PointF b) { return (double)a.X * b.Y - (double)a.Y * b.X; }
        private static PointF Unit(PointF a) { double d = Length(a); return new PointF((float)(a.X / d), (float)(a.Y / d)); }

        public static PolygonPath Round(IList<PointF> vertices, float radius)
        {
            if (!Finite(radius) || radius < 0) throw new ArgumentOutOfRangeException("radius");
            var p = new List<PointF>();
            foreach (var v in vertices)
            {
                if (!Finite(v.X) || !Finite(v.Y)) throw new ArgumentException("形状顶点包含无效坐标。");
                if (p.Count == 0 || Length(Sub(v, p[p.Count - 1])) > Epsilon) p.Add(v);
            }
            if (p.Count > 1 && Length(Sub(p[0], p[p.Count - 1])) < Epsilon) p.RemoveAt(p.Count - 1);
            // Adjusting a trapezoid to an extreme can produce duplicate or collinear vertices.
            bool removed = true;
            while (removed && p.Count >= 3)
            {
                removed = false;
                for (int i = 0; i < p.Count; i++)
                {
                    PointF a = Sub(p[i], p[(i + p.Count - 1) % p.Count]), b = Sub(p[(i + 1) % p.Count], p[i]);
                    if (Math.Abs(Cross(a, b)) < Epsilon * (Length(a) + Length(b)) && (double)a.X * b.X + (double)a.Y * b.Y >= 0)
                    { p.RemoveAt(i); removed = true; break; }
                }
            }
            if (p.Count < 3) throw new ArgumentException("形状过于扁平，无法设置圆角。");
            int n = p.Count, sign = 0;
            var edge = new PointF[n]; var length = new double[n]; var angle = new double[n]; var tangent = new double[n];
            for (int i = 0; i < n; i++)
            {
                PointF d = Sub(p[(i + 1) % n], p[i]); length[i] = Length(d); edge[i] = Unit(d);
            }
            for (int i = 0; i < n; i++)
            {
                PointF a = edge[(i + n - 1) % n], b = edge[i];
                double cross = Cross(a, b), dot = (double)a.X * b.X + (double)a.Y * b.Y;
                int turn = Math.Sign(cross);
                if (turn == 0 || (sign != 0 && turn != sign)) throw new ArgumentException("只支持具有有效面积的凸多边形。");
                sign = turn;
                angle[i] = Math.Atan2(Math.Abs(cross), dot);
                tangent[i] = Math.Tan(angle[i] / 2);
            }
            double effective = radius;
            for (int i = 0; i < n; i++) effective = Math.Min(effective, length[i] / (tangent[i] + tangent[(i + 1) % n]));
            var path = new PolygonPath { Radius = (float)effective };
            if (effective < Epsilon)
            {
                path.Start = p[0]; path.Radius = 0;
                for (int i = 1; i < n; i++) path.Segments.Add(new PolygonSegment { End = p[i] });
                return path;
            }
            var entry = new PointF[n]; var exit = new PointF[n];
            for (int i = 0; i < n; i++)
            {
                entry[i] = Add(p[i], edge[(i + n - 1) % n], -effective * tangent[i]);
                exit[i] = Add(p[i], edge[i], effective * tangent[i]);
            }
            path.Start = entry[0];
            for (int i = 0; i < n; i++)
            {
                PointF incoming = edge[(i + n - 1) % n];
                PointF center = Add(entry[i], new PointF(-incoming.Y, incoming.X), sign * effective);
                // COM coordinates are floats: a nominal 90-degree corner may drift either side.
                int parts = Math.Max(1, (int)Math.Ceiling(angle[i] / (Math.PI / 2) - 0.00001));
                double sweep = sign * angle[i] / parts, k = 4.0 / 3.0 * Math.Tan(sweep / 4);
                PointF start = entry[i];
                for (int j = 1; j <= parts; j++)
                {
                    PointF d = Sub(start, center);
                    PointF end = j == parts ? exit[i] : Add(center, new PointF(
                        (float)(d.X * Math.Cos(sweep) - d.Y * Math.Sin(sweep)),
                        (float)(d.X * Math.Sin(sweep) + d.Y * Math.Cos(sweep))), 1);
                    PointF e = Sub(end, center);
                    path.Segments.Add(new PolygonSegment { Curve = true,
                        Control1 = Add(start, new PointF(-d.Y, d.X), k),
                        Control2 = Add(end, new PointF(-e.Y, e.X), -k), End = end });
                    start = end;
                }
                if (i + 1 < n && Length(Sub(entry[i + 1], exit[i])) > Epsilon)
                    path.Segments.Add(new PolygonSegment { End = entry[i + 1] });
            }
            return path; // the final straight edge is the shape's existing closing segment
        }
    }
}
