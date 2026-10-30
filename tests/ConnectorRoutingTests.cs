using System;
using System.Drawing;
using System.Reflection;
using OrthoLink;

internal static class ConnectorRoutingTests
{
    // Exercise the geometry in the built add-in without starting PowerPoint or exposing internal types.
    private static readonly Type SideType = typeof(Connect).Assembly.GetType("OrthoLink.Side", true);
    private static readonly MethodInfo Plan = typeof(Connect).Assembly.GetType("OrthoLink.Geo", true).GetMethod("Plan");

    private static PointF Outward(string side)
    {
        switch (side)
        {
            case "Left": return new PointF(-1, 0);
            case "Right": return new PointF(1, 0);
            case "Top": return new PointF(0, -1);
            default: return new PointF(0, 1);
        }
    }

    private static PointF Offset(PointF p, PointF direction, float distance)
    {
        return new PointF(p.X + direction.X * distance, p.Y + direction.Y * distance);
    }

    private static RectangleF BoxAtPort(PointF port, PointF outward, float size)
    {
        var center = Offset(port, outward, -size / 2);
        return new RectangleF(center.X - size / 2, center.Y - size / 2, size, size);
    }

    private static void Expect(string expected, string source, string target, RectangleF a, RectangleF b,
                               PointF start, PointF end, string label)
    {
        object route = Plan.Invoke(null, new object[] { Enum.Parse(SideType, source), Enum.Parse(SideType, target), a, b, start, end });
        string actual = (string)route.GetType().GetField("Topo").GetValue(route);
        if (actual != expected)
            throw new Exception("FAIL: " + label + " " + source + " -> " + target + ": expected " + expected + ", got " + actual);
    }

    public static int Run()
    {
        int assertions = 0;
        foreach (string source in new[] { "Right", "Left", "Top", "Bottom" })
        {
            bool horizontal = source == "Left" || source == "Right";
            foreach (string target in horizontal ? new[] { "Top", "Bottom" } : new[] { "Left", "Right" })
            {
                PointF outward = Outward(source), targetOutward = Outward(target), start = PointF.Empty;
                RectangleF a = BoxAtPort(start, outward, 80);
                string elbow = horizontal ? "HV" : "VH", detour = horizontal ? "HVHV" : "VHVH";

                // Both short axes and crossing the old 20 pt threshold must keep a single elbow.
                foreach (float first in new[] { 0.25f, 5f, 19.99f, 20f, 20.01f, 80f })
                    foreach (float last in new[] { 0.25f, 5f, 19.99f, 20f, 20.01f, 80f })
                    {
                        PointF end = Offset(Offset(start, outward, first), targetOutward, -last);
                        Expect(elbow, source, target, a, BoxAtPort(end, targetOutward, 16), start, end,
                               "clear elbow with lengths " + first + ", " + last);
                        assertions++;
                    }

                // A zero or backwards segment cannot preserve the requested exit/entry direction.
                foreach (float length in new[] { 0f, -0.25f, -30f })
                    for (int axis = 0; axis < 2; axis++)
                    {
                        PointF end = Offset(Offset(start, outward, axis == 0 ? length : 80), targetOutward, axis == 1 ? -length : -80);
                        Expect(detour, source, target, a, BoxAtPort(end, targetOutward, 16), start, end, "invalid direction");
                        assertions++;
                    }

                // Real outline ports may sit inside the frame; a turn must still clear both frames.
                PointF recessedStart = Offset(start, outward, -15);
                PointF recessedEnd = Offset(Offset(recessedStart, outward, 10), targetOutward, -80);
                Expect(detour, source, target, a, BoxAtPort(recessedEnd, targetOutward, 16), recessedStart, recessedEnd, "source frame blocks elbow");
                PointF endInsideFrame = Offset(Offset(start, outward, 80), targetOutward, -10);
                RectangleF recessedTarget = BoxAtPort(Offset(endInsideFrame, targetOutward, 15), targetOutward, 16);
                Expect(detour, source, target, a, recessedTarget, start, endInsideFrame, "target frame blocks elbow");
                assertions += 2;
            }
        }

        Expect("HVH", "Right", "Left", new RectangleF(0, 0, 40, 40), new RectangleF(45, 10, 40, 40),
               new PointF(40, 20), new PointF(45, 30), "nearby facing horizontal ports");
        Expect("VHV", "Bottom", "Top", new RectangleF(0, 0, 40, 40), new RectangleF(10, 45, 40, 40),
               new PointF(20, 40), new PointF(30, 45), "nearby facing vertical ports");
        assertions += 2;
        Console.WriteLine("PASS connector routing (" + assertions + " assertions)");
        assertions += Radii();
        return assertions;
    }

    private static void ExpectRadius(string topo, PointF start, PointF end, float[] interior, float limit, float expected)
    {
        var method = typeof(Connect).Assembly.GetType("OrthoLink.Geo", true).GetMethod("EffectiveRadius");
        float actual = (float)method.Invoke(null, new object[] { topo, start, end, interior, limit });
        if (Math.Abs(actual - expected) > 0.001f)
            throw new Exception("FAIL: effective radius for " + topo + ": expected " + expected + ", got " + actual);
    }

    private static int Radii()
    {
        // Use explicit lengths, including a middle segment shared by two arcs and backwards detours.
        ExpectRadius("HV", PointF.Empty, new PointF(100, 3), new float[0], 8, 3);
        ExpectRadius("VH", PointF.Empty, new PointF(3, 100), new float[0], 8, 3);
        ExpectRadius("HV", new PointF(100, 20), new PointF(0, 17), new float[0], 8, 3);
        ExpectRadius("VH", new PointF(100, 20), new PointF(97, 120), new float[0], 8, 3);
        ExpectRadius("HV", PointF.Empty, new PointF(100, 30), new float[0], 8, 8);
        ExpectRadius("HV", PointF.Empty, new PointF(100, 3), new float[0], 0, 0);
        ExpectRadius("HVH", PointF.Empty, new PointF(100, 4), new[] { 50f }, 8, 2);
        ExpectRadius("VHV", PointF.Empty, new PointF(4, 100), new[] { 50f }, 8, 2);
        ExpectRadius("HVH", PointF.Empty, new PointF(100, 30), new[] { 1f }, 8, 1);
        ExpectRadius("HVHV", PointF.Empty, new PointF(100, 80), new[] { 15f, 2f }, 8, 1);
        ExpectRadius("VHVH", PointF.Empty, new PointF(80, 100), new[] { 15f, 2f }, 8, 1);
        ExpectRadius("HVHVH", PointF.Empty, new PointF(100, 80), new[] { -20f, -30f, 140f }, 24, 15);
        ExpectRadius("VHVHV", PointF.Empty, new PointF(80, 100), new[] { -20f, -30f, 140f }, 24, 15);
        ExpectRadius("HVH", PointF.Empty, new PointF(100, 0), new[] { 50f }, 8, 0);
        Console.WriteLine("PASS effective connector radius (14 assertions)");
        return 14;
    }
}
