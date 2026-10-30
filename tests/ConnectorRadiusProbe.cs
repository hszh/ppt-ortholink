using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using OrthoLink;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

internal static class ConnectorRadiusProbe
{
    private static object Call(object service, string method, params object[] args)
    {
        return service.GetType().GetMethod(method).Invoke(service, args);
    }

    [STAThread]
    private static int Main(string[] args)
    {
        PowerPoint.Application app = null;
        PowerPoint.Presentation pres = null;
        object service = null;
        float radius = args.Length == 0 ? 12f : float.Parse(args[0], System.Globalization.CultureInfo.InvariantCulture);
        bool arrows = Array.IndexOf(args, "--arrows") >= 0;
        bool sweep = Array.IndexOf(args, "--sweep") >= 0;
        bool verify = Array.IndexOf(args, "--verify") >= 0;
        string output = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "connector-radius-probe",
                                     "radius-" + radius.ToString(System.Globalization.CultureInfo.InvariantCulture) + (arrows ? "-arrows" : "") + (sweep ? "-sweep" : "") + (verify ? "-verify" : ""));
        Directory.CreateDirectory(output);
        try
        {
            app = new PowerPoint.Application();
            pres = app.Presentations.Add(Office.MsoTriState.msoFalse);
            pres.PageSetup.SlideWidth = 300;
            pres.PageSetup.SlideHeight = 180;
            var assembly = typeof(Connect).Assembly;
            var serviceType = assembly.GetType("OrthoLink.ConnectorService", true);
            var sideType = assembly.GetType("OrthoLink.Side", true);
            service = Activator.CreateInstance(serviceType, new object[] { app });
            var slide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            slide.FollowMasterBackground = Office.MsoTriState.msoFalse;
            slide.Background.Fill.Solid();
            slide.Background.Fill.ForeColor.RGB = 0xFFFFFF;
            var source = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 20, 60, 100, 60);
            var target = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, 208, 36, 24, 24);
            foreach (var shape in new[] { source, target })
            {
                shape.Fill.ForeColor.RGB = 0x805010;
                shape.Line.ForeColor.RGB = 0x402010;
            }
            var line = (PowerPoint.Shape)Call(service, "Create", slide, source, target);
            line = (PowerPoint.Shape)Call(service, "SetSide", slide, line, true, Enum.Parse(sideType, "Right"));
            line = (PowerPoint.Shape)Call(service, "SetSide", slide, line, false, Enum.Parse(sideType, "Bottom"));
            Call(service, "SetRadius", line, radius);
            line.Line.Weight = arrows ? 1.5f : 1f;
            line.Line.ForeColor.RGB = 0;
            line.Line.BeginArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadNone;
            line.Line.EndArrowheadStyle = arrows ? Office.MsoArrowheadStyle.msoArrowheadTriangle : Office.MsoArrowheadStyle.msoArrowheadNone;

            int step = 0;
            foreach (float height in new[] { 30f, 12f, 5f, 3f, 1f, 30f })
            {
                target.Top = 90 - target.Height - height;
                string name = line.Name;
                Call(service, "Follow", slide);
                line = slide.Shapes[name];
                string label = height.ToString(System.Globalization.CultureInfo.InvariantCulture).Replace('.', '_');
                if (step++ == 5) label += "-restored";
                foreach (string variant in sweep ? new[] { "requested", "half", "sharp", "short-arrow" } : new[] { "requested" })
                {
                    if (sweep)
                    {
                        float effective = variant == "half" ? Math.Min(radius, line.Height / 2) : variant == "sharp" ? 0 : radius;
                        Call(service, "SetRadius", line, effective);
                        line.Line.EndArrowheadLength = variant == "short-arrow" ? Office.MsoArrowheadLength.msoArrowheadShort : Office.MsoArrowheadLength.msoArrowheadLengthMedium;
                    }
                    string suffix = sweep ? "-" + variant : "";
                    float displayed = (float)serviceType.GetMethod("GetRadius").Invoke(null, new object[] { line });
                    float setting = line.Adjustments[1] / 0.127f;
                    Console.WriteLine("height=" + line.Height + " width=" + line.Width + " topology=" + line.Tags["OL_TOPO"] +
                                      " setting=" + setting + " displayed=" + displayed + " variant=" + variant);
                    if (verify)
                    {
                        float expected = Math.Min(radius, Math.Min(line.Width, line.Height));
                        if (Math.Abs(displayed - expected) > 0.001f)
                            throw new Exception("FAIL: displayed radius expected " + expected + ", got " + displayed);
                        if (Math.Abs(setting - radius) > 0.001f)
                            throw new Exception("FAIL: moving shapes changed the requested radius limit");
                    }
                    slide.Export(Path.Combine(output, "height-" + label + suffix + ".png"), "PNG", 1800, 1080);
                    try { line.Export(Path.Combine(output, "height-" + label + suffix + ".svg"), (PowerPoint.PpShapeFormat)6); }
                    catch (COMException ex) { Console.WriteLine("SVG export unavailable: " + ex.Message); }
                }
            }
            if (verify)
            {
                line = (PowerPoint.Shape)Call(service, "SetSide", slide, line, false, Enum.Parse(sideType, "Left"));
                if (line.Tags["OL_TOPO"] != "HVH") throw new Exception("FAIL: expected a three-segment handle test route");
                float width = line.Width, height = line.Height;
                line.Adjustments[2] = 0.02f;
                float expected = Math.Min(radius, width * 0.02f);
                float displayed = (float)serviceType.GetMethod("GetRadius").Invoke(null, new object[] { line });
                if (Math.Abs(displayed - expected) > 0.001f || line.Width != width || line.Height != height)
                    throw new Exception("FAIL: radius did not follow a yellow handle within the same frame");
                string saved = Path.Combine(output, "short-radius.pptx");
                string lineName = line.Name;
                pres.SaveAs(saved, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
                pres.Close();
                Marshal.ReleaseComObject(pres);
                pres = null;
                pres = app.Presentations.Open(saved, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
                line = pres.Slides[1].Shapes[lineName];
                float persistedRadius = (float)serviceType.GetMethod("GetRadius").Invoke(null, new object[] { line });
                if (Math.Abs(persistedRadius - expected) > 0.001f || Math.Abs(line.Adjustments[1] / 0.127f - radius) > 0.001f)
                    throw new Exception("FAIL: effective radius or requested limit changed on reopen");
                line.Adjustments[2] = 0.5f;
                displayed = (float)serviceType.GetMethod("GetRadius").Invoke(null, new object[] { line });
                if (Math.Abs(displayed - radius) > 0.001f) throw new Exception("FAIL: radius did not recover after moving the handle back");
                Console.WriteLine("PASS shrinking, recovery, yellow handles, requested limit and save/reopen");
            }
            pres.SaveAs(Path.Combine(output, "radius-probe.pptx"), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
            Console.WriteLine("Probe output: " + output);
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex); return 1; }
        finally
        {
            if (service != null) Call(service, "Shutdown");
            if (pres != null) { pres.Saved = Office.MsoTriState.msoTrue; pres.Close(); Marshal.ReleaseComObject(pres); }
            if (app != null) { if (app.Presentations.Count == 0) app.Quit(); Marshal.ReleaseComObject(app); }
            GC.Collect(); GC.WaitForPendingFinalizers();
        }
    }
}
