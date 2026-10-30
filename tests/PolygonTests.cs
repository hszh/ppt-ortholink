using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using OrthoLink;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

internal static class PolygonTests
{
    private static int assertions;
    private static void Check(bool condition, string message)
    {
        assertions++;
        if (!condition) throw new Exception("FAIL: " + message);
    }
    private static bool Near(float a, float b) { return Math.Abs(a - b) < 0.08f; }
    private static void Throws(Action action, string message)
    {
        bool threw = false; try { action(); } catch (ArgumentException) { threw = true; }
        Check(threw, message);
    }
    private static void Geometry()
    {
        var square = new[] { new PointF(0,0), new PointF(100,0), new PointF(100,100), new PointF(0,100) };
        var p = PolygonGeometry.Round(square, 10);
        Check(Near(p.Radius, 10) && Near(p.Start.X, 0) && Near(p.Start.Y, 10), "square tangent point");
        Check(p.Segments.Count == 7 && p.Segments[0].Curve, "square has four arcs and three explicit edges");
        Check(Near(p.Segments[0].Control1.X, 0) && Near(p.Segments[0].Control2.Y, 0), "arc tangency");
        Check(Near(p.Segments[0].End.X, 10) && Near(p.Segments[0].End.Y, 0), "arc endpoint");
        Check(Near(PolygonGeometry.Round(square, 200).Radius, 50), "oversized radius limited to half edge");
        Check(PolygonGeometry.Round(square, 0).Segments.TrueForAll(x => !x.Curve), "zero radius stays sharp");
        var triangle = new[] { new PointF(0,100), new PointF(80,0), new PointF(160,100) };
        p = PolygonGeometry.Round(triangle, 200);
        Check(p.Radius > 30 && p.Radius < 40, "triangle maximum is its inradius");
        foreach (var seg in p.Segments)
        {
            Check(PolygonGeometry.Finite(seg.End.X) && PolygonGeometry.Finite(seg.End.Y), "finite arc end");
            if (seg.Curve) Check(PolygonGeometry.Finite(seg.Control1.X) && PolygonGeometry.Finite(seg.Control2.Y), "finite controls");
        }
        Array.Reverse(triangle);
        Check(Near(PolygonGeometry.Round(triangle, 200).Radius, p.Radius), "clockwise and anticlockwise agree");
        Check(PolygonGeometry.Round(new[] { square[0], square[0], new PointF(50,0), square[1], square[2], square[3], square[0] }, 10).Segments.Count == 7, "duplicate and collinear vertices");
        Throws(() => PolygonGeometry.Round(square, float.NaN), "reject NaN");
        Throws(() => PolygonGeometry.Round(square, float.PositiveInfinity), "reject infinity");
        Throws(() => PolygonGeometry.Round(square, -1), "reject negative radius");
        Throws(() => PolygonGeometry.Round(new[] { new PointF(0,0),new PointF(10,0),new PointF(20,0) }, 2), "reject flat polygon");
        Throws(() => PolygonGeometry.Round(new[] { new PointF(0,0),new PointF(100,0),new PointF(30,20),new PointF(100,100),new PointF(0,100) }, 2), "reject concave polygon");
        Check(Connect.ParseRadius("8 pt") == 8 && Connect.ParseRadius("直角") == 0, "radius input units");
        Check(!Connect.ParseRadius("NaN").HasValue && !Connect.ParseRadius("Infinity").HasValue, "reject nonfinite input");
        Console.WriteLine("PASS geometry and input validation");
    }
    private static float[] Frame(PowerPoint.Shape s) { return new[] { s.Left, s.Top, s.Width, s.Height, s.Rotation }; }
    private static void SameFrame(float[] expected, PowerPoint.Shape s, string label)
    {
        float[] actual = Frame(s);
        for (int i = 0; i < expected.Length; i++) Check(Near(expected[i], actual[i]), label + " frame[" + i + "] expected=" + expected[i] + " actual=" + actual[i]);
    }
    private static readonly Office.MsoAutoShapeType[] Kinds = {
        Office.MsoAutoShapeType.msoShapeIsoscelesTriangle, Office.MsoAutoShapeType.msoShapeRightTriangle,
        Office.MsoAutoShapeType.msoShapeTrapezoid, Office.MsoAutoShapeType.msoShapeNonIsoscelesTrapezoid,
        Office.MsoAutoShapeType.msoShapeParallelogram, Office.MsoAutoShapeType.msoShapeDiamond,
        Office.MsoAutoShapeType.msoShapePentagon, Office.MsoAutoShapeType.msoShapeRegularPentagon,
        Office.MsoAutoShapeType.msoShapeHexagon, Office.MsoAutoShapeType.msoShapeHeptagon,
        Office.MsoAutoShapeType.msoShapeOctagon, Office.MsoAutoShapeType.msoShapeRectangle };

    private static void Edges(PowerPoint.Application app, PowerPoint.Presentation pres, PowerPoint.Slide slide, Connect addin)
    {
        var a = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeIsoscelesTriangle, 100, 80, 160, 100);
        var b = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeHexagon, 400, 250, 160, 100);
        var animation = slide.TimeLine.MainSequence.AddEffect(a, PowerPoint.MsoAnimEffect.msoAnimEffectFade);
        int animationCount = slide.TimeLine.MainSequence.Count;
        var native = slide.Shapes.AddConnector(Office.MsoConnectorType.msoConnectorStraight,0,0,100,100);
        native.ConnectorFormat.BeginConnect(a,1); native.ConnectorFormat.EndConnect(b,1);
        addin.SetShapeRadius(a,8);
        Check(animation.Shape.Id == a.Id && slide.TimeLine.MainSequence.Count == animationCount, "animation retained without duplicates");
        Check(native.ConnectorFormat.BeginConnected == Office.MsoTriState.msoTrue && native.ConnectorFormat.BeginConnectedShape.Id == a.Id, "native connector remains attached");
        a.Flip(Office.MsoFlipCmd.msoFlipHorizontal); a.Rotation = 67;
        var moved = Frame(a); addin.SetShapeRadius(a,12); addin.SetShapeRadius(a,8); SameFrame(moved,a,"flip and rotate after rounding");
        addin.SetShapeRadius(a,0);
        var group = slide.Shapes.Range(new object[] {a.Name,b.Name}).Group();
        Check(!addin.CanRoundShape(group), "group excluded");
        Check(!addin.CanRoundShape(group.GroupItems[1]), "child within group excluded");
        group.Ungroup();
        var text = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeTrapezoid,100,350,220,100);
        text.TextFrame.TextRange.Text = "Automatically sized text";
        text.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
        float[] before = Frame(text);
        addin.SetShapeRadius(text,8); addin.SetShapeRadius(text,0);
        Check(text.TextFrame.AutoSize == PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText, "text autosize setting retained");
        SameFrame(before,text,"text autofit roundtrip");
        var undoSlide = pres.Slides.Add(2,PowerPoint.PpSlideLayout.ppLayoutBlank);
        pres.Windows[1].Activate(); pres.Windows[1].View.GotoSlide(2);
        var undoShape = undoSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRegularPentagon,100,100,150,150);
        undoShape.Select(); int undoId = undoShape.Id;
        addin.RoundSelectedShapes(8);
        app.CommandBars.ExecuteMso("Undo");
        undoShape=undoSlide.Shapes[1];
        Check(undoShape.Id==undoId && undoShape.AutoShapeType==Office.MsoAutoShapeType.msoShapeRegularPentagon && undoShape.Tags["OL_POLY_VER"]=="", "single undo restores source");
        app.CommandBars.ExecuteMso("Redo");
        Check(addin.CanRoundShape(undoSlide.Shapes[1]) && Near(addin.GetShapeRadius(undoSlide.Shapes[1]),8), "redo restores radius and metadata");
        Console.WriteLine("PASS animation, native connectors, groups, autofit, undo and redo");
    }

    private static void Com(bool edgesOnly)
    {
        PowerPoint.Application app = null; PowerPoint.Presentation pres = null; Connect addin = null;
        Array custom = new object[0];
        string output = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "polygon-tests"); Directory.CreateDirectory(output);
        try
        {
            app = new PowerPoint.Application();
            pres = app.Presentations.Add(Office.MsoTriState.msoTrue);
            var slide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            addin = new Connect(); addin.OnConnection(app, Extensibility.ext_ConnectMode.ext_cm_Startup, null, ref custom);
            if (edgesOnly) { Edges(app,pres,slide,addin); return; }
            foreach (var kind in Kinds)
            {
                for (int variant = 0; variant < 4; variant++)
                {
                    var s = slide.Shapes.AddShape(kind, 140, 120, 160, 100);
                    if (variant == 3 && s.Adjustments.Count > 0) s.Adjustments[1] = 0.32f;
                    if (variant == 1 || variant == 3) s.Flip(Office.MsoFlipCmd.msoFlipHorizontal);
                    if (variant == 2 || variant == 3) s.Flip(Office.MsoFlipCmd.msoFlipVertical);
                    s.Rotation = variant * 37;
                    s.TextFrame.TextRange.Text = "Bold label\rSecond line";
                    s.TextFrame.TextRange.Characters(1,4).Font.Bold = Office.MsoTriState.msoTrue;
                    s.TextFrame.MarginLeft = 9;
                    s.Fill.ForeColor.RGB = 0xCC8844; s.Fill.BackColor.RGB = 0xFFFFFF;
                    s.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientHorizontal, 1);
                    s.Line.ForeColor.RGB = 0x553311; s.Line.Weight = 2.5f;
                    s.Tags.Add("USER_TAG", "KEEP"); s.AlternativeText = "Keep accessibility text";
                    s.LockAspectRatio = Office.MsoTriState.msoTrue;
                    int id = s.Id, z = s.ZOrderPosition, count = slide.Shapes.Count;
                    string name = s.Name; var frame = Frame(s); var h = s.HorizontalFlip; var v = s.VerticalFlip;
                    float[] adjustments = new float[s.Adjustments.Count];
                    for (int i=0; i<adjustments.Length; i++) adjustments[i] = s.Adjustments[i+1];
                    Check(addin.CanRoundShape(s), "supports " + kind);
                    Check(addin.SetShapeRadius(s, 8), "round " + kind);
                    Check(s.Type == Office.MsoShapeType.msoFreeform && s.Nodes.Count > 8, "editable curves");
                    Check(s.Id == id && s.Name == name && s.ZOrderPosition == z && slide.Shapes.Count == count,
                        "identity and z-order retained: id="+s.Id+"/"+id+" name="+s.Name+"/"+name+" z="+s.ZOrderPosition+"/"+z+" count="+slide.Shapes.Count+"/"+count);
                    Check(s.TextFrame.TextRange.Text == "Bold label\rSecond line" && s.TextFrame.TextRange.Characters(1,4).Font.Bold == Office.MsoTriState.msoTrue, "rich text retained");
                    Check(s.Fill.Type == Office.MsoFillType.msoFillGradient && s.Fill.ForeColor.RGB == 0xCC8844 && Near(s.Line.Weight, 2.5f), "gradient and outline retained");
                    Check(s.Tags["USER_TAG"] == "KEEP" && s.AlternativeText == "Keep accessibility text" && Near(s.TextFrame.MarginLeft, 9), "metadata and text margins retained");
                    Check(s.HorizontalFlip == h && s.VerticalFlip == v && Near(s.Rotation, frame[4]) && s.LockAspectRatio == Office.MsoTriState.msoTrue, "transform retained");
                    float[] rounded = Frame(s); int nodes = s.Nodes.Count;
                    addin.SetShapeRadius(s, 16); addin.SetShapeRadius(s, 8);
                    SameFrame(rounded, s, "reapply " + kind + " variant=" + variant);
                    Check(s.Nodes.Count == nodes, "no accumulating nodes: " + kind + " variant=" + variant + " expected=" + nodes + " actual=" + s.Nodes.Count);
                    Check(Near(addin.GetShapeRadius(s), 8), "radius stored");
                    addin.SetShapeRadius(s, 0);
                    Check(s.Type == Office.MsoShapeType.msoAutoShape && s.AutoShapeType == kind, "restore preset " + kind);
                    SameFrame(frame, s, "restore " + kind + " variant=" + variant);
                    for(int i=0;i<adjustments.Length;i++) Check(Near(adjustments[i],s.Adjustments[i+1]), "restore original adjustment");
                    Check(s.Id == id && s.Tags["OL_POLY_VER"] == "" && s.Tags["USER_TAG"] == "KEEP", "restore tags and identity");
                    s.Delete();
                }
                Console.WriteLine("PASS " + kind + " (normal, rotated, mirrored, adjusted)");
            }
            var a = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeIsoscelesTriangle, 100, 250, 160, 120);
            var b = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeHexagon, 400, 70, 160, 110);
            var ellipse = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, 600, 200, 80, 80);
            Check(!addin.CanRoundShape(ellipse) && !addin.SetShapeRadius(ellipse, 8), "unsupported shape skipped");
            Throws(() => addin.SetShapeRadius(a, float.NaN), "COM API rejects nonfinite radius");
            a.Select(); b.Select(Office.MsoTriState.msoFalse); ellipse.Select(Office.MsoTriState.msoFalse);
            Check(addin.GetHasPolygon(null), "ribbon enabled");
            Check(addin.RoundSelectedShapes(12) == 2, "mixed multi-selection rounds only supported shapes");
            Check(addin.GetShapeRadiusText(null) == "12 pt", "ribbon radius reflects selection");
            addin.SetShapeRadius(b, 6);
            a.Select(); b.Select(Office.MsoTriState.msoFalse);
            Check(addin.GetShapeRadiusText(null) == "多个值", "mixed radii shown explicitly");
            addin.RoundSelectedShapes(0);
            var line = (PowerPoint.Shape)addin.ConnectShapes(a, b);
            line = (PowerPoint.Shape)addin.SetConnectorSide(line, true, "Top");
            float oldBottom = line.Top + line.Height;
            int aId = a.Id, bId = b.Id;
            addin.SetShapeRadius(a, 16); addin.SetShapeRadius(b, 12);
            Check(line.Tags["OL_SRC"] == aId.ToString() && line.Tags["OL_DST"] == bId.ToString(), "connector IDs still valid");
            Check(line.Top + line.Height > oldBottom + 1, "connector follows rounded triangle apex");
            a.Left += 40; addin.FollowSlide(slide);
            Check(line.Tags["OL_SRC"] == a.Id.ToString(), "follow after moving rounded shape");
            Check(!addin.CanRoundShape(line), "connector excluded");
            float w = a.Width, ht = a.Height;
            a.LockAspectRatio = Office.MsoTriState.msoFalse; a.Width = w * 1.5f; a.Height = ht * 0.8f;
            addin.SetShapeRadius(a, 8); addin.SetShapeRadius(a, 0);
            Check(Near(a.Width, 240) && Near(a.Height, 96), "original polygon follows nonuniform scaling");
            addin.SetShapeRadius(a, 200);
            Check(addin.GetShapeRadius(a) > 0 && addin.GetShapeRadius(a) < 200, "large radius safely clamped");
            Console.WriteLine("PASS selection, connector follow, scaling and radius limit");
            string saved = Path.Combine(output, "rounded-polygons.pptx");
            pres.SaveAs(saved, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
            pres.Close(); pres = app.Presentations.Open(saved, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
            slide = pres.Slides[1];
            PowerPoint.Shape persisted = null;
            foreach (PowerPoint.Shape s in slide.Shapes) if (s.Id == bId) persisted = s;
            Check(persisted != null && addin.CanRoundShape(persisted) && Near(addin.GetShapeRadius(persisted),12), "rounding survives reopen");
            addin.SetShapeRadius(persisted, 6); addin.SetShapeRadius(persisted, 0);
            Check(persisted.AutoShapeType == Office.MsoAutoShapeType.msoShapeHexagon, "restore after reopen");
            Console.WriteLine("PASS save/reopen and restoring preset");
            var gallery = pres.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
            pres.PageSetup.SlideWidth = 960; pres.PageSetup.SlideHeight = 540;
            for (int i=0;i<Kinds.Length;i++)
            {
                float x = 25 + (i % 4) * 235, y = 30 + (i / 4) * 170;
                var sharp = gallery.Shapes.AddShape(Kinds[i],x,y+30,90,85);
                sharp.Fill.ForeColor.RGB = 0xE9E4DD; sharp.Line.ForeColor.RGB = 0x8F8170;
                var round = gallery.Shapes.AddShape(Kinds[i],x+110,y+30,90,85);
                round.Fill.ForeColor.RGB = 0xCC8844; round.Line.ForeColor.RGB = 0x553311;
                addin.SetShapeRadius(round,12);
                var label = gallery.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,x,y,220,25);
                label.TextFrame.TextRange.Text = Kinds[i].ToString().Replace("msoShape", ""); label.TextFrame.TextRange.Font.Size = 12;
            }
            gallery.Export(Path.Combine(output,"gallery.png"),"PNG",1920,1080);
            pres.SaveAs(Path.Combine(output,"rounded-polygons-gallery.pptx"),PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,Office.MsoTriState.msoFalse);
        }
        finally
        {
            if (addin != null) addin.OnDisconnection(Extensibility.ext_DisconnectMode.ext_dm_HostShutdown, ref custom);
            if (pres != null) { pres.Saved = Office.MsoTriState.msoTrue; pres.Close(); Marshal.ReleaseComObject(pres); }
            if (app != null) { if(app.Presentations.Count==0) app.Quit(); Marshal.ReleaseComObject(app); }
            GC.Collect(); GC.WaitForPendingFinalizers();
        }
    }
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            Geometry();
            assertions += ConnectorRoutingTests.Run();
            if (args.Length > 0 && (args[0] == "--com" || args[0] == "--com-edges")) Com(args[0] == "--com-edges");
            Console.WriteLine("PASS " + assertions + " assertions"); return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex.GetType().Name + ": " + ex.Message + "\n" + ex.StackTrace); return 1; }
    }
}
