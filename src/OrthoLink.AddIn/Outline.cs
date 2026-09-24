using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace OrthoLink
{
    /// <summary>
    /// The real outline of a shape, sampled from a rendering of it. Lets connector ends sit on the
    /// slanted side of a trapezoid or the curve of an oval instead of on the bounding box.
    /// Mask coordinates are normalised to the bounding box expanded by half the line width.
    /// </summary>
    internal sealed class Outline
    {
        private readonly bool[,] _opaque;   // [x, y]
        private readonly int _w, _h;
        private readonly float _pad;        // pt, on every side of the bbox

        private Outline(bool[,] opaque, int w, int h, float pad) { _opaque = opaque; _w = w; _h = h; _pad = pad; }

        public static Outline FromPng(string path, float padPt)
        {
            using (var src = new Bitmap(path))
            using (var bmp = new Bitmap(src.Width, src.Height, PixelFormat.Format32bppArgb))
            {
                // copy pixel-for-pixel; DrawImageUnscaled would rescale by the PNG's embedded DPI
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.Transparent);
                    g.DrawImage(src, new Rectangle(0, 0, src.Width, src.Height), 0, 0, src.Width, src.Height, GraphicsUnit.Pixel);
                }
                int w = bmp.Width, h = bmp.Height;
                var mask = new bool[w, h];
                var data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                try
                {
                    unsafe
                    {
                        byte* row = (byte*)data.Scan0;
                        for (int y = 0; y < h; y++, row += data.Stride)
                            for (int x = 0; x < w; x++) mask[x, y] = row[x * 4 + 3] > 40;
                    }
                }
                finally { bmp.UnlockBits(data); }
                return new Outline(mask, w, h, padPt);
            }
        }

        /// <summary>Point on the outline hit by a ray shot inward from the given side at the given position.
        /// Falls back to the bounding-box point when nothing is hit.</summary>
        public PointF Hit(RectangleF r, Side side, float pos)
        {
            float x0 = r.Left - _pad, y0 = r.Top - _pad, W = r.Width + 2 * _pad, H = r.Height + 2 * _pad;
            if (W <= 0 || H <= 0 || _w == 0 || _h == 0) return Geo.Port(r, side, pos);
            if (Geo.IsHorizontal(side))
            {
                float y = r.Top + r.Height * pos;
                int py = Clamp((int)((y - y0) / H * _h), 0, _h - 1);
                if (side == Side.Left)
                {
                    for (int px = 0; px < _w; px++) if (_opaque[px, py]) return new PointF(x0 + px / (float)_w * W, y);
                }
                else
                {
                    for (int px = _w - 1; px >= 0; px--) if (_opaque[px, py]) return new PointF(x0 + (px + 1) / (float)_w * W, y);
                }
            }
            else
            {
                float x = r.Left + r.Width * pos;
                int px = Clamp((int)((x - x0) / W * _w), 0, _w - 1);
                if (side == Side.Top)
                {
                    for (int py = 0; py < _h; py++) if (_opaque[px, py]) return new PointF(x, y0 + py / (float)_h * H);
                }
                else
                {
                    for (int py = _h - 1; py >= 0; py--) if (_opaque[px, py]) return new PointF(x, y0 + (py + 1) / (float)_h * H);
                }
            }
            return Geo.Port(r, side, pos);
        }

        private static int Clamp(int v, int lo, int hi) { return v < lo ? lo : v > hi ? hi : v; }
    }
}
