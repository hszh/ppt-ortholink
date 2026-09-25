using System;
using System.Drawing;
using System.Drawing.Imaging;

namespace OrthoLink
{
    /// <summary>
    /// The real outline of a shape, sampled from a rendering of it. Lets connector ends sit on the
    /// slanted side of a trapezoid or the curve of an oval instead of on the bounding box.
    /// The rendering is a whole slide: the shape drawn black on white, its bounding box starting
    /// <c>margin</c> pt from the slide's top-left corner.
    /// </summary>
    internal sealed class Outline
    {
        private readonly bool[,] _ink;      // [x, y]
        private readonly int _w, _h;
        private readonly float _kx, _ky;    // pixels per pt
        private readonly float _margin;     // pt

        private Outline(bool[,] ink, int w, int h, float kx, float ky, float margin)
        {
            _ink = ink; _w = w; _h = h; _kx = kx; _ky = ky; _margin = margin;
        }

        public static Outline FromSlidePng(string path, float slideWidth, float slideHeight, float margin)
        {
            using (var src = new Bitmap(path))
            using (var bmp = new Bitmap(src.Width, src.Height, PixelFormat.Format32bppArgb))
            {
                // copy pixel-for-pixel; DrawImageUnscaled would rescale by the PNG's embedded DPI
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.White);
                    g.DrawImage(src, new Rectangle(0, 0, src.Width, src.Height), 0, 0, src.Width, src.Height, GraphicsUnit.Pixel);
                }
                int w = bmp.Width, h = bmp.Height;
                var ink = new bool[w, h];
                var data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                try
                {
                    unsafe
                    {
                        byte* row = (byte*)data.Scan0;
                        for (int y = 0; y < h; y++, row += data.Stride)
                            for (int x = 0; x < w; x++) ink[x, y] = row[x * 4] + row[x * 4 + 1] + row[x * 4 + 2] < 3 * 128;
                    }
                }
                finally { bmp.UnlockBits(data); }
                return new Outline(ink, w, h, w / slideWidth, h / slideHeight, margin);
            }
        }

        /// <summary>Point on the outline hit by a ray shot inward from the given side at the given position.
        /// r is where the shape is now. Falls back to the bounding-box point when nothing is hit.</summary>
        public PointF Hit(RectangleF r, Side side, float pos)
        {
            float x0 = r.Left - _margin, y0 = r.Top - _margin;   // where the image's top-left corner lands
            if (_w == 0 || _h == 0) return Geo.Port(r, side, pos);
            if (Geo.IsHorizontal(side))
            {
                float y = r.Top + r.Height * pos;
                int py = Clamp((int)((y - y0) * _ky), 0, _h - 1);
                if (side == Side.Left)
                {
                    for (int px = 0; px < _w; px++) if (_ink[px, py]) return new PointF(x0 + px / _kx, y);
                }
                else
                {
                    for (int px = _w - 1; px >= 0; px--) if (_ink[px, py]) return new PointF(x0 + (px + 1) / _kx, y);
                }
            }
            else
            {
                float x = r.Left + r.Width * pos;
                int px = Clamp((int)((x - x0) * _kx), 0, _w - 1);
                if (side == Side.Top)
                {
                    for (int py = 0; py < _h; py++) if (_ink[px, py]) return new PointF(x, y0 + py / _ky);
                }
                else
                {
                    for (int py = _h - 1; py >= 0; py--) if (_ink[px, py]) return new PointF(x, y0 + (py + 1) / _ky);
                }
            }
            return Geo.Port(r, side, pos);
        }

        private static int Clamp(int v, int lo, int hi) { return v < lo ? lo : v > hi ? hi : v; }
    }
}
