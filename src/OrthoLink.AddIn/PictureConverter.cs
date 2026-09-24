using System.Drawing;
using System.Windows.Forms;

namespace OrthoLink
{
    /// <summary>Turns a GDI+ image into the COM picture type the ribbon wants for custom icons.</summary>
    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter() : base("59EE46BA-677D-4d20-BF10-8D8067CB8B33") { }

        public static stdole.IPictureDisp ToPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }
}
