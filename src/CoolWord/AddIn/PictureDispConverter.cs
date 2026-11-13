using System;
using System.Drawing;
using System.Windows.Forms;

namespace CoolWord.AddIn
{
    /// <summary>
    /// Chuyển đổi System.Drawing.Image/Bitmap thành OLE IPictureDisp (stdole.IPictureDisp)
    /// chuẩn cho Office Fluent Ribbon getImage callbacks.
    /// Kế thừa từ System.Windows.Forms.AxHost để sử dụng phương thức được bảo vệ GetIPictureDispFromPicture.
    /// </summary>
    public sealed class PictureDispConverter : AxHost
    {
        private PictureDispConverter() : base(string.Empty)
        {
        }

        public static stdole.IPictureDisp ToIPictureDisp(Image image)
        {
            if (image == null) return null;
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }
}
