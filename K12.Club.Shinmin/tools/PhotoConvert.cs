using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.IO;

namespace K12.Club.Shinmin
{
    static class PhotoConvert
    {
        public static Bitmap ConvertFromBase64Encoding(string base64, int maxWidth, int maxHeight)
        {
            byte[] bs = Convert.FromBase64String(base64);
            MemoryStream ms = new MemoryStream(bs);
            Bitmap bm = new Bitmap(ms);
            return Resize(bm, maxWidth, maxHeight);
        }

        public static Bitmap Resize(Bitmap photo, int maxWidth, int maxHeight)
        {
            Size s = GetResize(photo, maxWidth, maxHeight);
            return new Bitmap(photo, s);
        }

        public static Size GetResize(Bitmap photo, int maxWidth, int maxHeight)
        {
            int width = photo.Width;
            int height = photo.Height;
            Size newSize;

            if (width < maxWidth && height < maxHeight)
                return new Size(width, height);

            decimal maxW = Convert.ToDecimal(maxWidth);
            decimal maxH = Convert.ToDecimal(maxHeight);

            decimal mp = decimal.Divide(maxW, maxH);
            decimal p = decimal.Divide(width, height);


            // 若長寬比預設比例較寬, 則以傳入之長為縮放基準
            if (mp > p)
            {
                decimal hp = decimal.Divide(maxH, height);
                decimal newWidth = decimal.Multiply(hp, width);
                newSize = new Size(decimal.ToInt32(newWidth), maxHeight);
            }
            else
            {
                decimal wp = decimal.Divide(maxW, width);
                decimal newHeight = decimal.Multiply(wp, height);
                newSize = new Size(maxWidth, decimal.ToInt32(newHeight));
            }

            return newSize;
        }
    }
}
