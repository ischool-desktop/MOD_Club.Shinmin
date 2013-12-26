using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;

namespace K12.Club.Shinmin
{
    static class ResizeX1
    {
        public const int MAX_WIDTH = 640;
        public const int MAX_HEIGHT = 480;

        /// <summary>
        /// 調整圖片壓縮解析度
        /// </summary>
        public static string ToBase64String_x1(Bitmap newBmp, ImageFormat iFormat)
        {
            MemoryStream ms = new MemoryStream();

            //save
            ImageCodecInfo jgpEncoder = GetEncoder(ImageFormat.Jpeg);
            System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.Quality;
            EncoderParameters myEncoderParameters = new EncoderParameters(1);
            EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 30L);
            myEncoderParameters.Param[0] = myEncoderParameter;
            newBmp.Save(ms, jgpEncoder, myEncoderParameters);

            //Bitmap new2Bmp = new Bitmap(ms);
            //new2Bmp.Save(@"c:\test002.jpg", iFormat);

            //newBmp.Save(ms, iFormat);
            ms.Seek(0, SeekOrigin.Begin);

            //建立一個byte
            //並且定義長度為MemoryStream的長度
            byte[] bytes = new byte[ms.Length];
            //MemoryStream讀取資料
            ms.Read(bytes, 0, (int)ms.Length);
            ms.Close();

            return Convert.ToBase64String(bytes);
        }

        public static string ToBase64String(Bitmap newBmp, ImageFormat iFormat)
        {
            MemoryStream ms = new MemoryStream();

            //save
            newBmp.Save(ms, iFormat);

            //Bitmap new2Bmp = new Bitmap(ms);
            //new2Bmp.Save(@"c:\test001.jpg", iFormat);

            ms.Seek(0, SeekOrigin.Begin);

            //建立一個byte
            //並且定義長度為MemoryStream的長度
            byte[] bytes = new byte[ms.Length];
            //MemoryStream讀取資料
            ms.Read(bytes, 0, (int)ms.Length);
            ms.Close();

            return Convert.ToBase64String(bytes);
        }

        static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();

            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        public static Bitmap Resize(Bitmap photo)
        {
            Size s = GetResize(photo);
            return new Bitmap(photo, s);
        }

        public static Size GetResize(Bitmap photo)
        {
            return GetResize(photo, MAX_WIDTH, MAX_HEIGHT);
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
