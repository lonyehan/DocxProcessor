using DocumentFormat.OpenXml.Packaging;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace DocxProcessor.Tools
{
    public class Converter
    {
        #region Converter: Byte[] To MemoryStream
        public static MemoryStream ByteArrayToMemoryStream(byte[] bytes)
        {
            MemoryStream origin = new MemoryStream(bytes, true);
            MemoryStream destination = new MemoryStream();

            origin.CopyToAsync(destination);
            origin.Close();

            return destination;
        }
        #endregion 

        #region Converter: Filepath To Byte[]
        public static byte[] FilePathToByteArray(string FilePath)
        {
            // Filepath to Byte Array
            MemoryStream stream = new MemoryStream();

            using (FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
            {
                fs.CopyTo(stream);
            }

            stream.Seek(0, SeekOrigin.Begin);
            return stream.ToArray();
        }
        #endregion        

        #region Get: Get Image Type
        /// <summary>
        /// 獲得Input Image的Type
        /// </summary>
        /// <param name="Bitmap"></param>
        /// <returns></returns>
        public static ImagePartType GetImageType(Bitmap bitmap)
        {
            if (bitmap.RawFormat.Equals(ImageFormat.Jpeg)) //It's a JPEG;
                return ImagePartType.Jpeg;
            else if (bitmap.RawFormat.Equals(ImageFormat.Png)) //It's a PNG;
                return ImagePartType.Png;
            else if (bitmap.RawFormat.Equals(ImageFormat.Bmp)) //It's a BMP;
                return ImagePartType.Bmp;
            else if (bitmap.RawFormat.Equals(ImageFormat.Gif)) //It's a Gif;
                return ImagePartType.Gif;

            throw new ApplicationException($"Unsupported image type！");
        }
        #endregion
    }
}
