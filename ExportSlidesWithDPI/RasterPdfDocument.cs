using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace ExportSlidesWithDPIDoing
{
    /// <summary>
    /// Dependency-free, streaming PDF writer for cropped-slide export.
    /// Pages use lossless 24-bit RGB Flate encoding; only one compressed page exists at a time.
    /// </summary>
    internal sealed class RasterPdfDocument : IDisposable
    {
        private readonly FileStream stream;
        private readonly List<long> offsets = new List<long> { 0 };
        private readonly int pageCount;
        private int writtenPages;
        private bool completed;

        internal RasterPdfDocument(string outputPath, int pageCount)
        {
            if (pageCount <= 0) throw new ArgumentOutOfRangeException("pageCount");
            this.pageCount = pageCount;
            stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);

            WriteAscii("%PDF-1.4\n%\u00e2\u00e3\u00cf\u00d3\n");
            WriteObject(1, "<< /Type /Catalog /Pages 2 0 R >>");

            var kids = new StringBuilder();
            for (int i = 0; i < pageCount; i++)
                kids.AppendFormat(CultureInfo.InvariantCulture, "{0} 0 R ", 5 + i * 3);
            WriteObject(2, string.Format(CultureInfo.InvariantCulture,
                "<< /Type /Pages /Count {0} /Kids [ {1}] >>", pageCount, kids));
        }

        internal void AddPage(Image image, int dpi)
        {
            if (completed) throw new InvalidOperationException("PDF 已完成写入。");
            if (writtenPages >= pageCount) throw new InvalidOperationException("写入页数超过预期。");

            int pixelWidth = image.Width;
            int pixelHeight = image.Height;
            int safeDpi = Math.Max(1, dpi);
            string compressedPath = CreateCompressedRgbFile(image);
            try
            {
                int imageObject = 3 + writtenPages * 3;
                int contentObject = imageObject + 1;
                int pageObject = imageObject + 2;
                double widthPoints = pixelWidth * 72.0 / safeDpi;
                double heightPoints = pixelHeight * 72.0 / safeDpi;

                WriteFileStreamObject(imageObject, string.Format(CultureInfo.InvariantCulture,
                    "<< /Type /XObject /Subtype /Image /Width {0} /Height {1} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length {2} >>",
                    pixelWidth, pixelHeight, new FileInfo(compressedPath).Length), compressedPath);

                string content = string.Format(CultureInfo.InvariantCulture,
                    "q\n{0:0.###} 0 0 {1:0.###} 0 0 cm\n/Im0 Do\nQ\n", widthPoints, heightPoints);
                WriteByteStreamObject(contentObject,
                    string.Format(CultureInfo.InvariantCulture, "<< /Length {0} >>", Encoding.ASCII.GetByteCount(content)),
                    Encoding.ASCII.GetBytes(content));
                WriteObject(pageObject, string.Format(CultureInfo.InvariantCulture,
                    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {0:0.###} {1:0.###}] /Resources << /XObject << /Im0 {2} 0 R >> >> /Contents {3} 0 R >>",
                    widthPoints, heightPoints, imageObject, contentObject));
                writtenPages++;
            }
            finally
            {
                if (File.Exists(compressedPath)) File.Delete(compressedPath);
            }
        }

        internal void Complete()
        {
            if (completed) return;
            if (writtenPages != pageCount)
                throw new InvalidOperationException("PDF 页数不完整，未写入正式文件。");

            long xrefOffset = stream.Position;
            int objectCount = 3 + pageCount * 3;
            WriteAscii(string.Format(CultureInfo.InvariantCulture, "xref\n0 {0}\n", objectCount));
            WriteAscii("0000000000 65535 f \n");
            for (int i = 1; i < objectCount; i++)
                WriteAscii(string.Format(CultureInfo.InvariantCulture, "{0:0000000000} 00000 n \n", offsets[i]));
            WriteAscii(string.Format(CultureInfo.InvariantCulture,
                "trailer\n<< /Size {0} /Root 1 0 R >>\nstartxref\n{1}\n%%EOF\n", objectCount, xrefOffset));
            completed = true;
        }

        public void Dispose()
        {
            stream.Dispose();
        }

        private static string CreateCompressedRgbFile(Image image)
        {
            string path = Path.Combine(Path.GetTempPath(), "exportslides-" + Guid.NewGuid().ToString("N") + ".rgb.flate");
            using (var flattened = new Bitmap(image.Width, image.Height, PixelFormat.Format24bppRgb))
            using (var graphics = Graphics.FromImage(flattened))
            {
                graphics.Clear(Color.White);
                graphics.DrawImageUnscaled(image, 0, 0);

                var rectangle = new Rectangle(0, 0, flattened.Width, flattened.Height);
                var data = flattened.LockBits(rectangle, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                try
                {
                    int stride = Math.Abs(data.Stride);
                    var bgrRow = new byte[stride];
                    var rgbRow = new byte[flattened.Width * 3];
                    using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    {
                        // PDF FlateDecode requires a zlib wrapper. DeflateStream emits raw Deflate.
                        file.WriteByte(0x78);
                        file.WriteByte(0xDA);
                        uint adler32 = 1;
                        using (var compressor = new DeflateStream(file, CompressionLevel.Optimal, true))
                        {
                            for (int y = 0; y < flattened.Height; y++)
                            {
                                var pointer = new IntPtr(data.Scan0.ToInt64() + (long)y * data.Stride);
                                System.Runtime.InteropServices.Marshal.Copy(pointer, bgrRow, 0, stride);
                                for (int x = 0; x < flattened.Width; x++)
                                {
                                    int source = x * 3;
                                    rgbRow[source] = bgrRow[source + 2];
                                    rgbRow[source + 1] = bgrRow[source + 1];
                                    rgbRow[source + 2] = bgrRow[source];
                                }
                                adler32 = UpdateAdler32(adler32, rgbRow);
                                compressor.Write(rgbRow, 0, rgbRow.Length);
                            }
                        }
                        file.WriteByte((byte)(adler32 >> 24));
                        file.WriteByte((byte)(adler32 >> 16));
                        file.WriteByte((byte)(adler32 >> 8));
                        file.WriteByte((byte)adler32);
                    }
                }
                finally
                {
                    flattened.UnlockBits(data);
                }
            }
            return path;
        }

        private static uint UpdateAdler32(uint current, byte[] data)
        {
            const uint Modulus = 65521;
            uint a = current & 0xffff;
            uint b = current >> 16;
            int index = 0;
            while (index < data.Length)
            {
                // 5552 is the largest safe chunk before the running sums approach UInt32 overflow.
                int end = Math.Min(index + 5552, data.Length);
                while (index < end)
                {
                    a += data[index++];
                    b += a;
                }
                a %= Modulus;
                b %= Modulus;
            }
            return (b << 16) | a;
        }

        private void WriteObject(int number, string body)
        {
            offsets.Add(stream.Position);
            WriteAscii(number.ToString(CultureInfo.InvariantCulture) + " 0 obj\n" + body + "\nendobj\n");
        }

        private void WriteFileStreamObject(int number, string dictionary, string dataPath)
        {
            offsets.Add(stream.Position);
            WriteAscii(number.ToString(CultureInfo.InvariantCulture) + " 0 obj\n" + dictionary + "\nstream\n");
            using (var data = new FileStream(dataPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                data.CopyTo(stream, 64 * 1024);
            WriteAscii("\nendstream\nendobj\n");
        }

        private void WriteByteStreamObject(int number, string dictionary, byte[] data)
        {
            offsets.Add(stream.Position);
            WriteAscii(number.ToString(CultureInfo.InvariantCulture) + " 0 obj\n" + dictionary + "\nstream\n");
            stream.Write(data, 0, data.Length);
            WriteAscii("\nendstream\nendobj\n");
        }

        private void WriteAscii(string value)
        {
            byte[] bytes = Encoding.GetEncoding(1252).GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
        }
    }
}
