using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace PageOrientationEngine.Helpers
{
    public static class BitmapUtils
    {
        #region Internal static class NativeMethods
        internal static class NativeMethods
        {
            #region Struct Bitmapinfo
            [StructLayout(LayoutKind.Sequential)]
            public struct Bitmapinfo
            {
                public uint biSize;
                public int biWidth, biHeight;
                public short biPlanes, biBitCount;
                public uint biCompression, biSizeImage;
                public int biXPelsPerMeter, biYPelsPerMeter;
                public uint biClrUsed, biClrImportant;
                [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
                public uint[] cols;
            }
            #endregion

            [DllImport("gdi32.dll")]
            public static extern bool DeleteObject(IntPtr hObject);

            [DllImport("user32.dll")]
            public static extern int InvalidateRect(IntPtr hwnd, IntPtr rect, int bErase);

            [DllImport("user32.dll")]
            public static extern IntPtr GetDC(IntPtr hwnd);

            [DllImport("gdi32.dll")]
            public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

            [DllImport("user32.dll")]
            public static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

            [DllImport("gdi32.dll")]
            public static extern int DeleteDC(IntPtr hdc);

            [DllImport("gdi32.dll")]
            public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

            [DllImport("gdi32.dll")]
            public static extern int BitBlt(IntPtr hdcDst, int xDst, int yDst, int w, int h, IntPtr hdcSrc, int xSrc,
                                            int ySrc, int rop);

            [DllImport("gdi32.dll")]
            public static extern IntPtr CreateDIBSection(IntPtr hdc, ref Bitmapinfo bmi, uint usage, out IntPtr bits,
                                                          IntPtr hSection, uint dwOffset);
        }
        #endregion

        #region Fields
        private const int Srccopy = 0x00CC0020;
        private const uint BiRgb = 0;
        private const uint DibRgbColors = 0;
        #endregion

        #region MakeRGB
        private static uint MakeRGB(int r, int g, int b)
        {
            return ((uint)(b & 255)) | ((uint)((r & 255) << 8)) | ((uint)((g & 255) << 16));
        }
        #endregion

        #region CopyToBpp
        /// <summary>
        ///     Copies a bitmap into a 1bpp/8bpp bitmap of the same dimensions, fast
        /// </summary>
        /// <param name="bitmap">The original bitmap</param>
        /// <param name="bpp">The 1 or 8, target bpp</param>
        /// <returns>A 1bpp copy of the bitmap</returns>
        public static Bitmap CopyToBpp(Bitmap bitmap, int bpp)
        {
            if (bpp != 1 && bpp != 8)
                throw new ArgumentException("1 or 8", "bpp");

            // Plan: built into Windows GDI is the ability to convert
            // bitmaps from one format to another. Most of the time, this
            // job is actually done by the graphics hardware accelerator card
            // and so is extremely fast. The rest of the time, the job is done by
            // very fast native code.
            // We will call into this GDI functionality from C#. Our plan:
            // (1) Convert our Bitmap into a GDI hbitmap (ie. copy unmanaged->managed)
            // (2) Create a GDI monochrome hbitmap
            // (3) Use GDI "BitBlt" function to copy from hbitmap into monochrome (as above)
            // (4) Convert the monochrone hbitmap into a Bitmap (ie. copy unmanaged->managed)

            var width = bitmap.Width;
            var height = bitmap.Height;

            var bitmapHandle = bitmap.GetHbitmap(); // this is step (1)

            // Step (2): create the monochrome bitmap.
            // "BITMAPINFO" is an interop-struct which we define below.
            // In GDI terms, it's a BITMAPHEADERINFO followed by an array of two RGBQUADs
            var bitmapInfo = new NativeMethods.Bitmapinfo
            {
                biSize = 40,
                biWidth = width,
                biHeight = height,
                biPlanes = 1,
                biBitCount = (short)bpp,
                biCompression = BiRgb,
                biSizeImage = (uint)(((width + 7) & 0xFFFFFFF8) * height / 8),
                biXPelsPerMeter = 1000000,
                biYPelsPerMeter = 1000000
            };

            // Now for the colour table.
            var ncols = (uint)1 << bpp; // 2 colours for 1bpp; 256 colours for 8bpp
            bitmapInfo.biClrUsed = ncols;
            bitmapInfo.biClrImportant = ncols;
            bitmapInfo.cols = new uint[256]; // The structure always has fixed size 256, even if we end up using fewer colours
            if (bpp == 1)
            {
                bitmapInfo.cols[0] = MakeRGB(0, 0, 0);
                bitmapInfo.cols[1] = MakeRGB(255, 255, 255);
            }
            else
            {
                for (var i = 0; i < ncols; i++) bitmapInfo.cols[i] = MakeRGB(i, i, i);
            }

            // For 8bpp we've created an palette with just greyscale colours.
            // You can set up any palette you want here. Here are some possibilities:
            // greyscale: for (int i=0; i<256; i++) bmi.cols[i]=MAKERGB(i,i,i);
            // rainbow: bmi.biClrUsed=216; bmi.biClrImportant=216; int[] colv=new int[6]{0,51,102,153,204,255};
            //          for (int i=0; i<216; i++) bmi.cols[i]=MAKERGB(colv[i/36],colv[(i/6)%6],colv[i%6]);
            // optimal: a difficult topic: http://en.wikipedia.org/wiki/Color_quantization
            // 
            // Now create the indexed bitmap "hbm0"
            IntPtr bits0; // not used for our purposes. It returns a pointer to the raw bits that make up the bitmap.
            var hbm0 = NativeMethods.CreateDIBSection(IntPtr.Zero, ref bitmapInfo, DibRgbColors, out bits0, IntPtr.Zero, 0);

            //
            // Step (3): use GDI's BitBlt function to copy from original hbitmap into monocrhome bitmap
            // GDI programming is kind of confusing... nb. The GDI equivalent of "Graphics" is called a "DC".
            var sdc = NativeMethods.GetDC(IntPtr.Zero); // First we obtain the DC for the screen

            // Next, create a DC for the original hbitmap
            var hdc = NativeMethods.CreateCompatibleDC(sdc);
            NativeMethods.SelectObject(hdc, bitmapHandle);

            // and create a DC for the monochrome hbitmap
            var hdc0 = NativeMethods.CreateCompatibleDC(sdc);
            NativeMethods.SelectObject(hdc0, hbm0);

            // Now we can do the BitBlt:
            NativeMethods.BitBlt(hdc0, 0, 0, width, height, hdc, 0, 0, Srccopy);

            // Step (4): convert this monochrome hbitmap back into a Bitmap:
            var b0 = Image.FromHbitmap(hbm0);

            // Finally some cleanup.
            NativeMethods.DeleteDC(hdc);
            NativeMethods.DeleteDC(hdc0);
            NativeMethods.ReleaseDC(IntPtr.Zero, sdc);
            NativeMethods.DeleteObject(bitmapHandle);
            NativeMethods.DeleteObject(hbm0);
            
            return b0;
        }
        #endregion
    }
}
