using System;
using System.Drawing;
using System.Drawing.Imaging;

namespace PageOrientationEngine.Helpers
{
    /// <summary>
    /// This class allows to work with bitmap in an unsafe way
    /// <code>
    /// UnsafeBitmap your_fast_bitmap = new UnsafeBitmap(old_safe_bitmap);
    /// your_fast_bitmap.LockBitmap();
    /// PixelData pixel = your_fast_bitmap.getPixel(3, 4);
    /// your_fast_bitmap.UnlockBitmap();
    /// </code>
    /// </summary>
    public unsafe class UnsafeBitmap : IDisposable
    {
        #region Struct PixelData
        /// <summary>
        /// The RGB values from a pixel
        /// </summary>
        public struct PixelData
        {
            /// <summary>
            /// The pixel red value
            /// </summary>
            public byte Red;

            /// <summary>
            /// The pixel green value
            /// </summary>
            public byte Green;
            
            /// <summary>
            /// The pixel blue value
            /// </summary>
            public byte Blue;
        }
        #endregion

        #region Fields
        private readonly Bitmap _bitmap;
        private BitmapData _bitmapData;
        private Byte* _pBase = null;
        private int _width;
        private bool _bitmapLocked;
        #endregion

        #region Properties
        public Bitmap Bitmap
        {
            get { return (_bitmap); }
        }

        /// <summary>
        /// Returns the width of the <see cref="_bitmap"/>
        /// </summary>
        public int Width
        {
            get { return _bitmap.Width; }
        }

        /// <summary>
        /// Returns the height of the <see cref="_bitmap"/>
        /// </summary>
        public int Height
        {
            get { return _bitmap.Height; }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Creates this object for the given <paramref name="bitmap"/>
        /// </summary>
        /// <param name="bitmap"></param>
        public UnsafeBitmap(Image bitmap)
        {
            _bitmap = new Bitmap(bitmap);
        }

        /// <summary>
        /// Creates this object with a new empty <see cref="_bitmap"/>
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public UnsafeBitmap(int width, int height)
        {
            _bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
        }

        #endregion

        #region LockBitmap
        /// <summary>
        /// Locks the bitmap in memory so that it can be edited
        /// </summary>
        public void LockBitmap()
        {
            var unit = GraphicsUnit.Pixel;
            var boundsF = _bitmap.GetBounds(ref unit);
            var bounds = new Rectangle((int) boundsF.X,
                                       (int) boundsF.Y,
                                       (int) boundsF.Width,
                                       (int) boundsF.Height);

            // Figure out the number of bytes in a row
            // This is rounded up to be a multiple of 4
            // bytes, since a scan line in an image must always be a multiple of 4 bytes
            // in length. 
            _width = (int) boundsF.Width*sizeof (PixelData);
            if (_width%4 != 0)
                _width = 4*(_width/4 + 1);

            _bitmapData = _bitmap.LockBits(bounds, ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            _pBase = (Byte*) _bitmapData.Scan0.ToPointer();
            _bitmapLocked = true;
        }
        #endregion

        #region GetPixel
        /// <summary>
        /// Returns the pixel on the <paramref name="x"/> and <paramref name="y"/> position
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public PixelData GetPixel(int x, int y)
        {
            var returnValue = *PixelAt(x, y);
            return returnValue;
        }
        #endregion

        #region SetPixel
        /// <summary>
        /// Sets the pixel on the <paramref name="x"/> and <paramref name="y"/> position
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="colour"></param>
        public void SetPixel(int x, int y, PixelData colour)
        {
            var pixel = PixelAt(x, y);
            *pixel = colour;
        }
        #endregion

        #region UnlockBitmap
        /// <summary>
        /// Unlocks the <see cref="_bitmap"/>
        /// </summary>
        public void UnlockBitmap()
        {
            if (!_bitmapLocked) return;
            _bitmap.UnlockBits(_bitmapData);
            _bitmapData = null;
            _pBase = null;
            _bitmapLocked = false;
        }
        #endregion

        #region PixelAt
        /// <summary>
        /// Returns the <see cref="PixelData"/> for the pixel at the <paramref name="x"/> and <paramref name="y"/> position
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        public PixelData* PixelAt(int x, int y)
        {
            return (PixelData*) (_pBase + y*_width + x*sizeof (PixelData));
        }
        #endregion

        #region Dispose
        /// <summary>
        /// Disposes this object
        /// </summary>
        public void Dispose()
        {
            if (_bitmap == null) return;
            UnlockBitmap();
            _bitmap.Dispose();
        }
        #endregion
    }
}