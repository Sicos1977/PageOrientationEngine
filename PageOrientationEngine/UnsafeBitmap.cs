using System;
using System.Drawing;
using System.Drawing.Imaging;

namespace DocumentServices.Server.Utilities
{
    /// <summary>
    /// Via deze classe kan een bitmap unsafe (dus d.m.v. pointers) worden bewerkt.
    /// Dit versnelt het werken met bitmaps heel erg veel
    /// 
    /// UnsafeBitmap your_fast_bitmap = new UnsafeBitmap(old_safe_bitmap);
    /// your_fast_bitmap.LockBitmap();
    /// PixelData pixel = your_fast_bitmap.getPixel(3, 4);
    //// your_fast_bitmap.UnlockBitmap();
    /// </summary>
    public unsafe class UnsafeBitmap : IDisposable
    {
        #region Struct PixelData
        public struct PixelData
        {
            public byte Blue;
            public byte Green;
            public byte Red;
        }
        #endregion

        #region Fields
        private readonly Bitmap _bitmap;
        private BitmapData _bitmapData;
        private Byte* _pBase = null;
        private int _width;
        private bool _bitmapLocked;
        #endregion

        #region UnsafeBitmap

        public UnsafeBitmap(Bitmap bitmap)
        {
            _bitmap = new Bitmap(bitmap);
        }

        public UnsafeBitmap(int width, int height)
        {
            _bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
        }

        #endregion

        #region Dispose
        public void Dispose()
        {
            if (_bitmap == null) return;
            UnlockBitmap();
            _bitmap.Dispose();
        }
        #endregion

        #region Bitmap
        public Bitmap Bitmap
        {
            get { return (_bitmap); }
        }
        #endregion

        //#region PixelSize
        //private Point PixelSize
        //{
        //    get
        //    {
        //        var unit = GraphicsUnit.Pixel;
        //        var bounds = _bitmap.GetBounds(ref unit);

        //        return new Point((int) bounds.Width, (int) bounds.Height);
        //    }
        //}
        //#endregion

        #region LockBitmap
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
        public PixelData GetPixel(int x, int y)
        {
            var returnValue = *PixelAt(x, y);
            return returnValue;
        }
        #endregion

        #region SetPixel
        public void SetPixel(int x, int y, PixelData colour)
        {
            var pixel = PixelAt(x, y);
            *pixel = colour;
        }
        #endregion

        #region UnlockBitmap
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
        public PixelData* PixelAt(int x, int y)
        {
            return (PixelData*) (_pBase + y*_width + x*sizeof (PixelData));
        }
        #endregion

        #region Width
        public int Width
        {
            get { return _bitmap.Width; }
        }
        #endregion

        #region Height
        public int Height
        {
            get { return _bitmap.Height; }
        }
        #endregion
    }
}