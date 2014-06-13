using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using PageOrientationEngine.Helpers;
using Tesseract;

namespace PageOrientationEngine
{
    #region DocumentInspectorPageOrientation
    /// <summary>
    /// The orientation of the text on a page
    /// </summary>
    public enum DocumentInspectorPageOrientation
    {
        /// <summary>
        /// The text on page is correctly orientated
        /// </summary>
        PageCorrect,

        /// <summary>
        /// The text on the page is upside down
        /// </summary>
        PageUpsideDown,

        /// <summary>
        /// The text on the page is rotated to the left
        /// </summary>
        PageRotatedLeft,

        /// <summary>
        /// The text on the page is rotated to the right
        /// </summary>
        PageRotatedRight,

        /// <summary>
        /// The quality of text on the page is to bad to determine the orientation of the page
        /// </summary>
        Undetectable
    }
    #endregion

    public class DocumentInspector
    {
        #region Internal class WorkQueueItem
        /// <summary>
        /// Used to track all the workqueue items
        /// </summary>
        private class WorkQueueItem
        {
            /// <summary>
            /// The number of the page
            /// </summary>
            public int PageNumber { get; private set; }

            /// <summary>
            /// The page as a memory stream
            /// </summary>
            public MemoryStream MemoryStream { get; private set; }

            /// <summary>
            /// Creates this object
            /// </summary>
            /// <param name="pageNumber">The number of the page</param>
            /// <param name="memoryStream">The page as a memory stream</param>
            public WorkQueueItem(int pageNumber, MemoryStream memoryStream)
            {
                PageNumber = pageNumber;
                MemoryStream = memoryStream;
            }
        }
        #endregion

        #region Fields
        /// <summary>
        /// The load queue
        /// </summary>
        ConcurrentQueue<WorkQueueItem> _workQueue;

        /// <summary>
        /// Contains all the result of the page orientation detection
        /// </summary>
        private ConcurrentDictionary<int, DocumentInspectorPageOrientation> _detectionResult;
        #endregion

        #region Properties
        /// <summary>
        /// The path to the Tesseract data files
        /// </summary>
        public string TesseractDataPath { get; private set; }

        /// <summary>
        /// The language that needs to be used by Tesseract
        /// </summary>
        public string TesseractLanguage { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        /// Creates this object with the default <see cref="TesseractLanguage"/> language set to English.
        /// Use the <see cref="TesseractLanguage"/>
        /// </summary>
        public DocumentInspector(string dataPath, string language)
        {
            TesseractDataPath = dataPath;
            TesseractLanguage = language;
        }
        #endregion

        #region CopyToBpp
        private const int Srccopy = 0x00CC0020;
        private const uint BiRgb = 0;
        private const uint DibRgbColors = 0;

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
        private static extern IntPtr CreateDIBSection(IntPtr hdc, ref Bitmapinfo bmi, uint usage, out IntPtr bits,
                                                      IntPtr hSection, uint dwOffset);

        private static uint Makergb(int r, int g, int b)
        {
            return ((uint)(b & 255)) | ((uint)((r & 255) << 8)) | ((uint)((g & 255) << 16));
        }

        /// <summary>
        ///     Copies a bitmap into a 1bpp/8bpp bitmap of the same dimensions, fast
        /// </summary>
        /// <param name="b">original bitmap</param>
        /// <param name="bpp">1 or 8, target bpp</param>
        /// <returns>a 1bpp copy of the bitmap</returns>
        private static Bitmap CopyToBpp(Bitmap b, int bpp)
        {
            if (bpp != 1 && bpp != 8)
                // ReSharper disable LocalizableElement
                throw new ArgumentException("1 or 8", "bpp");
                // ReSharper restore LocalizableElement

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

            int w = b.Width, h = b.Height;
            var hbm = b.GetHbitmap(); // this is step (1)
            //
            // Step (2): create the monochrome bitmap.
            // "BITMAPINFO" is an interop-struct which we define below.
            // In GDI terms, it's a BITMAPHEADERINFO followed by an array of two RGBQUADs
            var bmi = new Bitmapinfo
            {
                biSize = 40,
                biWidth = w,
                biHeight = h,
                biPlanes = 1,
                biBitCount = (short)bpp,
                biCompression = BiRgb,
                biSizeImage = (uint)(((w + 7) & 0xFFFFFFF8) * h / 8),
                biXPelsPerMeter = 1000000,
                biYPelsPerMeter = 1000000
            };

            // Now for the colour table.
            var ncols = (uint)1 << bpp; // 2 colours for 1bpp; 256 colours for 8bpp
            bmi.biClrUsed = ncols;
            bmi.biClrImportant = ncols;
            bmi.cols = new uint[256]; // The structure always has fixed size 256, even if we end up using fewer colours
            if (bpp == 1)
            {
                bmi.cols[0] = Makergb(0, 0, 0);
                bmi.cols[1] = Makergb(255, 255, 255);
            }
            else
            {
                for (var i = 0; i < ncols; i++) bmi.cols[i] = Makergb(i, i, i);
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
            var hbm0 = CreateDIBSection(IntPtr.Zero, ref bmi, DibRgbColors, out bits0, IntPtr.Zero, 0);

            //
            // Step (3): use GDI's BitBlt function to copy from original hbitmap into monocrhome bitmap
            // GDI programming is kind of confusing... nb. The GDI equivalent of "Graphics" is called a "DC".
            var sdc = GetDC(IntPtr.Zero); // First we obtain the DC for the screen

            // Next, create a DC for the original hbitmap
            var hdc = CreateCompatibleDC(sdc);
            SelectObject(hdc, hbm);

            // and create a DC for the monochrome hbitmap
            var hdc0 = CreateCompatibleDC(sdc);
            SelectObject(hdc0, hbm0);

            // Now we can do the BitBlt:
            BitBlt(hdc0, 0, 0, w, h, hdc, 0, 0, Srccopy);

            // Step (4): convert this monochrome hbitmap back into a Bitmap:
            var b0 = Image.FromHbitmap(hbm0);

            //
            // Finally some cleanup.
            DeleteDC(hdc);
            DeleteDC(hdc0);
            ReleaseDC(IntPtr.Zero, sdc);
            DeleteObject(hbm);
            DeleteObject(hbm0);
            //
            return b0;
        }
        #endregion

        #region DetectPageOrientation
        /// <summary>
        /// Detecteert de text orientatie op opgegeven afbeeldingen. Deze functie werkt alleen correct
        /// met afbeeldingen waarop over het algemeen alleen maar text staat. Het is niet mogelijk
        /// om met deze functie te detecteren of een afbeelding correct staat
        /// </summary>
        /// <param name="memoryStreams">Lijst van memorystreams die TIFF bitmappen bevatten</param>
        /// <returns>Lijst met pagina orientaties</returns>
        public List<DocumentInspectorPageOrientation> DetectPageOrientation(List<MemoryStream> memoryStreams)
        {
            _detectionResult = new ConcurrentDictionary<int, DocumentInspectorPageOrientation>();
            _workQueue = new ConcurrentQueue<WorkQueueItem>();

            // WorkQueue vullen
            var i = 1;
            foreach (var memoryStream in memoryStreams)
            {
                _workQueue.Enqueue(new WorkQueueItem(i, memoryStream));
                i++;
            }

            //var procCount = Environment.ProcessorCount;
            var procCount = 4;

            if (_workQueue.Count < procCount)
                procCount = _workQueue.Count;

            var threads = new Thread[procCount];

            // Threads spawnen
            for (i = 0; i < procCount; i++)
            {
                ThreadStart threadStart = ProcessWorkQueue;
                threads[i] = new Thread(threadStart);
                threads[i].Start();
            }

            var workDone = false;

            while (!workDone)
            {
                for (i = 0; i < procCount; i++)
                    workDone = threads[i].Join(10);
            }

            return _detectionResult.Select(detectionResult => detectionResult.Value).ToList();
        }

        /// <summary>
        /// Verwerkt inhoud uit de _workQueue en zet het resultaat terug in de _detectionResult 
        /// dictionary
        /// </summary>
        public void ProcessWorkQueue()
        {
            WorkQueueItem workQueueItem;
            while (_workQueue.TryDequeue(out workQueueItem))
            {
                using (var bitmap = Image.FromStream(workQueueItem.MemoryStream) as Bitmap)
                {
                    var succes = false;
                    while (!succes)
                    {
                        succes = _detectionResult.TryAdd(workQueueItem.PageNumber,
                                                         DetectPageOrientation(bitmap));
                        Thread.Sleep(10);
                    }
                }
            }
        }

        /// <summary>
        /// Detecteert de text orientatie op opgegeven afbeeldingen. Deze functie werkt alleen correct
        /// met afbeeldingen waarop over het algemeen alleen maar text staat. Het is niet mogelijk
        /// om met deze functie te detecteren of een afbeelding correct staat
        /// </summary>
        /// <param name="inputFile">Het pad naar het TIFF bestand</param>
        /// <returns>Lijst met pagina orientaties</returns>
        public List<DocumentInspectorPageOrientation> DetectPageOrientation(string inputFile)
        {
            return DetectPageOrientation(TiffUtils.SplitTiffImage(inputFile));
        }

        /// <summary>
        /// Detecteert de text orientatie op de afbeelding. Deze functie werkt alleen correct
        /// met afbeeldingen waarop over het algemeen alleen maar text staat. Het is niet mogelijk
        /// om met deze functie te detecteren of een afbeelding correct staat
        /// </summary>
        /// <param name="bitmap">Een tiff afbeelding als bitmap object</param>
        /// <returns>De pagina orientatie</returns>
        public DocumentInspectorPageOrientation DetectPageOrientation(Bitmap bitmap)
        {
            if (bitmap == null)
                throw new NullReferenceException("The bitmap parameter is not set");

            if (bitmap.PixelFormat == PixelFormat.Format1bppIndexed)
                bitmap = CopyToBpp(bitmap, 8);

            using (var engine = new TesseractEngine(TesseractDataPath, TesseractLanguage))
            {
                // Kan gebruikt worden om ervoor te zorgen dan alleen onderstaande karakters gedetecteerd worden
                //engine.;SetVariable("tessedit_char_whitelist", "AEO");

                var rect = new Rect();

                using (var image = PixConverter.ToPix(bitmap))
                {
                    using (var page = engine.Process(image, PageSegMode.AutoOsd))
                    {
                        var pageIterator = page.AnalyseLayout();
                        pageIterator.Begin();

                        while (pageIterator.Next(PageIteratorLevel.Block))
                        {
                            var found = false;

                            while (pageIterator.Next(PageIteratorLevel.Para))
                            {
                                var counter = 0;

                                while (pageIterator.Next(PageIteratorLevel.TextLine, PageIteratorLevel.Word))
                                    counter++;

                                if (counter < 5) continue;
                                found = pageIterator.TryGetBoundingBox(PageIteratorLevel.TextLine, out rect);
                                break;
                            }

                            var croppedRect = new Rectangle(rect.X1, rect.Y1, rect.Width, rect.Height);
                            if (rect.Height == 0)
                                return DocumentInspectorPageOrientation.Undetectable;

                            //if (rect.Height > 50)
                            //    continue;

                            var croppedImage = found ? bitmap.Clone(croppedRect, bitmap.PixelFormat) : bitmap.Clone() as Bitmap;

                            //croppedImage.Save(@"d:\\Crop area normal.tif", System.Drawing.Imaging.ImageFormat.Tiff);

                            // De OCR zekerheidsgraad bij de eerste OCR run
                            float firstMeanConfedence;

                            // De OCR zekerheidsgraad bij de tweede OCR run
                            float secondMeanConfedence;

                            using (var engineCroppedImage = new TesseractEngine(TesseractDataPath, TesseractLanguage))
                            {
                                using (var imageNormal = PixConverter.ToPix(croppedImage))
                                {
                                    using (var pageNormal = engineCroppedImage.Process(imageNormal))
                                    {
                                        firstMeanConfedence = pageNormal.GetMeanConfidence();
                                    }
                                }

                                if (firstMeanConfedence > 0.75)
                                    return DocumentInspectorPageOrientation.PageCorrect;

                                // Dezelfde afbeelding maar nu 180 graden geroteerd
                                //croppedImage.RotateFlip(RotateFlipType.Rotate180FlipNone);
                                //croppedImage.Save(@"d:\\Crop area flipped.tif", System.Drawing.Imaging.ImageFormat.Tiff);

                                using (var imageRotated180 = PixConverter.ToPix(croppedImage))
                                {
                                    using (var pageRotated180 = engineCroppedImage.Process(imageRotated180))
                                    {
                                        secondMeanConfedence = pageRotated180.GetMeanConfidence();
                                    }
                                }
                            }

                            if (croppedImage != null) croppedImage.Dispose();

                            if (firstMeanConfedence > 0.40 && secondMeanConfedence > 0.40)
                                return firstMeanConfedence >= secondMeanConfedence
                                            ? DocumentInspectorPageOrientation.PageCorrect
                                            : DocumentInspectorPageOrientation.PageUpsideDown;

                        }
                    }
                    return DocumentInspectorPageOrientation.Undetectable;
                }
            }
        }
        #endregion
    }
}
