using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace PageOrientationEngine.Helpers
{
    /// <summary>
    /// Deze classe bevat code om verschillende handelingen uit te kunnen voeren met Tiff bestanden.
    /// </summary>
    public class TiffUtils
    {
        #region GetPageCount
        /// <summary>
        /// Geeft het aantal pagina's dat het Tiff bestand bevat
        /// <param name="inputFile">Het te openen Tiff bestand</param>
        /// </summary>
        public int GetPageCount(string inputFile)
        {
            var image = Image.FromFile(inputFile);
            return GetPageCount(image, true);
        }

        private static int GetPageCount(Image image, bool dispose)
        {
            var guid = image.FrameDimensionsList[0];
            var dimension = new FrameDimension(guid);

            var pages = image.GetFrameCount(dimension);

            if (dispose)
                image.Dispose();

            return pages;
        }
        #endregion

        #region SplitTiffImage
        /// <summary>
        /// Splits het geopende tiff bestand in losse bestanden
        /// </summary>
        /// <param name="inputFile">Het te splitsen Tiff bestand</param>
        /// <param name="outputFolder">De output folder (opgegeven zonder slash "\" aan het einde</param>
        /// <returns>List met de output bestanden</returns>
        public static List<string> SplitTiffImage(string inputFile, string outputFolder)
        {
            var image = Image.FromFile(inputFile);

            var outputFile = outputFolder + Path.DirectorySeparatorChar + Path.GetFileNameWithoutExtension(inputFile) +
                                "_";
            var output = new List<string>();

            var guid = image.FrameDimensionsList[0];
            var fdim = new FrameDimension(guid);

            var pageCount = GetPageCount(image, false);

            var ici = GetEncoderInfo("image/tiff");
            var enc = Encoder.Compression;
            var eps = new EncoderParameters(1);

            // Save the bitmap as a TIFF file with LZW compression.
            var ep = new EncoderParameter(enc, (long)EncoderValue.CompressionCCITT4);
            eps.Param[0] = ep;

            for (var i = 0; i < pageCount; i++)
            {
                image.SelectActiveFrame(fdim, i);
                //Save the master bitmap
                var fileName = string.Format("{0}_pagina{1}.tif", outputFile, i + 1);
                image.Save(fileName, ici, eps);
                output.Add(fileName);
            }

            image.Dispose();

            return output;
        }

        /// <summary>
        /// Splits het geopende tiff bestand in losse bestanden en geeft deze als een lijst van streams terug
        /// </summary>
        /// <param name="inputFile">Het te splitsen Tiff bestand</param>
        /// <returns>List met de output streams</returns>
        public static List<MemoryStream> SplitTiffImage(string inputFile)
        {
            var image = Image.FromFile(inputFile);
            var output = new List<MemoryStream>();

            var guid = image.FrameDimensionsList[0];
            var fdim = new FrameDimension(guid);

            var pageCount = GetPageCount(image, false);

            var ici = GetEncoderInfo("image/tiff");
            var enc = Encoder.Compression;
            var eps = new EncoderParameters(1);

            // Save the bitmap as a TIFF file with LZW compression.
            var ep = new EncoderParameter(enc, (long)EncoderValue.CompressionCCITT4);
            eps.Param[0] = ep;

            for (var i = 0; i < pageCount; i++)
            {
                image.SelectActiveFrame(fdim, i);
                var str = new MemoryStream();
                image.Save(str, ici, eps);
                output.Add(str);
            }

            image.Dispose();
            return output;
        }

        /// <summary>
        /// Splits het geopende tiff bestand in losse bestanden en geeft deze als een lijst van streams terug
        /// </summary>
        /// <param name="inputFile">Het te splitsen Tiff bestand als memorystream</param>
        /// <returns>List met de output streams</returns>
        public static List<MemoryStream> SplitTiffImage(MemoryStream inputFile)
        {
            var image = Image.FromStream(inputFile);
            var output = new List<MemoryStream>();

            var guid = image.FrameDimensionsList[0];
            var fdim = new FrameDimension(guid);

            var pageCount = GetPageCount(image, false);

            var ici = GetEncoderInfo("image/tiff");
            var enc = Encoder.Compression;
            var eps = new EncoderParameters(1);

            // Save the bitmap as a TIFF file with LZW compression.
            var ep = new EncoderParameter(enc, (long)EncoderValue.CompressionCCITT4);
            eps.Param[0] = ep;

            for (int i = 0; i < pageCount; i++)
            {
                image.SelectActiveFrame(fdim, i);
                var str = new MemoryStream();
                image.Save(str, ici, eps);
                output.Add(str);
            }

            image.Dispose();
            return output;
        }
        #endregion

        #region ConcatTiffImages
        /// <summary>
        /// Voegt de opgegeven lijst met Tiff bestanden samen tot 1 multi-page tiff bestand
        /// </summary>
        /// <param name="inputFiles">Lijst met de samen te voegen bestanden</param>
        /// <param name="outputFile">Het output bestand</param>
        /// <returns>Het samengevoegde bestand (inclusief pad)</returns>
        public static void ConcatTiffImages(List<string> inputFiles, string outputFile)
        {
            var ici = GetEncoderInfo("image/tiff");

            var enc = Encoder.SaveFlag;
            var eps = new EncoderParameters(2);
            eps.Param[0] = new EncoderParameter(enc, (long)EncoderValue.MultiFrame);
            eps.Param[1] = new EncoderParameter(Encoder.Compression, (long)EncoderValue.CompressionCCITT4);

            Image firstImage = null;

            foreach (var inputFile in inputFiles)
            {
                if (firstImage == null)
                {
                    firstImage = Image.FromFile(inputFile);
                    firstImage.Save(outputFile, ici, eps);
                }
                else
                {
                    eps.Param[0] = new EncoderParameter(enc, (long)EncoderValue.FrameDimensionPage);

                    var image = Image.FromFile(inputFile);
                    firstImage.SaveAdd(image, eps);
                    image.Dispose();
                }
            }

            if (firstImage != null)
                firstImage.Dispose();
        }

        /// <summary>
        /// Voegt de opgegeven lijst met Tiff bestanden samen tot 1 multi-page tiff bestand
        /// </summary>
        /// <param name="inputFiles">Lijst met de samen te voegen bestanden</param>
        /// <returns>Het samengevoegde bestand als memorystream</returns>
        public static MemoryStream ConcatTiffImages(List<string> inputFiles)
        {
            var ici = GetEncoderInfo("image/tiff");

            var enc = Encoder.SaveFlag;
            var eps = new EncoderParameters(2);
            eps.Param[0] = new EncoderParameter(enc, (long)EncoderValue.MultiFrame);
            eps.Param[1] = new EncoderParameter(Encoder.Compression, (long)EncoderValue.CompressionCCITT4);

            Image firstImage = null;
            var str = new MemoryStream();

            foreach (string inputFile in inputFiles)
            {
                if (firstImage == null)
                {
                    firstImage = Image.FromFile(inputFile);
                    firstImage.Save(str, ici, eps);
                }
                else
                {
                    eps.Param[0] = new EncoderParameter(enc, (long)EncoderValue.FrameDimensionPage);

                    var image = Image.FromFile(inputFile);
                    firstImage.SaveAdd(image, eps);
                    image.Dispose();
                }
            }

            if (firstImage != null)
                firstImage.Dispose();

            return str;
        }
        #endregion

        #region ChangeTiffResolution
        /// <summary>
        /// Past de resolutie aan van het opgegeven single-page of multi-page tiff bestand
        /// </summary>
        /// <param name="inputFile">Het aan te pasen bestand</param>
        /// <param name="outputFile">Het output bestand</param>
        /// <param name="resolution">De gewenste resolutie in DPI</param>
        public static void ChangeTiffResolution(string inputFile, string outputFile, int resolution)
        {
            //Eerst het originele tiff bestand splitsen in memory streams
            var streams = SplitTiffImage(inputFile);

            var ici = GetEncoderInfo("image/tiff");
            var enc = Encoder.SaveFlag;
            var eps = new EncoderParameters(2);

            eps.Param[0] = new EncoderParameter(enc, (long)EncoderValue.MultiFrame);
            eps.Param[1] = new EncoderParameter(Encoder.Compression, (long)EncoderValue.CompressionCCITT4);

            Bitmap outputBitmap = null;

            foreach (var memoryStream in streams)
            {
                if (outputBitmap == null)
                {
                    outputBitmap = (Bitmap)Image.FromStream(memoryStream);
                    outputBitmap.SetResolution(resolution, resolution);
                    outputBitmap.Save(outputFile, ici, eps);
                }
                else
                {
                    eps.Param[0] = new EncoderParameter(enc, (long)EncoderValue.FrameDimensionPage);
                    var bp = (Bitmap)Image.FromStream(memoryStream);
                    bp.SetResolution(resolution, resolution);
                    outputBitmap.SaveAdd(bp, eps);
                    bp.Dispose();
                }
            }

            if (outputBitmap != null)
                outputBitmap.Dispose();
        }
        #endregion

        #region GetEncodeInfo
        /// <summary>
        /// Geeft de ondersteunde encoder formaten
        /// </summary>
        /// <param name="mimeType">beschijving van het mime type</param>
        /// <returns>image codec informatie</returns>
        private static ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            var encoders = ImageCodecInfo.GetImageEncoders();
            foreach (var encoder in encoders)
            {
                if (encoder.MimeType == mimeType)
                    return encoder;
            }

            throw new Exception(mimeType + " mime type not found in ImageCodecInfo");
        }
        #endregion
    }
}