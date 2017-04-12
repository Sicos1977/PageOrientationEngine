using System;
using System.IO;
using PageOrientationEngine;

namespace TestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var poe = new DocumentInspector(@"c:\Program Files (x86)\Tesseract-OCR\tessdata\", "eng");
            var results = poe.DetectPageOrientation(@"d:\\example.tif");
            foreach(var result in results)
                Console.WriteLine(result.ToString());

            Console.ReadKey();
        }
    }
}
