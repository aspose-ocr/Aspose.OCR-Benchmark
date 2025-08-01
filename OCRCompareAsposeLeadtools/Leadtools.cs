using Leadtools;
using Leadtools.Document.Writer;
using Leadtools.Ocr;
using System.Text;
using System.IO;

namespace OCRCompareAsposeLeadtools
{
    /// <summary>
    /// Class for performing OCR using LEADTOOLS.
    /// </summary>
    public static class LeadtoolsTest
    {
        static IOcrEngine ocrEngine = null;
        public static void Init()
        {
            // Set LEADTOOLS license
            RasterSupport.SetLicense(
                Path.Combine(BenchmarkAsposeLeadtools.FullPathToData, "LEADTOOLSEvaluationLicense", "LEADTOOLS.lic"),
                File.ReadAllText(Path.Combine(BenchmarkAsposeLeadtools.FullPathToData, "LEADTOOLSEvaluationLicense", "LEADTOOLS.lic.key"))
            );

            if (RasterSupport.KernelExpired)
                Console.WriteLine("License file invalid or expired.");

            ocrEngine = OcrEngineManager.CreateEngine(OcrEngineType.LEAD);
        }

        public static void Dispose()
        {
            ocrEngine.Dispose();
        }

        /// <summary>
        /// Performs OCR on an image using LEADTOOLS.
        /// </summary>
        /// <param name="imagePath">Path to the image file.</param>
        /// <param name="language">Recognition language (e.g., "eng", "ru", "zh-Hant").</param>
        /// <returns>Recognized text or an error message.</returns>
        public static string RunOcr(string imagePath, string language)
        {
            // LEADTOOLS does not support Hindi
            if (language == "hin")
                return "not supported";
         

            // Select the OCR runtime path based on the language
            string ocrRuntimePath = language switch
            {
                //"zh-Hant" => @"C:\Users\Admin\.nuget\packages\leadtools.ocr.languages.asian.net\23.0.0.4\content\OcrLEADRuntime",
                //"ru" => @"C:\Users\Admin\.nuget\packages\leadtools.ocr.languages.additional.net\23.0.0.4\content\OcrLEADRuntime",
                _ => Path.Combine(BenchmarkAsposeLeadtools.FullPathToData, "OcrLEADRuntime")
            };

            string result = string.Empty;
            // Use 'using' to automatically release resources

            ocrEngine.Startup(null, null, null, ocrRuntimePath);
            ocrEngine.LanguageManager.EnableLanguages(new string[] { language });

            // Create an OCR document and add pages
            IOcrDocument ocrDocument = ocrEngine.DocumentManager.CreateDocument();
            ocrDocument.Pages.AddPages(imagePath, 1, -1, null);

            // Automatically detect zones
            ocrDocument.Pages.AutoZone(null);

            // Example: if you need to exclude the first zone from recognition (usually not required)
            // You can comment out the following block if you want to recognize everything:
            /*
            if (ocrDocument.Pages[0].Zones.Count > 0)
            {
                OcrZone ocrZone = ocrDocument.Pages[0].Zones[0];
                ocrZone.ZoneType = OcrZoneType.Graphic;
                ocrDocument.Pages[0].Zones[0] = ocrZone;
            }
            */

            // Recognize text
            ocrDocument.Pages.Recognize(null);

            // Save the result to a temporary file
            string tempFile = Path.GetTempFileName();
            ocrDocument.Save(tempFile, DocumentFormat.Text, null);

            // Read the result
            result = File.ReadAllText(tempFile, Encoding.GetEncoding("windows-1251"));

            // Delete the temporary file
            File.Delete(tempFile);

            return result.Trim();
        }
    }
}
