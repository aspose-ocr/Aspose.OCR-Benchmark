using Aspose.AI.LLM.Abstractions.Parameters;
using Aspose.OCR;
using Aspose.OCR.AI;
using Aspose.Pdf.Forms;
using Leadtools.Ocr;
using License = Aspose.OCR.License;

namespace OCRCompareAsposeLeadtools
{
    /// <summary>
    /// Class for performing OCR using Aspose.OCR.
    /// </summary>
    public class AsposeTest
    {
        static AsposeOcr ocrEngine = null;
        AsposeAI ai = null;
        public void Init()
        {
            // Set Aspose.OCR license
            License lic = new License();
            lic.SetLicense(Path.Combine(BenchmarkAsposeLeadtools.FullPathToData, "Aspose.OCR.NET2025_2.lic"));

            ocrEngine = new AsposeOcr();
            AsposeLLMModelConfig config = null;
            if (BenchmarkAsposeLeadtools.ModelFilePath != null || BenchmarkAsposeLeadtools.HuggingFaceRepoId != null || BenchmarkAsposeLeadtools.HuggingFaceQuantization != null)
            {
                config = new AsposeLLMModelConfig();
                config.SourceParameters.ModelFilePath = BenchmarkAsposeLeadtools.ModelFilePath;
                config.SourceParameters.HuggingFaceRepoId = BenchmarkAsposeLeadtools.HuggingFaceRepoId;
                config.SourceParameters.HuggingFaceQuantization = BenchmarkAsposeLeadtools.HuggingFaceQuantization;
                ai = new AsposeAI(true, null, null, config);
            }
            else
            {
                ai = new AsposeAI();
            }

            Console.WriteLine($"AI run with {BenchmarkAsposeLeadtools.ModelFilePath} {BenchmarkAsposeLeadtools.HuggingFaceRepoId} {BenchmarkAsposeLeadtools.HuggingFaceQuantization}");
           
            ai.AddPostProcessor(new SpellCheckAIProcessor());
            Console.WriteLine("Model downloaded and AI initialized");
        }

        public void Dispose()
        {
            ocrEngine.Dispose();
            ai.Dispose();
        }

        /// <summary>
        /// Performs OCR on an image using Aspose.OCR.
        /// </summary>
        /// <param name="imagePath">Path to the image file.</param>
        /// <param name="language">Recognition language (Aspose.OCR.Language enum).</param>
        /// <returns>Recognized text.</returns>
        public OcrOutput RunOcr(string imagePath, Language language)
        {
            // Prepare input for OCR
            OcrInput input = new OcrInput(InputType.SingleImage);
            input.Add(imagePath);

            // Perform OCR recognition
            var resultAspose = ocrEngine.Recognize(input, new RecognitionSettings
            {
                Language = language
            });

            // Clean up resources
            input.Clear();  
            return resultAspose;
        }

        /// <summary>
        /// Performs OCR on an image using Aspose.OCR.
        /// </summary>
        /// <param name="imagePath">Path to the image file.</param>
        /// <param name="language">Recognition language (Aspose.OCR.Language enum).</param>
        /// <returns>Recognized text.</returns>
        public string RunAIOcr(OcrOutput resultAspose)
        {
            ai.RunPostprocessor(resultAspose);
            return resultAspose[0].RecognitionText.Trim();
        }
    }
}
