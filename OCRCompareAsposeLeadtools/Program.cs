using Aspose.OCR;

namespace OCRCompareAsposeLeadtools
{
    /// <summary>
    /// Entry point for the OCR benchmarking application.
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// Main method to run the OCR benchmark.
        /// </summary>
        /// <param name="args">Command-line arguments (not used).</param>
        static void Main(string[] args)
        {
            // SET PATH TO RESULTS EXCEL FILE
            string excelFileName = @"benchmark_leadtools.xlsx";

            // SET PATH TO DATA FOLDER
            string relativePath = Path.Combine("..", "..", "..", "DATA");
            BenchmarkAsposeLeadtools.FullPathToData = Path.GetFullPath(relativePath);

            Console.WriteLine("Enter full path to the model or [n] to use default");
            string model = Console.ReadLine();
            if(model != "n" && model != "N")
            {
                BenchmarkAsposeLeadtools.ModelFilePath = model;
            }

            // SET PATH TO MODEL OR HUGGINGFACE REPO ID AND QUANTIZATION
            //BenchmarkAsposeLeadtools.HuggingFaceRepoId = "lmstudio-community/Llama-3.2-3B-Instruct-GGUF";
            //BenchmarkAsposeLeadtools.HuggingFaceQuantization = "q4_k_m";
            //BenchmarkAsposeLeadtools.ModelFilePath = @".\models\lmstudio-community_Qwen3-14B-GGUF\Qwen3-14B-Q4_K_M.gguf";


            // Run the benchmark for Aspose and Leadtools OCR engines
            BenchmarkAsposeLeadtools.RunCompetitorsBenchmark(
                excelFileName,
                "images", // Worksheet name
                Path.Combine(BenchmarkAsposeLeadtools.FullPathToData, "DATASET"), // Directory with images
                Language.Latin, // Aspose OCR language
                "en",           // Leadtools OCR language
                isWithImage: false, // Do not include images in Excel
                isWithTexts: true  // Do not include recognized/reference texts in Excel
            );
        }
    }
}
