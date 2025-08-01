using Aspose.OCR;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace OCRCompareAsposeLeadtools
{
    /// <summary>
    /// Class for benchmarking OCR results and performance between Aspose and Leadtools.
    /// </summary>
    public static class BenchmarkAsposeLeadtools
    {
        public static string FullPathToData = "";
        public static string HuggingFaceRepoId = null;
        public static string HuggingFaceQuantization = null;
        public static string ModelFilePath = null;

        /// <summary>
        /// Runs a benchmark comparing Aspose and Leadtools OCR engines on a set of images.
        /// </summary>
        /// <param name="excelFileName">Path to the Excel file for results.</param>
        /// <param name="listName">Worksheet name.</param>
        /// <param name="directoryWithImages">Directory containing image subfolders.</param>
        /// <param name="asposeLanguage">Aspose OCR language.</param>
        /// <param name="leadtoolLang">Leadtools OCR language.</param>
        /// <param name="isWithImage">Whether to include images in the Excel output.</param>
        /// <param name="isWithTexts">Whether to include recognized and reference texts in the Excel output.</param>
        public static void RunCompetitorsBenchmark(
            string excelFileName,
            string listName,
            string directoryWithImages,
            Language asposeLanguage, string leadtoolLang, bool isWithImage, bool isWithTexts)
        {
            int columnWidth = isWithTexts ? 30 : 10;
            CreateExcel.AddListToWorkBook(listName, excelFileName, new string[] { "Image Name", "Picture", "Ethalon", "Aspose", "", "","Aspose AI", "", "",
                "Leadtools", "", "" }, columnWidth, isWithImage);
            CreateExcel.AddDataOnList(listName, excelFileName, null, new object[] { "", "Result", "Time, ms", "Lev, %",
                "Result", "Time, ms", "Lev, %", "Time, ms", "Lev, %" });

            int counterImages = 1;
            double levAspTotal = 0;
            double levAspAITotal = 0;
            double levLeadTotal = 0;
            double timeAspTotal = 0;
            double timeAspAITotal = 0;
            double timeLeadTotal = 0;

            LeadtoolsTest.Init();
            AsposeTest aspose = new AsposeTest();
            aspose.Init();

            foreach (string subdir in Directory.EnumerateDirectories(directoryWithImages))
            {
                foreach (string imageName in Directory.EnumerateFiles(subdir).Where(f => !f.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)))
                {
                    Console.WriteLine(imageName);
                    // Read reference (ethalon) text
                    string ethalonFile = Path.ChangeExtension(imageName, ".txt");
                    string ethalonText = File.ReadAllText(ethalonFile);

                    // --- Aspose OCR ---
                    Console.WriteLine("Aspose start");
                    Stopwatch st = new Stopwatch();
                    st.Start();
                    // var outputAspose = aspose.RunOcr(imageName, asposeLanguage);
                    string resultAspose = "";// outputAspose[0].RecognitionText.Trim();
                    st.Stop();
                    var timeAspose = st.ElapsedMilliseconds;
                    var lA = 100 - LevenshteinDistance(ethalonText, resultAspose) * 100 / (double)ethalonText.Length;
                    levAspTotal += lA;
                    timeAspTotal += timeAspose;

                    // --- Aspose OCR with AI---
                    Console.WriteLine("AI start");
                    st.Restart();
                    var resultAsposeAI = "";// aspose.RunAIOcr(outputAspose);
                    st.Stop();
                    var timeAsposeAI = st.ElapsedMilliseconds;
                    var lAI = 100 - LevenshteinDistance(ethalonText, resultAsposeAI) * 100 / (double)ethalonText.Length;
                    levAspAITotal += lAI;
                    timeAspAITotal += timeAsposeAI;

                    // --- Leadtools OCR ---
                    Console.WriteLine("Leadtools start");
                    st.Restart();
                    string resultLeadtools = LeadtoolsTest.RunOcr(imageName, leadtoolLang);
                    st.Stop();
                    var timeLeadtools = st.ElapsedMilliseconds;
                    var lL = 100 - LevenshteinDistance(ethalonText, resultLeadtools) * 100 / (double)ethalonText.Length;
                    levLeadTotal += lL;
                    timeLeadTotal += timeLeadtools;

                    // --- Save results to Excel ---
                    if (!isWithTexts)
                    {
                        ethalonText = null;
                        resultAspose = null;
                        resultAsposeAI = null;
                        resultLeadtools = null;
                    }

                    CreateExcel.AddDataOnList(listName, excelFileName, imageName, new object[] { ethalonText,
                        resultAspose, timeAspose, lA,
                        resultAsposeAI, timeAsposeAI, lAI,
                        resultLeadtools, timeLeadtools, lL }, isWithImage);

                    counterImages++;
                }
            }

            LeadtoolsTest.Dispose();
            aspose.Dispose();

            // Add summary statistics to Excel
            CreateExcel.AddDataOnList(listName, excelFileName, null, new object[] { "",
                "AVG", timeAspTotal / counterImages, levAspTotal / counterImages,
                "AVG", timeAspAITotal / counterImages, levAspAITotal / counterImages,
                "AVG", timeLeadTotal / counterImages, levLeadTotal / counterImages});
            CreateExcel.AddDataOnList(listName, excelFileName, null, new object[] { "images amount", counterImages });
        }

        /// <summary>
        /// Calculates the Levenshtein distance between two strings.
        /// </summary>
        private static int LevenshteinDistance(string _firstWord, string _secondWord)
        {
            var firstWord = CleanTextContent(_firstWord);
            var secondWord = CleanTextContent(_secondWord);
            var n = firstWord.Length + 1;
            var m = secondWord.Length + 1;
            var matrixD = new int[n, m];

            const int deletionCost = 1;
            const int insertionCost = 1;

            for (var i = 0; i < n; i++)
                matrixD[i, 0] = i;
            for (var j = 0; j < m; j++)
                matrixD[0, j] = j;

            for (var i = 1; i < n; i++)
            {
                for (var j = 1; j < m; j++)
                {
                    var substitutionCost = firstWord[i - 1] == secondWord[j - 1] ? 0 : 1;
                    matrixD[i, j] = Minimum(matrixD[i - 1, j] + deletionCost,          // deletion
                        matrixD[i, j - 1] + insertionCost,         // insertion
                        matrixD[i - 1, j - 1] + substitutionCost); // substitution
                }
            }
            return matrixD[n - 1, m - 1];
        }

        static string CleanTextContent(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            // Remove carriage returns (\r)
            string cleaned = input.Replace("\r", "\n");

            // Remove quotes (both single and double)
            cleaned = cleaned.Replace("\"", "").Replace("'", "");

            // Replace multiple line breaks with single space
            cleaned = Regex.Replace(cleaned, @"\n+", "\n");

            // Replace multiple spaces with single space
            cleaned = Regex.Replace(cleaned, @"[ \t]+", " ");

            // Trim leading and trailing whitespace
            cleaned = cleaned.Trim();

            return cleaned;
        }

        /// <summary>
        /// Returns the minimum of three integer values.
        /// </summary>
        private static int Minimum(int a, int b, int c) => (a = a < b ? a : b) < c ? a : c;
    }
}
