using Aspose.Drawing;
using Aspose.Drawing.Drawing2D;
using Aspose.Drawing.Imaging;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace OCRCompareAsposeLeadtools
{
    /// <summary>
    /// Utility class for creating and updating Excel files with images and data rows.
    /// </summary>
    public static class CreateExcel
    {
        /// <summary>
        /// Creates or updates an Excel file with a worksheet and header columns.
        /// </summary>
        /// <param name="listName">Worksheet name.</param>
        /// <param name="docExcelName">Excel file path.</param>
        /// <param name="columnNames">Header column names.</param>
        /// <param name="columnWidth">Default column width.</param>
        /// <param name="isWithImages">Whether to include image columns.</param>
        public static void AddListToWorkBook(string listName, string docExcelName, string[] columnNames, int columnWidth, bool isWithImages)
        {
            if (string.IsNullOrWhiteSpace(listName) || string.IsNullOrWhiteSpace(docExcelName) || columnNames == null || columnNames.Length == 0)
                throw new ArgumentException("Invalid arguments for creating Excel worksheet.");

            var fileInfo = new FileInfo(docExcelName);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[listName];
                if (worksheet == null)
                {
                    worksheet = excelPackage.Workbook.Worksheets.Add(listName);
                }

                // Write header columns
                for (int i = 0; i < columnNames.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = columnNames[i];
                    worksheet.Column(i + 1).Width = columnWidth;
                    worksheet.Column(i + 1).Style.WrapText = true;
                    worksheet.Column(i + 1).Style.Font.Size = 10;
                }

                // Set width for image path column
                worksheet.Column(1).Width = 50;

                // Set width for image column if needed
                if (isWithImages)
                {
                    worksheet.Column(2).Width = 40; // Picture column width
                }

                // Apply border to header row
                using (var range = worksheet.Cells[1, 1, 1, columnNames.Length])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    range.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                }

                excelPackage.Save();
            }
        }

        /// <summary>
        /// Adds a data row to the worksheet, optionally embedding an image in the row.
        /// </summary>
        /// <param name="listName">Worksheet name.</param>
        /// <param name="fileName">Excel file path.</param>
        /// <param name="imageName">Image file path (optional).</param>
        /// <param name="columns">Data columns.</param>
        /// <param name="isWithImage">Whether to embed the image.</param>
        public static void AddDataOnList(string listName, string fileName, string imageName, object[] columns, bool isWithImage = true)
        {
            if (string.IsNullOrWhiteSpace(listName) || string.IsNullOrWhiteSpace(fileName) || columns == null)
                throw new ArgumentException("Invalid arguments for adding data row.");

            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(fileName)))
            {
                var worksheet = excelPackage.Workbook.Worksheets[listName];
                if (worksheet == null)
                    throw new ArgumentException($"Worksheet '{listName}' not found in file '{fileName}'.");

                int lastRow = worksheet.Dimension?.End.Row + 1 ?? 2; // If empty, start at row 2
                worksheet.Cells[lastRow, 1].Value = imageName ?? string.Empty;
                for (int i = 0; i < columns.Length; i++)
                {
                    worksheet.Cells[lastRow, i + 3].Value = columns[i];
                }

                // Optionally embed image in the row
                if (isWithImage && !string.IsNullOrWhiteSpace(imageName) && File.Exists(imageName))
                {
                    try
                    {
                        using (var img = Aspose.Drawing.Image.FromFile(imageName))
                        {
                            int newWidth = 200;
                            int newHeight = (int)(img.Height * ((float)newWidth / img.Width));
                            using (var resizedBitmap = new Aspose.Drawing.Bitmap(newWidth, newHeight))
                            using (var graphics = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                            {
                                graphics.CompositingQuality = CompositingQuality.HighQuality;
                                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                graphics.SmoothingMode = SmoothingMode.HighQuality;
                                graphics.DrawImage(img, 0, 0, newWidth, newHeight);
                                using (var stream = new MemoryStream())
                                {
                                    resizedBitmap.Save(stream, Aspose.Drawing.Imaging.ImageFormat.Png);
                                    stream.Position = 0;
                                    var picture = worksheet.Drawings.AddPicture($"img_{lastRow}", stream);
                                    picture.SetPosition(lastRow - 1, 0, 1, 0); // Place image in column 2
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Optionally log or handle image embedding error
                        // For client delivery, consider logging to a file or system log
                        // Example: LogError($"Failed to embed image: {ex.Message}");
                    }
                }

                excelPackage.Save();
            }
        }
    }
} 