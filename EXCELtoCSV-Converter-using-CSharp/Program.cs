using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using OfficeOpenXml;

namespace EXCELtoCSV_Converter_using_CSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set the EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string sourceFolder = ConfigurationManager.AppSettings["SourceFolder"];
            string destinationFolder = ConfigurationManager.AppSettings["DestinationFolder"];
            string connectionString = ConfigurationManager.ConnectionStrings["FileMoveHistoryDB"].ConnectionString;

            // Check if source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder {sourceFolder} does not exist.");
                return;
            }

            // Ensure destination folder exists
            if (!Directory.Exists(destinationFolder))
            {
                try
                {
                    Directory.CreateDirectory(destinationFolder);
                    Console.WriteLine($"Created folder: {destinationFolder}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error creating folder {destinationFolder}: {ex.Message}");
                    return;
                }
            }

            // Get all Excel files (xls and xlsx) from the source folder
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xls");
            string[] excelXlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx");
            List<string> files = new List<string>(excelFiles);
            files.AddRange(excelXlsxFiles);
            List<FileMoveRecord> moveRecords = new List<FileMoveRecord>();

            foreach (string file in files)
            {
                try
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    string destFile = Path.Combine(destinationFolder, $"{fileName}.csv");

                    // Convert Excel to CSV
                    if (ConvertExcelToCsv(file, destFile))
                    {
                        File.Delete(file);
                        Console.WriteLine($"Converted and moved {file} to {destFile}");
                        moveRecords.Add(new FileMoveRecord(fileName, "CSV", sourceFolder, destFile));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error moving file {file}: {ex.Message}");
                }
            }

            // Insert move history into SQL table
            DAL.InsertMoveHistory(moveRecords, connectionString);

            // Print summary
            Console.WriteLine("\nMove Process Summary:");
            Console.WriteLine($"Total files converted and moved: {moveRecords.Count}");
        }

        static bool ConvertExcelToCsv(string sourceFile, string destinationFile)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(sourceFile)))
                {
                    // Ensure the package contains at least one worksheet
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        Console.WriteLine($"No worksheets found in file {sourceFile}");
                        return false;
                    }

                    var worksheet = package.Workbook.Worksheets[0];
                    var dimension = worksheet.Dimension;

                    if (dimension == null)
                    {
                        Console.WriteLine($"No data found in the worksheet of file {sourceFile}");
                        return false;
                    }

                    int totalRows = dimension.End.Row;
                    int totalCols = dimension.End.Column;

                    using (var writer = new StreamWriter(destinationFile))
                    {
                        for (int row = 1; row <= totalRows; row++)
                        {
                            var rowContent = new List<string>();

                            for (int col = 1; col <= totalCols; col++)
                            {
                                var cellValue = worksheet.Cells[row, col].Text;
                                rowContent.Add(cellValue);
                            }

                            writer.WriteLine(string.Join(",", rowContent));
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting file {sourceFile} to CSV: {ex.Message}");
                return false;
            }
        }
    }

    public class FileMoveRecord
    {
        public string FileName { get; }
        public string FileType { get; }
        public string SourcePath { get; }
        public string DestinationPath { get; }

        public FileMoveRecord(string fileName, string fileType, string sourcePath, string destinationPath)
        {
            FileName = fileName;
            FileType = fileType;
            SourcePath = sourcePath;
            DestinationPath = destinationPath;
        }
    }
}