using System;
using OfficeOpenXml;
using System.IO;
using ReadExcel.Enums;
using ReadExcel.Models;
using ReadExcel.Services;

namespace ExcelConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "/Users/saidumallela/Dropbox (ASU)/Mac/Documents/RaviProject/Test.xlsm"; // Replace with the actual path to your Excel file
            string sheetName = "MetaData";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelService excelService = new ExcelService();
                ExcelWorksheet excelWorksheet = excelService.GetWorksheet(package, sheetName);

                List<string> categories = excelService.GetColumnAValues(excelWorksheet);
                Console.WriteLine(categories);

                String selectedCategory = excelService.FindValueInColumnB(excelWorksheet);
                Console.WriteLine(selectedCategory);

                List<string> sheetsForCategory = excelService.GetColumnEValues(excelWorksheet);
                Console.WriteLine(sheetsForCategory);

                MakeOperationsService makeOperationsService = new MakeOperationsService();
                makeOperationsService.MakeOperations(package, filePath, excelWorksheet, sheetsForCategory);

               // excelService.CopyColumnsJandKtoLM(excelWorksheet);
                package.Save();
                Console.WriteLine("Columns copied successfully.");

                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
        }  
    }
}
