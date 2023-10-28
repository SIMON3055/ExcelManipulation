using System;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using ReadExcel.Services;
using ReadExcel.Services;

namespace ReadExcel.Services
{
    public class MakeOperationsService
    {
        public MakeOperationsService()
        {

        }

        public void MakeOperations(ExcelPackage package, String filePath, ExcelWorksheet excelWorksheet, List<string> sheetsForCategory)
        {
            for (int row = excelWorksheet.Dimension.Start.Row+1; row <= excelWorksheet.Dimension.Start.Row+1; row++)
            {
                String sheetName = excelWorksheet.Cells[row, 8].Value.ToString().Trim();
                String commandName = excelWorksheet.Cells[row, 9].Value.ToString().Trim();
                String originStart = excelWorksheet.Cells[row, 10].Value.ToString().Trim();
                String originEnd = excelWorksheet.Cells[row, 11].Value.ToString().Trim();
                String DestStart = excelWorksheet.Cells[row, 12].Value.ToString().Trim();
                String DestEnd = excelWorksheet.Cells[row, 13].Value.ToString().Trim();

                if (sheetsForCategory.Contains(sheetName))
                {
                    DoOperation(package,excelWorksheet, filePath, sheetName, commandName, originStart, originEnd, DestStart, DestEnd);
                }
            }
        }

        private void DoOperation(ExcelPackage package, ExcelWorksheet excelWorksheet,String filePath, String sheetName, String commandName, String originStart, String originEnd, String destStart, String destEnd)
        {
            ExcelService excelService = new ExcelService();
            ExcelWorksheet excelWorksheet1 = excelService.GetWorksheet(package, sheetName);
            string command = commandName.Replace(" ", "");
                if(command == "RollDates" && sheetName=="Inputs")
                {
                    RollDates(originStart, originEnd, destStart, destEnd, excelWorksheet1);
                }
               
        }
        public  void RollDates(string originStart, string originEnd, string destStart, string destEnd, ExcelWorksheet worksheet)
        {
            worksheet.Cells[destStart].Value = worksheet.Cells[originStart].Value;
            worksheet.Cells[destEnd].Value = worksheet.Cells[originEnd].Value;
        }




    }
}

