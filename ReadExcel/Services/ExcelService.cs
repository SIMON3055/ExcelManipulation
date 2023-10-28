using System;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using ReadExcel.Services;

namespace ReadExcel.Services
{
	public class ExcelService
	{
            public ExcelWorksheet GetWorksheet(ExcelPackage package, string sheetName)
            {
                if (package.Workbook.Worksheets[sheetName] != null)
                {
                    return package.Workbook.Worksheets[sheetName];
                }
                else
                {
                    throw new ArgumentException("Worksheet with the provided name does not exist.");
                }
            }

            public List<string> GetColumnAValues(ExcelWorksheet worksheet)
            {
                List<string> columnAList = new List<string>();

                for (int row = worksheet.Dimension.Start.Row + 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    var cellValue = worksheet.Cells[row, 1].Value;
                    if (cellValue != null)
                    {
                        columnAList.Add(cellValue.ToString());
                    }
                }

                return columnAList;
            }

            public string FindValueInColumnB(ExcelWorksheet worksheet)
            {
                String category = "no Such Category";
                for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                {
                    var cellValueB = worksheet.Cells[row, 2].Value;
                    if (cellValueB != null && cellValueB.ToString() == "x")
                    {
                        var cellValueA = worksheet.Cells[row, 1].Value;
                        return cellValueA?.ToString();
                    }
                }

                return category; // If no value 'x' is found in column B
            }

            public void CopyColumnsJandKtoLM(ExcelWorksheet worksheet)
            {
                for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                {
                    var cellValueJ = worksheet.Cells[row, 10].Value;
                    var cellValueK = worksheet.Cells[row, 11].Value;

                    worksheet.Cells[row, 12].Value = cellValueJ; // Copying to column L
                    worksheet.Cells[row, 13].Value = cellValueK; // Copying to column M
                }
            }

            public List<string> GetColumnEValues(ExcelWorksheet excelWorksheet)
            {
                List<string> columnEList = new List<string>();

                for (int row = excelWorksheet.Dimension.Start.Row + 1; row <= excelWorksheet.Dimension.End.Row; row++)
                {
                var cellValueF = excelWorksheet.Cells[row, 6].Value;
                if (cellValueF != null && cellValueF.ToString() == "x")
                {
                    var cellValueE = excelWorksheet.Cells[row, 5].Value;
                    columnEList.Add(cellValueE.ToString());
                }
            }

                return columnEList;
            }
    }
}

