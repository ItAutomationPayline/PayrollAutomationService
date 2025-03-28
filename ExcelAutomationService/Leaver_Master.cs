using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAutomationService
{
    public class Leaver_Master
    {
        public static void LeaverMaster(string ascendcodes, string filePath, string destinationFolder)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int IP = Service1.getSheetNumber(filePath, "Leavers Data");
                    var inputWorkSheet = package.Workbook.Worksheets[IP];// Assuming the data is in the first worksheet
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    using (var outputPackage = new ExcelPackage())
                    {
                        int employeenumber = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                        int dateofLeaving = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Payroll End Date");
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Leaver_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "Date Of Leaving (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 3].Value = "Payroll Code";
                        outputWorksheet.Cells[1, 4].Value = "InactiveID";
                        outputWorksheet.Cells[1, 5].Value = "Date Of Resign (YYYY-MM-DD)";
                        for (int row = 2; row <= lastRow; row++)
                        {
                            var HRID = inputWorkSheet.Cells[row, employeenumber].GetValue<string>();
                            var DateOfLeaving = inputWorkSheet.Cells[row, dateofLeaving].GetValue<string>();
                            DateOfLeaving = DateOfLeaving.Replace(" ", "");
                            if (DateOfLeaving.Length == 10)
                            {
                                outputWorksheet.Cells[row, 2].Value = DateOfLeaving;
                            }
                            outputWorksheet.Cells[row, 1].Value = HRID;
                            outputWorksheet.Cells[row, 3].Value = "9";
                            outputWorksheet.Cells[row, 4].Value = "3";
                            outputWorksheet.Cells[row, 5].Value = DateOfLeaving;
                        }
                        string newFileName = Path.Combine(destinationFolder, Service1.FileCount+ "]Leaver_Master " + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        string cellValue = outputWorksheet.Cells[2, 1].GetValue<string>();
                        Service1.ShrinkString(cellValue);
                        if ((cellValue != null) && (cellValue != " ") && (cellValue != " "))
                        {
                            outputPackage.SaveAs(newFileInfo);
                            outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                            Service1.Log("Leaver Master Created Successfully");
                            Service1.FileCount++;
                        }
                        else
                        {
                            Service1.PathLog("No Leavers found.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Service1.ErrorCount++;
                Service1.Log($"An error occurred: {ex.Message}");
            }
        }
    }
}
