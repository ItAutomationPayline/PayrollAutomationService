using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.SqlServer.Server;
using OfficeOpenXml;

namespace ExcelAutomationService
{
    public class MasterCTC
    {
        public static void CTC_Master(string ctccodes, string filePath, string destinationFolder)
        {
            try
            {
                //string vlookuppath = @"E:\PAYROLL_SERVER\Automation\Config\" + "["+ Path.GetFileName(filePath)+"]";
                //Service1.PathLog("VlookupPath:"+vlookuppath);
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int IP = Service1.getSheetNumber(filePath, "Payments and Deductions");
                    var inputWorkSheet = package.Workbook.Worksheets[IP];
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int hrid = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                    int payelementdescription = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Element Description");
                    int payelementshortcode = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Element Short Code");
                    int witheffectfrom = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Start Date");
                    int payfreq = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Frequency");
                    int frequencyamount = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Frequency Amount");
                    int amount = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Amount");

                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Ctc_New_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "With effect From(YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 3].Value = "Annual CTC";
                        using (var CTCPackage = new ExcelPackage(ctccodes)) 
                        {
                            var ctcsheet= CTCPackage.Workbook.Worksheets[0];
                            int ctclength = ctcsheet.Dimension.End.Row;
                            int row = 1;
                            int coll=0;
                            for ( coll = 4; coll<= ctclength+3; coll++)
                            {
                                outputWorksheet.Cells[1,coll].Value= ctcsheet.Cells[row,1].Text;
                                //if (ctcsheet.Cells[2, coll].Text == ""|| ctcsheet.Cells[2, coll].Text == "0")
                                //    outputWorksheet.Cells[2, coll].Formula = ctcsheet.Cells[row, 2].Formula;
                                if (!string.IsNullOrEmpty(ctcsheet.Cells[row, 2].Formula)) {
                                    if (ctcsheet.Cells[row, 2].Formula.Contains("VLOOKUP"))
                                    {
                                        outputWorksheet.Cells[2, coll].Formula = @"=VLOOKUP(A2,'E:\PAYROLL_SERVER\Automation\Config\[1temp.xlsx]Joiner and Changes '!$B$2:$AR$29,43,0)";
                                        //outputWorksheet.Cells[2, coll].Formula = ctcsheet.Cells[row, 2].Formula.Replace(@"E:\PAYROLL_SERVER\Automation\Input\[McAfee_Mcafee Software (India) Pvt Ltd._India_February_20250211_Payroll_Data.xlsx]", Path.GetFileName(filePath));
                                        Service1.PathLog(outputWorksheet.Cells[2, coll].Formula);
                                    }
                                    else { 
                                    outputWorksheet.Cells[2, coll].Formula = ctcsheet.Cells[row, 2].Formula;
                                    }

                                }
                                else 
                                {
                                    outputWorksheet.Cells[2, coll].Value= ctcsheet.Cells[row, 2].Text;
                                }
                                row++;
                            }
                        }
                        
                        HashSet<string> HRID = new HashSet<string>();
                        for (int row = 2; row <= lastRow; row++)
                        {
                            var cell = inputWorkSheet.Cells[row, hrid];
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                            {
                                HRID.Add(inputWorkSheet.Cells[row, hrid].GetValue<string>());
                            }
                        }
                        int row2 = 2;
                        foreach (string t in HRID)
                        {
                            outputWorksheet.Cells[row2, 1].Value = t;
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row, hrid].GetValue<string>() == t)
                                {
                                    outputWorksheet.Cells[row2, 2].Value = inputWorkSheet.Cells[row, witheffectfrom].GetValue<string>();
                                    if (inputWorkSheet.Cells[row, payelementshortcode].GetValue<string>().ToLower().Contains("basic") && inputWorkSheet.Cells[row, payfreq].GetValue<string>() == "Annual") 
                                    {
                                        outputWorksheet.Cells[row2, 3].Value = inputWorkSheet.Cells[row, amount].GetValue<double>();
                                    }
                                }
                            }
                            row2++;
                        }
                        int outputrows = outputWorksheet.Dimension.End.Row;
                        int outputcolumns = outputWorksheet.Dimension.End.Column;
                        int col = 4;

                        for (col = 4; col <= outputcolumns; col++)
                        {
                            // Get the base formula from row 2 (template)
                            var baseFormula = outputWorksheet.Cells[2, col].Formula;
                            for (row2 = 3; row2 <= outputrows; row2++)
                            {
                                if (!string.IsNullOrWhiteSpace(baseFormula))
                                {
                                    // Adjust the base formula to the current row
                                    string adjustedFormula = AdjustFormulaToRow(baseFormula, 2, row2);

                                    outputWorksheet.Cells[row2, col].Formula = adjustedFormula;
                                }
                                else
                                {
                                    var value = outputWorksheet.Cells[2, col].Text;
                                    if (value.All(char.IsDigit))
                                    {
                                        outputWorksheet.Cells[row2, col].Value = value;
                                    }
                                }
                            }
                            //else 
                            //{
                            //    var value = outputWorksheet.Cells[2, col].Text;
                            //    if (value.All(char.IsDigit)){
                            //        outputWorksheet.Cells[row2, col].Value = value;
                            //    }
                            //}
                        }
                        string AdjustFormulaToRow(string formula, int baseRow, int targetRow)
                        {
                            if ((formula.Replace(".","")).All(char.IsDigit)) 
                            {
                                return formula;
                            }
                           
                            return Regex.Replace(formula, @"(\$?[A-Z]{1,3})(\$?\d+)", match =>
                            {
                                string col2 = match.Groups[1].Value;      // Could be like "$C"
                                string rowStr = match.Groups[2].Value;   // Could be like "$2"

                                if (!rowStr.StartsWith("$") && int.TryParse(rowStr, out int row))
                                {
                                    if (row == baseRow)
                                    {
                                        return col2 + targetRow.ToString(); // Replace with same column, new row
                                    }
                                }

                                return match.Value; // Leave constants and absolute refs unchanged
                            });
                        }
                        row2 = 2;
                        lastRow = inputWorkSheet.Dimension.End.Row;
                        outputcolumns = outputWorksheet.Dimension.End.Column;
                        for (col=2; col<=outputcolumns;col++)
                        {
                            if (outputWorksheet.Cells[2,col].Text.ToLower()== "fixed") 
                            {
                                outputWorksheet.Cells[2, col].Value = 0;
                                Service1.PathLog("Got fixed");
                                for (int row3 = 2; row3 <= lastRow; row3++)
                                {
                                    for (int row = 2; row <= lastRow; row++)
                                    {
                                        if ((Service1.ShrinkString(inputWorkSheet.Cells[row, hrid].Text) == Service1.ShrinkString(outputWorksheet.Cells[row2, 1].Text)&& (Service1.ShrinkString(inputWorkSheet.Cells[row, payelementshortcode].Text) == Service1.ShrinkString(outputWorksheet.Cells[1, col].Text))))
                                        {
                                            outputWorksheet.Cells[row2, col].Value = inputWorkSheet.Cells[row, amount].GetValue<double>();
                                            row2++;
                                        }
                                    }
                                    if (outputWorksheet.Cells[row2, col].Text == ""&& outputWorksheet.Cells[row2, 1].Text != "") 
                                    {
                                        outputWorksheet.Cells[row2, col].Value = 0;
                                    }
                                    row2++;
                                }
                            }
                            row2 = 2;
                        }
                        outputWorksheet.Column(17).Style.Numberformat.Format = "0.00";
                        string newFileName = Path.Combine(destinationFolder, Service1.FileCount + "]NEW_Joiners_CTC" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        string cellValue = outputWorksheet.Cells[2, 1].GetValue<string>();
                        Service1.ShrinkString(cellValue);
                        if ((cellValue != null) && (cellValue != " ") && (cellValue != " "))
                        {
                            outputPackage.SaveAs(newFileInfo);
                            outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                            Service1.FileCount++;
                        }
                    }
                    Console.WriteLine("CTC Excel file created successfully!");
                }
            }
            catch (Exception ex)
            {
                Service1.PathLog($"An error occurred: {ex.Message}");
            }
        }
    }
}