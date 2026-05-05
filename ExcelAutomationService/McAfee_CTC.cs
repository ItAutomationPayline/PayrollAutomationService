using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelAutomationService
{
    public class McAfee_CTC
    {
        public static void CTC_Master(string ascendcodes, string filePath, string destinationFolder)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var inputWorkSheet = package.Workbook.Worksheets["Payments and Deductions"];
                    var joinerandChangesSheet = package.Workbook.Worksheets["Joiner and Changes "];
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int lastRow2 = joinerandChangesSheet.Dimension.End.Row;
                    int employee_Number = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");

                    int payelementdescription = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Element Description");
                    int town = Service1.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "PT Location");
                    int PayScale = Service1.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "Employee Grade");
                    int hrid = Service1.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "Hr id");
                    int doj = Service1.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "Payroll Start Date");
                    // int payelementdescription = Service1.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "Pay Element Short Code");
                    int witheffectfrom = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Start Date");
                    int annualctc = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Amount");
                    int payfreq = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Frequency");
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Ctc_Master");
                        outputWorksheet.Cells[1, 1].Value = "EE Code";
                        outputWorksheet.Cells[1, 2].Value = "Doj";
                        outputWorksheet.Cells[1, 3].Value = "Annual CTC";
                        outputWorksheet.Cells[1, 4].Value = "Pay Scale Desc";
                        outputWorksheet.Cells[1, 5].Value = "001 Basic Salary";
                        outputWorksheet.Cells[1, 6].Value = "002 Residual Pay";
                        outputWorksheet.Cells[1, 7].Value = "003 Special Allowance";
                        outputWorksheet.Cells[1, 8].Value = "006 Stipend";
                        outputWorksheet.Cells[1, 9].Value = "007 Employer NPS";
                        outputWorksheet.Cells[1, 10].Value = "Employer PF";
                        HashSet<string> HRID = new HashSet<string>();
                        int row2 = 2;
                        for (int row = 2; row <= lastRow2; row++)
                        {
                            var cell = joinerandChangesSheet.Cells[row, hrid];
                            // Get the background color of the cell
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF", StringComparison.OrdinalIgnoreCase))
                            {
                                outputWorksheet.Cells[row2, 1].Value = joinerandChangesSheet.Cells[row, hrid].GetValue<string>();
                                outputWorksheet.Cells[row2, 2].Value = joinerandChangesSheet.Cells[row, doj].GetValue<string>();
                                string grade = joinerandChangesSheet.Cells[row, PayScale].GetValue<string>();
                                grade = grade.Substring(grade.Length - 2);
                                outputWorksheet.Cells[row2, 4].Value = grade;
                                row2++;
                            }
                        }
                        row2 = 2;
                        double annual;
                        double percentage;
                        int OutPutWorkSheetlastRow = outputWorksheet.Dimension.End.Row;
                        for (int row = 2; row <= OutPutWorkSheetlastRow; row++)
                        {
                            for (row2 = 2; row2 <= lastRow; row2++)
                            {
                                var cell = outputWorksheet.Cells[row, 1];
                                // Get the background color of the cell
                                var bgColor = cell.Style.Fill.BackgroundColor;
                                if (outputWorksheet.Cells[row, 1].GetValue<string>().Equals(inputWorkSheet.Cells[row2, employee_Number].GetValue<string>()))
                                {
                                    // outputWorksheet.Cells[row2, 2].Value = inputWorkSheet.Cells[row, witheffectfrom].GetValue<string>();
                                    if (inputWorkSheet.Cells[row2, payfreq].GetValue<string>().Equals("Annual"))
                                    {
                                        //annual = inputWorkSheet.Cells[row, annualctc].GetValue<double>();
                                        outputWorksheet.Cells[row, 3].Value = inputWorkSheet.Cells[row2, annualctc].GetValue<double>();
                                    }
                                }
                            }
                        }
                        for (int row = 2; row <= OutPutWorkSheetlastRow; row++)
                        {
                            annual = outputWorksheet.Cells[row, 3].GetValue<double>();
                            int grade = outputWorksheet.Cells[row, 4].GetValue<int>();
                            using (var package2 = new ExcelPackage(new FileInfo(ascendcodes)))
                            {
                                int ip = Service1.getSheetNumber(ascendcodes, "Pay Scale");
                                var AscendPayScaleSheet = package2.Workbook.Worksheets[ip];
                                int AscendEmpGrade = Service1.getColumnNumber(ascendcodes, AscendPayScaleSheet.ToString(), "Employee Grade");
                                int AscendPercentage = Service1.getColumnNumber(ascendcodes, AscendPayScaleSheet.ToString(), "Percentage");
                                int AscendPayScaleSheetEndRow = AscendPayScaleSheet.Dimension.End.Row;
                                outputWorksheet.Cells[row, 7].Value = 0;
                                outputWorksheet.Cells[row, 8].Value = 0;
                                for (int row3 = 2; row3 <= AscendPayScaleSheetEndRow; row3++)
                                {
                                    if (AscendPayScaleSheet.Cells[row3, AscendEmpGrade].GetValue<string>().Equals(outputWorksheet.Cells[row, 4].GetValue<string>()))
                                    {
                                        percentage = AscendPayScaleSheet.Cells[row3, AscendPercentage].GetValue<double>();
                                        outputWorksheet.Cells[row, 5].Value = Math.Round((annual / 12) * percentage, 2);
                                    }
                                }
                                outputWorksheet.Cells[row, 6].Value = Math.Round((annual / 12) - outputWorksheet.Cells[row, 5].GetValue<double>(), 2);
                                if (outputWorksheet.Cells[row, 5].GetValue<string>() != null || outputWorksheet.Cells[row, 5].GetValue<string>() != "")
                                {
                                    outputWorksheet.Cells[row, 10].Value = Math.Round(Convert.ToDouble(outputWorksheet.Cells[row, 5].GetValue<string>()) * (12.0 / 100));
                                }
                                ip = Service1.getSheetNumber(ascendcodes, "Special Allowance");
                                var AscendSpecialAllowanceSheet = package2.Workbook.Worksheets[ip];
                                AscendEmpGrade = Service1.getColumnNumber(ascendcodes, AscendSpecialAllowanceSheet.ToString(), "Employee Grade");
                                int AscendSpecialAllowance = Service1.getColumnNumber(ascendcodes, AscendSpecialAllowanceSheet.ToString(), "Special Allowance");
                                int AscendSpecialAllowanceSheetEndRow = AscendSpecialAllowanceSheet.Dimension.End.Row;
                                for (int row3 = 2; row3 <= AscendSpecialAllowanceSheetEndRow; row3++)
                                {
                                    if (AscendSpecialAllowanceSheet.Cells[row3, AscendEmpGrade].GetValue<string>().Equals(outputWorksheet.Cells[row, 4].GetValue<string>()))
                                    {
                                        outputWorksheet.Cells[row, 7].Value = Math.Round(AscendSpecialAllowanceSheet.Cells[row3, AscendSpecialAllowance].GetValue<double>(), 2);
                                        if (Math.Round(AscendSpecialAllowanceSheet.Cells[row3, AscendSpecialAllowance].GetValue<double>(), 2) == 0)
                                        {
                                            outputWorksheet.Cells[row, 8].Value = outputWorksheet.Cells[row, 3].GetValue<double>() / 12.0;
                                            outputWorksheet.Cells[row, 6].Value = 0;
                                            outputWorksheet.Cells[row, 7].Value = 0;
                                            outputWorksheet.Cells[row, 10].Value = 0;
                                        }
                                    }
                                }
                                percentage = 0;
                                outputWorksheet.Cells[row, 9].Value = 0;
                            }
                        }
                        outputWorksheet.Column(3).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(5).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(6).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(7).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(8).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(9).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(10).Style.Numberformat.Format = "0.00";
                        string newFileName = Path.Combine(destinationFolder, Service1.FileCount + "]NEW_Joiners_CTC" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        Service1.PathLog("CTCFILENAME:" + newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        string cellValue = outputWorksheet.Cells[2, 1].Text;
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
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
