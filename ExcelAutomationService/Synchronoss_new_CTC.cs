using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ExcelAutomationService
{
    public class Synchronoss_new_CTC
    {
        public static void CTC_Master(string ascendcodes, string filePath, string destinationFolder)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int IP = Service1.getSheetNumber(filePath, "Payments and Deductions");
                    var inputWorkSheet = package.Workbook.Worksheets[IP];
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int hrid = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                    int payelementdescription = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Element Short Code");
                    int witheffectfrom = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Start Date");
                    int payfreq = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Frequency");
                    int amount = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Frequency Amount");

                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Ctc_New_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "With effect From(YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 3].Value = "Annual CTC";
                        string[] headers = {
                            "001 Basic", "002 HRA", "003 Special Allowance", "004 Project Allowance", "005 Food Allowance",
                            "006 LTA", "PF_ER", "008 Stipend", "521 Employer NPS", "154 Telephone", "155 Petrol",
                            "503 Provident Fund", "504 Profession Tax", "506 Voluntary Provident Fund"
                        };
                        for (int i = 0; i < headers.Length; i++)
                        {
                            outputWorksheet.Cells[1, i + 4].Value = headers[i];
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
                            for (int col = 3; col < 3 + headers.Length + 1; col++) // 3 to 17
                            {
                                outputWorksheet.Cells[row2, col].Value = 0; // Default value for all fields
                            }
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row, hrid].GetValue<string>() == t)
                                {
                                    outputWorksheet.Cells[row2, 2].Value = inputWorkSheet.Cells[row, witheffectfrom].GetValue<string>();
                                    string description = inputWorkSheet.Cells[row, payelementdescription].GetValue<string>();
                                    double freqAmount = inputWorkSheet.Cells[row, amount].GetValue<double>();

                                    for (int col = 4; col < 4 + headers.Length; col++)
                                    {
                                        if (Service1.ShrinkString(outputWorksheet.Cells[1, col].GetValue<string>()) == Service1.ShrinkString(description))
                                        {
                                            outputWorksheet.Cells[row2, col].Value = freqAmount;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (outputWorksheet.Cells[row2,4].GetValue<double>()!=0.00) 
                            {
                                outputWorksheet.Cells[row2, 15].Value = 1;
                                outputWorksheet.Cells[row2, 16].Value = 1;
                            }
                            row2++;
                        }
                        for (int col = 3; col<= 14; col++) { 
                            outputWorksheet.Column(col).Style.Numberformat.Format = "0.00";
                        }
                        for (int col = 1; col < headers.Length; col++)
                        {
                            string temp = Service1.ShrinkString(outputWorksheet.Cells[1, col].GetValue<string>());
                            if (temp.Contains("pf") || temp.Contains("provident"))
                            {
                                outputWorksheet.Cells[1, col].Value = "007 PF Employer";
                            }
                            if (temp.Contains("nps"))
                            {
                                outputWorksheet.Cells[1, col].Value = "009 Employer NPS";
                            }
                            if (temp.Contains("food"))
                            {
                                outputWorksheet.Cells[1, col].Value = "005 Food Allowance";
                            }
                            if (temp.Contains("basic"))
                            {
                                outputWorksheet.Cells[1, col].Value = "001 Basic";
                            }
                        }
                        outputWorksheet.Column(17).Style.Numberformat.Format = "0.00";
                        string newFileName = Path.Combine(destinationFolder,Service1.FileCount+ "]NEW_Joiners_CTC" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        string cellValue = outputWorksheet.Cells[2, 1].GetValue<string>();
                        Service1.ShrinkString(cellValue);
                        if ((cellValue != null) && (cellValue != " ") && (cellValue != " ")) { 
                            outputPackage.SaveAs(newFileInfo);
                            outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                            Service1.FileCount++;
                        }
                    }
                    Console.WriteLine("CTC Excel file created successfully!");
                }
                #region secondfile
                using (var package2 = new ExcelPackage(new FileInfo(filePath)))
                {
                    int IP = Service1.getSheetNumber(filePath, "Payments and Deductions");
                    var inputWorkSheet = package2.Workbook.Worksheets[IP];
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int hrid = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                    int payelementdescription = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Element Short Code");
                    int witheffectfrom = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Start Date");
                    int payfreq = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Frequency");
                    int amount = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Frequency Amount");
                    int enddate= Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "End Date");
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Ctc_Existing_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "With effect From(YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 3].Value = "Annual CTC";
                        string[] headers = {
                            "001 Basic", "002 HRA", "003 Special Allowance", "004 Project Allowance", "005 Food Allowance",
                            "006 LTA", "PF_ER", "008 Stipend", "521 Employer NPS", "154 Telephone", "155 Petrol",
                            "503 Provident Fund", "504 Profession Tax", "506 Voluntary Provident Fund"
                        };
                        for (int i = 0; i < headers.Length; i++)
                        {
                            outputWorksheet.Cells[1, i + 4].Value = headers[i];
                        }
                        HashSet<string> HRID = new HashSet<string>();
                        for (int row = 2; row <= lastRow; row++)
                        {
                            var cell = inputWorkSheet.Cells[row, hrid];
                            var cell2= inputWorkSheet.Cells[row, enddate];
                            var bgColor2 = cell2.Style.Fill.BackgroundColor;
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if ((string.IsNullOrEmpty(bgColor.Rgb) || bgColor.Rgb.Equals("FFFFFF")) && inputWorkSheet.Cells[row, enddate].Text=="" )
                            {
                                HRID.Add(inputWorkSheet.Cells[row, hrid].GetValue<string>());
                            }
                        }
                        int row2 = 2;
                        foreach (string t in HRID)
                        {
                            outputWorksheet.Cells[row2, 1].Value = t;
                            for (int col = 3; col < 3 + headers.Length + 1; col++) // 3 to 17
                            {
                                outputWorksheet.Cells[row2, col].Value = 0; // Default value for all fields
                            }
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row, hrid].GetValue<string>() == t)
                                {
                                    outputWorksheet.Cells[row2, 2].Value = inputWorkSheet.Cells[row, witheffectfrom].GetValue<string>();
                                    string description = inputWorkSheet.Cells[row, payelementdescription].GetValue<string>();
                                    double freqAmount = inputWorkSheet.Cells[row, amount].GetValue<double>();

                                    for (int col = 4; col < 4 + headers.Length; col++)
                                    {
                                        if (Service1.ShrinkString(outputWorksheet.Cells[1, col].GetValue<string>()) == Service1.ShrinkString(description))
                                        {
                                            outputWorksheet.Cells[row2, col].Value = freqAmount;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (outputWorksheet.Cells[row2, 4].GetValue<double>() != 0.00)
                            {
                                outputWorksheet.Cells[row2, 15].Value = 1;
                                outputWorksheet.Cells[row2, 16].Value = 1;
                            }
                            row2++;
                        }
                        for (int col = 3; col <= 14; col++)
                        {
                            outputWorksheet.Column(col).Style.Numberformat.Format = "0.00";
                        }
                        for (int col = 1; col < headers.Length; col++)
                        {
                            string temp = Service1.ShrinkString(outputWorksheet.Cells[1, col].GetValue<string>());
                            if (temp.Contains("pf")|| temp.Contains("provident")) {
                                outputWorksheet.Cells[1, col].Value = "007 PF Employer";
                            }
                            if (temp.Contains("nps"))
                            {
                                outputWorksheet.Cells[1, col].Value = "009 Employer NPS";
                            }
                            if (temp.Contains("food"))
                            {
                                outputWorksheet.Cells[1, col].Value = "005 Food Allowance";
                            }
                            if (temp.Contains("basic"))
                            {
                                outputWorksheet.Cells[1, col].Value = "001 Basic";
                            }
                        }
                        outputWorksheet.Column(17).Style.Numberformat.Format = "0.00";
                        string newFileName = Path.Combine(destinationFolder,Service1.FileCount+ "]Existing_CTC_Changes" + Path.GetFileName(filePath));
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
                #endregion
                #region thirdfile
                using (var package3 = new ExcelPackage(new FileInfo(filePath)))
                {
                    int IP = Service1.getSheetNumber(filePath, "Payments and Deductions");
                    var inputWorkSheet = package3.Workbook.Worksheets[IP];
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int hrid = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                    int payelementdescription = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Element Short Code");
                    int witheffectfrom = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Start Date");
                    int payfreq = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Frequency");
                    int amount = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Frequency Amount");
                    int enddate = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "End Date");
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Ctc_Leavers_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "With effect From(YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 3].Value = "Annual CTC";
                        string[] headers = {
                            "001 Basic", "002 HRA", "003 Special Allowance", "004 Project Allowance", "005 Food Allowance",
                            "006 LTA", "PF_ER", "008 Stipend", "521 Employer NPS", "154 Telephone", "155 Petrol",
                            "503 Provident Fund", "504 Profession Tax", "506 Voluntary Provident Fund"
                        };
                        for (int i = 0; i < headers.Length; i++)
                        {
                            outputWorksheet.Cells[1, i + 4].Value = headers[i];
                        }
                        HashSet<string> HRID = new HashSet<string>();
                        for (int row = 2; row <= lastRow; row++)
                        {
                            var cell = inputWorkSheet.Cells[row, hrid];
                            var cell2 = inputWorkSheet.Cells[row, enddate];
                            var bgColor2 = cell2.Style.Fill.BackgroundColor;
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if ((string.IsNullOrEmpty(bgColor.Rgb) || bgColor.Rgb.Equals("FFFFFF")) && Service1.ShrinkString(inputWorkSheet.Cells[row, enddate].Text) != "")
                            {
                                HRID.Add(inputWorkSheet.Cells[row, hrid].GetValue<string>());
                            }
                        }
                        int row2 = 2;
                        foreach (string t in HRID)
                        {
                            outputWorksheet.Cells[row2, 1].Value = t;
                            for (int col = 3; col < 3 + headers.Length + 1; col++) // 3 to 17
                            {
                                outputWorksheet.Cells[row2, col].Value = 0; // Default value for all fields
                            }
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row, hrid].GetValue<string>() == t)
                                {
                                    outputWorksheet.Cells[row2, 2].Value = inputWorkSheet.Cells[row, witheffectfrom].GetValue<string>();
                                    string description = inputWorkSheet.Cells[row, payelementdescription].GetValue<string>();
                                    double freqAmount = inputWorkSheet.Cells[row, amount].GetValue<double>();

                                    for (int col = 4; col < 4 + headers.Length; col++)
                                    {
                                        if (Service1.ShrinkString(outputWorksheet.Cells[1, col].GetValue<string>()) == Service1.ShrinkString(description))
                                        {
                                            outputWorksheet.Cells[row2, col].Value = freqAmount;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (outputWorksheet.Cells[row2, 4].GetValue<double>() != 0.00)
                            {
                                outputWorksheet.Cells[row2, 15].Value = 1;
                                outputWorksheet.Cells[row2, 16].Value = 1;
                            }
                            row2++;
                        }
                        for (int col = 3; col <= 14; col++)
                        {
                            outputWorksheet.Column(col).Style.Numberformat.Format = "0.00";
                        }
                        for (int col = 1; col < headers.Length; col++)
                        {
                            string temp = Service1.ShrinkString(outputWorksheet.Cells[1, col].GetValue<string>());
                            if (temp.Contains("pf") || temp.Contains("provident"))
                            {
                                outputWorksheet.Cells[1, col].Value = "007 PF Employer";
                            }
                            if (temp.Contains("nps"))
                            {
                                outputWorksheet.Cells[1, col].Value = "009 Employer NPS";
                            }
                            if (temp.Contains("food"))
                            {
                                outputWorksheet.Cells[1, col].Value = "005 Food Allowance";
                            }
                            if (temp.Contains("basic"))
                            {
                                outputWorksheet.Cells[1, col].Value = "001 Basic";
                            }
                        }
                        outputWorksheet.Column(17).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Cells[1, 18].Value = "Kindly check with existing ctc structure.";
                        string newFileName = Path.Combine(destinationFolder,Service1.FileCount+ "]Leavers_CTC_Changes" + Path.GetFileName(filePath));
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
                #endregion
            }
            catch (Exception ex)
            {
                Service1.PathLog($"An error occurred: {ex.Message}");
            }
        }
    }
}