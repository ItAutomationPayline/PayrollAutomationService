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
            try{
                string basicnotation="";
                //string vlookuppath = @"E:\PAYROLL_SERVER\Automation\Config\" + "["+ Path.GetFileName(filePath)+"]";
                //Service1.PathLog("VlookupPath:"+vlookuppath);
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int jcsheet= Service1.getSheetNumber(filePath, "Joiner and Changes ");
                    var joinerandchangessheet= package.Workbook.Worksheets[jcsheet];
                    int hrid2 = Service1.getColumnNumber(filePath, joinerandchangessheet.ToString(), "HR ID");
                    int designation = Service1.getColumnNumber(filePath, joinerandchangessheet.ToString(), "Job Title");
                    int ptloc = Service1.getColumnNumber(filePath, joinerandchangessheet.ToString(), " PT Location");
                    int IP = Service1.getSheetNumber(filePath, "Payments and Deductions");
                    var inputWorkSheet = package.Workbook.Worksheets[IP];
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int lastRow2 = joinerandchangessheet.Dimension.End.Row;
                    int hrid = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                    int payelementdescription = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Element Description");
                    int payelementshortcode = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Element Short Code");
                    int witheffectfrom = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Start Date");
                    int payfreq = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Frequency");
                    int frequencyamount = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Frequency Amount");
                    int amount = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Amount");
                    string ctccolumnname, ctcpayfreq, ctcamountcolumn;
                    double maxwage =0;
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Ctc_New_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "With effect From(YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 3].Value = "Annual CTC";
                        using (var CTCPackage = new ExcelPackage(ctccodes))
                        {
                            var ctcsheet = CTCPackage.Workbook.Worksheets[0];
                            int ctclength = ctcsheet.Dimension.End.Row;
                            int row = 2;
                            int coll = 0;
                            ctccolumnname = ctcsheet.Cells[2, 2].Text;
                            ctcpayfreq = ctcsheet.Cells[2, 3].Text;
                            ctcamountcolumn = ctcsheet.Cells[2, 4].Text;
                            outputWorksheet.Cells[1, 3].Value = ctcsheet.Cells[2, 1].Text;
                            for (coll = 3; coll <= ctclength + 1; coll++)
                            {
                                outputWorksheet.Cells[1, coll].Value = ctcsheet.Cells[row, 1].Text;
                                //if (ctcsheet.Cells[2, coll].Text == ""|| ctcsheet.Cells[2, coll].Text == "0")
                                //    outputWorksheet.Cells[2, coll].Formula = ctcsheet.Cells[row, 2].Formula;
                                if (!string.IsNullOrEmpty(ctcsheet.Cells[row, 2].Formula))
                                {
                                    if (ctcsheet.Cells[row, 2].Formula.Contains("VLOOKUP"))
                                    {
                                        outputWorksheet.Cells[2, coll].Formula = @"=VLOOKUP(A2,'E:\PAYROLL_SERVER\Automation\Config\[1temp.xlsx]Joiner and Changes '!$B$2:$AR$29,43,0)";
                                        //outputWorksheet.Cells[2, coll].Formula = ctcsheet.Cells[row, 2].Formula.Replace(@"E:\PAYROLL_SERVER\Automation\Input\[McAfee_Mcafee Software (India) Pvt Ltd._India_February_20250211_Payroll_Data.xlsx]", Path.GetFileName(filePath));
                                        //Service1.PathLog(outputWorksheet.Cells[2, coll].Formula);
                                    }
                                    else
                                    {
                                        outputWorksheet.Cells[2, coll].Formula = ctcsheet.Cells[row, 2].Formula;
                                    }

                                }
                                else
                                {
                                    outputWorksheet.Cells[2, coll].Value = ctcsheet.Cells[row, 2].Text;
                                }
                                row++;
                            }
                            if (outputWorksheet.Cells[2, 4].Text.Contains("fixed"))
                            {
                                basicnotation = outputWorksheet.Cells[2, 4].Text;
                            }

                        }

                        HashSet<string> HRID = new HashSet<string>();
                        Dictionary<string, string> PtLoc = new Dictionary<string, string>();
                        Dictionary<string, string> Role = new Dictionary<string, string>();
                        Dictionary<string, Double> MinimumWages = new Dictionary<string, Double>();
                        Dictionary<string, Double> KarnatakaMinimumWages = new Dictionary<string, Double>();
                        if (Directory.Exists(Service1.MinimumWagesAct))
                        {
                            string[] referencefile = Directory.GetFiles((Service1.MinimumWagesAct), "*.xlsx");
                            string WagesFile = Service1.MinimumWagesAct + "/" + Path.GetFileName(referencefile[0]);
                            using (var MinimumWagesfile = new ExcelPackage(WagesFile))
                            {
                                var wagessheet = MinimumWagesfile.Workbook.Worksheets[0];
                                var karnatakawagessheet = MinimumWagesfile.Workbook.Worksheets[1];
                                int endRow = wagessheet.Dimension.End.Row;
                                int endCol = wagessheet.Dimension.End.Column;
                                for (int row = 2; row <= endRow; row++)
                                {
                                    MinimumWages.Add(Service1.ShrinkString(wagessheet.Cells[row, 1].Text), wagessheet.Cells[row, 2].GetValue<double>());
                                    if (maxwage < wagessheet.Cells[row, 2].GetValue<double>())
                                    {
                                        maxwage = wagessheet.Cells[row, 2].GetValue<double>();
                                    }
                                }
                                endRow = karnatakawagessheet.Dimension.End.Row;
                                endCol = karnatakawagessheet.Dimension.End.Column;
                                for (int row = 2; row <= endRow; row++)
                                {
                                    KarnatakaMinimumWages.Add(Service1.ShrinkString(karnatakawagessheet.Cells[row, 1].Text), karnatakawagessheet.Cells[row, 2].GetValue<double>());
                                    if (maxwage < karnatakawagessheet.Cells[row, 2].GetValue<double>())
                                    {
                                        maxwage = karnatakawagessheet.Cells[row, 2].GetValue<double>();
                                    }
                                }
                            }
                        }
                        for (int row = 2; row <= lastRow2; row++)
                        {
                            var cell = joinerandchangessheet.Cells[row, hrid2];
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF") && (!joinerandchangessheet.Cells[row, designation].Text.ToLower().Contains("intern") && !joinerandchangessheet.Cells[row, designation].Text.ToLower().Contains("trainee")))
                            {
                                HRID.Add(joinerandchangessheet.Cells[row, hrid2].GetValue<string>());
                                PtLoc.Add(joinerandchangessheet.Cells[row, hrid2].GetValue<string>(), Service1.ShrinkString(joinerandchangessheet.Cells[row, ptloc].GetValue<string>()));
                                Role.Add(joinerandchangessheet.Cells[row, hrid2].GetValue<string>(), Service1.ShrinkString(joinerandchangessheet.Cells[row, designation].GetValue<string>()));
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
                                    if (Service1.ShrinkString(inputWorkSheet.Cells[row, payelementshortcode].GetValue<string>()).Equals(Service1.ShrinkString(ctccolumnname)) && Service1.ShrinkString(inputWorkSheet.Cells[row, payfreq].GetValue<string>()) == Service1.ShrinkString(ctcpayfreq))
                                    {
                                        if (Service1.ShrinkString(ctcamountcolumn) == "amount")
                                        {
                                            outputWorksheet.Cells[row2, 3].Value = inputWorkSheet.Cells[row, amount].GetValue<double>();
                                        }
                                        if (Service1.ShrinkString(ctcamountcolumn) == "frequencyamount")
                                        {
                                            outputWorksheet.Cells[row2, 3].Value = inputWorkSheet.Cells[row, frequencyamount].GetValue<double>();
                                        }
                                    }
                                }
                            }
                            row2++;
                        }
                        int outputrows = outputWorksheet.Dimension.End.Row;
                        int outputcolumns = outputWorksheet.Dimension.End.Column;
                        int col = 4;

                        for (col = 3; col <= outputcolumns; col++)
                        {
                            // Get the base formula from row 2 (template)
                            var baseFormula = outputWorksheet.Cells[2, col].Formula;
                            for (row2 = 3; row2 <= outputrows; row2++)
                            {
                                if (!string.IsNullOrWhiteSpace(baseFormula))
                                {
                                    //Service1.PathLog("Formula executed for " + outputWorksheet.Cells[1, col].Text + " Column.");
                                    // Adjust the base formula to the current row
                                    string adjustedFormula = AdjustFormulaToRow(baseFormula, 2, row2);

                                    outputWorksheet.Cells[row2, col].Formula = adjustedFormula;
                                }
                                else if (outputWorksheet.Cells[2, col].Text.ToLower().Contains("pickup"))
                                {
                                    //Service1.PathLog("Pickup executed for " + outputWorksheet.Cells[1, col].Text + " Column.");
                                    for (int r = 3; r <= outputrows; r++)
                                    {
                                        outputWorksheet.Cells[r, col].Value = outputWorksheet.Cells[2, col].Text;
                                    }
                                }
                                //if (outputWorksheet.Cells[row2, col].Text.ToLower().Contains("pickup")){
                                //    outputWorksheet.Cells[row2, col].Value = PickupValueByEmployeeId(outputWorksheet.Cells[row2, 1].Text, outputWorksheet.Cells[row2, col].Text);
                                //}
                                else
                                {
                                    var value = outputWorksheet.Cells[2, col].Text;
                                    if (value.ToLower().Contains("pickup"))
                                    {
                                        outputWorksheet.Cells[row2, col].Value = PickupValueByEmployeeId(outputWorksheet.Cells[row2, 1].Text, outputWorksheet.Cells[row2, col].Text);
                                    }
                                    else
                                    {
                                        if (value.All(char.IsDigit) && outputWorksheet.Cells[row2, col].Text == "")
                                            outputWorksheet.Cells[row2, col].Value = value;
                                    }
                                }
                                //else
                                //{
                                //    var value = outputWorksheet.Cells[2, col].Text;
                                //    if (value.All(char.IsDigit)&& !string.IsNullOrWhiteSpace(outputWorksheet.Cells[2,col].Formula))
                                //    {
                                //        Service1.PathLog("As it is value used for " + outputWorksheet.Cells[1, col].Text + " Column.");
                                //        outputWorksheet.Cells[row2, col].Value = value;
                                //    }
                                //    if (value.ToLower().Contains("pickup"))
                                //    {
                                //        outputWorksheet.Cells[row2, col].Value = PickupValueByEmployeeId(outputWorksheet.Cells[row2, 1].Text, outputWorksheet.Cells[row2, col].Text);
                                //    }
                                //}
                            }
                            //else 
                            //{
                            //    var value = outputWorksheet.Cells[2, col].Text;
                            //    if (value.All(char.IsDigit)){
                            //        outputWorksheet.Cells[row2, col].Value = value;
                            //    }
                            //}
                        }
                        for (col = 4; col <= outputcolumns; col++)
                        {
                            var baseText = outputWorksheet.Cells[2, col].Text;
                            for (row2 = 2; row2 <= outputrows; row2++)
                            {
                                if (outputWorksheet.Cells[row2, col].Text.ToLower().Contains("pickup"))
                                {
                                    outputWorksheet.Cells[row2, col].Value = PickupValueByEmployeeId(outputWorksheet.Cells[row2, 1].Text, outputWorksheet.Cells[row2, col].Text);
                                }
                            }
                        }
                        string PickupValueByEmployeeId(string employeeId, string pickupString)
                        {
                            // Parse the pickup string
                            var parts = pickupString.Split(new[] { "[", "]" }, StringSplitOptions.RemoveEmptyEntries);
                            //Service1.PathLog(parts[0]+" " + parts[1] + " " + parts[2] + " " + parts[3]);
                            if (parts.Length != 4)
                                throw new ArgumentException("Invalid pickup string format. Expected: pickup[filePath][sheetName][columnName]");

                            string fP = parts[1];
                            if (fP.ToLower() == "input")
                            {
                                fP = filePath;
                            }
                            string sheetName = parts[2];
                            string columnName = parts[3];

                            if (!File.Exists(filePath))
                            {
                                Service1.PathLog("Input file not found" + filePath);
                                throw new FileNotFoundException("Input file not found", filePath);
                            }

                            using (var vlookuppackage = new ExcelPackage(new FileInfo(filePath)))
                            {
                                var worksheet = vlookuppackage.Workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                                if (worksheet == null)
                                {
                                    Service1.PathLog($"Sheet '{sheetName}' not found.");
                                    throw new ArgumentException($"Sheet '{sheetName}' not found.");
                                }

                                int startRow = worksheet.Dimension.Start.Row;
                                int endRow = worksheet.Dimension.End.Row;
                                int startCol = worksheet.Dimension.Start.Column;
                                int endCol = worksheet.Dimension.End.Column;

                                // Find column indexes for "Employee ID" and target column
                                int empIdCol = -1, targetCol = -1;
                                empIdCol = Service1.getColumnNumber(filePath, worksheet.ToString(), "HR ID");
                                targetCol = Service1.getColumnNumber(filePath, worksheet.ToString(), columnName);
                                //for (int col2 = startCol; col2 <= endCol; col2++)
                                //{
                                //    string header = worksheet.Cells[startRow, col2].Text.Trim();
                                //    if (header.Equals("HR ID", StringComparison.OrdinalIgnoreCase))
                                //        empIdCol = col2;
                                //    if (header.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                                //        targetCol = col2;
                                //}

                                //if (empIdCol == -1) {
                                //    Service1.PathLog("'Employee ID' column not found.");
                                //    throw new ArgumentException("'Employee ID' column not found.");
                                //}
                                //if (targetCol == -1) {
                                //    Service1.PathLog($"Column '{columnName}' not found.");
                                //    throw new ArgumentException($"Column '{columnName}' not found.");
                                //}

                                // Find the matching employee ID row
                                for (int row = startRow + 1; row <= endRow; row++)
                                {
                                    string currentEmpId = worksheet.Cells[row, empIdCol].Text.Trim();
                                    if (currentEmpId.Equals(employeeId, StringComparison.OrdinalIgnoreCase))
                                    {
                                        return worksheet.Cells[row, targetCol].Text;
                                    }
                                }

                                return null; // Employee ID not found
                            }
                        }
                        string AdjustFormulaToRow(string formula, int baseRow, int targetRow)
                        {
                            if ((formula.Replace(".", "")).All(char.IsDigit))
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
                        int outputRows = outputWorksheet.Dimension.End.Row;
                        for (col = 2; col <= outputcolumns; col++)
                        {
                            if ((outputWorksheet.Cells[2, col].Text.ToLower() == "fixedamount" || outputWorksheet.Cells[2, col].Text.ToLower() == "fixed") && outputWorksheet.Cells[1, col].Text != "")
                            {
                                outputWorksheet.Cells[2, col].Value = 0;
                                for (int row3 = 2; row3 <= lastRow; row3++)
                                {
                                    for (int row = 2; row <= lastRow; row++)
                                    {
                                        if ((Service1.ShrinkString(inputWorkSheet.Cells[row, hrid].Text) == Service1.ShrinkString(outputWorksheet.Cells[row3, 1].Text) && (Service1.ShrinkString(inputWorkSheet.Cells[row, payelementshortcode].Text) == Service1.ShrinkString(outputWorksheet.Cells[1, col].Text))))
                                        {
                                            outputWorksheet.Cells[row3, col].Value = inputWorkSheet.Cells[row, amount].GetValue<double>();
                                        }
                                    }
                                }
                            }
                            if ((outputWorksheet.Cells[2, col].Text.ToLower() == "fixedfrequencyamount") && outputWorksheet.Cells[1, col].Text != "")
                            {
                                outputWorksheet.Cells[2, col].Value = 0;
                                for (int row3 = 2; row3 <= lastRow; row3++)
                                {
                                    for (int row = 2; row <= lastRow; row++)
                                    {
                                        if ((Service1.ShrinkString(inputWorkSheet.Cells[row, hrid].Text) == Service1.ShrinkString(outputWorksheet.Cells[row3, 1].Text) && (Service1.ShrinkString(inputWorkSheet.Cells[row, payelementshortcode].Text) == Service1.ShrinkString(outputWorksheet.Cells[1, col].Text))))
                                        {
                                            outputWorksheet.Cells[row3, col].Value = inputWorkSheet.Cells[row, frequencyamount].GetValue<double>();
                                        }
                                    }
                                    //if (outputWorksheet.Cells[row2, col].Text == "" && outputWorksheet.Cells[row2, 1].Text != "")
                                    //{
                                    //    outputWorksheet.Cells[row2, col].Value = 0;

                                    //}
                                }
                                //Service1.PathLog("Got fixed");
                                //for (int row3 = 2; row3 <= lastRow; row3++)
                                //{
                                //    for (int row = 2; row <= lastRow; row++)
                                //    {
                                //        if ((Service1.ShrinkString(inputWorkSheet.Cells[row, hrid].Text) == Service1.ShrinkString(outputWorksheet.Cells[row2, 1].Text) && (Service1.ShrinkString(inputWorkSheet.Cells[row, payelementshortcode].Text) == Service1.ShrinkString(outputWorksheet.Cells[1, col].Text))))
                                //        {
                                //            outputWorksheet.Cells[row2, col].Value = inputWorkSheet.Cells[row, frequencyamount].GetValue<double>();
                                //            row2++;
                                //        }
                                //    }
                                //    if (outputWorksheet.Cells[row2, col].Text == "" && outputWorksheet.Cells[row2, 1].Text != "")
                                //    {
                                //        outputWorksheet.Cells[row2, col].Value = 0;
                                //        row2++;
                                //    }

                                //}
                            }
                            row2 = 2;
                            for (row2 = 2; row2 <= outputRows; row2++)
                            {
                                if (outputWorksheet.Cells[row2, col].Text == "" && string.IsNullOrEmpty(outputWorksheet.Cells[row2, col].Formula))
                                    outputWorksheet.Cells[row2, col].Value = 0;
                            }
                        }
                        if (Directory.Exists(Service1.MinimumWagesAct))
                        {
                            for (int row = 2; row <= outputrows; row++)
                            {
                                string empid = outputWorksheet.Cells[row, 1].GetValue<string>();
                                double basic = outputWorksheet.Cells[row, 4].GetValue<double>();
                                double basic2 = (outputWorksheet.Cells[row, 3].GetValue<double>() / 12) * 40 / 100;
                                if (string.IsNullOrWhiteSpace(empid) || !PtLoc.ContainsKey(empid))
                                    continue; // Skip if HRID is missing or not in PtLoc
                                double minWage = 0;
                                string location = Service1.ShrinkString(PtLoc[empid]);
                                string jobtitle = Service1.ShrinkString(Role[empid]);
                                Service1.PathLog("Maxwage:" + maxwage);
                                //note below.
                                if (Service1.ShrinkString(location).Contains("karnataka") || Service1.ShrinkString(location).Contains("bengaluru") || Service1.ShrinkString(location).Contains("bangalore"))
                                {
                                    if (maxwage < basic || (maxwage < basic2))
                                    {
                                        continue;
                                    }
                                    bool containsdesignation = KarnatakaMinimumWages.Keys.Any(k => jobtitle.Contains(k));
                                    if (!containsdesignation)
                                    {
                                        if (maxwage > basic || maxwage > basic2) 
                                        {
                                            Service1.PathLog("Designation not found in karnataka wages:" + jobtitle);
                                            var rowRange = outputWorksheet.Cells[row, 1, row, outputWorksheet.Dimension.End.Column];
                                            rowRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            rowRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                            Service1.PathLog("HRID:" + empid + " Kindly confirm the CTC structre of new joiner is as per Minimum Wages Act or not.");
                                            continue; // Skip if location not found in MinimumWages
                                        }
                                        continue;
                                    }
                                    string matcheddesignation = KarnatakaMinimumWages.Where(kvp => jobtitle.Contains(kvp.Key)).Select(kvp => kvp.Key).FirstOrDefault();  // returns null if not found
                                    minWage = KarnatakaMinimumWages[matcheddesignation];
                                    if (basicnotation.Contains("fixed"))
                                    {
                                        Service1.PathLog("FIXED EXECUTED");
                                        Service1.PathLog("BASICNOTATION:" + basicnotation);
                                        basic = outputWorksheet.Cells[row, 4].GetValue<double>();
                                        if (basic < minWage)
                                        {
                                            var rowRange = outputWorksheet.Cells[row, 1, row, outputWorksheet.Dimension.End.Column];
                                            rowRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            rowRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                            Service1.PathLog("HRID:" + empid + " Kindly confirm the CTC structre of new joiner is as per Minimum Wages Act or not.");
                                        }
                                    }
                                    else
                                    {
                                        basic2 = (outputWorksheet.Cells[row, 3].GetValue<double>() / 12) * 40 / 100;
                                        if (basic2 < minWage)
                                        {
                                            outputWorksheet.Cells[row, 4].Value = minWage;
                                            var rowRange = outputWorksheet.Cells[row, 1, row, outputWorksheet.Dimension.End.Column];
                                            rowRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            rowRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.GreenYellow);
                                            Service1.PathLog("HRID:" + empid + " Kindly confirm updated CTC structre by Automation of new joiner is as per Minimum Wages Act or not.");
                                        }
                                    }
                                }
                                else
                                {
                                    if (maxwage < basic || (maxwage < basic2))
                                    {
                                        continue;
                                    }
                                    if (!MinimumWages.ContainsKey(location))
                                    {
                                        if (maxwage > basic || maxwage > basic2) 
                                        {
                                            var rowRange = outputWorksheet.Cells[row, 1, row, outputWorksheet.Dimension.End.Column];
                                            rowRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            rowRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                            Service1.PathLog("HRID:" + empid + " Kindly confirm the CTC structre of new joiner is as per Minimum Wages Act or not.");
                                            continue; // Skip if location not found in MinimumWages
                                        }
                                    }
                                    minWage = MinimumWages[location];
                                    if (basicnotation.Contains("fixed"))
                                    {
                                        Service1.PathLog("FIXED EXECUTED");
                                        Service1.PathLog("BASICNOTATION:" + basicnotation);
                                        basic = outputWorksheet.Cells[row, 4].GetValue<double>();
                                        if (basic < minWage)
                                        {
                                            var rowRange = outputWorksheet.Cells[row, 1, row, outputWorksheet.Dimension.End.Column];
                                            rowRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            rowRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                            Service1.PathLog("HRID:" + empid + " Kindly confirm the CTC structre of new joiner is as per Minimum Wages Act or not.");
                                        }
                                    }
                                    else
                                    {
                                        basic2 = (outputWorksheet.Cells[row, 3].GetValue<double>() / 12) * 40 / 100;
                                        if (basic2 < minWage)
                                        {
                                            outputWorksheet.Cells[row, 4].Value = minWage;
                                            var rowRange = outputWorksheet.Cells[row, 1, row, outputWorksheet.Dimension.End.Column];
                                            rowRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            rowRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.GreenYellow);
                                            Service1.PathLog("HRID:" + empid + " Kindly confirm updated CTC structre by Automation of new joiner is as per Minimum Wages Act or not.");
                                        }
                                    }
                                }
                            }
                        }
                        //outputWorksheet.Column(17).Style.Numberformat.Format = "0.00";
                        string newFileName = Path.Combine(destinationFolder, Service1.FileCount + "]NEW_Joiners_CTC" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        Service1.PathLog("CTCFILENAME:"+newFileName);
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