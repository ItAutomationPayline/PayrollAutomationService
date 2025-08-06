using Microsoft.Office.Interop.Excel;
using NCalc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelAutomationService
{
    public class Variable
    {
        public static void Variable_Pay_Inputs_Data(string ascendcodes, string filePath, string destinationFolder)
        {
            try
            {
                string outputFilePath = Path.Combine(destinationFolder, "Variable_Pay_Summary.xlsx");
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int IP = Service1.getSheetNumber(filePath, "Variable Pay Inputs Data");
                    var inputWorkSheet = package.Workbook.Worksheets[IP];
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    // Get column numbers for relevant headers
                    int hridCol = Service1.getColumnNumber(filePath, inputWorkSheet.Name, "HR ID");
                    int comment= Service1.getColumnNumber(filePath, inputWorkSheet.Name, "Additional Comment");
                    int payElementCol = Service1.getColumnNumber(filePath, inputWorkSheet.Name, "Pay Element Short Code");
                    int amountCol = Service1.getColumnNumber(filePath, inputWorkSheet.Name, "Amount");

                    //Data structures to store unique pay elements and employee data
                    var employeeData = new Dictionary<string, Dictionary<string, double>>();
                    var payElementCodes = new HashSet<string>();
                    HashSet<string> NewHrid = new HashSet<string>();
                    List<string> CautionId = new List<string>();
                    List<string> CautionDesc = new List<string>();
                    List<double> CautionAmt = new List<double>();
                    List<string> CautionId2 = new List<string>();
                    List<string> CautionDesc2 = new List<string>();
                    List<double> CautionAmt2 = new List<double>();
                    List<string> CautionComment2 = new List<string>();
                    List<string> unitsnotfoundId = new List<string>();
                    List<string> unitsnotfoundDesc = new List<string>();
                    List<double> unitsnotfoundAmt = new List<double>();
                    // Read data from input sheet
                    for (int row = 2; row <= lastRow; row++)
                    {
                        var cell = inputWorkSheet.Cells[row, hridCol];
                        // Get the background color of the cell
                        var bgColor = cell.Style.Fill.BackgroundColor;
                        if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                        {
                            NewHrid.Add(cell.Text);
                        }
                        string hrid = inputWorkSheet.Cells[row, hridCol].GetValue<string>();
                        string payElement = inputWorkSheet.Cells[row, payElementCol].GetValue<string>();
                        string amountText = inputWorkSheet.Cells[row, amountCol].GetValue<string>();
                        double amount = double.TryParse(amountText, out var parsedAmount) ? parsedAmount : 0;
                        // Add pay element to the set
                        payElementCodes.Add(payElement);
                        // Add or update employee data
                        if (!employeeData.ContainsKey(hrid))
                        {
                            employeeData[hrid] = new Dictionary<string, double>();
                        }
                        if (!employeeData[hrid].ContainsKey(payElement))
                        {
                            employeeData[hrid][payElement] = 0;
                        }
                        employeeData[hrid][payElement] += amount;
                        if (inputWorkSheet.Cells[row, comment].Text != "" && !inputWorkSheet.Cells[row, comment].Text.Contains("Amount"))
                        {
                            CautionId2.Add(hrid);
                            CautionDesc2.Add(payElement);
                            CautionAmt2.Add(amount);
                            CautionComment2.Add(inputWorkSheet.Cells[row, comment].Text);
                        }
                    }
                    // Write the output file
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Summary");
                        // Write headers
                        outputWorksheet.Cells[1, 1].Value = "HR ID";
                        int colIndex = 2;
                        var payElementList = new List<string>(payElementCodes);
                        foreach (var payElement in payElementList)
                        {
                            outputWorksheet.Cells[1, colIndex].Value = payElement;
                            colIndex++;
                        }
                        // Write employee data
                        int rowIndex = 2;
                        foreach (var kvp in employeeData)
                        {
                            string hrid = kvp.Key;
                            outputWorksheet.Cells[rowIndex, 1].Value = hrid;

                            for (int i = 0; i < payElementList.Count; i++)
                            {
                                string payElement = payElementList[i];
                                double amount = kvp.Value.ContainsKey(payElement) ? kvp.Value[payElement] : 0;
                                outputWorksheet.Cells[rowIndex, i + 2].Value = amount;
                                
                                if (amount >= 200000.00 && NewHrid.Contains(hrid))
                                {
                                    CautionId.Add(hrid);
                                    CautionDesc.Add(payElement);
                                    CautionAmt.Add(amount);
                                    Service1.PathLog(hrid +" : Kindly confirm the amount: "+amount+" in variable file for a new joiner.");
                                    //string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ":Automation Alert: Amount in variable pay";
                                    //string body ="HRID:"+ hrid + " :<br>The amount: " + amount + " in variable sheet of client:"+ Path.GetFileName(filePath) + "<br>is high for a new joiner.<br>Please take necessary actions.<br><br>Regards,<br>Automation Team";
                                    //Service1.SendEmails(Service1.recipients, subject, body);
                                }
                            }
                            rowIndex++;
                        }
                        if (CautionId2.Count!=0) 
                        {
                            StringBuilder htmlTable = new StringBuilder();
                            htmlTable.Append("<table border='1' style='border-collapse: collapse;'>");
                            // Add table headers
                            htmlTable.Append("<tr>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>HRID</th>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>PayElement Code</th>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>Amount</th>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>Additional Comment</th>");
                            htmlTable.Append("</tr>");
                            Service1.PathLog("Below are the Additional comments of variable file:");
                            // Add table rows
                            for (int row8 = 0; row8 <= CautionId2.Count - 1; row8++)
                            {
                                Service1.PathLog("HRID:"+ CautionId2[row8]+" Description:"+ CautionDesc2[row8]+" Amount:"+ CautionAmt2[row8]+" Comment:"+ CautionComment2[row8]);
                                htmlTable.Append("<tr>");
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionId2[row8]);
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionDesc2[row8]);
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionAmt2[row8].ToString("N2"));
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionComment2[row8]);
                                htmlTable.Append("</tr>");
                            }
                            htmlTable.Append("</table>");
                            string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alert";
                            string body = "Below are the Additional Comments in Variable File of Client:" + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br>";
                            if (Service1.subject == "")
                            {
                                Service1.subject = subject;
                            }
                            Service1.body = Service1.body + body;
                        }
                        if (CautionId.Count != 0)
                        {
                            StringBuilder htmlTable = new StringBuilder();
                            htmlTable.Append("<table border='1' style='border-collapse: collapse;'>");
                            // Add table headers
                            htmlTable.Append("<tr>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>HRID</th>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>PayElement Code</th>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>Amount</th>");
                            htmlTable.Append("</tr>");
                            // Add table rows
                            for (int row8 = 0; row8 <= CautionId.Count - 1; row8++)
                            {
                                htmlTable.Append("<tr>");
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionId[row8]);
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionDesc[row8]);
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionAmt[row8].ToString("N2"));
                                htmlTable.Append("</tr>");
                            }
                            htmlTable.Append("</table>");
                            string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alert";
                            string body = "Kindly confirm the variable pay amounts for a new joinner in input file: " + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br>";
                            if (Service1.subject == "")
                            {
                                Service1.subject = subject;
                            }
                            Service1.body = Service1.body + body;
                            //Service1.SendEmails(Service1.recipients, subject, body);
                            //QuerySheet.VariableAmountQuery(destinationFolder, CautionId, CautionDesc, CautionAmt);
                        }
                        using (var ascentcodes = new ExcelPackage(new FileInfo(ascendcodes)))
                        {
                            bool sheetExists = ascentcodes.Workbook.Worksheets["Variable Units Calculation"] != null;
                            if (sheetExists)
                            {
                                Service1.FetchEmployeeMaster();
                            }
                        }
                        for (int column = 1; column <= payElementCodes.Count + 1; column++)
                        {
                            outputWorksheet.Column(column).Style.Numberformat.Format = "0.00";
                            // Get the header value of the current column
                            string temp = outputWorksheet.Cells[1, column].GetValue<string>(); // Correctly reference the column header
                            temp = Service1.ShrinkString(temp);
                            // Check for specific keywords
                            bool containsEncashment = temp.Contains("encashment");
                            bool containsHoliday = temp.Contains("holiday");
                            bool containsOvertime = temp.Contains("overtime");
                            bool containsShift = temp.Contains("shift");
                            bool extrahourspay = temp.Contains("extrahours");
                            bool severencepay = temp.Contains("severance");
                            bool noticepayout = temp.Contains("noticepayout");
                            bool ot = temp.Contains("otamount");
                            if (containsEncashment || containsHoliday || containsOvertime || containsShift|| extrahourspay || severencepay || noticepayout||ot)
                            {
                                int OutputLastRow = outputWorksheet.Dimension.End.Row;
                                string formula="";
                                using (var ascentcodes = new ExcelPackage(new FileInfo(ascendcodes)))
                                {
                                    bool sheetExists = ascentcodes.Workbook.Worksheets["Variable Units Calculation"] != null;
                                    if (sheetExists){
                                        var unitsheet = ascentcodes.Workbook.Worksheets["Variable Units Calculation"];
                                        int endrow = unitsheet.Dimension.End.Row;
                                        for (int row4 = 2;row4<=endrow;row4++){
                                            if (Service1.ShrinkString(unitsheet.Cells[row4,1].Text)==temp){
                                                Service1.PathLog(unitsheet.Cells[row4, 2].Text);
                                                formula = unitsheet.Cells[row4, 2].Text;
                                            }
                                        }
                                        for (int row2=2;row2<=OutputLastRow;row2++){
                                            if (outputWorksheet.Cells[row2,column].GetValue<double>()!=0 && formula!="")          
                                            {
                                                try
                                                {
                                                    Expression ex = new Expression(formula);
                                                    string d = outputWorksheet.Cells[row2, column].GetValue<string>();
                                                    ex.Parameters["value"] = outputWorksheet.Cells[row2, column].GetValue<double>();
                                                    ex.Parameters["previousmonthtotaldays"] = Service1.GetPreviousMonthTotalDays();
                                                    if (formula.Contains("previousmonthannualctc")) 
                                                    { 
                                                        double ctc= Service1.GetPreviousMonthAnnualCTC(outputWorksheet.Cells[row2, 1].Text);
                                                        ex.Parameters["previousmonthannualctc"] = ctc;
                                                        if (ctc == 1)
                                                        {
                                                            int targetRow = row2; // Example row
                                                            var rowRange = outputWorksheet.Cells[targetRow, 1, targetRow, 99];
                                                            unitsnotfoundId.Add(outputWorksheet.Cells[row2, 1].Text);
                                                            unitsnotfoundDesc.Add(outputWorksheet.Cells[1, column].Text);
                                                            unitsnotfoundAmt.Add(outputWorksheet.Cells[row2, column].GetValue<double>());
                                                            rowRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                            rowRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                                        }
                                                    }
                                                    if (formula.Contains("previousmonthhbctc"))
                                                    {
                                                        double ctc = Service1.GetPreviousMonthHBCTC(outputWorksheet.Cells[row2, 1].Text);
                                                        ex.Parameters["previousmonthhbctc"] = ctc;
                                                        if (ctc == 1)
                                                        {
                                                            int targetRow = row2; // Example row
                                                            var rowRange = outputWorksheet.Cells[targetRow, 1, targetRow, outputWorksheet.Dimension.End.Row];
                                                            unitsnotfoundId.Add(outputWorksheet.Cells[row2, 1].Text);
                                                            unitsnotfoundDesc.Add(outputWorksheet.Cells[1, column].Text);
                                                            unitsnotfoundAmt.Add(outputWorksheet.Cells[row2, column].GetValue<double>());
                                                            rowRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                            rowRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                                        }
                                                    }
                                                    if (formula.Contains("previousmonthymediactc"))
                                                    {
                                                        double ctc = Service1.GetPreviousMonthymediaCTC(outputWorksheet.Cells[row2, 1].Text);
                                                        ex.Parameters["previousmonthymediactc"] = ctc;
                                                        if (ctc == 1)
                                                        {
                                                            int targetRow = row2; // Example row
                                                            var rowRange = outputWorksheet.Cells[targetRow, 1, targetRow, outputWorksheet.Dimension.End.Row];
                                                            unitsnotfoundId.Add(outputWorksheet.Cells[row2, 1].Text);
                                                            unitsnotfoundDesc.Add(outputWorksheet.Cells[1, column].Text);
                                                            unitsnotfoundAmt.Add(outputWorksheet.Cells[row2, column].GetValue<double>());
                                                            rowRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                            rowRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                                        }
                                                    }
                                                    outputWorksheet.Cells[row2, column].Value = ex.Evaluate();
                                                }
                                                catch (Exception e)
                                                {
                                                    Service1.PathLog(e.ToString());
                                                    throw;
                                                }
                                            }
                                        }
                                    }
                                Service1.PathLog("check for encashment/holiday/Overtime/shift/extrahours is in units or amount in variable file.");
                                
                                // Define the range for the entire column
                                var columnRange = outputWorksheet.Cells[1, column, OutputLastRow, column];
                                // Apply fill color
                                columnRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                columnRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                }
                            }
                            if (containsEncashment && filePath.ToLower().Contains("twilio"))
                            {
                                int OutputLastRow = outputWorksheet.Dimension.End.Row;

                                using (var newPackage = new ExcelPackage())
                                {
                                    var newWorksheet = newPackage.Workbook.Worksheets.Add("EncashmentData");

                                    // Copy header
                                    newWorksheet.Cells[1, 1].Value = outputWorksheet.Cells[1, 1].Value;
                                    newWorksheet.Cells[1, 2].Value = outputWorksheet.Cells[1, column].Value;

                                    // Copy data
                                    for (int row = 2; row <= OutputLastRow; row++)
                                    {
                                        newWorksheet.Cells[row, 1].Value = outputWorksheet.Cells[row, 1].Value;
                                        newWorksheet.Cells[row, 2].Value = outputWorksheet.Cells[row, column].Value;
                                    }

                                    // Optional: format
                                    newWorksheet.Column(1).AutoFit();
                                    newWorksheet.Column(2).AutoFit();

                                    // Save file
                                    string encashmentFilePath = Path.Combine(destinationFolder, $"Variable EncashmentData_{Path.GetFileName(filePath)}");
                                    newPackage.SaveAs(new FileInfo(encashmentFilePath));
                                    Service1.PathLog("Encashment data file created: " + encashmentFilePath);
                                }
                            }
                        }
                        if (unitsnotfoundId.Count != 0)
                        {
                            StringBuilder htmlTable = new StringBuilder();
                            htmlTable.Append("<table border='1' style='border-collapse: collapse;'>");
                            // Add table headers
                            htmlTable.Append("<tr>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>HRID</th>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>PayElement Code</th>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>Amount</th>");
                            htmlTable.Append("</tr>");
                            // Add table rows
                            for (int row8 = 0; row8 <= unitsnotfoundId.Count - 1; row8++)
                            {
                                htmlTable.Append("<tr>");
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", unitsnotfoundId[row8]);
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", unitsnotfoundDesc[row8]);
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", unitsnotfoundAmt[row8].ToString("N2"));
                                htmlTable.Append("</tr>");
                            }
                            htmlTable.Append("</table>");
                            string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alert";
                            string body = "Below are employees which were not found in employee master for Units conversion.<br>Ensure that previous month employee master is pasted in Employee Master folder.<br> Kindly ignore if these are new joinners.<br><br>"+ htmlTable.ToString()+"<br>";
                            if (Service1.subject == "")
                            {
                                Service1.subject = subject;
                            }
                            Service1.body = Service1.body + body;
                            //Service1.SendEmails(Service1.recipients, subject, body);
                            //QuerySheet.VariableAmountQuery(destinationFolder, CautionId, CautionDesc, CautionAmt);
                        }
                        // Save output file
                        string newFileName = Path.Combine(destinationFolder,Service1.FileCount+ "]Variable_" + Path.GetFileName(filePath));
                        // outputPackage.SaveAs(new FileInfo(outputFilePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        Service1.EmployeeMaster.Clear();
                        string cellValue = outputWorksheet.Cells[2, 1].GetValue<string>();
                        Service1.ShrinkString(cellValue);
                        if ((cellValue != null) && (cellValue != " ") && (cellValue != " "))
                        {
                            outputPackage.SaveAs(newFileInfo);
                            outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                            Service1.Log("Variable Created Successfully");
                            Service1.FileCount++;
                        }
                        else
                        {
                            Service1.PathLog("no variable file created");
                        }
                    }
                }
                //Service1.Log($"Variable Excel file created successfully at {outputFilePath}!");
            }
            catch (Exception ex)
            {
                Service1.ErrorCount++;
                Service1.Log($"An error occurred: {ex.Message}");
            }
        }
    }
}