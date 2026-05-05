using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelAutomationService
{
    public class Benefeciaries_Data
    {
        public static void Beneficiaries_Data(string ascendcodes,string filePath, string destinationFolder)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    bool sheetExists = package.Workbook.Worksheets["Beneficiaries Data"] != null;
                    if (sheetExists){ 
                    int IP = Service1.getSheetNumber(filePath, "Beneficiaries Data");
                    var inputWorkSheet = package.Workbook.Worksheets[IP];// Assuming the data is in the first worksheet
                    int lastRow = inputWorkSheet.Dimension.End.Row;

                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Benificieries Data");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "Primary NameAsPerBank";
                        outputWorksheet.Cells[1, 3].Value = "Primary Bank A / c No";
                        outputWorksheet.Cells[1, 4].Value = "Primary IFSC";
                        outputWorksheet.Cells[1, 5].Value = "Primary Bank Code";
                        outputWorksheet.Cells[1, 6].Value = "Bank Name";
                        int hrid = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                        int bn = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Bank Name");
                        int benefeciaryname = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Beneficiary Name");
                        int acno = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "account number");
                        int ifsc = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "sort code");
                        int bic = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "BIC / SWIFT");
                        Dictionary<string, string> AscentBanksDetailed = new Dictionary<string, string>();
                        using (var package2 = new ExcelPackage(new FileInfo(ascendcodes)))
                        {
                            var BankSheet = package2.Workbook.Worksheets[Service1.getSheetNumber(ascendcodes, "Banks Detailed")];
                            int BankLastRow = BankSheet.Dimension.End.Row;
                            int description = Service1.getColumnNumber(ascendcodes, BankSheet.ToString(), "Name of Bank");
                            int code = Service1.getColumnNumber(ascendcodes, BankSheet.ToString(), "code");

                            for (int row2 = 2; row2 <= BankLastRow; row2++)
                            {
                                string bankname = BankSheet.Cells[row2, description].Text;
                                bankname = Service1.ShrinkString(bankname);
                                bankname = bankname.Replace("ltd", "");
                                bankname = bankname.Replace("limited", "");
                                bankname = bankname.Replace("pvt", "");
                                bankname = bankname.Replace(".", "");
                                bool containsBank = bankname.Contains("bank");
                                if (!containsBank)
                                {
                                    bankname = bankname + "bank";
                                }
                                if (!AscentBanksDetailed.ContainsKey(bankname))
                                    AscentBanksDetailed.Add(bankname, BankSheet.Cells[row2, code].Text);
                            }
                        }
                        int row3 = 2;
                        for (int row = 2; row <= lastRow; row++)
                        {
                            //string whiteColour = "16777215";
                            //string whiteColorHex = "FFFFFF";  // HEX representation of white color
                            var cell = inputWorkSheet.Cells[row, hrid]; // Example cell to check color
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            // Check if cell has a background color and compare it
                            if (!string.IsNullOrEmpty(bgColor.Rgb) || bgColor.Theme != null)
                            {
                                continue;
                            }
                            var HRID = inputWorkSheet.Cells[row, hrid].GetValue<string>();
                            var BENEFICIARYNAME = inputWorkSheet.Cells[row, benefeciaryname].Text;
                            var PrimaryBankAcNO = inputWorkSheet.Cells[row, acno].Text;
                            var IFSC = inputWorkSheet.Cells[row, ifsc].Text;
                            IFSC = Service1.ValidateIFSC(inputWorkSheet.ToString(), HRID, IFSC);
                            var BicSwift = inputWorkSheet.Cells[row, bic].Text;
                                if (IFSC=="") { 
                                    BicSwift = Service1.ValidateIFSC(inputWorkSheet.ToString(), HRID, BicSwift);
                                }
                                outputWorksheet.Cells[row3, 1].Value = HRID;
                            Regex validCharsRegex = new Regex("[^a-zA-Z ]");
                            BENEFICIARYNAME = validCharsRegex.Replace(BENEFICIARYNAME, "");
                            if (BENEFICIARYNAME == "")
                            {
                                Service1.PathLog(HRID + " comment:hrid's benefeciary name is not available.");
                            }
                            outputWorksheet.Cells[row3, 2].Value = BENEFICIARYNAME;
                            if ((PrimaryBankAcNO.All(char.IsDigit)))
                            {
                                outputWorksheet.Cells[row3, 3].Value = PrimaryBankAcNO;
                            }
                            if (IFSC != "")
                                outputWorksheet.Cells[row3, 4].Value = IFSC;
                            else
                                outputWorksheet.Cells[row3, 4].Value = BicSwift;

                            var bankname = inputWorkSheet.Cells[row, bn].GetValue<string>();
                            outputWorksheet.Cells[row3, 6].Value = bankname;
                            row3++;
                        }
                        row3 = 2;
                        int endRow = outputWorksheet.Dimension.End.Row;
                        using (var package2 = new ExcelPackage(new FileInfo(ascendcodes)))
                        {
                            int n = Service1.getSheetNumber(ascendcodes, "Banks Detailed");
                            var Ascendsheet = package2.Workbook.Worksheets[n];
                            int bankcode = Service1.getColumnNumber(ascendcodes, Ascendsheet.ToString(), "Code");
                            int Bankname = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Bank Name");
                            int bankname2 = Service1.getColumnNumber(ascendcodes, Ascendsheet.ToString(), "Name of Bank");
                            int lastRow2 = Ascendsheet.Dimension.End.Row;
                            for (int row = 2; row <= endRow; row++)
                            {
                                string bankname = outputWorksheet.Cells[row, 6].GetValue<string>();
                                bankname = Service1.ShrinkString(bankname);
                                bankname = bankname.Replace("ltd", "");
                                bankname = bankname.Replace("limited", "");
                                bankname = bankname.Replace("pvt", "");
                                bankname = bankname.Replace("branch", "");
                                bankname = bankname.Replace(".", "");
                                bool containsBank = bankname.Contains("bank");
                                if (!containsBank)
                                {
                                    bankname = bankname + "bank";
                                }
                                if (AscentBanksDetailed.ContainsKey(bankname))
                                {
                                    outputWorksheet.Cells[row, 5].Value = AscentBanksDetailed[bankname];
                                }
                            }
                            outputWorksheet.DeleteColumn(6);
                        }
                        string newFileName = Path.Combine(destinationFolder, Service1.FileCount + "]Beneficiaries Data_" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        string cellValue = outputWorksheet.Cells[2, 1].GetValue<string>();
                        Service1.ShrinkString(cellValue);
                        if ((cellValue != null) && (cellValue != " ") && (cellValue != " "))
                        {
                            outputPackage.SaveAs(newFileInfo);
                            outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                            Service1.Log("Beneficiaries data Created Successfully");
                            Service1.FileCount++;
                        }
                        else
                        {
                            Service1.PathLog("no existing benefeciaries file created");
                        }
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