using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using OfficeOpenXml;

namespace ExcelAutomationService
{
    public class Musarubra_CTC
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
                    int empgrade = Service1.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "Employee Grade");
                    int hrid = Service1.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "Hr id");
                    // int payelementdescription = Service1.getColumnNumber(filePath, joinerandChangesSheet.ToString(), "Pay Element Short Code");
                    int witheffectfrom = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Start Date");
                    int annualctc = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Amount");
                    int frequencyamount = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Frequency Amount");
                    int payfreq = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Pay Frequency");

                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Ctc_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "With effect From(YYYY - MM - DD)";
                        outputWorksheet.Cells[1, 3].Value = "Annual CTC";
                        outputWorksheet.Cells[1, 4].Value = "001 Basic";
                        outputWorksheet.Cells[1, 5].Value = "Location";
                        outputWorksheet.Cells[1, 6].Value = "002 Residual Pay";
                        outputWorksheet.Cells[1, 7].Value = "003 Special Allowance";
                        outputWorksheet.Cells[1, 8].Value = "004 House Rent Allowance";
                        outputWorksheet.Cells[1, 9].Value = "005 Leave Travel Allowance";
                        outputWorksheet.Cells[1, 10].Value = "006 Stipend";
                        outputWorksheet.Cells[1, 11].Value = "007 Employer NPS";
                        outputWorksheet.Cells[1, 12].Value = "010 Meal Allowance";
                        outputWorksheet.Cells[1, 13].Value = "Employee Grade";
                        HashSet<string> HRID = new HashSet<string>();
                        List<string> ptloc = new List<string>();
                        List<string> EmpGrade = new List<string>();

                        int row2 = 2;
                        for (int row = 2; row <= lastRow2; row++)
                        {
                            var cell = joinerandChangesSheet.Cells[row, hrid];
                            // Get the background color of the cell
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                            {
                                HRID.Add(joinerandChangesSheet.Cells[row, employee_Number].GetValue<string>().Replace(" ", ""));
                                ptloc.Add(joinerandChangesSheet.Cells[row, town].Text);
                                EmpGrade.Add(joinerandChangesSheet.Cells[row, empgrade].Text);
                            }
                        }
                        row2 = 2;
                        foreach (string t in HRID)
                        {
                            outputWorksheet.Cells[row2, 1].Value = t;
                            outputWorksheet.Cells[row2, 5].Value = ptloc[row2 - 2];
                            outputWorksheet.Cells[row2, 13].Value = EmpGrade[row2 - 2];
                            row2++;
                        }
                        row2 = 2;
                        foreach (string t in HRID)
                        {
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row, employee_Number].Text.Replace(" ", "").Equals(t))
                                {
                                    if (inputWorkSheet.Cells[row, payelementdescription].Text.ToLower().Contains("meal"))
                                    {
                                        outputWorksheet.Cells[row2, 12].Value = inputWorkSheet.Cells[row, annualctc].Text;

                                    }

                                }
                            }
                            if (outputWorksheet.Cells[row2, 12].Text=="") 
                            {
                                outputWorksheet.Cells[row2, 12].Value = 0;
                            }
                            row2++;
                        }
                        row2 = 2;
                        foreach (string t in HRID)
                        {
                            for (int row = 2; row <= lastRow; row++)
                            {
                                if (inputWorkSheet.Cells[row, employee_Number].Text.Replace(" ", "").Equals(t) && inputWorkSheet.Cells[row, payelementdescription].Text.ToLower().Contains("basic") && inputWorkSheet.Cells[row, payfreq].Text.ToLower().Contains("annual"))
                                {
                                    outputWorksheet.Cells[row2, 2].Value = inputWorkSheet.Cells[row, witheffectfrom].Text;
                                    outputWorksheet.Cells[row2, 3].Value = inputWorkSheet.Cells[row, annualctc].Text;
                                    double monthly40 = (inputWorkSheet.Cells[row, annualctc].GetValue<double>()) / 30.0;
                                    double monthlyctc= outputWorksheet.Cells[row2, 3].GetValue<double>()/12.0;
                                    monthly40 = Math.Round(monthly40);
                                    outputWorksheet.Cells[row2, 4].Value = (outputWorksheet.Cells[row2, 3].GetValue<double>()/12)*0.5;
                                    outputWorksheet.Cells[row2, 7].Value = 0;
                                    outputWorksheet.Cells[row2, 9].Value = (outputWorksheet.Cells[row2, 4].GetValue<double>() * 0.2);
                                    outputWorksheet.Cells[row2, 10].Value = 0;
                                    outputWorksheet.Cells[row2, 11].Value = 0;
                                    
                                    //outputWorksheet.Cells[row2, 8].Value = monthly40;
                                    outputWorksheet.Cells[row2, 8].Value = Math.Round(outputWorksheet.Cells[row2, 4].GetValue<double>() * 0.4);
                                    string city = outputWorksheet.Cells[row2, 5].GetValue<string>();
                                    city = Service1.ShrinkString(city);
                                    bool delhi = city.Contains("delhi");
                                    bool mumbai = city.Contains("mumbai");
                                    bool maharashtra = city.Contains("maharashtra");
                                    bool tamilnadu = city.Contains("tamilnadu");
                                    bool newdelhi = city.Contains("newdelhi");
                                    bool chennai = city.Contains("chennai");
                                    bool kolkata = city.Contains("kolkata");
                                    bool calcutta = city.Contains("calcutta");
                                    if (delhi || mumbai || newdelhi || chennai || kolkata || maharashtra || tamilnadu)
                                    {
                                        outputWorksheet.Cells[row2, 8].Value = outputWorksheet.Cells[row2, 4].GetValue<double>() / 2;
                                    }
                                    outputWorksheet.Cells[row2, 6].Value = ((outputWorksheet.Cells[row2, 3].GetValue<double>() / 12) - outputWorksheet.Cells[row2, 4].GetValue<double>()) - (outputWorksheet.Cells[row2, 8].GetValue<double>() + outputWorksheet.Cells[row2, 9].GetValue<double>() + outputWorksheet.Cells[row2, 10].GetValue<double>() + outputWorksheet.Cells[row2, 11].GetValue<double>() + outputWorksheet.Cells[row2, 12].GetValue<double>());
                                    //outputWorksheet.Cells[row2, 7].Value = Math.Round(outputWorksheet.Cells[row2, 6].GetValue<double>());
                                    //outputWorksheet.Cells[row2, 8].Value = Math.Round(((outputWorksheet.Cells[row2, 3].GetValue<double>()) / 12.0) - outputWorksheet.Cells[row2, 4].GetValue<double>() - outputWorksheet.Cells[row2, 6].GetValue<double>() - outputWorksheet.Cells[row2, 7].GetValue<double>() - outputWorksheet.Cells[row2, 9].GetValue<double>());
                                    row2++;
                                }
                            }
                        }
                        int endrow = outputWorksheet.Dimension.End.Row;
                        for (row2=2;row2<=endrow; row2++){
                            if (outputWorksheet.Cells[row2,13].Text.Contains("98")){
                                outputWorksheet.Cells[row2, 4].Value = 0;
                                outputWorksheet.Cells[row2, 6].Value = 0;
                                outputWorksheet.Cells[row2, 7].Value = 0;
                                outputWorksheet.Cells[row2, 8].Value = 0;
                                outputWorksheet.Cells[row2, 9].Value = 0;
                                outputWorksheet.Cells[row2, 10].Value = outputWorksheet.Cells[row2, 3].GetValue<double>()/12;
                            }
                        }
                        outputWorksheet.Column(3).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(4).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(6).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(7).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(8).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(9).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(10).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(11).Style.Numberformat.Format = "0.00";
                        outputWorksheet.Column(12).Style.Numberformat.Format = "0.00";
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
