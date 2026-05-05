using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Common;
using System.Windows.Media.Imaging;

namespace ExcelAutomationService
{
    public class Existing_Changes_Master
    {
        public static void Existing_changes_Master(string ascendcodes, string filePath, string destinationFolder)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int row, row2;
                    int IP = Service1.getSheetNumber(filePath, "Joiner and Changes ");
                    var inputWorkSheet = package.Workbook.Worksheets[IP];
                    IP = Service1.getSheetNumber(filePath, "Beneficiaries Data");
                    var BenefeciariesDataSheet = package.Workbook.Worksheets[IP];
                    int ManagerTier = 200;
                    //var OrgAssignmentsDataSheet = package.Workbook.Worksheets["Org Assignments"];
                    //int OrgAssignmentsDataSheetLastRow = OrgAssignmentsDataSheet.Dimension.End.Row;
                    int employeenumber = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                    int fn = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "firstname");
                    int mn = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "middlename");
                    int ln = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "surname");
                    int gender = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Gender");
                    int erelation = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "relation");
                    //int dateofleaving = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "payroll end date");
                    int add1 = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 01");
                    int add2 = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 02");
                    int add3 = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 03");
                    int town = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "town");
                    int pincode = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "ZIP / Postal Code");
                    int marriedornot = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "marital status");
                    int dob = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "date of birth");
                    int payrollstartdate = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Payroll Start Date");
                    int jobtitle = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "job title");
                    int pancard = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Permanent Account Number (PAN)");
                    int emailid = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Email Address");
                    int pension = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Employee Pension Scheme");
                    int nationality = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Nationality");
                    int Aadhar = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Aadhaar Card Number");
                    int uan = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Universal Account Number (UAN)");
                    int ptlocation = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "PT Location");
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int FatherorHusbandName = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Father or Husband Name");
                    int lastColumn = inputWorkSheet.Dimension.End.Column;
                    int EmployeeGrade = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Employee Grade");
                    int BenefeciarieslastRow = BenefeciariesDataSheet.Dimension.End.Row;
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Existing_Changes_Master");
                        outputWorksheet.Cells[1, 1].Value = "Employee Number";
                        outputWorksheet.Cells[1, 2].Value = "Gender -M/F/T";
                        outputWorksheet.Cells[1, 3].Value = "First Name";
                        outputWorksheet.Cells[1, 4].Value = "Middle Name";
                        outputWorksheet.Cells[1, 5].Value = "lastName";
                        outputWorksheet.Cells[1, 6].Value = "Fathers/Husband Name";
                        outputWorksheet.Cells[1, 7].Value = "EmpRelation";
                        outputWorksheet.Cells[1, 8].Value = "Display Name";
                        outputWorksheet.Cells[1, 9].Value = "Marital Status (B/S/M/W)";
                        outputWorksheet.Cells[1, 10].Value = "Spouse Name";
                        outputWorksheet.Cells[1, 11].Value = "No. of Children";
                        outputWorksheet.Cells[1, 12].Value = "Date Of Leaving (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 13].Value = "Reason Of Leaving (S/R/C/D/P)";
                        outputWorksheet.Cells[1, 14].Value = "Present Address 1";
                        outputWorksheet.Cells[1, 15].Value = "Present Address 2";
                        outputWorksheet.Cells[1, 16].Value = "Present Address 3";
                        outputWorksheet.Cells[1, 17].Value = "Present City";
                        outputWorksheet.Cells[1, 18].Value = "Present State Code";
                        outputWorksheet.Cells[1, 19].Value = "Present PinCode";
                        outputWorksheet.Cells[1, 20].Value = "Present Phone";
                        outputWorksheet.Cells[1, 21].Value = "Permanent Address 1";
                        outputWorksheet.Cells[1, 22].Value = "Permanent Address 2";
                        outputWorksheet.Cells[1, 23].Value = "Permanent Address 3";
                        outputWorksheet.Cells[1, 24].Value = "Permanent City";
                        outputWorksheet.Cells[1, 25].Value = "Permanent State Code";
                        outputWorksheet.Cells[1, 26].Value = "Permanent PinCode";
                        outputWorksheet.Cells[1, 27].Value = "Permanent Phone";
                        outputWorksheet.Cells[1, 28].Value = "Primary Bank Code";
                        outputWorksheet.Cells[1, 29].Value = "Primary IFSC";
                        outputWorksheet.Cells[1, 30].Value = "Primary Bank A/c No";
                        outputWorksheet.Cells[1, 31].Value = "Secondary Bank Code";
                        outputWorksheet.Cells[1, 32].Value = "Secondary IFSC";
                        outputWorksheet.Cells[1, 33].Value = "Secondary Bank A/c No";
                        outputWorksheet.Cells[1, 34].Value = "Payroll Code";
                        outputWorksheet.Cells[1, 35].Value = "Date Of Joining (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 36].Value = "Training Start Date (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 37].Value = "Probation Start Date (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 38].Value = "Date of Confirmation (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 39].Value = "Date of Retirement (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 40].Value = "Date Of Birth (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 41].Value = "Marriage Anniversary Date (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 42].Value = "Super Annuation Number";
                        outputWorksheet.Cells[1, 43].Value = "SA wef Dt. (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 44].Value = "Super Annuation Percent";
                        outputWorksheet.Cells[1, 45].Value = "Super Annuation Max Limit";
                        outputWorksheet.Cells[1, 46].Value = "Gratuity Number";
                        outputWorksheet.Cells[1, 47].Value = "Gratuity wef Dt. (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 48].Value = "Gratuity %";
                        outputWorksheet.Cells[1, 49].Value = "FPS Number";
                        outputWorksheet.Cells[1, 50].Value = "Category Code";
                        outputWorksheet.Cells[1, 51].Value = "Status Code";
                        outputWorksheet.Cells[1, 52].Value = "Grade Code";
                        outputWorksheet.Cells[1, 53].Value = "Designation";
                        outputWorksheet.Cells[1, 54].Value = "Cost Centre Code";
                        outputWorksheet.Cells[1, 55].Value = "Business Area Code";
                        outputWorksheet.Cells[1, 56].Value = "Location Code";
                        outputWorksheet.Cells[1, 57].Value = "Leave Approver";
                        outputWorksheet.Cells[1, 58].Value = "Occupation Code";
                        outputWorksheet.Cells[1, 59].Value = "Qualification";
                        outputWorksheet.Cells[1, 60].Value = "Permanent A/c No.";
                        outputWorksheet.Cells[1, 61].Value = "P.F. Registration Code";
                        outputWorksheet.Cells[1, 62].Value = "P.F. A/c No.(10 Digits)";
                        outputWorksheet.Cells[1, 63].Value = "PF wef Dt. (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 64].Value = "E.S.I. No.";
                        outputWorksheet.Cells[1, 65].Value = "ESIC Clinic";
                        outputWorksheet.Cells[1, 66].Value = "Blood Group";
                        outputWorksheet.Cells[1, 67].Value = "Emergency Phone No.";
                        outputWorksheet.Cells[1, 68].Value = "Emergency Contact Person";
                        outputWorksheet.Cells[1, 69].Value = "Email ID";
                        outputWorksheet.Cells[1, 70].Value = "Reports To Emp";
                        outputWorksheet.Cells[1, 71].Value = "Passport No";
                        outputWorksheet.Cells[1, 72].Value = "Passport Validity (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 73].Value = "Mobile No";
                        outputWorksheet.Cells[1, 74].Value = "Web User Name";
                        outputWorksheet.Cells[1, 75].Value = "Web User Password";
                        outputWorksheet.Cells[1, 76].Value = "User Profiles";
                        outputWorksheet.Cells[1, 77].Value = "Voluntary PF %";
                        outputWorksheet.Cells[1, 78].Value = "Date Of Resign (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 79].Value = "Photo File Path";
                        outputWorksheet.Cells[1, 80].Value = "MICR";
                        outputWorksheet.Cells[1, 81].Value = "MICR2";
                        outputWorksheet.Cells[1, 82].Value = "IsLocalAuthentication";
                        outputWorksheet.Cells[1, 83].Value = "Userdefined 1 Code";
                        if (filePath.ToLower().Contains("mcafee") || filePath.ToLower().Contains("musarubra"))
                        {
                            outputWorksheet.Cells[1, 83].Value = "Pay Scale";
                        }
                        outputWorksheet.Cells[1, 84].Value = "Userdefined 2 Code";
                        outputWorksheet.Cells[1, 85].Value = "Note";
                        if (filePath.ToLower().Contains("anthology") || filePath.ToLower().Contains("blackboard"))
                        {
                            outputWorksheet.Cells[1, 85].Value = "Manager Tier";
                            ManagerTier = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Manager Tier");
                        }
                        outputWorksheet.Cells[1, 86].Value = "Userdefined 4 Code";
                        outputWorksheet.Cells[1, 87].Value = "Userdefined 5 Code";
                        outputWorksheet.Cells[1, 88].Value = "Userdefined 6";
                        outputWorksheet.Cells[1, 89].Value = "Userdefined 7 Code";
                        outputWorksheet.Cells[1, 90].Value = "Userdefined 8 Code";
                        outputWorksheet.Cells[1, 91].Value = "Userdefined 9";
                        outputWorksheet.Cells[1, 92].Value = "Userdefined 10 Code";
                        outputWorksheet.Cells[1, 93].Value = "Userdefined 11 Code";
                        outputWorksheet.Cells[1, 94].Value = "Aadhaar Card No";
                        outputWorksheet.Cells[1, 95].Value = "Primary Name As Per Bank";
                        outputWorksheet.Cells[1, 96].Value = "Secondary Name As Per Bank";
                        outputWorksheet.Cells[1, 97].Value = "InactiveID";
                        outputWorksheet.Cells[1, 98].Value = "Inactive Notes";
                        outputWorksheet.Cells[1, 99].Value = "Process Last Month In FNF";
                        outputWorksheet.Cells[1, 100].Value = "UAN";
                        outputWorksheet.Cells[1, 101].Value = "Pension Scheme";
                        outputWorksheet.Cells[1, 102].Value = "Personal Email ID";
                        outputWorksheet.Cells[1, 103].Value = "Group Joining Date (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 104].Value = "PRAN";
                        outputWorksheet.Cells[1, 105].Value = "Nationality";
                        outputWorksheet.Cells[1, 106].Value = "Religion";
                        outputWorksheet.Cells[1, 107].Value = "Personal Mobile No";
                        outputWorksheet.Cells[1, 108].Value = "Company Superannuation No";
                        outputWorksheet.Cells[1, 109].Value = "Notice Period Confirmed Days/Months";
                        outputWorksheet.Cells[1, 110].Value = "Notice Period Confirmed Type";
                        outputWorksheet.Cells[1, 111].Value = "Notice Period Probation Days/Months";
                        outputWorksheet.Cells[1, 112].Value = "Notice Period Probation Type";
                        outputWorksheet.Cells[1, 113].Value = "Labour Indentification No";
                        outputWorksheet.Cells[1, 114].Value = "Division Code";
                        outputWorksheet.Cells[1, 115].Value = "Training End Date (YYYY-MM-DD)";
                        outputWorksheet.Cells[1, 116].Value = "Probation End Date (YYYY-MM-DD)";
                        int row7 = 2;
                        Dictionary<string, string> AscentLocations=new Dictionary<string, string>();
                        Dictionary<string, string> AscentGrades = new Dictionary<string, string>();
                        using (var package2 = new ExcelPackage(new FileInfo(ascendcodes)))
                        {
                            var LocationSheet = package2.Workbook.Worksheets[Service1.getSheetNumber(ascendcodes, "Locations")];
                            int LocationsLastRow = LocationSheet.Dimension.End.Row;
                            int description = Service1.getColumnNumber(ascendcodes, LocationSheet.ToString(), "description");
                            int code = Service1.getColumnNumber(ascendcodes, LocationSheet.ToString(), "code");

                            for (int row3 = 2; row3 <= LocationsLastRow; row3++)
                            {
                                if (!AscentLocations.ContainsKey(Service1.ShrinkString(LocationSheet.Cells[row3, description].Text)) && LocationSheet.Cells[row3, description].Text != "")
                                    AscentLocations.Add(Service1.ShrinkString(LocationSheet.Cells[row3,description].Text), LocationSheet.Cells[row3, code].Text);
                            }
                            var GradeSheet = package2.Workbook.Worksheets[Service1.getSheetNumber(ascendcodes, "Grades")];
                            int GradeLastRow = GradeSheet.Dimension.End.Row;
                            description = Service1.getColumnNumber(ascendcodes, GradeSheet.ToString(), "description");
                            code = Service1.getColumnNumber(ascendcodes, GradeSheet.ToString(), "code");

                            for (int row3 = 2; row3 <= GradeLastRow; row3++)
                            {
                                if (!AscentGrades.ContainsKey(GradeSheet.Cells[row3, description].Text) && GradeSheet.Cells[row3, description].Text != "")
                                    AscentGrades.Add(GradeSheet.Cells[row3, description].Text, GradeSheet.Cells[row3, code].Text);
                            }
                        }
                        for (row = 2; row <= lastRow; row++)
                        {
                            var cell = inputWorkSheet.Cells[row, employeenumber];
                            // Get the background color of the cell
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if (string.IsNullOrEmpty(bgColor.Rgb) || bgColor.Rgb.Equals("FFFFFF"))
                            {
                                var HRID = inputWorkSheet.Cells[row, employeenumber].GetValue<string>();
                                outputWorksheet.Cells[row7, 1].Value = HRID;

                                var Firstname = inputWorkSheet.Cells[row, fn].GetValue<string>();
                                var LastName = inputWorkSheet.Cells[row, ln].GetValue<string>();
                                cell = inputWorkSheet.Cells[row, fn];
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 3].Value = Firstname;
                                    outputWorksheet.Cells[row7, 8].Value = Firstname+" "+LastName;
                                    if (Firstname==LastName) 
                                    {
                                        outputWorksheet.Cells[row7, 8].Value = Firstname;
                                    }
                                }
                                cell = inputWorkSheet.Cells[row, mn];
                                var MiddleName = inputWorkSheet.Cells[row, mn].GetValue<string>();
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 4].Value = MiddleName;
                                }
                                
                                cell = inputWorkSheet.Cells[row, ln];
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 5].Value = LastName;
                                    outputWorksheet.Cells[row7, 8].Value = Firstname + " " + LastName;
                                }
                                var columnRange = outputWorksheet.Cells[1, 55, lastRow, 55];
                                columnRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                columnRange.Style.Fill.BackgroundColor.SetColor(Color.Red);


                                var MaritalStatus = inputWorkSheet.Cells[row, marriedornot].Text;
                                MaritalStatus = MaritalStatus.ToUpper();
                                MaritalStatus = MaritalStatus.Replace(" ", "");
                                cell = inputWorkSheet.Cells[row, marriedornot];
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                   MaritalStatus=Service1.ValidateMaritalStatus(MaritalStatus);
                                   outputWorksheet.Cells[row7, 9].Value = MaritalStatus;
                                }
                                if (filePath.ToLower().Contains("anthology") || filePath.ToLower().Contains("blackboard"))
                                {
                                    if (Service1.ShrinkString(inputWorkSheet.Cells[row, ManagerTier].Text) != "")
                                    {
                                        outputWorksheet.Cells[row7, 85].Value = inputWorkSheet.Cells[row, ManagerTier].Text;
                                    }
                                }
                                var UAN = inputWorkSheet.Cells[row, uan].Text;
				                cell = inputWorkSheet.Cells[row, uan];
				                bgColor = cell.Style.Fill.BackgroundColor;
				                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
				                {
   				                    outputWorksheet.Cells[row7, 100].Value = UAN;
				                }

                                var Gender = inputWorkSheet.Cells[row, gender].Text;
                                Gender = Gender.ToLower();
                                Gender = Gender.Replace(" ", "");
                                cell = inputWorkSheet.Cells[row, gender];
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    Gender = Service1.ValidateGender(Gender);
                                    outputWorksheet.Cells[row7, 2].Value = Gender;
                                }

                                var email = inputWorkSheet.Cells[row, emailid].Text;
                                cell = inputWorkSheet.Cells[row, emailid];
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 69].Value = email;
                                }

                                var date = inputWorkSheet.Cells[row, dob].Text;
                                cell = inputWorkSheet.Cells[row, dob];
                                bgColor = cell.Style.Fill.BackgroundColor;
                                date = date.Replace(" ", "");
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF")) { 
                                    if ((date.Length == 10) && (date[4] == '-'))
                                    {
                                    outputWorksheet.Cells[row7, 40].Value = date;
                                    }
                                }

                                date = inputWorkSheet.Cells[row, payrollstartdate].Text;
                                date = date.Replace(" ", "");
                                cell = inputWorkSheet.Cells[row, payrollstartdate];
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF")) { 
                                    if ((date.Length == 10) && (date[4] == '-'))
                                    {
                                        outputWorksheet.Cells[row7, 35].Value = date;
                                        outputWorksheet.Cells[row7, 63].Value = date;
                                        outputWorksheet.Cells[row7, 103].Value = date;
                                    }
                                }

                                var pan = inputWorkSheet.Cells[row, pancard].Text;
                                cell = inputWorkSheet.Cells[row, pancard];
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                        outputWorksheet.Cells[row7, 60].Value = Service1.ValidatePAN(inputWorkSheet.ToString(), HRID.ToString(), pan.ToString());
                                }

                                var adhaar = inputWorkSheet.Cells[row, Aadhar].Text;
                                cell = inputWorkSheet.Cells[row, Aadhar];
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 94].Value = Service1.ValidateAadhar(inputWorkSheet.ToString(), HRID.ToString(), adhaar.ToString());
                                }

                                cell = inputWorkSheet.Cells[row, jobtitle];
                                var JobTitle = inputWorkSheet.Cells[row, jobtitle].Text;
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 53].Value = JobTitle;
                                }

                                #region Fatherorhusband & Emprelation
                                cell = inputWorkSheet.Cells[row, FatherorHusbandName];
                                var forh = inputWorkSheet.Cells[row, FatherorHusbandName].Text;
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 6].Value = forh;
                                }

                                cell = inputWorkSheet.Cells[row, erelation];
                                var Emprelation = inputWorkSheet.Cells[row, erelation].Text;
                                Emprelation = Service1.ShrinkString(Emprelation);
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    switch (Emprelation)
                                    {
                                        case "father":
                                            Emprelation = "F";
                                            break;
                                        case "husband":
                                            Emprelation = "H";
                                            break;
                                        case "f":
                                            Emprelation = "F";
                                            break;
                                        case "h":
                                            Emprelation = "H";
                                            break;
                                    }
                                    outputWorksheet.Cells[row7, 7].Value = Emprelation;
                                }
                                #endregion

                                #region PTLocationChange
                                var PTLocation = inputWorkSheet.Cells[row, ptlocation].Text;
                                cell = inputWorkSheet.Cells[row, ptlocation];
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 56].Value = PTLocation;
                                    if (AscentLocations.ContainsKey(Service1.ShrinkString(PTLocation)))
                                    {
                                        outputWorksheet.Cells[row7, 56].Value=AscentLocations[Service1.ShrinkString(PTLocation)];
                                    }
                                }
                                #endregion

                                #region grade for others
                                var grd = inputWorkSheet.Cells[row, EmployeeGrade].Text;
                                cell = inputWorkSheet.Cells[row, EmployeeGrade];
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    outputWorksheet.Cells[row7, 52].Value = grd;
                                    if (AscentGrades.ContainsKey(grd))
                                    {
                                        outputWorksheet.Cells[row7, 52].Value = AscentGrades[grd];
                                    }
                                }
                                #endregion

                                #region pay scale for musarubra and McAfee
                                if (filePath.ToLower().Contains("mcafee") || filePath.ToLower().Contains("musarubra"))
                                {
                                    var grade = inputWorkSheet.Cells[row, EmployeeGrade].Text;
                                    string gr = grade.ToLower();
                                    cell = inputWorkSheet.Cells[row, EmployeeGrade];
                                    // Get the background color of the cell
                                    bgColor = cell.Style.Fill.BackgroundColor;
                                    if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF")) { 
                                        if (gr.Contains("grade"))
                                        {
                                            grade = grade.Substring(grade.Length - 2);
                                        }
                                    outputWorksheet.Cells[row7, 83].Value = grade;
                                    outputWorksheet.Cells[row7, 52].Value = "";
                                    }
                                }
                                #endregion

                                var Nationality = inputWorkSheet.Cells[row, nationality].GetValue<string>();
                                cell = inputWorkSheet.Cells[row, nationality];
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                     outputWorksheet.Cells[row7, 105].Value = Nationality;
                                }

                                var Pension = inputWorkSheet.Cells[row, pension].GetValue<string>();
                                cell = inputWorkSheet.Cells[row, pension];
                                // Get the background color of the cell
                                bgColor = cell.Style.Fill.BackgroundColor;
                                if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                                {
                                    Pension = Service1.ValidatePension(Pension);
                                    outputWorksheet.Cells[row7, 101].Value = Pension;
                                }
                                row7++;
                            }
                        }
                        lastColumn = outputWorksheet.Dimension.End.Column;
                        lastRow = outputWorksheet.Dimension.End.Row;
                        for (int i = 1; i <= lastRow; i++)
                        {
                            int j;
                            for (j = 1; j <= lastColumn; j++)
                            {
                                string cellValue = outputWorksheet.Cells[i, j].GetValue<string>();
                                if ((cellValue == null) || (cellValue == "") || (cellValue == " "))
                                {
                                    outputWorksheet.Cells[i, j].Value = "!!!";
                                }
                            }
                        }
                        for (int j = 2; j <= lastColumn; j++)
                        {
                            var cell = inputWorkSheet.Cells[1, j];
                            // Get the background color of the cell
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            if (string.IsNullOrEmpty(bgColor.Rgb) || bgColor.Rgb.Equals("FFFFFF"))
                            {
                                string sheetname = outputWorksheet.Cells[1, j].GetValue<string>();
                                var outputWorksheets = outputPackage.Workbook.Worksheets.Add(sheetname);
                                outputWorksheets.Cells[1, 1].Value = "HR ID";
                                outputWorksheets.Cells[1, 2].Value = sheetname;
                                int m = 2;
                                int k = 2;
                                for (int l = 2; l <= lastRow; l++)
                                {
                                    if (outputWorksheet.Cells[l, j].GetValue<string>() != "!!!")
                                    {
                                        outputWorksheets.Cells[m, k - 1].Value = outputWorksheet.Cells[l, 1].GetValue<string>();
                                        outputWorksheets.Cells[m, k].Value = outputWorksheet.Cells[l, j].GetValue<string>();
                                        outputWorksheets.Cells[outputWorksheets.Dimension.Address].AutoFitColumns();
                                        m++;
                                    }
                                    if (sheetname.ToLower().Contains("pension") || sheetname.ToLower().Contains("uan") || sheetname.ToLower().Contains("birth") 
                                        || sheetname.ToLower().Contains("date of joining") || sheetname.ToLower().Contains("pf wef dt") 
                                        || sheetname.ToLower().Contains("payroll start date") || sheetname.ToLower().Contains("group joining date") 
                                        || sheetname.ToLower().Contains("gender"))
                                    {
                                        int OutputLastRow = outputWorksheets.Dimension.End.Row;
                                        outputWorksheets.Cells[2, 3].Value = "Before processing kindly confirm with client.";
                                        var columnRange = outputWorksheets.Cells[1, 2, OutputLastRow, 2];
                                        columnRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        columnRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                    }
                                    if (sheetname.ToLower().Contains("location"))
                                    {
                                        int OutputLastRow = outputWorksheets.Dimension.End.Row;
                                        outputWorksheets.Cells[2, 3].Value = "Before processing kindly change pf reg code if applicable.";
                                        var columnRange = outputWorksheets.Cells[1, 2, OutputLastRow, 2];
                                        columnRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        columnRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                    }
                                    if (sheetname.ToLower().Contains("pay scale"))
                                    {
                                        int OutputLastRow = outputWorksheets.Dimension.End.Row;
                                        outputWorksheets.Cells[2, 3].Value = "Check wef date for salary structure";
                                        var columnRange = outputWorksheets.Cells[1, 2, OutputLastRow, 2];
                                        columnRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        columnRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                    }
                                }
                            }
                        }
                        //Logic To delete the Empty Sheets.
                        int worksheetCount = outputPackage.Workbook.Worksheets.Count;
                        for (int i = worksheetCount - 1; i >= 1; i--)
                        {
                            var testworksheet = outputPackage.Workbook.Worksheets[i];
                            string temp = testworksheet.Cells[2, 1].GetCellValue<string>();
                            if ((temp == null) || (temp == ""))
                            {
                                outputPackage.Workbook.Worksheets.Delete(i);
                            }
                        }
                        worksheetCount = outputPackage.Workbook.Worksheets.Count;
                        for (int i = worksheetCount - 1; i >= 1; i--)
                        {
                            var testworksheet = outputPackage.Workbook.Worksheets[i];
                            StringBuilder htmlTable = new StringBuilder();
                            htmlTable.Append("<table border='1' style='border-collapse: collapse;'>");
                            int rows = testworksheet.Dimension.Rows;
                            int cols = 2;
                            string sheetname = testworksheet.Cells[1, 2].GetCellValue<string>();
                            if (sheetname.ToLower().Contains("date of joining"))
                            {
                                Dictionary<string, string> doj = new Dictionary<string, string>();
                                for (int row8 = 1; row8 <= rows; row8++) 
                                {
                                    doj.Add(testworksheet.Cells[row8,1].Text,testworksheet.Cells[row8, 2].Text);
                                }
                                
                                for (int row8 = 1; row8 <= rows; row8++)
                                {
                                    htmlTable.Append("<tr>");
                                    for (int col = 1; col <= cols; col++)
                                    {
                                        string cellValue = testworksheet.Cells[row8, col].Text;
                                        if (row8 == 1) // Header row
                                            htmlTable.Append($"<th style='background-color:lightgray;padding:5px;'>{cellValue}</th>");
                                        else
                                            htmlTable.Append($"<td style='padding:5px;'>{cellValue}</td>");
                                    }
                                    htmlTable.Append("</tr>");
                                }
                                htmlTable.Append("</table>");
                                string subject = Service1.CapitalizeEachWord(Service1.ClientName)+": Automation Alert";
                                //string table = Service1.ReadExcelAsHtml(outputPackage, sheetname);
                                string body = "Date Of Joining change is requested in the client input file: " + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br>";
                                if (Service1.subject == "")
                                {
                                    Service1.subject = subject;
                                }
                                Service1.body = Service1.body + body;
                                //Service1.SendEmails(Service1.recipients, subject, body);
                                Service1.PathLog("Date Of Joining change is reuested");
                                QuerySheet.DojQuery(destinationFolder, doj);
                            }
                            if (sheetname.ToLower().Contains("birth"))
                            {
                                Dictionary<string, string> dobc = new Dictionary<string, string>();
                                for (int row8 = 1; row8 <= rows; row8++)
                                {
                                    dobc.Add(testworksheet.Cells[row8, 1].Text, testworksheet.Cells[row8, 2].Text);
                                }
                                
                                for (int row8 = 1; row8 <= rows; row8++)
                                {
                                    htmlTable.Append("<tr>");
                                    for (int col = 1; col <= cols; col++)
                                    {
                                        string cellValue = testworksheet.Cells[row8, col].Text;
                                        if (row8 == 1) // Header row
                                            htmlTable.Append($"<th style='background-color:lightgray;padding:5px;'>{cellValue}</th>");
                                        else
                                            htmlTable.Append($"<td style='padding:5px;'>{cellValue}</td>");
                                    }
                                    htmlTable.Append("</tr>");
                                }
                                htmlTable.Append("</table>");
                                string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alert";
                                string body = "Date Of Birth change is requested in the client input file: " + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br>";
                                if (Service1.subject == "")
                                {
                                    Service1.subject = subject;
                                }
                                Service1.body = Service1.body + body;
                                //Service1.SendEmails(Service1.recipients, subject, body);
                                Service1.PathLog("Date Of Birth change is reuested");
                                QuerySheet.DobQuery(destinationFolder, dobc);
                            }
                            if (sheetname.ToLower().Contains("gender"))
                            {
                                Dictionary<string, string> gc = new Dictionary<string, string>();
                                for (int row8 = 1; row8 <= rows; row8++)
                                {
                                    gc.Add(testworksheet.Cells[row8, 1].Text, testworksheet.Cells[row8, 2].Text);
                                }
                                
                                for (int row8 = 1; row8 <= rows; row8++)
                                {
                                    htmlTable.Append("<tr>");
                                    for (int col = 1; col <= cols; col++)
                                    {
                                        string cellValue = testworksheet.Cells[row8, col].Text;
                                        if (row8 == 1) // Header row
                                            htmlTable.Append($"<th style='background-color:lightgray;padding:5px;'>{cellValue}</th>");
                                        else
                                            htmlTable.Append($"<td style='padding:5px;'>{cellValue}</td>");
                                    }
                                    htmlTable.Append("</tr>");
                                }
                                htmlTable.Append("</table>");
                                string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alert: Gender change request";
                                string body = "Gender change is requested in the client input file: " + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br>";
                                if (Service1.subject == "")
                                {
                                    Service1.subject = subject;
                                }
                                Service1.body = Service1.body + body;
                                //Service1.SendEmails(Service1.recipients, subject, body);
                                Service1.PathLog("Gender change is reuested");
                                QuerySheet.GenderChangeQuery(destinationFolder, gc);
                            }
                            if (sheetname.ToLower().Contains("nationality"))
                            {
                                Dictionary<string, string> nc = new Dictionary<string, string>();
                                for (int row8 = 1; row8 <= rows; row8++)
                                {
                                    nc.Add(testworksheet.Cells[row8, 1].Text, testworksheet.Cells[row8, 2].Text);
                                }
                                
                                for (int row8 = 1; row8 <= rows; row8++)
                                {
                                    htmlTable.Append("<tr>");
                                    for (int col = 1; col <= cols; col++)
                                    {
                                        string cellValue = testworksheet.Cells[row8, col].Text;
                                        if (row8 == 1) // Header row
                                            htmlTable.Append($"<th style='background-color:lightgray;padding:5px;'>{cellValue}</th>");
                                        else
                                            htmlTable.Append($"<td style='padding:5px;'>{cellValue}</td>");
                                    }
                                    htmlTable.Append("</tr>");
                                }
                                htmlTable.Append("</table>");
                                string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alert: Nationality change request";
                                string body = "Nationality change is requested in the client input file: " + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br>";
                                if (Service1.subject == "")
                                {
                                    Service1.subject = subject;
                                }
                                Service1.body = Service1.body + body;
                                //Service1.SendEmails(Service1.recipients, subject, body);
                                Service1.PathLog("Nationality change is reuested");
                                QuerySheet.ExistingNationalityQuery(destinationFolder, nc);
                            }
                        }
                        string newFileName = Path.Combine(destinationFolder,Service1.FileCount+ "]Existing_Changes_Master" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        outputPackage.Workbook.Worksheets.Delete(0);
                        if (outputPackage.Workbook.Worksheets.Count != 0)
                        {
                            outputPackage.SaveAs(newFileInfo);
                            outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                            Service1.FileCount++;
                            Service1.Log("Existing_Changes_Master Excel file created successfully!");
                        }
                        else
                        {
                            Service1.PathLog("No existing employee changes");
                        }
                        
                    }
                }
            }
            catch (Exception ex)
            {
                Service1.Log("Existing error");
                Service1.ErrorCount++;
                Service1.Log($"An error occurred: {ex.Message}");
            }
        }
    }
}