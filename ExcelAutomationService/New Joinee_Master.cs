using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Extensions.Options;

namespace ExcelAutomationService
{
    public class New_Joinee_Master
    {
        public static void NewJoinee_Master(string ascendcodes, string filePath, string destinationFolder)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int row, row2;

                    int IP = Service1.getSheetNumber(filePath, "Joiner and Changes ");
                    var inputWorkSheet = package.Workbook.Worksheets[IP];

                    //var OrgAssignmentsDataSheet = package.Workbook.Worksheets["Org Assignments"];
                    //int OrgAssignmentsDataSheetLastRow = OrgAssignmentsDataSheet.Dimension.End.Row;
                    int employeenumber = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "HR ID");
                    int EventType = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Event Type");
                    int Aadhar = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Aadhaar Card Number");
                    int uan = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Universal Account Number (UAN)");
                    int PreferredName = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Preferred Name");
                    int fn = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "firstname");
                    int mn = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "middlename");
                    int ln = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "surname");
                    int gender = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Gender");
                    int erelation = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "relation");
                    // int dateofleaving = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "payroll end date");
                    int add1 = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 01");
                    int add2 = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 02");
                    int add3 = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Address line 03");
                    int town = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "town");
                    int pincode = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "ZIP / Postal Code");
                    int marriedornot = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "marital status");
                    //int ifsccode = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "sort code");
                    //int acno = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "account number");
                    int dob = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "date of birth");
                    int payrollstartdate = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Payroll Start Date");
                    int jobtitle = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "job title");
                    int pancard = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Permanent Account Number (PAN)");
                    int emailid = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Email Address");
                    int ManagerTier = 200;
                    int pension = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Employee Pension Scheme");
                    //int bfid = Service1.getColumnNumber(filePath, BenefeciariesDataSheet.ToString(), "Hr id");

                    int nationality = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Nationality");
                    int ptlocation = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), " PT Location");
                    int FatherorHusbandName = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Father or Husband Name");
                    int relation = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Relation");
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int EmployeeGrade = Service1.getColumnNumber(filePath, inputWorkSheet.ToString(), "Employee Grade");

                    List<string> CautionId = new List<string>();
                    List<string> InvalidPanHRId = new List<string>();
                    List<string> InvalidPan = new List<string>();
                    List<string> CautionNationality = new List<string>();
                    List<string> WrongPTLOCId = new List<string>();
                    List<string> WrongPTLOC = new List<string>();
                    string pfregistrationcode, OccupationsCode, CategoriesCode;
                    Dictionary<string, Dictionary<string, string>> BeneficiariesData = new Dictionary<string, Dictionary<string, string>>();
                    bool benefsheetExists = package.Workbook.Worksheets["Beneficiaries Data"] != null;
                    if (benefsheetExists)
                    {
                        IP = Service1.getSheetNumber(filePath, "Beneficiaries Data");
                        var BenefeciariesDataSheet = package.Workbook.Worksheets[IP];
                        int BenefeciarieslastRow = BenefeciariesDataSheet.Dimension.End.Row;
                        int primarynameasperbank = Service1.getColumnNumber(filePath, BenefeciariesDataSheet.ToString(), "Beneficiary Name");
                        int bfbankname = Service1.getColumnNumber(filePath, BenefeciariesDataSheet.ToString(), "Bank Name");
                        int bfifsc = Service1.getColumnNumber(filePath, BenefeciariesDataSheet.ToString(), "Sort Code");
                        int bfacno = Service1.getColumnNumber(filePath, BenefeciariesDataSheet.ToString(), "Account Number");
                        int bfhrid = Service1.getColumnNumber(filePath, BenefeciariesDataSheet.ToString(), "HR ID");
                        int bflastrow = BenefeciariesDataSheet.Dimension.End.Row;
                        for (row = 2; row <= bflastrow; row++)
                        {
                            if (!BeneficiariesData.ContainsKey(BenefeciariesDataSheet.Cells[row, bfhrid].Text))
                            {
                                Dictionary<string, string> empData = new Dictionary<string, string>();
                                empData.Add("Beneficiary Name", BenefeciariesDataSheet.Cells[row, primarynameasperbank].Text);
                                empData.Add("Account Number", BenefeciariesDataSheet.Cells[row, bfacno].Text);
                                empData.Add("Sort Code", Service1.ValidateIFSC(BenefeciariesDataSheet.ToString(), BenefeciariesDataSheet.Cells[row, bfhrid].Text, BenefeciariesDataSheet.Cells[row, bfifsc].Text));
                                string bankname = BenefeciariesDataSheet.Cells[row, bfbankname].Text;
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
                                empData.Add("Bank Name", bankname);
                                BeneficiariesData.Add(BenefeciariesDataSheet.Cells[row, bfhrid].Text, empData);
                            }
                        }
                    }
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("New Joinee_Master");
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
                        Dictionary<string, string> AscentLocations = new Dictionary<string, string>();
                        Dictionary<string, string> AscentGrades = new Dictionary<string, string>();
                        Dictionary<string, string> AscentBanksDetailed = new Dictionary<string, string>();
                       
                        using (var package2 = new ExcelPackage(new FileInfo(ascendcodes)))
                        {
                            var pfsheet= package2.Workbook.Worksheets[Service1.getSheetNumber(ascendcodes, "P.F. Registration Code")];
                            pfregistrationcode = pfsheet.Cells[2, 2].Text;
                            var OccupationsSheet = package2.Workbook.Worksheets[Service1.getSheetNumber(ascendcodes, "Occupations")];
                            OccupationsCode = OccupationsSheet.Cells[2, 1].Text;
                            var CategoriesSheet = package2.Workbook.Worksheets[Service1.getSheetNumber(ascendcodes, "Categories")];
                            CategoriesCode = CategoriesSheet.Cells[2, 1].Text;
                            var LocationSheet = package2.Workbook.Worksheets[Service1.getSheetNumber(ascendcodes, "Locations")];
                            int LocationsLastRow = LocationSheet.Dimension.End.Row;
                            int description = Service1.getColumnNumber(ascendcodes, LocationSheet.ToString(), "description");
                            int code = Service1.getColumnNumber(ascendcodes, LocationSheet.ToString(), "code");

                            for (int row3 = 2; row3 <= LocationsLastRow; row3++)
                            {
                                if (!AscentLocations.ContainsKey(Service1.ShrinkString(LocationSheet.Cells[row3, description].Text)) && LocationSheet.Cells[row3, description].Text != "")
                                    AscentLocations.Add(Service1.ShrinkString( LocationSheet.Cells[row3, description].Text), LocationSheet.Cells[row3, code].Text);
                            }
                            var GradeSheet = package2.Workbook.Worksheets[Service1.getSheetNumber(ascendcodes, "Grades")];
                            int GradeLastRow = GradeSheet.Dimension.End.Row;
                            description = Service1.getColumnNumber(ascendcodes, GradeSheet.ToString(), "description");
                            code = Service1.getColumnNumber(ascendcodes, GradeSheet.ToString(), "code");

                            for (int row3 = 2; row3 <= GradeLastRow; row3++)
                            {
                                if (!AscentGrades.ContainsKey(GradeSheet.Cells[row3, description].Text) && GradeSheet.Cells[row3,description].Text!="")
                                    AscentGrades.Add(GradeSheet.Cells[row3, description].Text, GradeSheet.Cells[row3, code].Text);
                            }
                            var BankSheet = package2.Workbook.Worksheets[Service1.getSheetNumber(ascendcodes, "Banks Detailed")];
                            int BankLastRow = BankSheet.Dimension.End.Row;
                            description = Service1.getColumnNumber(ascendcodes, BankSheet.ToString(), "Name of Bank");
                            code = Service1.getColumnNumber(ascendcodes, BankSheet.ToString(), "code");
                            
                            for (int row3 = 2; row3 <= BankLastRow; row3++)
                            {
                                string bankname = BankSheet.Cells[row3, description].Text;
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
                                    AscentBanksDetailed.Add(bankname, BankSheet.Cells[row3, code].Text);
                            }
                        }
                        for (row = 2; row <= lastRow; row++)
                        {
                            var cell = inputWorkSheet.Cells[row, employeenumber];
                            // Get the background color of the cell
                            var bgColor = cell.Style.Fill.BackgroundColor;
                            string eventtype = inputWorkSheet.Cells[row, EventType].Text;
                            eventtype=Service1.ShrinkString(eventtype);
                            if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF")&&eventtype!= "re-hire")
                            {
                                var HRID = inputWorkSheet.Cells[row, 2].Text;
                                outputWorksheet.Cells[row7, 1].Value = HRID;
                                var Firstname = inputWorkSheet.Cells[row, fn].Text;
                                var MiddleName = inputWorkSheet.Cells[row, mn].Text;
                                var LastName = inputWorkSheet.Cells[row, ln].Text;
                                // var FatherOrHusband = inputWorkSheet.Cells[row, 7].GetValue<string>();
                                //var Relation = inputWorkSheet.Cells[row, 7].GetValue<string>();
                                outputWorksheet.Cells[row7, 3].Value = Firstname;
                                outputWorksheet.Cells[row7, 4].Value = MiddleName;
                                Regex validCharsRegex = new Regex("[^a-zA-Z ]");
                                LastName = validCharsRegex.Replace(LastName, "");
                                if (LastName == "")
                                {
                                    outputWorksheet.Cells[row7, 5].Value = ".";
                                }
                                else
                                {
                                    outputWorksheet.Cells[row7, 5].Value = LastName;
                                }
                                outputWorksheet.Cells[row7, 8].Value = Firstname + " " + LastName;
                                if (Service1.ShrinkString(Firstname) == Service1.ShrinkString(LastName))
                                {
                                    outputWorksheet.Cells[row7, 8].Value = Firstname;
                                }
                                var Gender = inputWorkSheet.Cells[row, gender].GetValue<string>();
                                //Method to Validate Gender
                                Gender = Service1.ValidateGender(Gender);
                                outputWorksheet.Cells[row7, 2].Value = Gender;
                                if (!benefsheetExists){
                                    outputWorksheet.Cells[row7, 28].Value = "00000";
                                }
                                var fatherhusband = inputWorkSheet.Cells[row, FatherorHusbandName].Text;
                                outputWorksheet.Cells[row7, 6].Value = fatherhusband;
                                var Emprelation = inputWorkSheet.Cells[row, relation].Text;
                                Emprelation = Service1.ShrinkString(Emprelation);
                                switch (Emprelation)
                                {
                                    case "father":
                                        Emprelation = "F";
                                        outputWorksheet.Cells[row7, 7].Value = Emprelation;
                                        break;
                                    case "":
                                        Service1.PathLog(HRID + ":" + "EmpRelation is not present in " + inputWorkSheet.ToString() + " sheet.");
                                        break;
                                    case "husband":
                                        Emprelation = "H";
                                        outputWorksheet.Cells[row7, 7].Value = Emprelation;
                                        break;
                                    case "f":
                                        Emprelation = "F";
                                        outputWorksheet.Cells[row7, 7].Value = Emprelation;
                                        break;
                                    case "h":
                                        Emprelation = "H";
                                        outputWorksheet.Cells[row7, 7].Value = Emprelation;
                                        break;
                                }
                                var MaritalStatus = inputWorkSheet.Cells[row, marriedornot].GetValue<string>();
                                //Method to validate marital status
                                MaritalStatus = Service1.ValidateMaritalStatus(MaritalStatus);
                                outputWorksheet.Cells[row7, 9].Value = MaritalStatus;
                                var email = inputWorkSheet.Cells[row, emailid].GetValue<string>();
                                outputWorksheet.Cells[row7, 69].Value = email;
                                var empgr = inputWorkSheet.Cells[row, EmployeeGrade].Text;
                                outputWorksheet.Cells[row7, 52].Value = empgr;
                                outputWorksheet.Cells[row7, 74].Value = HRID;
                                outputWorksheet.Cells[row7, 31].Value = "00000";
                                var date = inputWorkSheet.Cells[row, dob].GetValue<string>();
                                date = date.Replace(" ", "");
                                if ((date.Length == 10) && (date[4] == '-'))
                                {
                                    outputWorksheet.Cells[row7, 40].Value = date;
                                }
                                date = inputWorkSheet.Cells[row, payrollstartdate].GetValue<string>();

                                date = date.Replace(" ", "");
                                if ((date.Length == 10) && (date[4] == '-'))
                                {
                                    outputWorksheet.Cells[row7, 35].Value = date;
                                    outputWorksheet.Cells[row7, 63].Value = date;
                                    outputWorksheet.Cells[row7, 103].Value = date;
                                }
                                var pan = inputWorkSheet.Cells[row, pancard].Text;

                                pan = Service1.ValidatePAN(inputWorkSheet.ToString(), HRID, pan);
                                outputWorksheet.Cells[row7, 60].Value = pan;
                                if (pan == "PANNOTAVBL"||pan==""){
                                    InvalidPanHRId.Add(HRID);
                                    InvalidPan.Add(pan);
                                }
                                var adhaar = (inputWorkSheet.Cells[row, Aadhar].Text).Replace(" ", "");

                                adhaar = Service1.ValidateAadhar(inputWorkSheet.ToString(), HRID, adhaar);
                                outputWorksheet.Cells[row7, 94].Value = adhaar;
                                var UAN = inputWorkSheet.Cells[row, uan].Text;
                                outputWorksheet.Cells[row7, 100].Value = UAN;
                                if (UAN == "" || UAN == null)
                                {
                                    Service1.PathLog(HRID + " :UAN is not present in " + inputWorkSheet.ToString() + " sheet.");
                                }
                                var JobTitle = inputWorkSheet.Cells[row, 13].Text;
                                outputWorksheet.Cells[row7, 53].Value = JobTitle;
                                var PTLocation = inputWorkSheet.Cells[row, ptlocation].Text;
                                outputWorksheet.Cells[row7, 56].Value = PTLocation;
                                string r = Service1.ShrinkString(filePath);
                                if (r.Contains("mcafee") || r.Contains("musarubra"))
                                {
                                    var grade = inputWorkSheet.Cells[row, EmployeeGrade].GetValue<string>();
                                    string gr = grade.ToLower();
                                    if (gr.Contains("grade"))
                                    {
                                        grade = grade.Substring(grade.Length - 2);
                                    }
                                    outputWorksheet.Cells[row7, 83].Value = grade;
                                }
                                var Nationality = inputWorkSheet.Cells[row, nationality].Text;
                                outputWorksheet.Cells[row7, 105].Value = Nationality;
                                if (Service1.ShrinkString(Nationality)!="ind" && Service1.ShrinkString(Nationality) != "") 
                                {
                                    CautionId.Add(HRID);
                                    CautionNationality.Add(Nationality);
                                }
                                var Pension = inputWorkSheet.Cells[row, pension].GetValue<string>();
                                Pension = Pension.ToLower();
                                Pension = Pension.Replace(" ", "");
                                switch (Pension)
                                {
                                    case "yes":
                                        Pension = "1";
                                        break;
                                    case "no":
                                        Pension = "0";
                                        break;
                                    case "0":
                                        Pension = "0";
                                        break;
                                    case "1":
                                        Pension = "1";
                                        break;
                                    default:
                                        Pension = "0";
                                        break;
                                }
                                //for (row2 = 2; row2 <= OrgAssignmentsDataSheetLastRow; row2++)
                                //{
                                //    var id = OrgAssignmentsDataSheet.Cells[row2, 1].GetValue<string>();
                                //    if (id.Equals(HRID))
                                //    {
                                //        //outputWorksheet.Cells[row7, 54].Value = OrgAssignmentsDataSheet.Cells[row2, 7].GetValue<string>(); ;

                                //    }
                                //}
                                //int n = Service1.getSheetNumber(filePath, "Cost Center code");
                                //var inputcostcentersheet = package.Workbook.Worksheets[n];
                                //int inputcostcenterLastRow = inputcostcentersheet.Dimension.End.Row;
                                //var locationcode = "!!!";
                                //for (int i = 2; i <= inputcostcenterLastRow; i++)
                                //{
                                //    var id = inputcostcentersheet.Cells[i, 1].GetValue<string>();
                                //    if (HRID.Equals(id))
                                //    {
                                //        locationcode = inputcostcentersheet.Cells[i, 2].GetValue<string>();
                                //        outputWorksheet.Cells[row7, 54].Value = locationcode;
                                //    }
                                //}
                                using (var package4 = new ExcelPackage(new FileInfo(ascendcodes)))
                                {
                                    int n = Service1.getSheetNumber(ascendcodes, "Status");
                                    var AscendStatusCode = package4.Workbook.Worksheets[n];
                                    outputWorksheet.Cells[row7, 51].Value = AscendStatusCode.Cells[2, 1].GetValue<string>();
                                    n = Service1.getSheetNumber(ascendcodes, "Grades");
                                    var AscendGradeCode = package4.Workbook.Worksheets[n];
                                    outputWorksheet.Cells[row7, 52].Value = AscendGradeCode.Cells[2, 1].GetValue<string>();
                                    n = Service1.getSheetNumber(ascendcodes, "Business Area");
                                    var AscendBusinessAreaCode = package4.Workbook.Worksheets[n];
                                    outputWorksheet.Cells[row7, 55].Value = AscendBusinessAreaCode.Cells[2, 1].GetValue<string>();
                                    outputWorksheet.Cells[row7, 50].Value = CategoriesCode;
                                    n = Service1.getSheetNumber(ascendcodes, "Cost Centers");
                                    var AscendCostCenterCode = package4.Workbook.Worksheets[n];
                                    outputWorksheet.Cells[row7, 54].Value = AscendCostCenterCode.Cells[2, 1].GetValue<string>();
                                    outputWorksheet.Cells[row7, 58].Value = OccupationsCode;
                                    n = Service1.getSheetNumber(ascendcodes, "Payroll Code");
                                    var AscendPayrollCode = package4.Workbook.Worksheets[n];
                                    outputWorksheet.Cells[row7, 34].Value = AscendPayrollCode.Cells[2, 1].GetValue<string>();
                                    if (filePath.ToLower().Contains("anthology")&& PTLocation.ToLower().Contains("chennai")) 
                                    {
                                        outputWorksheet.Cells[row7, 34].Value = AscendPayrollCode.Cells[3, 1].GetValue<string>();
                                    }
                                    outputWorksheet.Cells[row7, 61].Value = pfregistrationcode;
                                }
                                outputWorksheet.Cells[row7, 101].Value = Pension;
                                if (filePath.ToLower().Contains("anthology") || filePath.ToLower().Contains("blackboard"))
                                {
                                    if (Service1.ShrinkString(inputWorkSheet.Cells[row, ManagerTier].Text) != "")
                                    {
                                        outputWorksheet.Cells[row7, 85].Value = inputWorkSheet.Cells[row, ManagerTier].Text;
                                    }
                                }
                                
                                if (BeneficiariesData.ContainsKey(HRID))
                                {
                                    validCharsRegex = new Regex("[^a-zA-Z ]");
                                    outputWorksheet.Cells[row7, 95].Value = Service1.CapitalizeEachWord(validCharsRegex.Replace(BeneficiariesData[HRID]["Beneficiary Name"], ""));
                                    outputWorksheet.Cells[row7, 29].Value = BeneficiariesData[HRID]["Sort Code"];
                                    outputWorksheet.Cells[row7, 30].Value = BeneficiariesData[HRID]["Account Number"];
                                    string bankname= BeneficiariesData[HRID]["Bank Name"];
                                    if (AscentBanksDetailed.ContainsKey(bankname))
                                    {
                                        outputWorksheet.Cells[row7, 28].Value = AscentBanksDetailed[bankname];
                                    }
                                }
                                if (filePath.ToLower().Replace(" ", "").Contains("searchagency"))
                                {
                                    if (outputWorksheet.Cells[row7,28].Text!="") {
                                        outputWorksheet.Cells[row7, 83].Value = "N";
                                    }
                                    if (outputWorksheet.Cells[row7, 28].Text == "AXIS")
                                    {
                                        outputWorksheet.Cells[row7, 83].Value = "I";
                                    }
                                }
                                string loc = outputWorksheet.Cells[row7, 56].Text;
                                loc = loc.Replace("Remote - IND -", "");
                                loc = Service1.ShrinkString(loc);
                                if (AscentLocations.ContainsKey(loc)){
                                    outputWorksheet.Cells[row7, 56].Value = AscentLocations[loc];
                                }
                                else{
                                    WrongPTLOCId.Add(HRID);
                                    WrongPTLOC.Add(outputWorksheet.Cells[row7, 56].Text);
                                }
                                string grd = inputWorkSheet.Cells[row, EmployeeGrade].Text;

                                if (AscentGrades.ContainsKey(grd))
                                {
                                    outputWorksheet.Cells[row7, 52].Value = AscentGrades[grd];
                                }
                                //for (row2 = 2; row2 <= BenefeciarieslastRow; row2++)
                                //{
                                //    var id = BenefeciariesDataSheet.Cells[row2, bfhrid].GetValue<string>();
                                //    if (HRID.Equals(id))
                                //    {
                                //        validCharsRegex = new Regex("[^a-zA-Z ]");//logic to remove special characters
                                //        outputWorksheet.Cells[row7, 95].Value = Service1.CapitalizeEachWord(validCharsRegex.Replace(BenefeciariesDataSheet.Cells[row2, primarynameasperbank].Text, ""));
                                //        var ifsc = BenefeciariesDataSheet.Cells[row2, bfifsc].Text;
                                //        ifsc = ifsc.Replace(" ", "");
                                //        ifsc = Service1.ValidateIFSC(BenefeciariesDataSheet.ToString(), HRID, ifsc);
                                //        if (ifsc.Length == 11) { outputWorksheet.Cells[row7, 29].Value = ifsc; }
                                //        outputWorksheet.Cells[row7, 30].Value = BenefeciariesDataSheet.Cells[row2, bfacno].Text;
                                //        var bankname = BenefeciariesDataSheet.Cells[row2, bfbankname].Text;
                                //        bankname = Service1.ShrinkString(bankname);
                                //        bankname = bankname.Replace("ltd", "");
                                //        bankname = bankname.Replace("limited", "");
                                //        bankname = bankname.Replace("pvt", "");
                                //        bankname = bankname.Replace(".", "");
                                //        bankname = bankname.Replace("branch", "");
                                //        bool containsBank = bankname.Contains("bank");
                                //        if (!containsBank)
                                //        {
                                //            bankname = bankname + "bank";
                                //        }
                                //        using (var package2 = new ExcelPackage(new FileInfo(ascendcodes)))
                                //        {
                                //            int t = Service1.getSheetNumber(ascendcodes, "Locations");
                                //            var LocationSheet = package2.Workbook.Worksheets[t];
                                //            t = Service1.getSheetNumber(ascendcodes, "Banks Detailed");
                                //            var Ascendsheet = package2.Workbook.Worksheets[t];
                                //            t = Service1.getSheetNumber(ascendcodes, "Grades");
                                //            var GradeSheet = package2.Workbook.Worksheets[t];
                                //            int GradeLastRow = GradeSheet.Dimension.End.Row;
                                //            int AscendLastRow = Ascendsheet.Dimension.End.Row;
                                //            int LocationsLastRow = LocationSheet.Dimension.End.Row;
                                //            int description = Service1.getColumnNumber(ascendcodes, LocationSheet.ToString(), "description");
                                //            int code = Service1.getColumnNumber(ascendcodes, LocationSheet.ToString(), "code");
                                //            //for (int row5 = 2; row5 <= AscendLastRow; row5++)
                                //            //{
                                //            //    var bank = Ascendsheet.Cells[row5, 2].GetValue<string>();
                                //            //    bank = Service1.ShrinkString(bank);
                                //            //    containsBank = bank.Contains("bank");
                                //            //    if (!containsBank)
                                //            //    {
                                //            //        bank = bank + "bank";
                                //            //    }
                                //            //    if (bank.Equals(bankname))
                                //            //    {
                                //            //        outputWorksheet.Cells[row7, 28].Value = Ascendsheet.Cells[row5, 1].GetValue<string>();
                                //            //    }
                                //            //}
                                //            //if (AscentBanksDetailed.ContainsKey(bankname))
                                //            //{
                                //            //    outputWorksheet.Cells[row7, 28].Value = AscentBanksDetailed[bankname];
                                //            //}
                                //            //string loc = outputWorksheet.Cells[row7, 56].Text;
                                //            //loc = loc.Replace("Remote - IND -", "");
                                //            //loc = Service1.ShrinkString(loc);
                                //            //if (AscentLocations.ContainsKey(loc))
                                //            //{
                                //            //    outputWorksheet.Cells[row7, 56].Value = AscentLocations[loc];
                                //            //}
                                //            //string grd = inputWorkSheet.Cells[row, EmployeeGrade].Text;

                                //            //if (AscentGrades.ContainsKey(grd))
                                //            //{
                                //            //    outputWorksheet.Cells[row7, 52].Value = AscentGrades[grd];
                                //            //}
                                //        }
                                //    }
                                //}
                                row7++;
                            }
                        }
                        if (CautionId.Count != 0)
                        {
                            StringBuilder htmlTable = new StringBuilder();
                            htmlTable.Append("<table border='1' style='border-collapse: collapse;'>");
                            // Add table headers
                            htmlTable.Append("<tr>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>HRID</th>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>Nationality</th>");
                            htmlTable.Append("</tr>");
                            // Add table rows
                            for (int row8 = 0; row8 <= CautionId.Count - 1; row8++)
                            {
                                htmlTable.Append("<tr>");
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionId[row8]);
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionNationality[row8]);
                                htmlTable.Append("</tr>");
                            }
                            htmlTable.Append("</table>");
                            string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alerts";
                            string body = "The nationality of new joinners is different in input file: " + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString()+"<br>Please check the Expat PF contribution";
                            if (Service1.subject == "") 
                            {
                                Service1.subject = subject;
                            }
                            Service1.body = Service1.body+body;
                            // Service1.SendEmails(Service1.recipients, subject, body);
                            QuerySheet.NewJoinerNationalityQuery(destinationFolder, CautionId, CautionNationality);
                            foreach (string item in CautionId)
                            {
                                Service1.PathLog(item + ": Please check the Expat PF contribution as Nationality is not indian");
                            }
                        }
                        if (InvalidPanHRId.Count != 0)
                        {
                            StringBuilder htmlTable = new StringBuilder();
                            htmlTable.Append("<table border='1' style='border-collapse: collapse;'>");
                            // Add table headers
                            htmlTable.Append("<tr>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>HRID</th>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>PAN</th>");
                            htmlTable.Append("</tr>");
                            // Add table rows
                            for (int row8 = 0; row8 <= InvalidPan.Count - 1; row8++)
                            {
                                htmlTable.Append("<tr>");
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", InvalidPanHRId[row8]);
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", InvalidPan[row8]);
                                htmlTable.Append("</tr>");
                            }
                            htmlTable.Append("</table>");
                            string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alerts";
                            string body = "The PAN of new joinners is invalid or not given in input file: " + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br><br>";
                            if (Service1.subject == "")
                            {
                                Service1.subject = subject;
                            }
                            Service1.body = Service1.body + body;
                            // Service1.SendEmails(Service1.recipients, subject, body);
                        }
                        if (WrongPTLOCId.Count != 0)
                        {
                            StringBuilder htmlTable = new StringBuilder();
                            htmlTable.Append("<table border='1' style='border-collapse: collapse;'>");
                            // Add table headers
                            htmlTable.Append("<tr>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>HRID</th>");
                            htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>PT Location</th>");
                            htmlTable.Append("</tr>");
                            // Add table rows
                            for (int row8 = 0; row8 <= WrongPTLOCId.Count - 1; row8++)
                            {
                                htmlTable.Append("<tr>");
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", WrongPTLOCId[row8]);
                                htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", WrongPTLOC[row8]);
                                htmlTable.Append("</tr>");
                            }
                            htmlTable.Append("</table>");
                            string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alerts";
                            string body = "Kindly review below PT Locations of client: " + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br><br>";
                            if (Service1.subject == "")
                            {
                                Service1.subject = subject;
                            }
                            Service1.body = Service1.body + body;
                            // Service1.SendEmails(Service1.recipients, subject, body);
                        }
                        string newFileName = Path.Combine(destinationFolder, Service1.FileCount + "]New Joinee_Master" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        string cellValue = outputWorksheet.Cells[2, 1].GetValue<string>();
                        Service1.ShrinkString(cellValue);
                        if ((cellValue != null) && (cellValue != " ") && (cellValue != " "))
                        {
                            outputPackage.SaveAs(newFileInfo);
                            outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
                            Service1.Log("New Joiners Created Successfully");
                            Service1.FileCount++;
                        }
                        else
                        {
                            Service1.PathLog("no new joiners file created");
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