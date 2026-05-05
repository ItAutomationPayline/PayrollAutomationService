using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Timers;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using Timer = System.Timers.Timer;

namespace ExcelAutomationService
{
    public partial class Service1 : ServiceBase
    {
        public static string errors = @"E:/PAYROLL_SERVER/Automation/Errors";
        public static string archived = @"E:/PAYROLL_SERVER/Automation/Archived";
        public static int ErrorCount = 0;
        public static int FileCount = 1;
        public static string[] CtcClients;
        public static string[] recipients = { "dayaghan.limaye@paylineindia.com", "dhanashree.athavale@paylineindia.com", "tushar.chaudhari@paylineindia.com", "office12@yaminipanchwagh.com" };
       // public static string[] recipients = { "dayaghan.limaye@paylineindia.com"};
        public static string ClientName ="";
        Timer timer = new Timer();
        string sourceFolder = @"E:\PAYROLL_SERVER\Automation\Input";     // Folder to watch for Excel files
        public static string destination = @"E:/PAYROLL_SERVER/Automation/output";
        public static string ctcfolder = @"E:/PAYROLL_SERVER/Automation/output";
        public static string MinimumWagesAct = @"E:/PAYROLL_SERVER/Automation/output/Minimum Wages Act/";
        public static string employeemaster = @"E:/PAYROLL_SERVER/Automation/output";
        public static string payrollInputFile = "";
        string destinationFolder = @"E:/PAYROLL_SERVER/Automation/output";
        string ascendcodes = "E:/PAYROLL_SERVER/Automation/Twilio_Twilio Technology/Automation_Ascent_Codes/Ascent Codes.xlsx";
        public static string subject = "";
        public static string body = "";
        public static Dictionary<string, Dictionary<string, string>> EmployeeMaster = new Dictionary<string, Dictionary<string, string>>();

        public Service1()
        {
            InitializeComponent();
        }

        #region CTC file opening
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr WTSGetActiveConsoleSessionId();

        [DllImport("wtsapi32.dll", SetLastError = true)]
        private static extern bool WTSQueryUserToken(IntPtr SessionId, out IntPtr TokenHandle);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool CreateProcessAsUser(
            IntPtr hToken, string lpApplicationName, string lpCommandLine,
            IntPtr lpProcessAttributes, IntPtr lpThreadAttributes,
            bool bInheritHandles, uint dwCreationFlags,
            IntPtr lpEnvironment, string lpCurrentDirectory,
            ref STARTUPINFO lpStartupInfo, out PROCESS_INFORMATION lpProcessInformation);

        private struct STARTUPINFO
        {
            public int cb;
            public string lpReserved;
            public string lpDesktop;
            public string lpTitle;
            public int dwX;
            public int dwY;
            public int dwXSize;
            public int dwYSize;
            public int dwXCountChars;
            public int dwYCountChars;
            public int dwFillAttribute;
            public int dwFlags;
            public short wShowWindow;
            public short cbReserved2;
            public IntPtr lpReserved2;
            public IntPtr hStdInput;
            public IntPtr hStdOutput;
            public IntPtr hStdError;
        }

        private struct PROCESS_INFORMATION
        {
            public IntPtr hProcess;
            public IntPtr hThread;
            public int dwProcessId;
            public int dwThreadId;
        }

        public static void StartProcessAsUser(string applicationPath)
        {
            IntPtr tokenHandle;
            IntPtr sessionId = WTSGetActiveConsoleSessionId();

            if (WTSQueryUserToken(sessionId, out tokenHandle))
            {
                STARTUPINFO startupInfo = new STARTUPINFO
                {
                    cb = Marshal.SizeOf(typeof(STARTUPINFO)),
                    lpDesktop = "winsta0\\default" // This allows UI interaction
                };

                PROCESS_INFORMATION procInfo;
                bool result = CreateProcessAsUser(tokenHandle, applicationPath, null,
                    IntPtr.Zero, IntPtr.Zero, false, 0, IntPtr.Zero, null,
                    ref startupInfo, out procInfo);

                if (!result)
                {
                    Console.WriteLine($"Error: {Marshal.GetLastWin32Error()}");
                }
            }
        }
        #endregion
        public static void SendEmails(string[] emailAddresses, string subject, string body)
        {
            try
            {
                // SMTP server configuration
                string smtpHost = "smtp.office365.com"; // Replace with your SMTP server
                int smtpPort = 587; // Port number (e.g., 587 for TLS, 465 for SSL)
                string smtpUser = "donot_reply@paylineindia.com"; // Replace with your email
                string smtpPass = "D0n0t$ep!y"; // Replace with your email password
                // Initialize the SMTP client
                using (SmtpClient smtpClient = new SmtpClient(smtpHost, smtpPort))
                {
                    smtpClient.Credentials = new NetworkCredential(smtpUser, smtpPass);
                    smtpClient.EnableSsl = true; // Enable SSL/TLS for security
                    // Loop through the recipient emails
                    foreach (string recipientEmail in emailAddresses)
                    {
                        using (MailMessage mail = new MailMessage())
                        {
                            try
                            {
                                mail.From = new MailAddress(smtpUser); // Sender's email address
                                mail.To.Add(recipientEmail); // Add recipient email
                                mail.Subject = subject; // Email subject
                                mail.Body = body; // Email body
                                mail.IsBodyHtml = true; // Set to true if the body contains HTML content                        
                                smtpClient.Send(mail); // Send the email
                                Console.WriteLine($"Email sent to: {recipientEmail}");
                            }
                            catch (SmtpFailedRecipientException ex)
                            {
                                PathLog($"❌ Failed to deliver email to {ex.FailedRecipient}: {ex.Message}");
                                PathLog("Kindly inform managers that email alert was not sent successfully!!!");
                            }
                            catch (Exception ex)
                            {
                                Log($"Error sending emails: {ex.Message}");
                                PathLog($"Error sending emails: {ex.Message}");
                                PathLog($"❌ Failed to deliver email to "+recipientEmail+" Subject:"+ subject);
                                PathLog("Kindly inform managers that email alert was not sent successfully!!!");
                                // Return false if any email fails to send
                            }
                        }
                    }
                }
            }
            catch (FormatException ex)
            {
               PathLog($"🚨 Invalid Email Format: {ex.Message}");
               PathLog("Kindly inform managers that email alert was not sent successfully!!!");
            }
            catch (SmtpFailedRecipientException ex)
            {
                PathLog($"❌ Failed to deliver email to {ex.FailedRecipient}: {ex.Message}");
                PathLog("Kindly inform managers that email alert was not sent successfully!!!");
            }
            catch (SmtpException ex)
            {
                PathLog($"❌ SMTP Error: {ex.Message}");
                PathLog("Kindly inform managers that email alert was not sent successfully!!!");
            }
            catch (Exception ex)
            {
                Log($"Error sending emails: {ex.Message}");
                PathLog($"Error sending emails: {ex.Message}");
                PathLog("Kindly inform managers that email alert was not sent successfully!!!");
                // Return false if any email fails to send
            }
        }
        public static string CapitalizeEachWord(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return input; // Return the input as is if it's null, empty, or whitespace
            }
            // Split the string into words, capitalize each word, and join them back
            return string.Join(" ", input
                .Split(' ') // Split the string by spaces
                .Where(word => !string.IsNullOrWhiteSpace(word)) // Ignore extra spaces
                .Select(word => char.ToUpper(word[0]) + word.Substring(1).ToLower())); // Capitalize each word
        }
        
        public static int getColumnNumber(string filepath, string worksheetname, string columnname)
        {
            try
            {
                columnname = columnname.ToLower();
                columnname = columnname.Replace(" ", "");
                using (var package = new ExcelPackage(new FileInfo(filepath)))
                {
                    var inputWorkSheet = package.Workbook.Worksheets[worksheetname];
                    int col = 1;
                    int totalColumns = inputWorkSheet.Dimension.End.Column;
                    for (col = 1; col <= totalColumns; col++)
                    {
                        string temp = inputWorkSheet.Cells[1, col].Text.ToLower();
                        temp = temp.Replace(" ", "");
                        temp = temp.Replace("partner","");//this is done because of descrepency in some variable columns.
                        if (columnname.Equals(temp))
                        {
                            return col; // Return the column number if the header matches
                        }
                    }
                    col = -1;
                    if (col == -1)
                    {
                        PathLog(columnname + " column was not found in " + worksheetname + " of " + filepath + " file.");
                        ErrorCount++;
                    }
                    col = 999;
                    return col;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                PathLog(columnname+" column was not found in"+worksheetname+" of "+filepath+" file.");
                throw;
            }
        }
        
        public static int getSheetNumber(string filepath, string worksheetname)
        {
            try
            {
                worksheetname = ShrinkString(worksheetname);
                using (var package = new ExcelPackage(new FileInfo(filepath)))
                {
                    int worksheetCount = package.Workbook.Worksheets.Count;
                    int i = 0;
                    for (i = worksheetCount - 1; i >= 0; i--)
                    {
                        string temp = package.Workbook.Worksheets[i].Name;
                        temp = ShrinkString(temp);
                        if (temp.Equals(worksheetname))
                        {
                            return i;
                        }
                    }
                    i = 0;
                    if (i==0) {
                        PathLog(worksheetname + " sheet was not found in " + filepath);
                        ErrorCount++;
                    }
                    return i;
                }
            }
            catch (Exception e)
            {
                PathLog(e.Message);
                throw;
            }
        }
        
        public static string ValidateAadhar(string sheetname, string hrid, string adhaar)
        {
            adhaar = adhaar.Replace(" ", "");
            if (adhaar.Length == 0)
            {
                PathLog(hrid + "  aadhar number not given in " + sheetname + " sheet.");
                return "";
            }
            if ((adhaar.Length == 12) && (adhaar.All(char.IsDigit)&&(adhaar.Length != 0)))
                return adhaar;
            else
            {
                PathLog(hrid+"  adhaar number "+ adhaar + " is not valid in " + sheetname + " sheet.");
                return "";
            }
        }
        
        public static string ValidatePAN(string sheetname, string hrid, string pan)
        {
            pan = pan.Replace(" ", "");
            if ((pan.Length == 10) && (pan[3] == 'P'))
                return pan;
            if (pan.Length == 0)
            {
                PathLog( hrid + " PAN is empty in " + sheetname + " sheet.");
                return "PANNOTAVBL";
            }
            else
            {
                PathLog(hrid + "  pan number "+pan+" is not valid in " + sheetname + " sheet.");
                return "";
            }
        }
        public static string ValidateDate(string date)
        {
            return date;
        }
        
        public static string ValidateIFSC(string sheetname, string hrid, string ifsc)
        {
            ifsc = ifsc.Replace(" ", "");
            if (ifsc.Length == 0)
            {
                //PathLog(hrid+" IFSC code is not given in " + sheetname + " sheet.");
                return "";
            }
            if (ifsc.Length == 11)
                return ifsc;
            else
            {
                PathLog(hrid+" IFSC "+ ifsc +" is not valid in " + sheetname + " sheet.");
                return "";
            }
        }
        public static string ValidateGender(string Gender)
        {
            Gender=ShrinkString(Gender);
            switch (Gender)
            {
                case "male":
                    Gender = "M";
                    break;
                case "female":
                    Gender = "F";
                    break;
                case "transgender":
                    Gender = "T";
                    break;
                case "m":
                    Gender = "M";
                    break;
                case "f":
                    Gender = "F";
                    break;
                default:
                    Gender = "";
                    break;
            }
            return Gender;
        }
        public static string ValidatePension(string Pension)
        {
            Pension = ShrinkString(Pension);
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
            return Pension;
        }
        public static string ValidateMaritalStatus(string MaritalStatus)
        {
            MaritalStatus=MaritalStatus.ToUpper();
            switch (MaritalStatus)
            {
                case "BACHELOR":
                    MaritalStatus = "B";
                    break;
                case "B":
                    MaritalStatus = "B";
                    break;
                case "BACHLOR":
                    MaritalStatus = "B";
                    break;
                case "MARRIED":
                    MaritalStatus = "M";
                    break;
                case "M":
                    MaritalStatus = "M";
                    break;
                case "WIDOW":
                    MaritalStatus = "W";
                    break;
                case "W":
                    MaritalStatus = "W";
                    break;
                case "WIDOWED":
                    MaritalStatus = "W";
                    break;
                default:
                    MaritalStatus = "B";
                    break;
            }
            return MaritalStatus;
        }
        public static string ShrinkString(string input)
        {
            if (input != null)
            {
                input = input.ToLower();
                input = input.Replace(" ", "");
                return input;
            }
            return "";
        }
        protected override void OnStart(string[] args)
        {
            timer.Interval = 1000;
            timer.Enabled = true;
            if (!Directory.Exists(sourceFolder) || !Directory.Exists(destinationFolder))
            {
                Console.WriteLine("Source or destination folder does not exist. Please check paths.");
                return;
            }
            Log("Service started");
            Log("Watching for Excel files in " + sourceFolder);
            FileSystemWatcher watcher = new FileSystemWatcher(sourceFolder, "*.xlsx")
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
            };
           

            watcher.Created += async (sender, e) => await ProcessFile(ascendcodes, e.FullPath, destinationFolder);
            watcher.EnableRaisingEvents = true;
            
            //Log("Press Enter to exit...");
            Console.ReadLine();
        }
        public static void GetAlertmails() {
            string Alerts = @"E:\PAYROLL_SERVER\Automation\Config\Alerts.txt";
            if (File.Exists(Alerts))
            {
                recipients = File.ReadAllLines(Alerts);
            }
            else
            {
                recipients= new string[] { "dayaghan.limaye@paylineindia.com", "dhanashree.athavale@paylineindia.com", "tushar.chaudhari@paylineindia.com", "office12@yaminipanchwagh.com"};
            }
        }
        public static void FetchEmployeeMaster()
        {
            if (Directory.Exists(employeemaster))
            {
                string[] referencefile = Directory.GetFiles((employeemaster), "*.xlsx");
                string empmaster = employeemaster + Path.GetFileName(referencefile[0]);
                using (var package = new ExcelPackage(new FileInfo(empmaster)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    int endRow = worksheet.Dimension.End.Row;
                    int endCol = worksheet.Dimension.End.Column;

                    for (int row = 2; row <= endRow; row++) // skip header row
                    {
                        string outerKey = worksheet.Cells[row, 1].Text;

                        if (string.IsNullOrWhiteSpace(outerKey))
                            continue;

                        var innerDict = new Dictionary<string, string>();

                        for (int col = 2; col <= endCol; col++) // from column 2 onward
                        {
                            string header = worksheet.Cells[1, col].Text;
                            string value = worksheet.Cells[row, col].Text;
                            innerDict[header] = value;
                        }
                        // Store in public static dictionary
                        EmployeeMaster[outerKey] = innerDict;
                    }
                }
            }
            else
            {
                Console.WriteLine($"Directory not found: {employeemaster}");
            }
        }
        public static void GetDOBAlert(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                int IP = getSheetNumber(filePath, "Joiner and Changes");
                var CTCWorkSheet = package.Workbook.Worksheets[IP];// Assuming the data is in the first worksheet
                int lastRow = CTCWorkSheet.Dimension.End.Row;
                List<string> CautionId = new List<string>();
                List<string> CautionDate = new List<string>();
                int hridCol = getColumnNumber(filePath, CTCWorkSheet.Name, "HR ID");
                int dobcol = getColumnNumber(filePath, CTCWorkSheet.Name, "Date of Birth");
                DateTime currentDate = DateTime.Now;
                for (int row = 2; row <= lastRow; row++)
                {
                    string dobString = ShrinkString(CTCWorkSheet.Cells[row, dobcol].Text);
                    if (dobString != "")
                    {
                        DateTime dob;
                        if (DateTime.TryParseExact(dobString, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dob))
                        {
                            // Check if the employee's age is less than 18 years
                            int age = currentDate.Year - dob.Year;
                            if (currentDate.Month < dob.Month || (currentDate.Month == dob.Month && currentDate.Day < dob.Day))
                            {
                                age--;
                            }

                            if (age < 18)
                            {
                                CautionId.Add(CTCWorkSheet.Cells[row, hridCol].Text);
                                CautionDate.Add(dobString);
                            }
                        }
                    }
                }
                if (CautionId.Count != 0)
                {
                    StringBuilder htmlTable = new StringBuilder();
                    htmlTable.Append("<table border='1' style='border-collapse: collapse;'>");
                    // Add table headers
                    htmlTable.Append("<tr>");
                    htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>HRID</th>");
                    htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>Date of Birth</th>");
                    htmlTable.Append("</tr>");
                    PathLog("Below are the Suspicious Date of births of Joiner and Chnages sheet:");
                    // Add table rows
                    for (int row8 = 0; row8 <= CautionId.Count - 1; row8++)
                    {
                        Service1.PathLog("HRID:" + CautionId[row8] + " Dob:" + CautionDate[row8]);
                        htmlTable.Append("<tr>");
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionId[row8]);
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionDate[row8]);
                        htmlTable.Append("</tr>");
                    }
                    htmlTable.Append("</table>");
                    string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alert";
                    string body = "Below are the Suspicious Date of Births of employees of Client:" + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br>";
                    if (Service1.subject == "")
                    {
                        Service1.subject = subject;
                    }
                    Service1.body = Service1.body + body;
                }
            }
        }
        public static void GetDOJMismatchAlert(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                int IP = getSheetNumber(filePath, "Joiner and Changes");
                var JoinerandchangesSheet = package.Workbook.Worksheets[IP];// Assuming the data is in the first worksheet
                int lastRow = JoinerandchangesSheet.Dimension.End.Row;
                Dictionary<string, string> EmpData = new Dictionary<string, string>();
                List<string> CautionId = new List<string>();
                List<string> CautionNewJoinerDoj = new List<string>();
                List<string> CautionPaymentsandDeductionDoj = new List<string>();
                int hridCol = getColumnNumber(filePath, JoinerandchangesSheet.Name, "HR ID");
                int dojCol = getColumnNumber(filePath, JoinerandchangesSheet.Name, "Payroll Start Date");
                DateTime currentDate = DateTime.Now;
                for (int row = 2; row <= lastRow; row++)
                {
                    var cell = JoinerandchangesSheet.Cells[row, hridCol];
                    var bgColor = cell.Style.Fill.BackgroundColor;
                    if (!string.IsNullOrEmpty(bgColor.Rgb) && !bgColor.Rgb.Equals("FFFFFF"))
                    {
                        string dojString = ShrinkString(JoinerandchangesSheet.Cells[row, dojCol].Text);
                        if (dojString != "")
                        {
                            EmpData.Add(JoinerandchangesSheet.Cells[row, hridCol].Text, dojString);
                        }
                    }
                }
                IP = getSheetNumber(filePath, "Payments and Deductions");
                var PaymentsandDeducionsSheet = package.Workbook.Worksheets[IP];// Assuming the data is in the first worksheet
                lastRow = PaymentsandDeducionsSheet.Dimension.End.Row;
                hridCol = getColumnNumber(filePath, PaymentsandDeducionsSheet.Name, "HR ID");
                dojCol = getColumnNumber(filePath, PaymentsandDeducionsSheet.Name, "Start Date");
                for (int row = 2; row <= lastRow; row++)
                {
                    string hrid = PaymentsandDeducionsSheet.Cells[row, hridCol].Text;
                    string doj = PaymentsandDeducionsSheet.Cells[row, dojCol].Text;
                    if (EmpData.ContainsKey(hrid)){
                        if (EmpData[hrid] != doj)
                        {
                            if (!CautionId.Contains(hrid))
                            {
                                CautionId.Add(hrid);
                                CautionNewJoinerDoj.Add(EmpData[hrid]);
                                CautionPaymentsandDeductionDoj.Add(doj);
                            }
                        }
                    }
                }
                if (CautionId.Count != 0)
                {
                    StringBuilder htmlTable = new StringBuilder();
                    htmlTable.Append("<table border='1' style='border-collapse: collapse;'>");
                    // Add table headers
                    htmlTable.Append("<tr>");
                    htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>HRID</th>");
                    htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>DOJ in Joiner & Changes</th>");
                    htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>DOJ in Payments & Deduction</th>");
                    htmlTable.Append("</tr>");
                   
                    // Add table rows
                    for (int row8 = 0; row8 <= CautionId.Count - 1; row8++)
                    {
                        htmlTable.Append("<tr>");
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionId[row8]);
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionNewJoinerDoj[row8]);
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionPaymentsandDeductionDoj[row8]);
                        htmlTable.Append("</tr>");
                    }
                    htmlTable.Append("</table>");
                    string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alert";
                    string body = "Kindly review the Date of Joinings of new joiners of Client:" + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br>";
                    if (Service1.subject == "")
                    {
                        Service1.subject = subject;
                    }
                    Service1.body = Service1.body + body;
                }
            }
        }
        public static void GetCtcAdditionalCommentAlert(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                int IP = getSheetNumber(filePath, "Payments and Deductions");
                var CTCWorkSheet = package.Workbook.Worksheets[IP];// Assuming the data is in the first worksheet
                int lastRow = CTCWorkSheet.Dimension.End.Row;
                List<string> CautionId = new List<string>();
                List<string> CautionDesc = new List<string>();
                List<double> CautionAmt = new List<double>();
                List<string> CautionComment = new List<string>();
                int hridCol= getColumnNumber(filePath, CTCWorkSheet.Name, "HR ID");
                int payElementCol= getColumnNumber(filePath, CTCWorkSheet.Name, "Pay Element Short Code");
                int amountCol = getColumnNumber(filePath, CTCWorkSheet.Name, "Amount");
                int commentcol = getColumnNumber(filePath, CTCWorkSheet.Name, "Additional Comment");
                for (int row = 2; row <= lastRow; row++)
                {
                    var commentText = CTCWorkSheet.Cells[row, commentcol].Text;
                    if (ShrinkString(CTCWorkSheet.Cells[row,commentcol].Text)!="" && !CautionComment.Contains(commentText))
                    {
                        CautionId.Add(CTCWorkSheet.Cells[row, hridCol].Text);
                        CautionDesc.Add(CTCWorkSheet.Cells[row, payElementCol].Text);
                        CautionAmt.Add(CTCWorkSheet.Cells[row, amountCol].GetValue<double>());
                        CautionComment.Add(CTCWorkSheet.Cells[row, commentcol].Text);
                    }
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
                    htmlTable.Append("<th style='background-color:lightgray;padding:5px;'>Additional Comment</th>");
                    htmlTable.Append("</tr>");
                    PathLog("Below are the Additional comments of Payments and Deductions sheet:");
                    // Add table rows
                    for (int row8 = 0; row8 <= CautionId.Count - 1; row8++)
                    {
                        Service1.PathLog("HRID:" + CautionId[row8] + " Description:" + CautionDesc[row8] + " Amount:" + CautionAmt[row8] + " Comment:" + CautionComment[row8]);
                        htmlTable.Append("<tr>");
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionId[row8]);
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionDesc[row8]);
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionAmt[row8].ToString("N2"));
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionComment[row8]);
                        htmlTable.Append("</tr>");
                    }
                    htmlTable.Append("</table>");
                    string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alert";
                    string body = "Below are the Additional Comments in Payments and Deductions File of Client:" + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br>";
                    if (Service1.subject == "")
                    {
                        Service1.subject = subject;
                    }
                    Service1.body = Service1.body + body;
                }
            }
        }
        public static void GetAlterCouponAlert(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                int IP = getSheetNumber(filePath, "Payments and Deductions");
                var CTCWorkSheet = package.Workbook.Worksheets[IP];// Assuming the data is in the first worksheet
                int lastRow = CTCWorkSheet.Dimension.End.Row;
                //HashSet<string> EmpId = new HashSet<string>();
                List<string> EmpId = new List<string>();
                List<string> CautionId = new List<string>();
                List<string> CautionDesc = new List<string>();
                List<string> CautionCode = new List<string>();
                List<double> CautionAmt = new List<double>();
                int hridCol = getColumnNumber(filePath, CTCWorkSheet.Name, "HR ID");
                int payElementCol = getColumnNumber(filePath, CTCWorkSheet.Name, "Pay Element Short Code");
                int amountCol = getColumnNumber(filePath, CTCWorkSheet.Name, "Amount");
                int commentcol = getColumnNumber(filePath, CTCWorkSheet.Name, "Additional Comment");
                for (int row = 2; row <= lastRow; row++)
                {
                    var cell = CTCWorkSheet.Cells[row, hridCol];
                    // Get the background color of the cell
                    var bgColor = cell.Style.Fill.BackgroundColor;
                    if (string.IsNullOrEmpty(bgColor.Rgb) || bgColor.Rgb.Equals("FFFFFF"))
                    {
                        if (CTCWorkSheet.Cells[row, payElementCol].Text.Contains("Meal"))
                        {

                            if (!EmpId.Contains(CTCWorkSheet.Cells[row, hridCol].Text))
                            {
                                EmpId.Add(CTCWorkSheet.Cells[row, hridCol].Text);
                            }
                            else
                            {
                                CautionId.Add(CTCWorkSheet.Cells[row, hridCol].Text);
                                CautionCode.Add(CTCWorkSheet.Cells[row, payElementCol].Text);
                                CautionAmt.Add(CTCWorkSheet.Cells[row, amountCol].GetValue<double>());
                            }
                        }
                    }
                }
                EmpId.Clear();
                for (int row = 2; row <= lastRow; row++)
                {
                    var cell = CTCWorkSheet.Cells[row, hridCol];
                    // Get the background color of the cell
                    var bgColor = cell.Style.Fill.BackgroundColor;
                    if (string.IsNullOrEmpty(bgColor.Rgb) || bgColor.Rgb.Equals("FFFFFF"))
                    {
                        if (CTCWorkSheet.Cells[row, payElementCol].Text.Contains("Fuel"))
                        {
                            if (!EmpId.Contains(CTCWorkSheet.Cells[row, hridCol].Text))
                            {
                                EmpId.Add(CTCWorkSheet.Cells[row, hridCol].Text);
                            }
                            else
                            {
                                CautionId.Add(CTCWorkSheet.Cells[row, hridCol].Text);
                                CautionCode.Add(CTCWorkSheet.Cells[row, payElementCol].Text);
                                CautionAmt.Add(CTCWorkSheet.Cells[row, amountCol].GetValue<double>());
                            }
                        }
                    }
                }
                EmpId.Clear();
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
                    //PathLog("Kindly review these employees which have repeated coupon");
                    // Add table rows
                    for (int row8 = 0; row8 <= CautionId.Count - 1; row8++)
                    {
                        htmlTable.Append("<tr>");
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionId[row8]);
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionCode[row8]);
                        htmlTable.AppendFormat("<td style='padding:5px;'>{0}</td>", CautionAmt[row8].ToString("N2"));
                        htmlTable.Append("</tr>");
                    }
                    htmlTable.Append("</table>");
                    string subject = Service1.CapitalizeEachWord(Service1.ClientName) + ": Automation Alert";
                    string body = "Kindly review the Pay element codes which are repeated for some employees of Client:" + Path.GetFileName(filePath) + "<br><br>" + htmlTable.ToString() + "<br>";
                    if (Service1.subject == "")
                    {
                        Service1.subject = subject;
                    }
                    Service1.body = Service1.body + body;
                }
                else 
                {
                    PathLog("No employee found which has repeated coupon");
                }
            }
        }
        public static string ReplaceAmountPlaceholders(string formula, string empId)
        {
            PathLog(formula);
            if (string.IsNullOrWhiteSpace(formula) || !formula.Contains("amount("))
                return formula;

            var matches = Regex.Matches(formula, @"amount\((.*?)\)");

            foreach (Match match in matches)
            {
                string fullMatch = match.Value;                 // e.g., amount(001 Basic Salary)
                string innerKey = match.Groups[1].Value.Trim(); // e.g., 001 Basic Salary
                string replacementValue = "0";                  // default fallback

                if (EmployeeMaster != null && EmployeeMaster.ContainsKey(empId))
                {
                    var innerDict = EmployeeMaster[empId];

                    if (innerDict.ContainsKey(innerKey))
                    {
                        replacementValue = innerDict[innerKey];

                        if (!double.TryParse(replacementValue, out _))
                        {
                            PathLog($"HRID: {empId} has invalid numeric value '{replacementValue}' for key '{innerKey}'.");
                            replacementValue = "0";
                        }
                    }
                    else
                    {
                        PathLog($"HRID: {empId} does not contain key '{innerKey}' in EmployeeMaster.");
                    }
                }
                else
                {
                    PathLog($"HRID: {empId} not found in EmployeeMaster.");
                }

                formula = formula.Replace(fullMatch, replacementValue);
            }
            PathLog(formula);
            return formula;
        }
        public static double GetPreviousMonthAnnualCTC(string empid)
        {
            if (EmployeeMaster != null && EmployeeMaster.ContainsKey(empid))
            {
                var inner = EmployeeMaster[empid];

                if (inner.ContainsKey("CTC"))
                {
                    string ctcText = inner["CTC"];

                    if (double.TryParse(ctcText, out double ctcValue))
                    {
                        return ctcValue;
                    }
                    else
                    {
                        PathLog("HRID: " + empid + " has invalid CTC format.");
                        return 1;
                    }
                }
                else
                {
                    PathLog("HRID: " + empid + " does not contain CTC column.");
                    return 1;
                }
            }

            PathLog("HRID: " + empid + " does not exist in employee master for units conversion. Kindly ensure that correct previous month employee master is loaded.");
            return 1;
        }
        public static (int year, int month) ExtractYearMonthFromFileName(string fileName)
        {
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            string[] parts = nameWithoutExtension.Split('_');

            // Find index of "Payroll_Data"
            int payrollIndex = Array.FindIndex(parts, p => p.Equals("Payroll", StringComparison.OrdinalIgnoreCase));

            if (payrollIndex < 2)
                throw new Exception("Invalid file name format");

            // Date is just before "Payroll"
            string datePart = parts[payrollIndex - 1];

            // Month is before date
            string monthName = parts[payrollIndex - 2];

            int year = int.Parse(datePart.Substring(0, 4));

            int month = DateTime.ParseExact(
                monthName,
                "MMMM",
                System.Globalization.CultureInfo.InvariantCulture
            ).Month;

            return (year, month);
        }
        public static int GetCurrentMonthTotalDays()
        {
            var (year, month) = ExtractYearMonthFromFileName(Path.GetFileName(payrollInputFile));
            //Program.currentmonthdays = DateTime.DaysInMonth(year, month);
            return DateTime.DaysInMonth(year, month);
        }
        public static int GetPreviousMonthTotalDays()
        {
            var (year, month) = ExtractYearMonthFromFileName(Path.GetFileName(payrollInputFile));

            if (month == 1)
            {
                month = 12;
                year -= 1;
            }
            else
            {
                month -= 1;
            }
            //Program.previousmonthdays=DateTime.DaysInMonth(year, month);
            return DateTime.DaysInMonth(year, month);
        }


        public static async Task ProcessFile(string ascendcodes, string filePath, string destinationFolder)
        {
            try
            {
                payrollInputFile = "";
                subject = "";
                body = "";
                GetAlertmails();
                DateTime now = DateTime.Now;
                // Format the month and year as "Month_Year"
                string formattedDate = $"{now:yyyyMMdd}";
                string foldername = Path.GetFileName(filePath);
                foldername = foldername.Replace(".xlsx", "");
                string filename = Path.GetFileName(filePath.ToLower());
                string[] directories = Directory.GetDirectories(destinationFolder);
                //method to find right destinationfolder
                string[] folderNames = Array.ConvertAll(directories, dir => Path.GetFileName(dir.ToLower()));
                foreach (string folderName in folderNames)
                {
                    if (!folderName.Contains(' '))
                    {
                        if (filename.ToLower().Contains(folderName.ToLower()))
                        {
                            Console.WriteLine(folderName);
                            destinationFolder = destinationFolder + "/" + folderName;
                            string[] referencefile=Directory.GetFiles((destinationFolder), "*.xlsx");
                            ascendcodes = destinationFolder + "/" + Path.GetFileName(referencefile[0]);
                            ctcfolder = destinationFolder + "/CTC_Structure/";
                            employeemaster = destinationFolder + "/Employee Master/";
                            destinationFolder = destinationFolder + "/" + formattedDate + " " + folderName;
                            destination = destinationFolder;
                            Console.WriteLine(ascendcodes);
                            ClientName=folderName;
                            break;
                        }
                    }
                    //in case of spaces in folder name
                    else
                    {
                        string[] parts = folderName.Split(' ');
                        int count = parts.Length;
                        int temp = 0;
                        foreach (string part in parts)
                        {
                            if (filename.ToLower().Contains(part.ToLower()))
                            {
                                temp++;
                            }
                        }
                        if (temp == count)
                        {
                            destinationFolder = destinationFolder + "/" + folderName;
                            string[] referencefile = Directory.GetFiles((destinationFolder), "*.xlsx");
                            ascendcodes = destinationFolder + "/" + Path.GetFileName(referencefile[0]);
                            ctcfolder = destinationFolder + "/CTC_Structure/";
                            employeemaster = destinationFolder + "/Employee Master/";
                            destinationFolder = destinationFolder + "/" + formattedDate + " " + folderName;
                            destination = destinationFolder;
                            Console.WriteLine(ascendcodes);
                            ClientName = folderName;
                            break;
                        }
                    }
                }
                if (!Directory.Exists(foldername))
                {
                    Directory.CreateDirectory(destinationFolder);
                }
                // Ensure file is fully available by checking in a loop until it's accessible
                for (int retries = 0; retries < 5; retries++)
                {
                    try
                    {
                        using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                        {
                            stream.Close();
                            break; // If accessible, break the loop
                        }
                    }
                    catch (IOException)
                    {
                        await Task.Delay(500); // Wait and retry if file is still being written
                    }
                }
                Log("File detected:"+ Path.GetFileName(filePath));
                payrollInputFile = filePath;
                New_Joinee_Master.NewJoinee_Master(ascendcodes, filePath, destinationFolder);
                Rehire_Master.rehire_Master(ascendcodes, filePath, destinationFolder);
                if (filePath.ToLower().Contains("synchronoss"))
                {
                    //Synchronoss_new_CTC.CTC_Master(ascendcodes, filePath, destinationFolder);
                }
                if (filePath.ToLower().Contains("alter"))
                {
                    GetAlterCouponAlert(filePath);
                    Variable.AlterDomusCoupunsVariable(ascendcodes, filePath, destinationFolder);
                }
                if (filePath.ToLower().Contains("mcafee"))
                {
                    McAfee_CTC.CTC_Master(ascendcodes, filePath, destinationFolder);
                }
                if (filePath.ToLower().Contains("musarubra"))
                {
                    Musarubra_CTC.CTC_Master(ascendcodes, filePath, destinationFolder);
                }
                if (Directory.Exists(ctcfolder))
                {
                    string[] referencefile = Directory.GetFiles((ctcfolder), "*.xlsx");
                    string ctccodes = ctcfolder + Path.GetFileName(referencefile[0]);
                    //string copypath = @"E:\PAYROLL_SERVER\Automation\Config\" +"1"+ "temp.xlsx";
                    MasterCTC.CTC_Master(ctccodes, filePath, destinationFolder);
                    InternCTC.CTC_Master(ctccodes, filePath, destinationFolder);
                    TraineeCTC.CTC_Master(ctccodes, filePath, destinationFolder);
                }
                Existing_Changes_Master.Existing_changes_Master(ascendcodes, filePath, destinationFolder);
                Benefeciaries_Data.Beneficiaries_Data(ascendcodes, filePath, destinationFolder);
                Variable.Variable_Pay_Inputs_Data(ascendcodes, filePath, destinationFolder);
                Leaver_Master.LeaverMaster(ascendcodes, filePath, destinationFolder);
                GetCtcAdditionalCommentAlert(filePath);
                GetDOBAlert(filePath);
                GetDOJMismatchAlert(filePath);
                EmployeeMaster.Clear();
                if (subject != "")
                {
                    SendEmails(recipients, subject, body+ "Please take necessary actions.<br><br> Regards,<br> Automation Team");
                    string wordFileName = Path.Combine(destinationFolder,Service1.FileCount 
                        + "]DO READ THIS & ACTION!!!" + ".docx");
                    SaveBodyToWord(body, wordFileName);
                }
                //action after processing
                FileCount = 1;//Setting Back File Count to 1 for new file!!!
                //using (var package = new ExcelPackage(new FileInfo(ascendcodes)))
                //{
                    //int n = getSheetNumber(ascendcodes, "P.F. Registration Code");
                    //var PfSheet = package.Workbook.Worksheets[n];
                    //string pp = PfSheet.Cells[2, 3].GetValue<string>();
                    //if (ShrinkString(pp) != "")
                    //{
                    //    //try
                    //    //{
                    //    //    StartProcessAsUser(@""+pp);
                    //    //}
                    //    //catch (Exception ex)
                    //    //{
                    //    //    Log($"Error opening file: {ex.Message}");
                    //    //    File.Delete(filePath);
                    //    //}
                    //}
                    //else 
                    //{
                        if (!Directory.Exists(archived))
                        {
                            Directory.CreateDirectory(archived);
                        }
                        if (!Directory.Exists(errors))
                        {
                            Directory.CreateDirectory(errors);
                        }
                        if (File.Exists(filePath))
                        {
                            if (ErrorCount == 0)
                            {
                                if (File.Exists(archived + "/" + Path.GetFileName(filePath)))
                                {
                                    File.Delete(filePath);
                                }
                                else
                                {
                                    File.Move(filePath, Path.Combine(archived, Path.GetFileName(filePath)));
                                }
                            }
                            else
                            {
                                if (File.Exists(errors + "/" + Path.GetFileName(filePath)))
                                {
                                    File.Delete(filePath);
                                }
                                else
                                {
                                    File.Move(filePath, Path.Combine(errors, Path.GetFileName(filePath)));
                                }
                                ErrorCount = 0;
                            }
                        }
                        //File.Delete(destPath);
                        Log($"Processed file: {Path.GetFileName(filePath)}\n\n");
                    //}
                //}
                
            }
            catch (Exception ex)
            {
               Log($"Error processing file {Path.GetFileName(filePath)}: {ex.Message}");
               File.Move(filePath, Path.Combine(errors, Path.GetFileName(filePath)));
            }
        }
        protected override void OnStop()
        {
            DateTime today = DateTime.Today;
            Log("Service stopped.");
            string[] recipients = { "dayaghan.limaye@paylineindia.com" };
            string subject = "Service Stopeed";
            string body = "PayrollAutomation service was stopped at:"+ $"{DateTime.Now}"+ "\n\nRegards,\nEmailService";
            //SendEmails(recipients, subject, body);
        }
        private static void SaveBodyToWord(string htmlBody, string outputPath)
        {
            // Step 1: Write directly to file (NOT MemoryStream - avoids flush issues)
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(
                outputPath, WordprocessingDocumentType.Document))
            {
                // Step 2: Add and fully initialize the main document part
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();

                // Step 3: CRITICAL - Create a complete, valid document skeleton first
                mainPart.Document = new Document(
                    new Body(
                        new Paragraph(new Run(new Text(""))) // placeholder, required
                    )
                );

                // Step 4: Convert HTML to OpenXml
                HtmlConverter converter = new HtmlConverter(mainPart);

                // Step 5: Clean up common HTML email issues before parsing
                string cleanHtml = htmlBody
                    .Replace("&nbsp;", " ")
                    .Replace("\r\n", " ")
                    .Replace("\n", " ");

                converter.ParseHtml(cleanHtml);

                // Step 6: CRITICAL - Ensure body has at least one paragraph (Word requires this)
                Body body = mainPart.Document.Body;
                if (!(body.LastChild is Paragraph))
                {
                    body.AppendChild(new Paragraph());
                }


                // Step 7: Save explicitly before closing
                mainPart.Document.Save();

                // wordDoc.Dispose() is called here by 'using' — file is complete
            }
        }
        public static void Log(string message)
        {
            try
            {
                DateTime today = DateTime.Today;
                string _logFilePath = @"E:\PAYROLL_SERVER\Automation\ServiceLogs\" + today.ToString("dd/MMMM/yyyy")+"_PayrollAutomationService.log";
                Directory.CreateDirectory(Path.GetDirectoryName(_logFilePath));
                File.AppendAllText(_logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                // Fail silently if logging fails to avoid crashing the service
            }
        }
        public static void PathLog(string message)
        {
            try
            {
                DateTime today = DateTime.Today;
                string _logFilePath = destination +"/"+"_PayrollAutomationService_" + today.ToString("dd/MMMM/yyyy") + ".log";
                Directory.CreateDirectory(Path.GetDirectoryName(_logFilePath));
                File.AppendAllText(_logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                // Fail silently if logging fails to avoid crashing the service
            }
        }
    }
}
