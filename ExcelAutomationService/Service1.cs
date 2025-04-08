using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Timer = System.Timers.Timer;
using OfficeOpenXml.Utils;
using System.Runtime.InteropServices;

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
        string destinationFolder = @"E:/PAYROLL_SERVER/Automation/output";
        string ascendcodes = "E:/PAYROLL_SERVER/Automation/Twilio_Twilio Technology/Automation_Ascent_Codes/Ascent Codes.xlsx";
        public static string subject = "";
        public static string body = "";

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
                PathLog(hrid+" IFSC code is not given in " + sheetname + " sheet.");
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
        public static void GetPlateCtcClients()
        {
            string Alerts = @"E:\PAYROLL_SERVER\Automation\Config\CTC_Breakup_By_Client.txt";
            if (File.Exists(Alerts))
            {
                CtcClients = File.ReadAllLines(Alerts);
            }
        }
        public static async Task ProcessFile(string ascendcodes, string filePath, string destinationFolder)
        {
            try
            {
                subject = "";
                body = "";
                GetAlertmails();
                GetPlateCtcClients();
                DateTime now = DateTime.Now;
                // Format the month and year as "Month_Year"
                string formattedDate = $"{now:dd_MMMM_yyyy}";
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
                            destinationFolder = destinationFolder + "/" + folderName + " " + formattedDate;
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
                            destinationFolder = destinationFolder + "/" + folderName + " " + formattedDate;
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
                // Call the relevant methods to process the file
                New_Joinee_Master.NewJoinee_Master(ascendcodes, filePath, destinationFolder);
                Rehire_Master.rehire_Master(ascendcodes, filePath, destinationFolder);
                if (filePath.ToLower().Contains("synchronoss"))
                {
                    Synchronoss_new_CTC.CTC_Master(ascendcodes, filePath, destinationFolder);
                }
                foreach (string t in CtcClients)
                {
                    if (Path.GetFileName(filePath).Contains(t))
                    {
                        //CtcByClientMaster.CtcByClient(ascendcodes, filePath, destinationFolder);
                    }
                }
                Existing_Changes_Master.Existing_changes_Master(ascendcodes, filePath, destinationFolder);
                Benefeciaries_Data.Beneficiaries_Data(ascendcodes, filePath, destinationFolder);
                Variable.Variable_Pay_Inputs_Data(ascendcodes, filePath, destinationFolder);
                Leaver_Master.LeaverMaster(ascendcodes, filePath, destinationFolder);
                //await Task.Run(() => Joiner_Leaver_Master.JoinerLeaverMaster(ascendcodes, filePath, destinationFolder));
                //await Task.Run(() => CTC_new_joiner.CTC_Master(ascendcodes, filePath, destinationFolder));
                if (subject != "")
                {
                    SendEmails(recipients, subject,body+ "Please take necessary actions.<br><br> Regards,<br> Automation Team");
                }
                //action after processing
                FileCount = 1;//Setting Back File Count to 1 for new file!!!
                using (var package = new ExcelPackage(new FileInfo(ascendcodes)))
                {
                    int n = getSheetNumber(ascendcodes, "P.F. Registration Code");
                    var PfSheet = package.Workbook.Worksheets[n];
                    string pp = PfSheet.Cells[2, 3].GetValue<string>();
                    if (ShrinkString(pp) != "")
                    {
                        try
                        {
                            StartProcessAsUser(@""+pp);
                        }
                        catch (Exception ex)
                        {
                            Log($"Error opening file: {ex.Message}");
                            File.Delete(filePath);
                        }
                    }
                    else 
                    {
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
                        Log($"Processed file: {Path.GetFileName(filePath)}\n\n");
                    }
                }
                
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
