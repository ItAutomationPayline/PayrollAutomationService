using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelAutomationService
{
    public class QuerySheet
    {
        public static void DobQuery(string destinationFolder,Dictionary<string,string> doj)
        {
            string filePath = Path.Combine(destinationFolder, "QuerySheet.xlsx");
            FileInfo fileInfo = new FileInfo(filePath);

            // Load existing file or create a new one
            using (var outputPackage = File.Exists(filePath) ? new ExcelPackage(fileInfo) : new ExcelPackage())         // Create new
            {
                var outputWorksheet = outputPackage.Workbook.Worksheets.Add("DOB");
                int row = 1;
                foreach (var ele in doj)
                {
                    outputWorksheet.Cells[row, 1].Value = ele.Key;
                    outputWorksheet.Cells[row, 2].Value = ele.Value;
                    row++;
                }
                string newFileName = Path.Combine(destinationFolder, "QuerySheet.xlsx");
                FileInfo newFileInfo = new FileInfo(newFileName);
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                //outputPackage.SaveAs(newFileInfo);
                //outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
            }
        }
        public static void DojQuery(string destinationFolder, Dictionary<string, string> doj) 
        {
            string filePath = Path.Combine(destinationFolder, "QuerySheet.xlsx");
            FileInfo fileInfo = new FileInfo(filePath);

            // Load existing file or create a new one
            using (var outputPackage = File.Exists(filePath) ? new ExcelPackage(fileInfo) : new ExcelPackage())         // Create new
            {
                var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Doj");
                int row = 1;
                foreach (var ele in doj)
                {
                    outputWorksheet.Cells[row, 1].Value = ele.Key;
                    outputWorksheet.Cells[row, 2].Value = ele.Value;
                    row++;
                }
                string newFileName = Path.Combine(destinationFolder, "QuerySheet.xlsx");
                FileInfo newFileInfo = new FileInfo(newFileName);
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                //outputPackage.SaveAs(newFileInfo);
                //outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
            }
        }
        public static void ExistingNationalityQuery(string destinationFolder, Dictionary<string, string> exn) 
        {
            string filePath = Path.Combine(destinationFolder, "QuerySheet.xlsx");
            FileInfo fileInfo = new FileInfo(filePath);

            // Load existing file or create a new one
            using (var outputPackage = File.Exists(filePath) ? new ExcelPackage(fileInfo) : new ExcelPackage())         // Create new
            {
                var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Existing Nationality");
                int row = 1;
                foreach (var ele in exn)
                {
                    outputWorksheet.Cells[row, 1].Value = ele.Key;
                    outputWorksheet.Cells[row, 2].Value = ele.Value;
                    row++;
                }
                string newFileName = Path.Combine(destinationFolder, "QuerySheet.xlsx");
                FileInfo newFileInfo = new FileInfo(newFileName);
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                //outputPackage.SaveAs(newFileInfo);
                //outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
            }
        }
        public static void NewJoinerNationalityQuery(string destinationFolder, List<string> id, List<string> nat) 
        {
            string filePath = Path.Combine(destinationFolder, "QuerySheet.xlsx");
            FileInfo fileInfo = new FileInfo(filePath);
            // Load existing file or create a new one
            using (var outputPackage = File.Exists(filePath) ? new ExcelPackage(fileInfo): new ExcelPackage())         // Create new
            {
                var outputWorksheet = outputPackage.Workbook.Worksheets.Add("New joiner's Nationality");
                string newFileName = Path.Combine(destinationFolder, "QuerySheet.xlsx");
                outputWorksheet.Cells[1, 1].Value = "HRID";
                outputWorksheet.Cells[1, 2].Value = "Nationality";
                int row = 2;
                for (int i=0;i<=id.Count-1;i++)
                {
                    outputWorksheet.Cells[row, 1].Value= id[i];
                    outputWorksheet.Cells[row, 2].Value = nat[i];
                    row++;
                }
                FileInfo newFileInfo = new FileInfo(newFileName);
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                //outputPackage.SaveAs(newFileInfo);
                //outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
            }
        }
        public static void GenderChangeQuery(string destinationFolder, Dictionary<string, string> gc)
        {
            string filePath = Path.Combine(destinationFolder, "QuerySheet.xlsx");
            FileInfo fileInfo = new FileInfo(filePath);

            // Load existing file or create a new one
            using (var outputPackage = File.Exists(filePath)? new ExcelPackage(fileInfo): new ExcelPackage())         // Create new
            {
                var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Gender");
                int row = 1;
                foreach (var ele in gc)
                {
                    outputWorksheet.Cells[row, 1].Value = ele.Key;
                    outputWorksheet.Cells[row, 2].Value = ele.Value;
                    row++;
                }
                string newFileName = Path.Combine(destinationFolder, "QuerySheet.xlsx");
                FileInfo newFileInfo = new FileInfo(newFileName);
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                //outputPackage.SaveAs(newFileInfo);
                //outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
            }
        }
        public static void VariableAmountQuery(string destinationFolder, List<string> Cautid, List<string> cautDesc, List<double> cautAmt)
        {
            string filePath = Path.Combine(destinationFolder, "QuerySheet.xlsx");
            FileInfo fileInfo = new FileInfo(filePath);

            // Load existing file or create a new one
            using (var outputPackage = File.Exists(filePath)? new ExcelPackage(fileInfo): new ExcelPackage())
            {
                var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Variable Amount");
                outputWorksheet.Cells[1, 1].Value = "HRID";
                outputWorksheet.Cells[1, 2].Value = "Description";
                outputWorksheet.Cells[1, 3].Value = "Amount";
                int row = 2;
                for (int i = 0; i <= Cautid.Count - 1; i++)
                {
                    outputWorksheet.Cells[row, 1].Value = Cautid[i];
                    outputWorksheet.Cells[row, 2].Value = cautDesc[i];
                    outputWorksheet.Cells[row, 3].Value = cautAmt[i];
                    row++;
                }
                string newFileName = Path.Combine(destinationFolder, "QuerySheet.xlsx");
                FileInfo newFileInfo = new FileInfo(newFileName);
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                //outputPackage.SaveAs(newFileInfo);
                //outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
            }
        }
        public static void RehireCasesQuery(string destinationFolder, List<string>CautionId, List<string> evtype, List<string> doj, List<string> nm, List<string> UACN)
        {
            string filePath = Path.Combine(destinationFolder, "QuerySheet.xlsx");
            FileInfo fileInfo = new FileInfo(filePath);
            using (var outputPackage = File.Exists(filePath) ? new ExcelPackage(fileInfo) : new ExcelPackage())
            {
                var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Re-Hire Cases");
                outputWorksheet.Cells[1, 1].Value = "HRID";
                outputWorksheet.Cells[1, 2].Value = "Event Type";
                outputWorksheet.Cells[1, 3].Value = "Date of joining";
                outputWorksheet.Cells[1, 4].Value = "Name";
                outputWorksheet.Cells[1, 5].Value = "UAN";
                int row = 2;
                for (int i = 0; i <= CautionId.Count - 1; i++)
                {
                    outputWorksheet.Cells[row, 1].Value = CautionId[i];
                    outputWorksheet.Cells[row, 2].Value = evtype[i];
                    outputWorksheet.Cells[row, 3].Value = doj[i];
                    outputWorksheet.Cells[row, 4].Value = nm[i];
                    outputWorksheet.Cells[row, 5].Value = UACN[i];
                    row++;
                }
                string newFileName = Path.Combine(destinationFolder, "QuerySheet.xlsx");
                FileInfo newFileInfo = new FileInfo(newFileName);
                outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                //outputPackage.SaveAs(newFileInfo);
                //outputPackage.SaveAsAsync(new FileInfo(destinationFolder));
            }
        }
    }
}
