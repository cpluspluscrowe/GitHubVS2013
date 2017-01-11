using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
namespace alt_sheets
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xl = new Excel.Application();
            xl.Visible = true;
            Excel.Workbook main = xl.Workbooks.Open(@"C:\Users\CCrowe\Documents\Traffic\Weiss\Save dbf files\PrioritizationDelayAccessibilityResults_For_Chad.xlsx");
            Excel.Worksheet template = main.Sheets["ALT_1Access15"];
            string folder = @"C:\Users\CCrowe\Documents\Traffic\Weiss\Save dbf files";
            DirectoryInfo di = new DirectoryInfo(folder);
            foreach(var file in di.GetFiles("*DBF*")){
                Excel.Workbook wb = xl.Workbooks.Open(file.FullName);
                Excel.Worksheet ws = wb.Sheets[1];
                template.Copy(After:main.Sheets[main.Sheets.Count]);
                Excel.Worksheet newWs = main.Sheets[main.Sheets.Count];
                newWs.Name = ws.Name;
                newWs.Range["B1:B" + newWs.UsedRange.Rows.Count.ToString()].ClearContents();
                newWs.Range["E814"].Select();
                foreach (Excel.Range cell in ws.Range["B1:B" + newWs.UsedRange.Rows.Count.ToString()].Cells)
                {
                    string address = cell.Address;
                    newWs.Range[cell.Address].Value = cell.Value;
                }
                wb.Close(false);
            }
        }
    }
}
