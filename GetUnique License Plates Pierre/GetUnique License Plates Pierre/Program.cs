using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace GetUnique_License_Plates_Pierre
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xl = new Excel.Application();
            xl.Visible = true;
            SortedSet<string> locations = new SortedSet<string>();
            string excelDocPath = @"C:\Users\CCrowe\Documents\Traffic\Carbee\Plate Matching\Pierre SDDOT Plate Matching.xlsm";
            Excel.Workbook wb = xl.Workbooks.Open(excelDocPath);
            Excel.Worksheet ws = wb.Sheets["Raw Data"];
            for (int i = 2; i <= ws.UsedRange.Rows.Count; i++)
            {
                if (ws.Range["B" + i.ToString()].Value != null)
                {
                    locations.Add(ws.Range["B" + i.ToString()].Value);
                }
            }
            int cnt = 2;
            foreach (var location in locations)
            {
                ws.Range["F" + cnt.ToString()].Value = location;
                cnt += 1;
            }
        }
    }
}
