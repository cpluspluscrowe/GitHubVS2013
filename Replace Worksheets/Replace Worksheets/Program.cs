using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
namespace Replace_Worksheets
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xl = new Excel.Application();
            xl.Visible = true;
            xl.Application.DisplayAlerts = false;
            DirectoryInfo diFac = new DirectoryInfo(@"C:\Users\CCrowe\Documents\AFCS Folder\Facilities");
            DirectoryInfo diNew = new DirectoryInfo(@"C:\Users\CCrowe\Documents\AFCS Folder\Fac_Replacements");
            foreach (var file in diFac.GetFiles("*"))
            {
                string facilityName = file.Name.Split(new[] { " - " }, StringSplitOptions.None).ElementAt(0);
                foreach (var replFile in diNew.GetFiles(String.Format("*{0}*", facilityName)))
                {
                    try
                    {
                        var exists1 = File.Exists(file.FullName);
                        var exists2 = File.Exists(replFile.FullName);
                        Excel.Workbook orig = xl.Workbooks.Open(file.FullName);
                        Excel.Workbook newWb = xl.Workbooks.Open(replFile.FullName);
                        Excel.Worksheet oldBom = orig.Sheets["Bill of Materials"];
                        Excel.Worksheet oldRes = orig.Sheets["Resources"];
                        Excel.Worksheet oldWBS = orig.Sheets["Work Breakdown Structure"];

                        Excel.Worksheet newBom = newWb.Sheets["Bill of Materials"];
                        Excel.Worksheet newRes = newWb.Sheets["Resources"];
                        Excel.Worksheet newWBS = newWb.Sheets["Work Breakdown Structure"];

                        oldBom.Name = "OldBOM";
                        oldRes.Name = "OldRes";
                        oldWBS.Name = "OldWBS";

                        newBom.Copy(oldBom);
                        newRes.Copy(oldRes);
                        newWBS.Copy(oldWBS);

                        oldBom.Delete();
                        oldRes.Delete();
                        oldWBS.Delete();

                        orig.Close(true);//change to true
                        newWb.Close(false);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(file.Name);
                    }
                }
            }
            Console.ReadLine();
        }
    }
}
