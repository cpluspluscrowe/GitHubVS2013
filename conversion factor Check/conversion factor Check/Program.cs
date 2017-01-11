using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using LinqToExcel;
using LinqToExcel.Attributes;
namespace conversion_factor_Check
{
    class CheckData
    {
        public string Description;
        public float Cube;
        public float Weight;
        public float Price;
        public string UOM;
        public int Conversion_Factor;
        public string UOI;

        public string Alert;
        public string NSNSheetValue;
        public string CurrentValue;
        public CheckData(string alert,string nsnSheetValue,string currentValue)
        {
            this.Alert = alert;
            this.NSNSheetValue = nsnSheetValue;
            this.CurrentValue = currentValue; 
        }
    }
    internal class NSNSheet
    {
        [ExcelColumn("AFCS ID")]
        public string AFCS_ID { get; set; }

        [ExcelColumn("FLIS NSN")]
        public string FLIS_NSN { get; set; }

        [ExcelColumn("FSC")]
        public string FSC { get; set; }

        [ExcelColumn("Country Code")]
        public string Country_Code { get; set; }

        [ExcelColumn("Item Number")]
        public string Item_Number { get; set; }

        [ExcelColumn("Type")]
        public string Type { get; set; }

        [ExcelColumn("AFCS Nomenclature")]
        public string AFCS_Nomenclature { get; set; }

        [ExcelColumn("Units")]
        public string Units { get; set; }

        [ExcelColumn("Unit of Measure")]
        public string Unit_of_Measure { get; set; }

        [ExcelColumn("Conversion Factor")]
        public string Conversion_Factor { get; set; }

        [ExcelColumn("Unit Of Issue")]
        public string Unit_Of_Issue { get; set; }

        [ExcelColumn("Price")]
        public string Price { get; set; }

        [ExcelColumn("Weight")]
        public string Weight { get; set; }

        [ExcelColumn("Volume")]
        public string Volume { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xl = new Excel.Application();
            xl.Visible = true;
            string facilityString = @"C:\Users\CCrowe\Documents\AFCS Folder\Old_Scope_Facilities";
            DirectoryInfo di = new DirectoryInfo(facilityString);
            var excel = new ExcelQueryFactory(@"C:\Users\CCrowe\Documents\AFCS Folder\Facilities\1211000AA - FORWARD AREA REFUELING WITH 1000 GALLON STORAGE.xlsm")
            {
                DatabaseEngine = LinqToExcel.Domain.DatabaseEngine.Ace,
                TrimSpaces = LinqToExcel.Query.TrimSpacesType.Both,
                UsePersistentConnection = true,
                ReadOnly = true
            };
            var planets = from p in excel.Worksheet<NSNSheet>("NSN")
                          select p;
            foreach (var file in di.GetFiles("*"))
            {
                Excel.Workbook wb = xl.Workbooks.Open(file.FullName);
                Excel.Worksheet bom = wb.Sheets["Bill of Materials"];
                Excel.Worksheet nsn = wb.Sheets["NSN"];
                int lr = bom.UsedRange.Rows.Count;
                for (int i = 3; i <= lr; i++)
                {
                    if (bom.Range["G" + i.ToString()].Value != null)
                    {
                        string itemNsn = bom.Range["G" + i.ToString()].Value;
                        List<CheckData> chkList = new List<CheckData>();
                        try
                        {
                            var row = planets.First(p => p.AFCS_ID == itemNsn);

                            chkList.Add(new CheckData("AFCS Description Error", bom.Range["H" + i.ToString()].Value.ToString(), row.AFCS_Nomenclature));
                            chkList.Add(new CheckData("AFCS Cube Error", bom.Range["I" + i.ToString()].Value.ToString(), row.Volume));
                            chkList.Add(new CheckData("AFCS Weight Error", bom.Range["J" + i.ToString()].Value.ToString(), row.Weight));
                            chkList.Add(new CheckData("AFCS Price Error", bom.Range["K" + i.ToString()].Value.ToString(), row.Price));
                            chkList.Add(new CheckData("AFCS Unit Of Measure Error", bom.Range["L" + i.ToString()].Value.ToString(), row.Unit_of_Measure));
                            chkList.Add(new CheckData("AFCS Conversion Factor Error", bom.Range["N" + i.ToString()].Value.ToString(), row.Conversion_Factor));
                            chkList.Add(new CheckData("AFCS Unit of Issue Error", bom.Range["O" + i.ToString()].Value.ToString(), row.Unit_Of_Issue));
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("NSN not found in worksheet:" + itemNsn);
                        }
                        foreach (var checkItem in chkList)
                        {
                            if (checkItem.CurrentValue != checkItem.NSNSheetValue)
                            {
                                if (checkItem.NSNSheetValue.ToString() != "0")
                                {
                                    Console.WriteLine(wb.Name + " ;row:" + i.ToString() + "; \n\t" + checkItem.Alert + " ;\n\t\tCurrentValue:" + checkItem.CurrentValue + " ;\n\t\tSheet Value:" + checkItem.NSNSheetValue);
                                }
                            }
                        }
                    }
                }
                wb.Close(false);
            }
            Console.ReadLine();
            xl.Quit();
        }
    }
}
