using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
namespace default_transcad_program
{
    class Program
    {
        static void Main(string[] args)
        {
            Open_Table();
        }
        /// <summary>
        /// Open a table in the tutorial folder, and select some rows by an SQL-like expression
        /// </summary>
        static void Open_Table()
        {
            DirectoryInfo di = new DirectoryInfo(@"C:\Users\CCrowe\Documents\Traffic\Weiss\TransCAD files\From_Jweiss");
            string fileLocation = @"C:\Users\CCrowe\Documents\Traffic\Weiss\TransCAD files\PrioritizationDelayAccessibilityResults.xlsx";
            Excel.Application xl = new Excel.Application();
            xl.Visible = true;
            Excel.Workbook wb = xl.Workbooks.Open(fileLocation);
            Excel.Worksheet main = wb.Sheets["Summary"];
            CaliperForm.Connection Conn = new CaliperForm.Connection { MappingServer = "TransCAD" };
            Boolean opened = false;
            try
            {
                opened = Conn.Open();
                if (opened)
                {
                    // You must declare dk as "dynamic" or the compiler will throw an error
                    dynamic dk = Conn.Gisdk;
                    string tutorial_folder = dk.Macro("G30 Tutorial Folder") as string;
                    // Open a table
                    foreach(var file in di.GetFiles("*.bin"))
                    {
                        Console.WriteLine(file.Name);
                        string table_name = dk.OpenTable("sales", "ffb", new Object[] { file.FullName, null });
                        object[] fields = dk.GetFields(table_name, "All");
                        var field_names = fields[0] as object[];
                        var field_specs = fields[1] as object[];
                        // select some records and get record values
                        dk.SetView(table_name);
                        int num_rows = dk.GetRecordCount(table_name, null);
                        string query = "select * where ID1 > -1";
                        int num_found = dk.SelectByQuery("large towns", "several", query, new Object[] { new Object[] { "Index Limit", 0 } });
                        if (num_found > 0)
                        {
                            string view_set = table_name + "|large towns";
                            object[] sort_order = null;
                            object[] options = null;
                            string order = "Row";
                            string first_record = dk.GetFirstRecord(view_set, null);
                            DataTable dt = new DataTable();
                            foreach (var name in field_names)
                            {
                                dt.Columns.Add(name.ToString(),typeof(float));
                            }
                            int cnt = 1;
                            List<List<object>> arr = new List<List<object>>();
                            foreach (object[] row in dk.GetRecordsValues(view_set, first_record, field_names, sort_order, num_found, order, null))
                            {
                                for (int i = 0; i <= row.Length-1; i++)
                                {
                                    if (Convert.IsDBNull(row[i]))
                                    {
                                        row[i] = 0.0;
                                    }
                                }
                                dt.Rows.Add(row.ToArray());
                            }
                            var Tot_V_Dist_T = dt.AsEnumerable().Sum(x => x.Field<float>(12));
                            var AB_VHT = dt.AsEnumerable().Sum(x => x.Field<float>(13));
                            var BA_VHT = dt.AsEnumerable().Sum(x => x.Field<float>(14));
                            var Tot_VHT = dt.AsEnumerable().Sum(x => x.Field<float>(15));
                            var AB_Delay = dt.AsEnumerable().Sum(x => x.Field<float>(16));
                            var BA_Delay = dt.AsEnumerable().Sum(x => x.Field<float>(17));
                            var Tot_Delay = dt.AsEnumerable().Sum(x => x.Field<float>(18));
                            string altString = "Alt " + file.Name.Replace("ALT_","").Replace("_LinkFlowTot#1.bin","");
                            int col;
                            for(int j = 2;j<=200;j++){
                                if(main.Cells[5,j].Value == altString){
                                    main.Cells[5, j].Style = "Good";
                                    main.Cells[6,j].Value = Tot_V_Dist_T;
                                    main.Cells[9,j].Value = AB_VHT;
                                    main.Cells[12,j].Value = BA_VHT;
                                    main.Cells[13,j].Value = Tot_VHT;
                                    main.Cells[14,j].Value = AB_Delay;
                                    main.Cells[15,j].Value = BA_Delay;
                                    main.Cells[16,j].Value = Tot_Delay;
                                    break;
                                }
                            }
                        }
                        dk.CloseView(table_name);
                        Console.Out.WriteLine();
                    }
                    Conn.Close();
                    Console.ReadLine();
                }
            }
            catch (System.Exception error)
            {
                Console.Out.WriteLine(error.Message);
            }
        }

    }
}
