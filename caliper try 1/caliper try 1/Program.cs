using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CaliperForm;

namespace caliper_try_1
{
    class Program
    {
        static void Main(string[] args)
        {
            CaliperForm.Connection conn = new CaliperForm.Connection { MappingServer = "TransCAD" };
            Boolean opened = false;
            opened = conn.Open();
            if (opened)
            {
                dynamic dk = conn.Gisdk;
                string table_name = dk.OpenTable("ALT_1_LinkFlowTot#1", "ffb", new Object[] { @"C:\Users\CCrowe\Documents\Traffic\Weiss\TransCAD files\From_Jweiss\ALT_1_LinkFlowTot#1.bin", null });
                object[] fields = dk.GetFields(table_name, "All");
                var field_names = fields[0] as object[];
                var field_specs = fields[1] as object[];
                object[] sort_order = null;
                object[] options = null;
                string order = "Row";
                string query = "select * where ID1 > -1";
                int num_found = dk.SelectByQuery("ID1", "subset", query, new Object[] { new Object[] { "Index Limit", 0 } });
                string view_set = table_name + "|ID1";
                string first_record = dk.GetFirstRecord(table_name + "|",new string[]{"ID1"});
                foreach (object[] row in dk.GetRecordsValues(view_set, first_record, field_names, sort_order, num_found, order, null))
                {
                    string row_values = string.Join(",", row);
                    int p = 5;
                }
                dk.CloseView(table_name);
            }
            conn.Close();
        }
    }
}
