using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.SqlClient;

namespace NSN_Check
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xl = new Excel.Application();
            xl.Visible = true;
            string directory = "C:\\Users\\CCrowe\\Documents\\Facilities";
            DirectoryInfo di = new DirectoryInfo(directory);
            foreach (var file in di.GetFiles())
            {
                string workbookName = file.FullName;
                facilityName = "";
                SqlConnection conn = new SqlConnection("Server=(LocalDB)\\MSSQLLocalDB;Database=JCMS_Local_41;Integrated Security = true");
                conn.Open();
                string sql = @"DECLARE @ele_id nvarchar(50);
DECLARE @ele_name nvarchar(100);
DECLARE @ele_descr nvarchar(100);
DECLARE @type nvarchar(50);
DECLARE @subtype nvarchar(50);
DECLARE @FetchStatus int
DECLARE CA_cursor CURSOR  
	FOR select Element_Id, Element_Nbr, Element_Descr FROM Element WHERE Element_Id in 
	(SELECT Element_Id FROM Element_Hierarchy WHERE Parent_Element_Id In 
	(SELECT Element_Id FROM Element WHERE Element_Nbr = @facility)) ORDER BY Element_Nbr ASC;
	
OPEN CA_cursor  
FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name,@ele_descr
WHILE @@FETCH_STATUS = 0
	BEGIN
        SELECT @ele_id as 'CA', @ele_name as 'Name', @ele_descr as 'Description';
		SELECT NSN,Item_Name FROM NSN WHERE NSN_ID IN (SELECT NSN_Id from CA_NSN where Element_Id = @ele_id);
		FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name, @ele_descr
	END
CLOSE CA_cursor;
DEALLOCATE CA_cursor;";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Add("@projectId", facilityName);
                    SqlDataReader reader = cmd.ExecuteReader();
                    do
                    {
                        while (reader.Read())
                        {
                            string[] colNames = new string[2];
                            reader.GetValues(colNames);
                            
                        }
                    } while (reader.NextResult());
                    reader.Close();
                }
            }
        }
    }
}
