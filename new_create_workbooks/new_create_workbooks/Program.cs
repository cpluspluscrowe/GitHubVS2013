using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Drawing;
//This program creates a majority of the workbooks for AFCS
//There is a folder called workbooks to create.xlsm located in your documents that it grabs facility names from

namespace new_create_workbooks
{
    public class Nsn
    {
        public string Number;
        public string Description;
        public string Uom;
        public string Quantity;

        public Nsn(string number, string description, string uom, string quantity)
        {
            this.Number = number;
            this.Description = description;
            this.Uom = uom;
            this.Quantity = quantity;
        }

        public Nsn()
        {
            //empty constructor
        }
    }

    public class We
    {
        public string Division;
        public string Section;
        public string LineItem;
        public string ShortCode;
        public string LongCode;
        public List<Nsn> NsnList = new List<Nsn>();
        public double Quantity;
        public string Uom;
        public string GeneralManHours;
        public string BuilderManHours;
        public string ElectricianManHours;
        public string EngServicesManHours;
        public string EquipOperatorManHours;
        public string SteelWorkerManHours;
        public string UtilityManHours;
        public string Equipment1Description;
        public string Equipment1Hours;
        public string Equipment2Description;
        public string Equipment2Hours;
        public string Equipment3Description;
        public string Equipment3Hours;
        public string Equipment4Description;
        public string Equipment4Hours;
        public string Equipment5Description;
        public string Equipment5Hours;
        public string Equipment6Description;
        public string Equipment6Hours;
        public string Equipment7Description;
        public string Equipment7Hours;
        public string Equipment8Description;
        public string Equipment8Hours;
        public string Equipment9Description;
        public string Equipment9Hours;
        public string Equipment10Description;
        public string Equipment10Hours;

        public We(string division, string section, string lineItem, string shortCode)
        {
            this.Division = division;
            this.Section = section;
            this.LineItem = lineItem;
            this.ShortCode = shortCode;
        }
    }

    class Ca
    {
        public string Number;
        public string Description;
        public List<We> WeList = new List<We>();

        public Ca(string number, string description)
        {
            this.Number = number;
            this.Description = description;
        }
    }

    class Program
    {
        public static Excel.Application xl;
        static void Main(string[] args)
        {
            xl = new Excel.Application();
            xl.AskToUpdateLinks = false;
            xl.Visible = true;
            xl.DisplayAlerts = false;
            Excel.Workbook wb = xl.Workbooks.Open(@"C:\Create_Workbooks\workbooks to create.xlsm");
            Excel.Worksheet ws = wb.Sheets["Sheet1"];
            Excel.Workbook exWb = xl.Workbooks.Open(@"C:\Users\CCrowe\Documents\AFCS Folder\Template\Template.xlsm");
            //This next part will be needed in all worksheets.  It will create a DataTable from the worksheet CostBook.  Must be done before I loop through all the worksheets
            //Looks into the already created and formatted workbook and examines its CostBook.  It pulls data from that CostBook
            Excel.Worksheet cb = exWb.Sheets["CostBook"];
            object[,] data = cb.Range["A2:E1965"].Value2;
            DataTable cbTable = new DataTable();
            DataRow dtRow;
            for (int cCnt = 1; cCnt <= 5; cCnt++) //This records the costbook data
            {
                DataColumn column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = (string) data[1, cCnt];
                cbTable.Columns.Add(column);
                for (int rCnt = 2; rCnt < data.GetLength(0); rCnt++)
                {
                    if (cCnt == 1)
                    {
                        dtRow = cbTable.NewRow();
                        dtRow[column.ColumnName] = (string) data[rCnt, cCnt];
                        cbTable.Rows.Add(dtRow);
                    }
                    else
                    {
                        dtRow = cbTable.Rows[rCnt - 2];
                        dtRow[column.ColumnName] = (string) data[rCnt, cCnt];
                    }

                }
            }
            int rc = cbTable.Rows.Count; //how many rows are in the above DataTable
            wb.Close(false);
            exWb.Close(false);
            for (int i = 2; i <= 26; i++) //These are the rows in "workbooks to create.xlsm" I will traverse
                // ws.UsedRange.Rows.Count //Edit this for which workbooks (names and descriptions) you wish to create
            {
                wb = xl.Workbooks.Open(@"C:\Create_Workbooks\workbooks to create.xlsm");
                //I close then reopen both files.  This is an attempt to help the Interop Services from losing resources or having memory problems
                ws = wb.Sheets["Sheet1"];
                exWb = xl.Workbooks.Open(@"C:\Users\CCrowe\Documents\AFCS Folder\Template\Template.xlsm");

                string fileNumber = ws.get_Range("A" + i.ToString()).Value2;
                //Get the JCMS Description
                SqlConnection conn =
                    new SqlConnection(
                        "Server=OME-CND6435DR5;Database=JCMS_Local_41;Integrated Security = true");
                conn.Open();
                string sql = @"SELECT Element_Descr
                FROM Element where Element_Nbr = @fileNumber;
                ";
                string fileDescription = "";
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Add("@fileNumber",fileNumber);
                    SqlDataReader reader = cmd.ExecuteReader();
                    do
                    {
                        while (reader.Read())
                        {
                            string[] colNames = new string[2];
                            reader.GetValues(colNames);
                            fileDescription = colNames[0];
                        }
                    } while (reader.NextResult());
                    reader.Close();
                }
                string fileName = fileNumber + " - " + fileDescription;
                string fullPath = Path.Combine(@"C:\Create_Workbooks\New Workbooks",
                    fileName,
                    fileName + ".xlsx");
                bool exists = File.Exists(fullPath);
                if (!exists)
                {

                    wb.Close(false);
                    if (!File.Exists(@"C:\Create_Workbooks\New Workbooks" + fileName))
                    {
                        Excel.Workbook newWorkbook = xl.Workbooks.Add();
                        newWorkbook.Worksheets.Add();
                        newWorkbook.Worksheets.Add();
                        newWorkbook.Worksheets.Add();
                        newWorkbook.Worksheets.Add();
                        newWorkbook.Worksheets.Add();
                        Excel.Worksheet facility = newWorkbook.Sheets["Sheet6"];
                        facility.Name = "Facility";
                        Excel.Worksheet wbs = newWorkbook.Sheets["Sheet5"];
                        wbs.Name = "Work Breakdown Structure";
                        Excel.Worksheet draw = newWorkbook.Sheets["Sheet4"];
                        draw.Name = "Drawings";
                        Excel.Worksheet bom = newWorkbook.Sheets["Sheet1"];
                        bom.Name = "Bill of Materials";
                        Excel.Worksheet res = newWorkbook.Sheets["Sheet2"];
                        res.Name = "Resources";
                        Excel.Worksheet togs = newWorkbook.Sheets["Sheet3"];
                        togs.Name = "TOGS";

                        //move the NSN, CostBook, and Equipment tabs into each worksheet
                        Excel.Worksheet nsnSheet = exWb.Sheets["NSN"];
                        Excel.Worksheet costBook = exWb.Sheets["CostBook"];
                        Excel.Worksheet equipment = exWb.Sheets["Equipment"];
                        nsnSheet.Copy(After: newWorkbook.Sheets[newWorkbook.Sheets.Count]);
                        costBook.Copy(After: newWorkbook.Sheets[newWorkbook.Sheets.Count]);
                        equipment.Copy(After: newWorkbook.Sheets[newWorkbook.Sheets.Count]);

                        Excel.Worksheet exFac = exWb.Sheets["Facility"];
                        Excel.Range source = exFac.Range["A1:F1"];
                        Excel.Range dest = facility.Range["A1:F1"];
                        source.Copy();
                        facility.Activate();
                        facility.get_Range("A1").Select();
                        facility.Paste();

                        //FACILITY
                        source = exFac.Range["A1:B36"];
                        dest = facility.Range["A1:B36"];
                        source.Copy();
                        facility.get_Range("A1").Select();
                        facility.Paste();
                        facility.Range["A2"].Value = fileNumber;
                        facility.Range["B2"].Value = fileDescription;
                        facility.Range["A13:B30"].Style = "Normal";
                        facility.Range["B2"].Style = "Normal";
                        facility.Range["B2"].Font.Bold = true;
                        facility.Range["A13:B30"].ClearContents();
                        Excel.Range dateComment = facility.Range["A13:B30"];
                        if (dateComment.Comment != null)
                        {
                            dateComment.Comment.Delete();
                        }
                        dateComment = facility.Range["B2"];
                        if (dateComment.Comment != null)
                        {
                            dateComment.Comment.Delete();
                        }

                        facility.Range["B4"].Formula = "='Bill of Materials'!T2";
                        facility.Range["B5"].Formula = "='Bill of Materials'!S2";
                        facility.Range["B6"].Formula = "='Bill of Materials'!R2";
                        facility.Range["B8"].Formula = "=Resources!J2";
                        facility.Range["B9"].Formula = "=Resources!K2 + Resources!L2 + Resources!O2 + Resources!P2";
                        facility.Range["B9"].Formula = "=Resources!K2 + Resources!L2 + Resources!O2 + Resources!P2";
                        facility.Range["B10"].Formula = "=Resources!N2";



                        facility.Range["A2"].Style = "Untouched from JCMS";
                        facility.Range["B4:B6"].Style = "Untouched from JCMS";
                        facility.Range["B8:B10"].Style = "Untouched from JCMS";
                        facility.Range["E2:F2"].Style = "Untouched from JCMS";
                        facility.Range["A1:F1"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                            Excel.XlLineStyle.xlContinuous;
                        facility.Columns["C:F"].Font.Bold = true;
                        //Fit columns to match the 95% widths
                        for (int c = 1; c <= 6; c++) //In Excel the first column is 1 not 0
                        {
                            double colWidth = exFac.Columns[c].ColumnWidth;
                            facility.Columns[c].ColumnWidth = colWidth;
                        }
                        //SQL for Facility
                        sql = string.Format("SELECT Element_Detail from Element WHERE Element_Nbr = '{0}'",
                            fileNumber);
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            SqlDataReader reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                object readValue = reader.GetString(0);
                                facility.Range["A13"].MergeArea.Style = "Untouched from JCMS";
                                facility.Range["A13"].Value2 = readValue.ToString();
                                facility.Range["A13:A30"].HorizontalAlignment = Excel.Constants.xlLeft;
                                facility.Range["A13:A30"].VerticalAlignment = Excel.Constants.xlTop;
                                facility.Range["A13"].WrapText = true;
                            }
                            reader.Close();
                        } //end of getting element_detail for the facility sheet

                        //Now get all drawings associated with the file (usually the facility, could be a component if it is a component file)
                        sql = string.Format(
                            @"DECLARE @ele_id nvarchar(50);
                DECLARE @type nvarchar(50);
                DECLARE @facility nvarchar(100);
                SET @facility = '{0}';
                SET @ele_id = (SELECT Element_Id from Element WHERE Element_Nbr = @facility);
                SET @type = (SELECT Element_Type from Element where Element_Nbr = @facility);
                SELECT DISTINCT File_Nbr, File_Descr from JCMS_File WHERE File_Id in 
                    (SELECT File_Id FROM JCMS_File_Owner WHERE File_Owner_Id = @ele_id  AND File_Owner_Obj_Type = @type) 
                    AND File_Class = 'DRAWING' ORDER BY File_Nbr ASC ;
                ", fileNumber);
                        int facCurRow1 = 2;
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            SqlDataReader reader = cmd.ExecuteReader();
                            do
                            {
                                while (reader.Read())
                                {
                                    string[] colNames = new string[2];
                                    reader.GetValues(colNames);
                                    facility.Cells[facCurRow1, 3].Value = colNames[0];
                                    facility.Cells[facCurRow1, 4].Value = colNames[1];
                                    facCurRow1 += 1;
                                }
                            } while (reader.NextResult());
                            reader.Close();
                        }


                        //Now try to find any support files
                        sql = string.Format(
                            @"DECLARE @ele_id nvarchar(50);
                            DECLARE @type nvarchar(50);
                            DECLARE @subtype nvarchar(50);
                            DECLARE @facility nvarchar(100);
                            SET @facility = '{0}';
                            SET @type = (SELECT Element_Type from Element where Element_Nbr = @facility);
                                            SELECT DISTINCT File_Nbr, File_Descr from JCMS_File WHERE File_Id in (SELECT File_Id FROM JCMS_File_Owner WHERE File_Owner_Id = (SELECT Element_Id FROM Element WHERE Element_Nbr = @facility) AND File_Owner_Obj_Type = @type )  AND File_Class = 'SUPPORT' ORDER BY File_Nbr ASC ;
				                            DECLARE @ele_name nvarchar(100);
                                            DECLARE @ele_descr nvarchar(100);
                                            DECLARE @FetchStatus int
                                            DECLARE CA_cursor CURSOR  
	                                            FOR select Element_Id, Element_Nbr, Element_Descr FROM Element WHERE Element_Id in 
	                                            (SELECT Element_Id FROM Element_Hierarchy WHERE Parent_Element_Id In 
	                                            (SELECT Element_Id FROM Element WHERE Element_Nbr = @facility)) ORDER BY Element_Nbr ASC;

                                            OPEN CA_cursor  
                                            FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name,@ele_descr
                                            WHILE @@FETCH_STATUS = 0
	                                            BEGIN
                                                    SELECT @ele_id, @ele_name, @ele_descr;
						                            SET @subtype = (SELECT Element_Type from Element where Element_Nbr = @ele_name);
		                                            SELECT DISTINCT File_Nbr, File_Descr from JCMS_File WHERE File_Id in (SELECT File_Id FROM JCMS_File_Owner WHERE File_Owner_Id = @ele_id AND File_Owner_Obj_Type = @subtype) AND File_Class = 'SUPPORT' ORDER BY File_Nbr ASC ;
                                                    FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name, @ele_descr
	                                            END
                                            CLOSE CA_cursor;
                                            DEALLOCATE CA_cursor;
                ", fileNumber);
                        facility.Range["E2"].Value = fileNumber;
                        facility.Range["F2"].Value = "MS Project Schedule";
                        facility.Range["E2:F2"].Style = "Complete";
                        int facCurRow = 3;
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            SqlDataReader reader = cmd.ExecuteReader();
                            do
                            {
                                while (reader.Read())
                                {
                                    cmd.CommandTimeout = 600;
                                    string[] colNames = new string[3];
                                    reader.GetValues(colNames);
                                    if (colNames[2] == null) //I only select two values for facilityings, three for CA
                                    {
                                        facility.Cells[facCurRow, 5].Value = colNames[0];
                                        facility.Cells[facCurRow, 6].Value = colNames[1];
                                        facCurRow += 1;
                                    }
                                    /*else
                            {
                                facility.Cells[facCurRow, 1].Value = colNames[1];
                                facility.Cells[facCurRow, 2].Value = colNames[2];
                                facCurRow += 1;
                            }*/
                                    //I did not specify which CA each of these support files came from.  I have just listed all of them on the facility worksheet
                                }
                            } while (reader.NextResult());
                            reader.Close();
                        }



                        //WORK BREAKDOWN STRUCTURE
                        Excel.Worksheet wbsSource = exWb.Sheets["Work Breakdown Structure"];
                        source = wbsSource.Range["A1:E1"];
                        dest = wbs.Range["A1:E1"];
                        source.Copy();
                        wbs.Activate();
                        wbs.Range["A1"].Select();
                        wbs.Paste();

                        source = wbsSource.Range["A115:B119"];
                        dest = wbs.Range["A396:B400"];
                        source.Copy();
                        wbs.Activate();
                        wbs.Range["A396"].Select();
                        wbs.Paste();


                        int topOfLegend = 391;
                        source = wbsSource.Range["A28:E28"];
                        dest = wbs.Range["A392:E392"];
                        source.Copy();
                        wbs.Activate();
                        wbs.Range["A392"].Select();
                        wbs.Paste();

                        wbs.Rows[392].RowHeight = wbsSource.Rows[28].RowHeight;
                        (wbs.Rows[392] as Excel.Range).WrapText = false;
                        for (int a = 2; a <= 91; a++)
                        {
                            wbs.Range["E" + a.ToString()].Formula = "=IF($D" + a.ToString() + "=\"\",\"\",VLOOKUP($D" +
                                                                    a.ToString() +
                                                                    ",CostBook!$A$2:$W$3000,4,FALSE)&\" - \"&VLOOKUP($D" +
                                                                    a.ToString() + ",CostBook!$A$2:$W$3000,5,FALSE))";
                        }
                        wbs.Columns[3].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        wbs.Columns["A:D"].Font.Bold = true;
                        //Fit columns to match the 95% widths
                        for (int c = 1; c <= 5; c++) //In Excel the first column is 1 not 0
                        {
                            double colWidth = wbsSource.Columns[c].ColumnWidth;
                            wbs.Columns[c].ColumnWidth = colWidth;
                        }
                        wbs.Range["A396:B396"].Font.Bold = false;
                        //DRAWINGS
                        Excel.Worksheet drawSource = exWb.Sheets["Drawings"];
                        source = drawSource.Range["A1:D1"];
                        dest = draw.Range["A1:D1"];
                        source.Copy();
                        draw.Activate();
                        draw.Range["A1"].Select();
                        draw.Paste();

                        source = drawSource.Range["A250:B254"];
                        dest = draw.Range["A258:B262"];
                        source.Copy();
                        draw.Activate();
                        draw.Range["A258"].Select();
                        draw.Paste();


                        topOfLegend = 253;
                        source = drawSource.Range["A17:D17"];
                        dest = draw.Range["A254:D254"];
                        source.Copy();
                        draw.Activate();
                        draw.Range["A254"].Select();
                        draw.Paste();
                        //Fit columns to match the 95% widths
                        draw.Columns["A:D"].Font.Bold = true;
                        for (int c = 1; c <= 4; c++) //In Excel the first column is 1 not 0
                        {
                            double colWidth = drawSource.Columns[c].ColumnWidth;
                            draw.Columns[c].ColumnWidth = colWidth;
                        }
                        //Sql statements to fill Drawings
                        sql = string.Format(
                            @"DECLARE @ele_id nvarchar(50);
                        DECLARE @ele_name nvarchar(100);
                        DECLARE @ele_descr nvarchar(100);
                        DECLARE @type nvarchar(50);
                        DECLARE @subtype nvarchar(50);
                        DECLARE @FetchStatus int
                        DECLARE CA_cursor CURSOR  
	                        FOR select Element_Id, Element_Nbr, Element_Descr FROM Element WHERE Element_Id in 
	                        (SELECT Element_Id FROM Element_Hierarchy WHERE Parent_Element_Id In 
	                        (SELECT Element_Id FROM Element WHERE Element_Nbr = '{0}')) ORDER BY Element_Nbr ASC;
	
                        OPEN CA_cursor  
                        FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name,@ele_descr
                        WHILE @@FETCH_STATUS = 0
	                        BEGIN
                                SELECT @ele_id, @ele_name, @ele_descr
                                SET @subtype = (SELECT Element_Type from Element where Element_Nbr = @ele_name);
		                        SELECT DISTINCT File_Nbr, File_Descr from JCMS_File WHERE File_Id in (SELECT File_Id FROM JCMS_File_Owner WHERE File_Owner_Id = @ele_id AND File_Owner_Obj_Type = @subtype) AND File_Class = 'DRAWING' ORDER BY File_Nbr ASC ;
		                        FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name, @ele_descr
	                        END
                        CLOSE CA_cursor;
                        DEALLOCATE CA_cursor;
                ", fileNumber);
                        int curRow = 2;
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            SqlDataReader reader = cmd.ExecuteReader();
                            do
                            {
                                while (reader.Read())
                                {
                                    string[] colNames = new string[3];
                                    reader.GetValues(colNames);
                                    if (colNames[2] == null) //I only select two values for drawings, three for CA
                                    {
                                        draw.Cells[curRow, 3].Value = colNames[0];
                                        draw.Cells[curRow, 4].Value = colNames[1];
                                        curRow += 1;
                                    }
                                    else
                                    {
                                        draw.Cells[curRow, 1].Value = colNames[1];
                                        draw.Cells[curRow, 2].Value = colNames[2];
                                        curRow += 1;
                                    }
                                }
                            } while (reader.NextResult());
                            reader.Close();
                        }
                        draw.Range["A58:B58"].Font.Bold = false;
                        RemoveEmptyRows(topOfLegend, draw);
                        draw.Range["A1:D1"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                            Excel.XlLineStyle.xlContinuous;
                        //BILL OF MATERIALS
                        Excel.Worksheet bomSource = exWb.Sheets["Bill of Materials"];
                        source = bomSource.Range["A1:T2"];
                        dest = bom.Range["A1:T2"];
                        source.Copy();
                        bom.Activate();
                        bom.Range["A1"].Select();
                        bom.Paste();
                        bom.Rows[1].RowHeight = bomSource.Rows[1].RowHeight;

                        source = bomSource.Range["A471:B475"];
                        dest = bom.Range["A604:B608"];
                        source.Copy();
                        bom.Activate();
                        bom.Range["A604"].Select();
                        bom.Paste();

                        topOfLegend = 600;

                        source = bomSource.Range["A462:T462"];
                        dest = bom.Range["A600:T600"];
                        source.Copy();
                        bom.Activate();
                        bom.Range["A600"].Select();
                        bom.Paste();

                        bom.Rows[600].RowHeight = bomSource.Rows[471].RowHeight;
                        (bom.Rows[600] as Excel.Range).WrapText = false;
                        for (int rn = 600; rn < 604; rn++)
                        {
                            bom.Rows[rn].RowHeight = exWb.Sheets["Bill of Materials"].Rows[rn + (471 - 471)].RowHeight;
                        }
                        bom.Range["R2"].Formula =
                            "=SUM(INDIRECT(\"R3:R\" & MATCH(\"Optional Construction Activities\", A:A, 0) - 1))";
                        bom.Range["S2"].Formula =
                            "=SUM(INDIRECT(\"S3:S\" & MATCH(\"Optional Construction Activities\", A:A, 0) - 1))";
                        bom.Range["T2"].Formula =
                            "=SUM(INDIRECT(\"T3:T\" & MATCH(\"Optional Construction Activities\", A:A, 0) - 1))";
                        bom.Columns[2].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        bom.Columns[4].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        bom.Columns[6].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        bom.Columns[16].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        bom.Columns[17].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        bom.Columns[20].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        bom.Columns["A:E"].Font.Bold = true;
                        bom.Columns["G"].Font.Bold = true;
                        bom.Columns["M"].Font.Bold = true;
                        //Fit columns to match the 95% widths
                        for (int c = 1; c <= 20; c++) //In Excel the first column is 1 not 0
                        {
                            double colWidth = bomSource.Columns[c].ColumnWidth;
                            bom.Columns[c].ColumnWidth = colWidth;
                        }

                        //Go into CA Workbooks, grab WE Information, and populate BOM
                        sql = string.Format(
                            @"DECLARE @ele_id nvarchar(50);
                DECLARE @ele_name nvarchar(100);
                DECLARE @ele_descr nvarchar(100);
                DECLARE @FetchStatus int
                DECLARE CA_cursor CURSOR  
	                FOR select Element_Id, Element_Nbr, Element_Descr FROM Element WHERE Element_Id in 
	                (SELECT Element_Id FROM Element_Hierarchy WHERE Parent_Element_Id In 
	                (SELECT Element_Id FROM Element WHERE Element_Nbr = '{0}')) ORDER BY Element_Nbr ASC;

                OPEN CA_cursor  
                FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name,@ele_descr
                WHILE @@FETCH_STATUS = 0
	                BEGIN
                        SELECT @ele_id, @ele_name, @ele_descr
		                FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name, @ele_descr
	                END
                CLOSE CA_cursor;
                DEALLOCATE CA_cursor;
                ", fileNumber);

                        int curBomRow = 3;
                        int curWbsRow = 2;
                        int curResRow = 3;
                        int maxFurthest = 3;
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            SqlDataReader reader = cmd.ExecuteReader();
                            do
                            {
                                while (reader.Read())
                                {
                                    string[] colNames = new string[3];
                                    reader.GetValues(colNames);
                                    Ca ca = new Ca(colNames[1], colNames[2]);
                                    DirectoryInfo caWeFolder =
                                        new DirectoryInfo(@"C:\Users\CCrowe\Documents\AFCS Folder\CA");
                                    foreach (var file in caWeFolder.GetFiles("*" + ca.Number + "*.xlsx"))
                                    {
                                        Excel.Workbook caWb = xl.Workbooks.Open(file.FullName);
                                        Excel.Worksheet weWs = caWb.Sheets["Work Element"];
                                        for (int row = 1; row <= weWs.UsedRange.Rows.Count; row++)
                                        {
                                            if (weWs.Range["B" + row.ToString()].Value != null)
                                            {
                                                System.Drawing.Color topCellColor =
                                                    System.Drawing.ColorTranslator.FromOle(
                                                        (int) ((double) weWs.Range["B" + row.ToString()].Interior.Color));
                                                System.Drawing.Color secondCellColor =
                                                    System.Drawing.ColorTranslator.FromOle(
                                                        (int)
                                                        ((double)
                                                            weWs.Range["B" + (row + 1).ToString()].Interior.Color));
                                                System.Drawing.Color thirdCellColor =
                                                    System.Drawing.ColorTranslator.FromOle(
                                                        (int)
                                                        ((double)
                                                            weWs.Range["B" + (row + 2).ToString()].Interior.Color));
                                                if (topCellColor == Color.FromArgb(141, 180, 226))
                                                {
                                                    if (secondCellColor ==
                                                        Color.FromArgb(141, 180, 226) /* Condition 1 */&&
                                                        thirdCellColor ==
                                                        Color.FromArgb(141, 180, 226) /* Condition 2 */&&
                                                        weWs.Range["B" + (row + 3).ToString()].Interior.Color == 65535)
                                                        //Condition 3
                                                    {
                                                        var v1 = CheckNull(weWs.Range["B" + row.ToString()].Value);
                                                        var v2 = CheckNull(weWs.Range["B" + (row + 1).ToString()].Value);
                                                        var v3 = CheckNull(weWs.Range["B" + (row + 2).ToString()].Value);
                                                        var v4 = CheckNull(weWs.Range["B" + (row + 3).ToString()].Value);

                                                        We we = new We(v1,
                                                            v2,
                                                            v3,
                                                            v4);
                                                        //Get Hour information
                                                        if (weWs.Range["G" + row.ToString()].Value != null)
                                                        {
                                                            var type = weWs.Range["G" + row.ToString()].Value.GetType();
                                                            if (type.ToString() == "System.String")
                                                            {
                                                                we.Uom = weWs.Range["G" + row.ToString()].Value;
                                                            }
                                                            else if (type.ToString() == "System.Double")
                                                            {
                                                                we.Uom = weWs.Range["G" + row.ToString()].ToString();
                                                            }
                                                        }
                                                        if (weWs.Range["H" + row.ToString()].Value != null)
                                                        {
                                                            var type = weWs.Range["H" + row.ToString()].Value.GetType();
                                                            if (type.ToString() == "System.String")
                                                            {
                                                                if (
                                                                    xl.WorksheetFunction
                                                                        .IsErr(weWs.Range["H" + row.ToString()].Value))
                                                                {
                                                                    we.Quantity = -1; //error
                                                                }
                                                                else
                                                                {
                                                                    var val = weWs.Range["H" + row.ToString()].Value;
                                                                    we.Quantity =(double)weWs.Range["H" + row.ToString()].Value;
                                                                }
                                                            }
                                                            else if (type.ToString() == "System.Double")
                                                            {
                                                                we.Quantity = weWs.Range["H" + row.ToString()].Value;
                                                            }
                                                        }
                                                        if (weWs.Range["J" + row.ToString()].Value != null)
                                                        {
                                                            we.GeneralManHours =
                                                                weWs.Range["J" + row.ToString()].Value.ToString();
                                                        }
                                                        if (weWs.Range["L" + row.ToString()].Value != null)
                                                        {
                                                            we.BuilderManHours =
                                                                weWs.Range["L" + row.ToString()].Value.ToString();
                                                        }
                                                        if (weWs.Range["N" + row.ToString()].Value != null)
                                                        {
                                                            we.ElectricianManHours =
                                                                weWs.Range["N" + row.ToString()].Value.ToString();
                                                        }
                                                        if (weWs.Range["P" + row.ToString()].Value != null)
                                                        {
                                                            we.EngServicesManHours =
                                                                weWs.Range["P" + row.ToString()].Value.ToString();
                                                        }

                                                        if (weWs.Range["R" + row.ToString()].Value != null)
                                                        {
                                                            we.EquipOperatorManHours =
                                                                weWs.Range["R" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["T" + row.ToString()].Value != null)
                                                        {
                                                            we.SteelWorkerManHours =
                                                                weWs.Range["T" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["V" + row.ToString()].Value != null)
                                                        {
                                                            we.UtilityManHours =
                                                                weWs.Range["V" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["X8"].Value != null)
                                                        {
                                                            we.Equipment1Description = weWs.Range["X8"].Value.ToString();

                                                        }
                                                        if (weWs.Range["Y" + row.ToString()].Value != null)
                                                        {
                                                            we.Equipment1Hours =
                                                                weWs.Range["Y" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["Z8"].Value != null)
                                                        {
                                                            we.Equipment2Description = weWs.Range["Z8"].Value.ToString();

                                                        }
                                                        if (weWs.Range["AA" + row.ToString()].Value != null)
                                                        {
                                                            we.Equipment2Hours =
                                                                weWs.Range["AA" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["AB8"].Value != null)
                                                        {
                                                            we.Equipment3Description =
                                                                weWs.Range["AB8"].Value.ToString();

                                                        }
                                                        if (weWs.Range["AC" + row.ToString()].Value != null)
                                                        {
                                                            we.Equipment3Hours =
                                                                weWs.Range["AC" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["AD8"].Value != null)
                                                        {
                                                            we.Equipment4Description =
                                                                weWs.Range["AD8"].Value.ToString();

                                                        }
                                                        if (weWs.Range["AE" + row.ToString()].Value != null)
                                                        {
                                                            we.Equipment4Hours =
                                                                weWs.Range["AE" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["AF8"].Value != null)
                                                        {
                                                            we.Equipment5Description =
                                                                weWs.Range["AF8"].Value.ToString();

                                                        }
                                                        if (weWs.Range["AG" + row.ToString()].Value != null)
                                                        {
                                                            we.Equipment5Hours =
                                                                weWs.Range["AG" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["AH8"].Value != null)
                                                        {
                                                            we.Equipment6Description =
                                                                weWs.Range["AH8"].Value.ToString();

                                                        }
                                                        if (weWs.Range["AI" + row.ToString()].Value != null)
                                                        {
                                                            we.Equipment6Hours =
                                                                weWs.Range["AI" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["AJ8"].Value != null)
                                                        {
                                                            we.Equipment7Description =
                                                                weWs.Range["AJ8"].Value.ToString();

                                                        }
                                                        if (weWs.Range["AK" + row.ToString()].Value != null)
                                                        {
                                                            we.Equipment7Hours =
                                                                weWs.Range["AK" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["AL8"].Value != null)
                                                        {
                                                            we.Equipment8Description =
                                                                weWs.Range["AL8"].Value.ToString();

                                                        }
                                                        if (weWs.Range["AM" + row.ToString()].Value != null)
                                                        {
                                                            we.Equipment8Hours =
                                                                weWs.Range["AM" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["AN8"].Value != null)
                                                        {
                                                            we.Equipment9Description =
                                                                weWs.Range["AN8"].Value.ToString();

                                                        }
                                                        if (weWs.Range["AO" + row.ToString()].Value != null)
                                                        {
                                                            we.Equipment9Hours =
                                                                weWs.Range["AO" + row.ToString()].Value.ToString();

                                                        }
                                                        if (weWs.Range["AP8"].Value != null)
                                                        {
                                                            we.Equipment10Description =
                                                                weWs.Range["AP8"].Value.ToString();

                                                        }
                                                        if (weWs.Range["AQ" + row.ToString()].Value != null)
                                                        {
                                                            we.Equipment10Hours =
                                                                weWs.Range["AQ" + row.ToString()].Value.ToString();

                                                        }
                                                        //Get NSN info and store into We.NsnList
                                                        Nsn nsn = new Nsn();
                                                        for (int nsnRow = row;
                                                            nsnRow <= weWs.UsedRange.Rows.Count;
                                                            nsnRow++)
                                                        {
                                                            if (weWs.Range["C" + nsnRow.ToString()].Value != null)
                                                            {
                                                                nsn =
                                                                    new Nsn(
                                                                        weWs.Range["C" + nsnRow.ToString()].Value
                                                                            .ToString(),
                                                                        weWs.Range["D" + nsnRow.ToString()].Value
                                                                            .ToString(),
                                                                        weWs.Range["E" + nsnRow.ToString()].Value
                                                                            .ToString(),
                                                                        weWs.Range["F" + nsnRow.ToString()].Value
                                                                            .ToString());
                                                                we.NsnList.Add(nsn);
                                                                nsn = null;
                                                            }
                                                            else
                                                            {
                                                                System.Drawing.Color cellColor =
                                                                    System.Drawing.ColorTranslator.FromOle(
                                                                        (int)
                                                                        ((double)
                                                                            weWs.Range["B" + (row + 1).ToString()]
                                                                                .Interior
                                                                                .Color));
                                                                if (cellColor == Color.FromArgb(141, 180, 226))
                                                                    //this is the start of a new We
                                                                {
                                                                    break;
                                                                }
                                                                break;
                                                            }
                                                        }
                                                        if (we != null)
                                                        {
                                                            ca.WeList.Add(we);
                                                            we = null;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        caWb.Close(false);
                                    }
                                    //Insert CA Information.  I now have the WE and NSN info for that single CA
                                    bom.Range["A" + curBomRow].Value = ca.Number;
                                    bom.Range["B" + curBomRow.ToString()].Value = ca.Description;
                                    bom.Range["C" + curBomRow.ToString()].Value = 1;
                                    //They will change this value in good time
                                    int caRow = curBomRow;
                                    curBomRow += 1;
                                    //Insert CA info into WBS
                                    wbs.Range["A" + curWbsRow.ToString()].Value = ca.Number;
                                    wbs.Range["B" + curWbsRow.ToString()].Value = ca.Description;
                                    wbs.Range["C" + curWbsRow.ToString()].Value = 1; //They will change this later on
                                    int caResRow = curResRow;
                                    curWbsRow += 1;
                                    //insert CA info into Resources
                                    res.Range["A" + curResRow.ToString()].Value = ca.Number;
                                    res.Range["B" + curResRow.ToString()].Value = ca.Description;
                                    res.Range["C" + curResRow.ToString()].Value = 1;
                                    curResRow += 1;
                                    foreach (var we in ca.WeList)
                                    {
                                        //find the we in the CostBook
                                        var longCode = from row in cbTable.AsEnumerable()
                                            where row.Field<string>("Division Name") == we.Division
                                                  && row.Field<string>("Section Name") == we.Section
                                                  && row.Field<string>("Work Element Description") == we.LineItem
                                            select row.ItemArray[0]; //row.Field<string>("Section Code");
                                        if (!longCode.Any())
                                        {
                                            Console.WriteLine(we.Division);
                                            Console.WriteLine(we.Section);
                                            Console.WriteLine(we.LineItem);
                                            Console.WriteLine();
                                        }
                                        else
                                        {

                                            we.LongCode = longCode.ElementAt(0).ToString();
                                            bom.Range["E" + curBomRow.ToString()].Value = we.LongCode;
                                            bom.Range["F" + curBomRow.ToString()].Formula = "=IF($E" +
                                                                                            curBomRow.ToString() +
                                                                                            "=\"\",\"\",VLOOKUP($E" +
                                                                                            curBomRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3000,4,FALSE)&\" - \"&VLOOKUP($E" +
                                                                                            curBomRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3000,5,FALSE))";
                                            curBomRow += 1;
                                            wbs.Range["D" + curWbsRow.ToString()].Value = we.LongCode;
                                            curWbsRow += 1; // + "\"\"" + 
                                            res.Range["E" + curResRow.ToString()].Value = we.LongCode;
                                            res.Range["H" + curResRow.ToString()].Value = we.Quantity;
                                            res.Range["D" + curResRow.ToString()].Value = "=IF(H" + curResRow.ToString() +
                                                                                          "=" +
                                                                                          "\"\"" + ", " + "\"\"" + ",C" +
                                                                                          caResRow.ToString() + ")";
                                            //caResRow
                                            res.Range["F" + curResRow.ToString()].Formula = "=IF($E" +
                                                                                            curResRow.ToString() +
                                                                                            "=\"\",\"\",VLOOKUP($E" +
                                                                                            curResRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3000,4,FALSE)&\" - \"&VLOOKUP($E" +
                                                                                            curResRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3000,5,FALSE))";
                                            res.Range["G" + curResRow.ToString()].Formula = "=IF($E" +
                                                                                            curResRow.ToString() +
                                                                                            "=\"\",\"\",VLOOKUP($E" +
                                                                                            curResRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3000,6,FALSE))";
                                            res.Range["I" + curResRow.ToString()].Formula = "=IF(H" +
                                                                                            curResRow.ToString() +
                                                                                            "=\"\",\"\",D" +
                                                                                            curResRow.ToString() +
                                                                                            "*H" + curResRow.ToString() +
                                                                                            ")";
                                            res.Range["J" + curResRow.ToString()].Formula = "=IF($I" +
                                                                                            curResRow.ToString() +
                                                                                            "=\"\",\"\",$I" +
                                                                                            curResRow.ToString() +
                                                                                            "*VLOOKUP($E" +
                                                                                            curResRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3019,11,FALSE))";
                                            res.Range["K" + curResRow.ToString()].Formula = "=IF($I" +
                                                                                            curResRow.ToString() +
                                                                                            "=\"\",\"\",$I" +
                                                                                            curResRow.ToString() +
                                                                                            "*VLOOKUP($E" +
                                                                                            curResRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3019,13,FALSE))";
                                            res.Range["L" + curResRow.ToString()].Formula = "=IF($I" +
                                                                                            curResRow.ToString() +
                                                                                            "=\"\",\"\",$I" +
                                                                                            curResRow.ToString() +
                                                                                            "*VLOOKUP($E" +
                                                                                            curResRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3019,15,FALSE))";
                                            res.Range["M" + curResRow.ToString()].Formula = "=IF($I" +
                                                                                            curResRow.ToString() +
                                                                                            "=\"\",\"\",$I" +
                                                                                            curResRow.ToString() +
                                                                                            "*VLOOKUP($E" +
                                                                                            curResRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3019,17,FALSE))";
                                            res.Range["N" + curResRow.ToString()].Formula = "=IF($I" +
                                                                                            curResRow.ToString() +
                                                                                            "=\"\",\"\",$I" +
                                                                                            curResRow.ToString() +
                                                                                            "*VLOOKUP($E" +
                                                                                            curResRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3019,19,FALSE))";
                                            res.Range["O" + curResRow.ToString()].Formula = "=IF($I" +
                                                                                            curResRow.ToString() +
                                                                                            "=\"\",\"\",$I" +
                                                                                            curResRow.ToString() +
                                                                                            "*VLOOKUP($E" +
                                                                                            curResRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3019,21,FALSE))";
                                            res.Range["P" + curResRow.ToString()].Formula = "=IF($I" +
                                                                                            curResRow.ToString() +
                                                                                            "=\"\",\"\",$I" +
                                                                                            curResRow.ToString() +
                                                                                            "*VLOOKUP($E" +
                                                                                            curResRow.ToString() +
                                                                                            ",CostBook!$A$2:$W$3019,23,FALSE))";
                                            //res.Range["R" + curResRow.ToString()].Formula = "";

                                        }
                                        int furthest = 3;
                                        if (we.GeneralManHours != "")
                                        {
                                            if (res.Range["J" + curResRow.ToString()].Value != null)
                                            {
                                                double d1 =
                                                    Math.Round(
                                                        Convert.ToDouble(res.Range["J" + curResRow.ToString()].Value), 2);
                                                double d2 = Math.Round(Convert.ToDouble(we.GeneralManHours), 2);
                                                if (
                                                    !Math.Round(res.Range["J" + curResRow.ToString()].Value, 2)
                                                        .Equals(Math.Round(Convert.ToDouble(we.GeneralManHours), 2)))
                                                {
                                                    //Assume no comment previously exists
                                                    res.Range["J" + curResRow.ToString()].AddComment(
                                                        "Construction Activity Workbook lists " + we.GeneralManHours);
                                                }
                                            }
                                            else
                                            {
                                                res.Range["J" + curResRow.ToString()].Value = we.GeneralManHours;
                                            }
                                        }
                                        if (we.BuilderManHours != "")
                                        {
                                            if (res.Range["K" + curResRow.ToString()].Value != null)
                                            {
                                                double d1 = Math.Round(res.Range["K" + curResRow.ToString()].Value, 2);
                                                double d2 = Math.Round(Convert.ToDouble(we.BuilderManHours), 2);
                                                if (
                                                    !Math.Round(res.Range["K" + curResRow.ToString()].Value, 2)
                                                        .Equals(Math.Round(Convert.ToDouble(we.BuilderManHours), 2)))
                                                {
                                                    //Assume no comment previously exists
                                                    res.Range["K" + curResRow.ToString()].AddComment(
                                                        "Construction Activity Workbook lists " + we.BuilderManHours);
                                                }
                                            }
                                            else
                                            {
                                                res.Range["K" + curResRow.ToString()].Value = we.BuilderManHours;
                                            }
                                        }
                                        if (we.ElectricianManHours != "")
                                        {
                                            if (res.Range["L" + curResRow.ToString()].Value != null)
                                            {
                                                if (
                                                    !Math.Round(res.Range["L" + curResRow.ToString()].Value, 2)
                                                        .Equals(Math.Round(Convert.ToDouble(we.ElectricianManHours), 2)))
                                                {
                                                    //Assume no comment previously exists
                                                    res.Range["L" + curResRow.ToString()].AddComment(
                                                        "Construction Activity Workbook lists " + we.ElectricianManHours);
                                                }
                                            }
                                            else
                                            {
                                                res.Range["L" + curResRow.ToString()].Value = we.ElectricianManHours;
                                            }
                                        }
                                        if (we.EngServicesManHours != "")
                                        {
                                            if (res.Range["M" + curResRow.ToString()].Value != null)
                                            {
                                                if (
                                                    !Math.Round(res.Range["M" + curResRow.ToString()].Value, 2)
                                                        .Equals(Math.Round(Convert.ToDouble(we.EngServicesManHours), 2)))
                                                {
                                                    //Assume no comment previously exists
                                                    res.Range["M" + curResRow.ToString()].AddComment(
                                                        "Construction Activity Workbook lists " + we.EngServicesManHours);
                                                }
                                            }
                                            else
                                            {
                                                res.Range["M" + curResRow.ToString()].Value = we.EngServicesManHours;
                                            }
                                        }
                                        if (we.EquipOperatorManHours != "")
                                        {
                                            if (res.Range["N" + curResRow.ToString()].Value != null)
                                            {
                                                if (
                                                    !Math.Round(res.Range["N" + curResRow.ToString()].Value, 2)
                                                        .Equals(Math.Round(Convert.ToDouble(we.EquipOperatorManHours), 2)))
                                                {
                                                    //Assume no comment previously exists
                                                    res.Range["N" + curResRow.ToString()].AddComment(
                                                        "Construction Activity Workbook lists " +
                                                        we.EquipOperatorManHours);
                                                }
                                            }
                                            else
                                            {
                                                res.Range["N" + curResRow.ToString()].Value = we.EquipOperatorManHours;
                                            }
                                        }
                                        if (we.SteelWorkerManHours != "")
                                        {
                                            if (res.Range["O" + curResRow.ToString()].Value != null)
                                            {
                                                if (
                                                    !Math.Round(res.Range["O" + curResRow.ToString()].Value, 2)
                                                        .Equals(Math.Round(Convert.ToDouble(we.SteelWorkerManHours), 2)))
                                                {
                                                    //Assume no comment previously exists
                                                    res.Range["O" + curResRow.ToString()].AddComment(
                                                        "Construction Activity Workbook lists " + we.SteelWorkerManHours);
                                                }
                                            }
                                            else
                                            {
                                                res.Range["O" + curResRow.ToString()].Value = we.SteelWorkerManHours;
                                            }
                                        }
                                        if (we.UtilityManHours != "")
                                        {
                                            if (res.Range["P" + curResRow.ToString()].Value != null)
                                            {
                                                if (
                                                    !Math.Round(res.Range["P" + curResRow.ToString()].Value, 2)
                                                        .Equals(Math.Round(Convert.ToDouble(we.UtilityManHours), 2)))
                                                {
                                                    //Assume no comment previously exists
                                                    res.Range["P" + curResRow.ToString()].AddComment(
                                                        "Construction Activity Workbook lists " + we.UtilityManHours);
                                                }
                                            }
                                            else
                                            {
                                                res.Range["P" + curResRow.ToString()].Value = we.UtilityManHours;
                                            }
                                        }

                                        //Now include the hours information
                                        if (we.Equipment1Description != null && we.Equipment1Hours != null)
                                        {
                                            res.Range["Q" + curResRow.ToString()].Value = we.Equipment1Description;
                                            res.Range["R" + curResRow.ToString()].Value = we.Equipment1Hours;
                                        }
                                        if (we.Equipment2Description != null && we.Equipment2Hours != null)
                                        {
                                            res.Range["S" + curResRow.ToString()].Value = we.Equipment2Description;
                                            res.Range["T" + curResRow.ToString()].Value = we.Equipment2Hours;
                                        }
                                        if (we.Equipment3Description != null && we.Equipment3Hours != null)
                                        {
                                            res.Range["U" + curResRow.ToString()].Value = we.Equipment1Description;
                                            res.Range["V" + curResRow.ToString()].Value = we.Equipment1Hours;
                                        }
                                        Excel.Worksheet resSource;
                                        if (we.Equipment4Description != null && we.Equipment4Hours != null)
                                        {
                                            furthest = 5;

                                            resSource = exWb.Sheets["Resources"];
                                            source = resSource.Range["S1:T2"];
                                            dest = res.Range["W1:X2"];
                                            source.Copy();
                                            res.Activate();
                                            res.Range["W1"].Select();
                                            res.Paste();

                                            //res.Range["S1:T2"].Copy(res.Range["W1:X2"]);
                                            res.Range["W1"].Value = "EQUIPMENT #4";
                                            res.Range["X1"].Value = "EQUIPMENT #4 HRS";
                                            res.Columns["W"].ColumnWidth = 13.29;
                                            res.Columns["X"].ColumnWidth = 17.71;
                                            res.Columns["W"].Font.Bold = true;
                                            res.Columns["X"].Font.Bold = true;
                                            res.Range["W" + curResRow.ToString()].Value = we.Equipment4Description;
                                            res.Range["X" + curResRow.ToString()].Value = we.Equipment4Hours;
                                        }
                                        if (we.Equipment5Description != null && we.Equipment5Hours != null)
                                        {
                                            furthest = 7;

                                            resSource = exWb.Sheets["Resources"];
                                            source = resSource.Range["S1:T2"];
                                            dest = res.Range["Y1:Z2"];
                                            source.Copy();
                                            res.Activate();
                                            res.Range["Y1"].Select();
                                            res.Paste();

                                            res.Range["S1:T2"].Copy(res.Range["Y1:Z2"]);
                                            res.Range["Y1"].Value = "EQUIPMENT #5";
                                            res.Range["Z1"].Value = "EQUIPMENT #5 HRS";
                                            res.Columns["Y"].ColumnWidth = 13.29;
                                            res.Columns["Z"].ColumnWidth = 17.71;
                                            res.Columns["Y"].Font.Bold = true;
                                            res.Columns["Z"].Font.Bold = true;
                                            res.Range["Y" + curResRow.ToString()].Value = we.Equipment5Description;
                                            res.Range["Z" + curResRow.ToString()].Value = we.Equipment5Hours;
                                        }
                                        if (we.Equipment6Description != null && we.Equipment6Hours != null)
                                        {
                                            furthest = 9;

                                            resSource = exWb.Sheets["Resources"];
                                            source = resSource.Range["S1:T2"];
                                            dest = res.Range["AA1:AB2"];
                                            source.Copy();
                                            res.Activate();
                                            res.Range["AA1"].Select();
                                            res.Paste();

                                            res.Range["S1:T2"].Copy(res.Range["AA1:AB2"]);
                                            res.Range["AA1"].Value = "EQUIPMENT #6";
                                            res.Range["AB1"].Value = "EQUIPMENT #6 HRS";
                                            res.Columns["AA"].ColumnWidth = 13.29;
                                            res.Columns["AB"].ColumnWidth = 17.71;
                                            res.Columns["AA"].Font.Bold = true;
                                            res.Columns["AB"].Font.Bold = true;
                                            res.Range["AA" + curResRow.ToString()].Value = we.Equipment6Description;
                                            res.Range["AB" + curResRow.ToString()].Value = we.Equipment6Hours;
                                        }
                                        if (we.Equipment7Description != null && we.Equipment7Hours != null)
                                        {
                                            furthest = 11;

                                            resSource = exWb.Sheets["Resources"];
                                            source = resSource.Range["S1:T2"];
                                            dest = res.Range["AC1:AD2"];
                                            source.Copy();
                                            res.Activate();
                                            res.Range["AC1"].Select();
                                            res.Paste();

                                            res.Range["S1:T2"].Copy(res.Range["AC1:AD2"]);
                                            res.Range["AC1"].Value = "EQUIPMENT #7";
                                            res.Range["AD1"].Value = "EQUIPMENT #7 HRS";
                                            res.Columns["AC"].ColumnWidth = 13.29;
                                            res.Columns["AD"].ColumnWidth = 17.71;
                                            res.Columns["AC"].Font.Bold = true;
                                            res.Columns["AD"].Font.Bold = true;
                                            res.Range["AC" + curResRow.ToString()].Value = we.Equipment7Description;
                                            res.Range["AD" + curResRow.ToString()].Value = we.Equipment7Hours;
                                        }
                                        if (we.Equipment8Description != null && we.Equipment8Hours != null)
                                        {
                                            furthest = 13;

                                            resSource = exWb.Sheets["Resources"];
                                            source = resSource.Range["S1:T2"];
                                            dest = res.Range["AE1:AF2"];
                                            source.Copy();
                                            res.Activate();
                                            res.Range["AE1"].Select();
                                            res.Paste();

                                            res.Range["S1:T2"].Copy(res.Range["AE1:AF2"]);
                                            res.Range["AE1"].Value = "EQUIPMENT #8";
                                            res.Range["AF1"].Value = "EQUIPMENT #8 HRS";
                                            res.Columns["AE"].ColumnWidth = 13.29;
                                            res.Columns["AF"].ColumnWidth = 17.71;
                                            res.Columns["AE"].Font.Bold = true;
                                            res.Columns["AF"].Font.Bold = true;
                                            res.Range["AE" + curResRow.ToString()].Value = we.Equipment8Description;
                                            res.Range["AF" + curResRow.ToString()].Value = we.Equipment8Hours;
                                        }
                                        if (we.Equipment9Description != null && we.Equipment9Hours != null)
                                        {
                                            furthest = 15;

                                            resSource = exWb.Sheets["Resources"];
                                            source = resSource.Range["S1:T2"];
                                            dest = res.Range["AG1:AH2"];
                                            source.Copy();
                                            res.Activate();
                                            res.Range["AG1"].Select();
                                            res.Paste();

                                            res.Range["S1:T2"].Copy(res.Range["AG1:AH2"]);
                                            res.Range["AG1"].Value = "EQUIPMENT #9";
                                            res.Range["AH1"].Value = "EQUIPMENT #9 HRS";
                                            res.Columns["AG"].ColumnWidth = 13.29;
                                            res.Columns["AH"].ColumnWidth = 17.71;
                                            res.Columns["AG"].Font.Bold = true;
                                            res.Columns["AH"].Font.Bold = true;
                                            res.Range["AG" + curResRow.ToString()].Value = we.Equipment9Description;
                                            res.Range["AH" + curResRow.ToString()].Value = we.Equipment9Hours;
                                        }
                                        if (we.Equipment10Description != null && we.Equipment10Hours != null)
                                        {
                                            furthest = 17;

                                            resSource = exWb.Sheets["Resources"];
                                            source = resSource.Range["S1:T2"];
                                            dest = res.Range["AI1:AJ2"];
                                            source.Copy();
                                            res.Activate();
                                            res.Range["AI1"].Select();
                                            res.Paste();

                                            res.Range["S1:T2"].Copy(res.Range["AI1:AJ2"]);
                                            res.Range["AI1"].Value = "EQUIPMENT #10";
                                            res.Range["AJ1"].Value = "EQUIPMENT #10 HRS";
                                            res.Columns["AI"].ColumnWidth = 13.29;
                                            res.Columns["AJ"].ColumnWidth = 17.71;
                                            res.Columns["AI"].Font.Bold = true;
                                            res.Columns["AJ"].Font.Bold = true;
                                            res.Range["AI" + curResRow.ToString()].Value = we.Equipment10Description;
                                            res.Range["AJ" + curResRow.ToString()].Value = we.Equipment10Hours;
                                        }
                                        if (furthest > maxFurthest)
                                        {
                                            maxFurthest = furthest;
                                        }

                                        curResRow += 1;
                                        //Now include the NSNs
                                        foreach (var nsn in we.NsnList)
                                        {
                                            string formattedNsn = nsn.Number.Substring(0, 4) + "-" +
                                                                  nsn.Number.Substring(4, 2) +
                                                                  "-" +
                                                                  nsn.Number.Substring(6, 3) + "-" +
                                                                  nsn.Number.Substring(9, nsn.Number.Length - 9);
                                            bom.Range["D" + curBomRow.ToString()].Formula = "=IF(G" +
                                                                                            curBomRow.ToString() +
                                                                                            "=" +
                                                                                            "\"\"" + "\"\"" + "," +
                                                                                            "\"\"" +
                                                                                            ",C$" + caRow + ")";
                                            bom.Range["G" + curBomRow.ToString()].Value = formattedNsn;
                                            bom.Range["M" + curBomRow.ToString()].Value = nsn.Quantity;
                                            bom.Range["H" + curBomRow.ToString()].Formula = "=IF($G" +
                                                                                            curBomRow.ToString() +
                                                                                            "=\"\",\"\",VLOOKUP($G" +
                                                                                            curBomRow.ToString() +
                                                                                            ",NSN!$A$2:$N$5000,7,FALSE))";
                                            bom.Range["I" + curBomRow.ToString()].Formula = "=IF($G" +
                                                                                            curBomRow.ToString() +
                                                                                            "=\"\",\"\",VLOOKUP($G" +
                                                                                            curBomRow.ToString() +
                                                                                            ",NSN!$A$2:$N$5000,14,FALSE))";
                                            bom.Range["J" + curBomRow.ToString()].Formula = "=IF($G" +
                                                                                            curBomRow.ToString() +
                                                                                            "=\"\",\"\",VLOOKUP($G" +
                                                                                            curBomRow.ToString() +
                                                                                            ",NSN!$A$2:$N$5000,13,FALSE))";
                                            bom.Range["K" + curBomRow.ToString()].Formula = "=IF($G" +
                                                                                            curBomRow.ToString() +
                                                                                            "=\"\",\"\",VLOOKUP($G" +
                                                                                            curBomRow.ToString() +
                                                                                            ",NSN!$A$2:$N$5000,12,FALSE))";
                                            bom.Range["L" + curBomRow.ToString()].Formula = "=IF($G" +
                                                                                            curBomRow.ToString() +
                                                                                            "=\"\",\"\",VLOOKUP($G" +
                                                                                            curBomRow.ToString() +
                                                                                            ",NSN!$A$2:$N$5000,9,FALSE))";
                                            bom.Range["N" + curBomRow.ToString()].Formula = "=IF($G" +
                                                                                            curBomRow.ToString() +
                                                                                            "=\"\",\"\",VLOOKUP($G" +
                                                                                            curBomRow.ToString() +
                                                                                            ",NSN!$A$2:$N$5000,10,FALSE))";
                                            bom.Range["O" + curBomRow.ToString()].Formula = "=IF($G" +
                                                                                            curBomRow.ToString() +
                                                                                            "=\"\",\"\",VLOOKUP($G" +
                                                                                            curBomRow.ToString() +
                                                                                            ",NSN!$A$2:$N$5000,11,FALSE))";
                                            bom.Range["P" + curBomRow.ToString()].Formula = "=IF($G" +
                                                                                            curBomRow.ToString() +
                                                                                            "=\"\",\"\",M" +
                                                                                            curBomRow.ToString() + "/N" +
                                                                                            curBomRow.ToString() + ")";
                                            bom.Range["Q" + curBomRow.ToString()].Formula = "=IF(P" +
                                                                                            curBomRow.ToString() +
                                                                                            "=\"\", \"\", D" +
                                                                                            curBomRow.ToString() + "*P" +
                                                                                            curBomRow.ToString() + ")";
                                            // + "\"\"" + 
                                            // + curBomRow.ToString() + 
                                            bom.Range["R" + curBomRow.ToString()].Formula = "=IF(P" +
                                                                                            curBomRow.ToString() +
                                                                                            "=" +
                                                                                            "\"\"" + ", " + "\"\"" +
                                                                                            ", Q" +
                                                                                            curBomRow.ToString() + "*I" +
                                                                                            curBomRow.ToString() +
                                                                                            "/40)";
                                            bom.Range["S" + curBomRow.ToString()].Formula = "=IF(P" +
                                                                                            curBomRow.ToString() +
                                                                                            "=" +
                                                                                            "\"\"" + ", " + "\"\"" +
                                                                                            ", Q" +
                                                                                            curBomRow.ToString() + "*J" +
                                                                                            curBomRow.ToString() +
                                                                                            "/2000)";
                                            bom.Range["T" + curBomRow.ToString()].Formula = "=IF(P" +
                                                                                            curBomRow.ToString() +
                                                                                            "=" +
                                                                                            "\"\"" + ", " + "\"\"" +
                                                                                            ", Q" +
                                                                                            curBomRow.ToString() + "*K" +
                                                                                            curBomRow.ToString() + ")";
                                            curBomRow += 1;
                                        }
                                        //cbTable.Select(String.Format("Division Name = '{0}' AND Section Name = '{1}' AND Work Element Description = '{2}", we.Division, we.Section, we.LineItem));
                                        //bom.Range["E" + curBomRow.ToString()].Value = we.
                                    }

                                }
                            } while (reader.NextResult());
                            reader.Close();
                        }

                        bom.Range["A600:B600"].Font.Bold = false;
                        bom.Range["B605:B608"].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight =
                            Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        RemoveEmptyRows(topOfLegend, bom);
                        //wbs stuff
                        wbs.Activate();
                        topOfLegend = 391; //for wbs

                        RemoveEmptyRows(topOfLegend, wbs); //Only now have we filled out the wbs
                        bom.Activate();
                        foreach (Excel.Range cell in bom.UsedRange.Cells)
                        {
                            if (cell.Row > 2 && cell.Column > 8 && cell.Row < bom.UsedRange.Rows.Count - 4)
                            {
                                if (cell.Value != null)
                                {
                                    double newVal;
                                    if (Double.TryParse(cell.Value.ToString(), out newVal))
                                        //has a decimal to be rounded to
                                    {
                                        int row = cell.Row;
                                        int col = cell.Column;

                                        var dec = newVal - Math.Truncate(newVal);
                                        dec = SetSigFigs(dec, 2);
                                        if (dec.ToString().Length > 0 && dec != 0)
                                        {
                                            if (cell.Column != 18 && cell.Column != 19 && cell.Column != 20)
                                                //since I am rounding these to 2-4 decimal places already
                                            {
                                                string formatString;
                                                if (newVal < 1)
                                                {
                                                    formatString = "0.";
                                                }
                                                else
                                                {
                                                    formatString = "#.";
                                                }
                                                string sdec = dec.ToString("0.################");
                                                for (int fc = 0;
                                                        fc < dec.ToString("0.################").Length - 2;
                                                        fc++)
                                                    //0. is two, so subtract from this
                                                {
                                                    formatString += "#";
                                                }
                                                cell.NumberFormat = formatString;
                                            }
                                            var roundedDec = Math.Round(dec, 4);
                                            if (cell.Column == 18 || cell.Column == 19)
                                            {
                                                string formatString;
                                                if (newVal < 1)
                                                {
                                                    formatString = "0.";
                                                }
                                                else
                                                {
                                                    formatString = "#.";
                                                }
                                                string sdec = dec.ToString("0.################");
                                                for (int fc = 0;
                                                        fc < 4;
                                                        fc++)
                                                    //0. is two, so subtract from this
                                                {
                                                    formatString += "#";
                                                }
                                                cell.NumberFormat = formatString;
                                                if (cell.Text.ToString().Contains("E-") || cell.Text.ToString() == "0.")
                                                {
                                                    cell.NumberFormat = "0";
                                                }
                                            }

                                            roundedDec = Math.Round(dec, 2);
                                            if (cell.Column == 20)
                                            {
                                                string formatString;
                                                if (newVal < 1)
                                                {
                                                    formatString = "0.";
                                                }
                                                else
                                                {
                                                    formatString = "#.";
                                                }
                                                string sdec = dec.ToString("0.################");
                                                for (int fc = 0;
                                                        fc < 2;
                                                        fc++)
                                                    //0. is two, so subtract from this
                                                {
                                                    formatString += "#";
                                                }
                                                cell.NumberFormat = formatString;
                                                if (cell.Text.ToString().Contains("E-") || cell.Text.ToString() == "0.")
                                                {
                                                    cell.NumberFormat = "0";
                                                }
                                            }

                                            if (cell.Text.ToString().Contains("E-") || cell.Text.ToString() == "0.")
                                            {
                                                cell.NumberFormat = "0";
                                            }
                                        }

                                    }
                                }
                            }
                        }
                        foreach (Excel.Range cell in bom.Range["R3:S" + bom.UsedRange.Rows.Count.ToString()].Cells)
                        {
                            if (cell.Value != null)
                            {
                                if (!xl.WorksheetFunction.IsError(cell))
                                {
                                    if (cell.Text.ToString().Contains("."))
                                    {
                                        int io = cell.Text.ToString().IndexOf(".");
                                        int io2 = cell.Text.ToString().Length;
                                        string ss = cell.Text.ToString().Substring(io + 1, io2 - io - 1);
                                        if (
                                            ss.Length > 4)
                                        {
                                            Console.WriteLine("Error in BOM col R/S number of decimals, row:" + cell.Row);
                                        }
                                    }
                                }
                            }
                        }
                        foreach (Excel.Range cell in bom.Range["T" + bom.UsedRange.Rows.Count.ToString()].Cells)
                        {
                            if (cell.Value != null)
                            {
                                if (!xl.WorksheetFunction.IsError(cell))
                                {
                                    if (cell.Text.ToString().Contains("."))
                                    {
                                        int io = cell.Text.ToString().IndexOf(".");
                                        int io2 = cell.Text.ToString().Length;
                                        string ss = cell.Text.ToString().Substring(io + 1, io2 - io - 1);
                                        if (
                                            ss.Length > 2)
                                        {
                                            Console.WriteLine("Error in BOM col T number of decimals, row:" + cell.Row);
                                        }
                                    }
                                }
                            }
                        }
                        //RESOURCES
                        Excel.Worksheet reSource = exWb.Sheets["Resources"];
                        source = reSource.Range["A1:V2"];
                        dest = res.Range["A1:V2"];
                        source.Copy();
                        res.Activate();
                        res.Range["A1"].Select();
                        res.Paste();

                        source = reSource.Range["A118:B122"];
                        dest = res.Range["A377:B380"];
                        source.Copy();
                        res.Activate();
                        res.Range["A377"].Select();
                        res.Paste();

                        res.Range["A377:B377"].Merge();

                        
                        source = reSource.Range["A29:V29"];
                        dest = res.Range["A373:V373"];
                        source.Copy();
                        res.Activate();
                        res.Range["A373"].Select();
                        res.Paste();
                        (res.Rows[373] as Excel.Range).WrapText = false;

                        for (int j = 22; j <= 19 + maxFurthest; j++)
                        {
                            source = res.Range["T373"];
                            dest = res.Cells[373, j];
                            source.Copy();
                            res.Cells[373, j].Select();
                            res.Paste();
                        }

                        for (int rn = 377; rn < 423; rn++)
                        {
                            res.Rows[rn].RowHeight = exWb.Sheets["Resources"].Rows[rn + (418 - 377)].RowHeight;
                        }
                        res.Range["F1"].Select();
                        res.Rows[373].autofit();
                        res.Columns[2].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        res.Columns[4].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        res.Columns[6].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        res.Columns[8].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        res.Columns[22].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        res.Columns[22].Borders(Excel.XlBordersIndex.xlEdgeRight).Weight =
                            Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        //Fit columns to match the 95% widths
                        res.Columns["A:E"].Font.Bold = true;
                        res.Columns["H"].Font.Bold = true;
                        res.Columns["Q:V"].Font.Bold = true;
                        for (int c = 1; c <= 22; c++) //In Excel the first column is 1 not 0
                        {
                            double colWidth = reSource.Columns[c].ColumnWidth;
                            res.Columns[c].ColumnWidth = colWidth;
                        }
                        int rememberCol = maxFurthest;


                        res.Rows[1].Rowheight = 30;
                        res.Rows[77].RowHeight = 21;
                        res.Range["A77:B77"].Font.Bold = false;
                        res.Range["B78:B81"].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight =
                            Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        res.Range["I73"].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone; //change col xledge
                        topOfLegend = 372;
                        RemoveEmptyRows(topOfLegend, res);
                        res.Columns["V"].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone; //change col xledge
                        res.Range["V1:V2"].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        res.Range["V1:V2"].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight =
                            Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                        res.Range["J2"].Formula =
                            "=SUM(INDIRECT(\"J3:J\" & MATCH(\"Optional Construction Activities\",A:A, 0) - 1))";
                        res.Range["K2"].Formula =
                            "=SUM(INDIRECT(\"K3:K\" & MATCH(\"Optional Construction Activities\",A:A, 0) - 1))";
                        res.Range["L2"].Formula =
                            "=SUM(INDIRECT(\"L3:L\" & MATCH(\"Optional Construction Activities\",A:A, 0) - 1))";
                        res.Range["M2"].Formula =
                            "=SUM(INDIRECT(\"M3:M\" & MATCH(\"Optional Construction Activities\",A:A, 0) - 1))";
                        res.Range["N2"].Formula =
                            "=SUM(INDIRECT(\"N3:N\" & MATCH(\"Optional Construction Activities\",A:A, 0) - 1))";
                        res.Range["O2"].Formula =
                            "=SUM(INDIRECT(\"O3:O\" & MATCH(\"Optional Construction Activities\",A:A, 0) - 1))";
                        res.Range["P2"].Formula =
                            "=SUM(INDIRECT(\"P3:P\" & MATCH(\"Optional Construction Activities\",A:A, 0) - 1))";
                        foreach (Excel.Range cell in res.UsedRange.Cells)
                        {
                            if (cell.Row > 2 && cell.Column > 7 && cell.Row < res.UsedRange.Rows.Count - 4)
                            {
                                if (cell.Value != null)
                                {
                                    double newVal;
                                    if (Double.TryParse(cell.Value.ToString(), out newVal))
                                        //has a decimal to be rounded to
                                    {
                                        var dec = newVal - Math.Truncate(newVal);
                                        dec = SetSigFigs(dec, 2);
                                        if (dec.ToString().Length > 0 && dec != 0)
                                        {
                                            string formatString;
                                            if (newVal < 1)
                                            {
                                                formatString = "0.";
                                            }
                                            else
                                            {
                                                formatString = "#.";
                                            }
                                            string sdec = dec.ToString("0.################");
                                            for (int fc = 0; fc < dec.ToString("0.################").Length - 2; fc++)
                                                //0. is two, so subtract from this
                                            {
                                                formatString += "#";
                                            }
                                            cell.NumberFormat = formatString;
                                        }
                                    }
                                }
                            }
                        }

                        int lastEdgeRow = 0;
                        for (int k = 1; k <= res.UsedRange.Rows.Count; k++)
                        {
                            if (res.Range["A" + k.ToString()].Value == "Key - Editable Columns are bold")
                            {
                                lastEdgeRow = k - 2;
                                res.Range["B" + (k + 1).ToString()].Font.Bold = false;
                                res.Range["B" + (k + 2).ToString()].Font.Bold = false;
                                res.Range["B" + (k + 3).ToString()].Font.Bold = false;
                                res.Range["B" + (k + 4).ToString()].Font.Bold = false;
                                break;
                            }
                        }
                        res.Range[res.Cells[1, 19 + maxFurthest], res.Cells[lastEdgeRow, 19 + rememberCol]].Borders(
                                Excel.XlBordersIndex.xlEdgeRight).Weight =
                            Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        res.Range[res.Cells[1, 19 + maxFurthest], res.Cells[lastEdgeRow, 19 + rememberCol]].Borders(
                                Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        lastEdgeRow = 0;
                        for (int k = 1; k <= bom.UsedRange.Rows.Count; k++)
                        {
                            if (bom.Range["A" + k.ToString()].Value == "Key - Editable Columns are bold")
                            {
                                lastEdgeRow = k - 2;
                                break;
                            }
                        }

                        bom.Range[bom.Cells[1, 20], bom.Cells[lastEdgeRow, 20]].Borders(Excel.XlBordersIndex.xlEdgeRight)
                                .Weight =
                            Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        bom.Range[bom.Cells[1, 20], bom.Cells[lastEdgeRow, 20]].Borders(Excel.XlBordersIndex.xlEdgeRight)
                                .LineStyle =
                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        bom.Activate();
                        bom.Range[bom.Cells[1, 22], bom.Cells[lastEdgeRow, 22]].Select();
                        //bom.Range["T1:T" + lastEdgeRow.ToString()].Borders(Excel.XlBordersIndex.xlEdgeRight).Weight =
                        //  Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                        //bom.Range["T1:T" + lastEdgeRow.ToString()].Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle =
                        //Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                        //TOGS
                        Excel.Worksheet togSource = exWb.Sheets["TOGS"];
                        source = togSource.Range["A1:D1"];
                        dest = togs.Range["A1:D1"];
                        source.Copy();
                        togs.Activate();
                        togs.Range["A1"].Select();
                        togs.Paste();

                        source = togSource.Range["A292:B296"];
                        dest = togs.Range["A265:B269"];
                        source.Copy();
                        togs.Activate();
                        togs.Range["A265"].Select();
                        togs.Paste();


                        topOfLegend = 260;
                        source = togSource.Range["A17:D17"];
                        dest = togs.Range["A261:D261"];
                        source.Copy();
                        togs.Activate();
                        togs.Range["A261"].Select();
                        togs.Paste();
                        togs.Columns["A:D"].Font.Bold = true;
                        //Fit columns to match the 95% widths
                        for (int c = 1; c <= 4; c++) //In Excel the first column is 1 not 0
                        {
                            double colWidth = togSource.Columns[c].ColumnWidth;
                            togs.Columns[c].ColumnWidth = colWidth;
                        }

                        //FILL OUT TOGS INFORMATION
                        sql = string.Format(
                            @"DECLARE @ele_id nvarchar(50);
                DECLARE @ele_name nvarchar(100);
                DECLARE @ele_descr nvarchar(100);
                DECLARE @subtype nvarchar(50);
                DECLARE @FetchStatus int
                DECLARE CA_cursor CURSOR  
	                FOR select Element_Id, Element_Nbr, Element_Descr FROM Element WHERE Element_Id in 
	                (SELECT Element_Id FROM Element_Hierarchy WHERE Parent_Element_Id In 
	                (SELECT Element_Id FROM Element WHERE Element_Nbr = '{0}')) ORDER BY Element_Nbr ASC;

                OPEN CA_cursor  
                FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name,@ele_descr
                WHILE @@FETCH_STATUS = 0
	                BEGIN
                        SELECT @ele_id, @ele_name, @ele_descr
                        SET @subtype = (SELECT Element_Type from Element where Element_Nbr = @ele_name);
		                SELECT DISTINCT File_Nbr, File_Descr from JCMS_File WHERE File_Id in (SELECT File_Id FROM JCMS_File_Owner WHERE File_Owner_Id = @ele_id AND File_Owner_Obj_Type = @subtype) AND File_Class = 'TOGS' ORDER BY File_Nbr ASC ;
                        FETCH NEXT FROM CA_cursor INTO @ele_id, @ele_name, @ele_descr
	                END
                CLOSE CA_cursor;
                DEALLOCATE CA_cursor;
                ", fileNumber);
                        curRow = 2;
                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            SqlDataReader reader = cmd.ExecuteReader();
                            do
                            {
                                while (reader.Read())
                                {
                                    string[] colNames = new string[3];
                                    reader.GetValues(colNames);
                                    if (colNames[2] == null) //I only select two values for togsings, three for CA
                                    {
                                        togs.Cells[curRow, 3].NumberFormat = "@"; //text format the togs
                                        togs.Cells[curRow, 4].NumberFormat = "@";
                                        togs.Cells[curRow, 3].Value = colNames[0];
                                        togs.Cells[curRow, 4].Value = colNames[1];
                                        curRow += 1;
                                    }
                                    else
                                    {
                                        togs.Cells[curRow, 1].NumberFormat = "@";
                                        togs.Cells[curRow, 2].NumberFormat = "@";
                                        togs.Cells[curRow, 1].Value = colNames[1];
                                        togs.Cells[curRow, 2].Value = colNames[2];
                                        curRow += 1;
                                    }
                                }
                            } while (reader.NextResult());
                            reader.Close();
                        }
                        //find where Key - Editable Columns are Bold legend exists on the sheet
                        togs.Range["A65:B65"].Font.Bold = false;
                        RemoveEmptyRows(topOfLegend, togs);
                        string[] wsNames =
                        {
                            "Facility", "Work Breakdown Structure", "Drawings", "Bill of Materials",
                            "Resources", "TOGS"
                        };


                        foreach (Excel.Worksheet sh in newWorkbook.Sheets)
                        {
                            if (wsNames.Contains(sh.Name))
                            {
                                sh.Select();
                                if (sh.Name == "Facility")
                                {
                                    (sh.Rows[1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["A1:F1"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                }
                                else if (sh.Name == "Work Breakdown Structure")
                                {
                                    lastEdgeRow = 0;
                                    for (int k = 1; k <= sh.UsedRange.Rows.Count; k++)
                                    {
                                        if (sh.Range["A" + k.ToString()].Value == "Key - Editable Columns are bold")
                                        {
                                            lastEdgeRow = k - 2;
                                            sh.Range["B" + (k + 1).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 2).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 3).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 4).ToString()].Font.Bold = false;
                                        }
                                    }
                                    (sh.Columns[3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["C1:C" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                                    (sh.Rows[1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["A1:E1"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                }
                                else if (sh.Name == "Drawings")
                                {
                                    (sh.Rows[1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["A1:D1"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                    for (int k = 1; k <= sh.UsedRange.Rows.Count; k++)
                                    {
                                        if (sh.Range["A" + k.ToString()].Value == "Key - Editable Columns are bold")
                                        {
                                            sh.Range["B" + (k + 1).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 2).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 3).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 4).ToString()].Font.Bold = false;
                                        }
                                    }
                                }
                                else if (sh.Name == "TOGS")
                                {
                                    (sh.Rows[1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["A1:D1"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                    for (int k = 1; k <= sh.UsedRange.Rows.Count; k++)
                                    {
                                        if (sh.Range["A" + k.ToString()].Value == "Key - Editable Columns are bold")
                                        {
                                            sh.Range["B" + (k + 1).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 2).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 3).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 4).ToString()].Font.Bold = false;
                                        }
                                    }
                                }
                                else if (sh.Name == "Bill of Materials")
                                {
                                    lastEdgeRow = 0;
                                    (sh.Rows[2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["A2:T2"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                    sh.Range["A2:T2"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight =
                                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

                                    for (int k = 1; k <= sh.UsedRange.Rows.Count; k++)
                                    {
                                        if (sh.Range["A" + k.ToString()].Value == "Key - Editable Columns are bold")
                                        {
                                            lastEdgeRow = k - 2;
                                            sh.Range["B" + (k + 1).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 2).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 3).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 4).ToString()].Font.Bold = false;
                                        }
                                    }
                                    (sh.Columns[2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["B1:B" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                                    sh.Range["B" + (lastEdgeRow + 3) + ":B" + (lastEdgeRow + 6)].Select();
                                    sh.Range["B" + (lastEdgeRow + 3) + ":B" + (lastEdgeRow + 6)].Borders[
                                            Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                    sh.Range["B" + (lastEdgeRow + 3) + ":B" + (lastEdgeRow + 6)].Borders[
                                            Excel.XlBordersIndex.xlEdgeRight].Weight =
                                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;


                                    (sh.Columns[4] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["D1:D" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                                    (sh.Columns[6] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["F1:F" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                                    (sh.Columns[16] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle
                                        =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["P1:P" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                                    (sh.Columns[17] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle
                                        =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["Q1:Q" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                                    (sh.Columns[20] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle
                                        =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["T1:T" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                                    sh.Range["T1:T" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight =
                                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

                                }
                                else if (sh.Name == "Resources")
                                {
                                    lastEdgeRow = 0;
                                    for (int k = 1; k <= sh.UsedRange.Rows.Count; k++)
                                    {
                                        if (sh.Range["A" + k.ToString()].Value == "Key - Editable Columns are bold")
                                        {
                                            lastEdgeRow = k - 2;
                                            sh.Range["B" + (k + 1).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 2).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 3).ToString()].Font.Bold = false;
                                            sh.Range["B" + (k + 4).ToString()].Font.Bold = false;
                                        }
                                    }
                                    int lastCol = 0;
                                    for (int k = 1; k <= sh.UsedRange.Columns.Count; k++)
                                    {
                                        string s = sh.Cells[1, k].Value;
                                        if (sh.Cells[1, k].Value == null)
                                        {
                                            lastCol = k - 1;
                                            break;
                                        }
                                    }
                                    if (lastCol == 0)
                                    {
                                        lastCol = sh.UsedRange.Rows.Count;
                                    }
                                    if (lastCol < 22 || lastCol > 30)
                                    {
                                        lastCol = 22;
                                    }
                                    /*(sh.Rows[2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range[sh.Cells[2, 1], sh.Cells[2, lastCol]].Borders[
                                        Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                    sh.Range[sh.Cells[2, 1], sh.Cells[2, lastCol]].Borders[
                                        Excel.XlBordersIndex.xlEdgeBottom].Weight =
                                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;*/


                                    (sh.Columns[2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["B1:B" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                                    sh.Range["B" + (lastEdgeRow + 3) + ":B" + (lastEdgeRow + 6)].Select();
                                    sh.Range["B" + (lastEdgeRow + 3) + ":B" + (lastEdgeRow + 6)].Borders[
                                            Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                    sh.Range["B" + (lastEdgeRow + 3) + ":B" + (lastEdgeRow + 6)].Borders[
                                            Excel.XlBordersIndex.xlEdgeRight].Weight =
                                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;

                                    (sh.Columns[4] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["D1:D" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                                    (sh.Columns[6] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["F1:F" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;


                                    (sh.Columns[8] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["H1:H" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                                    (sh.Columns[9] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    sh.Range["I1:I" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                                    if (sh.Range["W1"].Value != "EQUIPMENT #4")
                                    {
                                        sh.Range["V1:V" + lastEdgeRow].Borders[Excel.XlBordersIndex.xlEdgeRight]
                                                .LineStyle =
                                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                        sh.Range["V1:V2"].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight =
                                            Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                                    }
                                    else
                                    {
                                        lastCol = 24;
                                        if (sh.Range["Y1"].Value == "EQUIPMENT #5")
                                        {
                                            lastCol = 26;
                                        }
                                        if (sh.Range["Y1"].Value == "EQUIPMENT #6")
                                        {
                                            lastCol = 28;
                                        }
                                        if (sh.Range["Y1"].Value == "EQUIPMENT #7")
                                        {
                                            lastCol = 30;
                                        }
                                        if (sh.Range["Y1"].Value == "EQUIPMENT #8")
                                        {
                                            lastCol = 32;
                                        }
                                        if (sh.Range["Y1"].Value == "EQUIPMENT #9")
                                        {
                                            lastCol = 34;
                                        }
                                        if (sh.Range["Y1"].Value == "EQUIPMENT #10")
                                        {
                                            lastCol = 36;
                                        }
                                        (sh.Rows[2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle
                                            =
                                            Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        sh.Range[sh.Cells[2, 1], sh.Cells[2, lastCol]].Borders[
                                                Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                                            Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                        sh.Range[sh.Cells[2, 1], sh.Cells[2, lastCol]].Borders[
                                                Excel.XlBordersIndex.xlEdgeBottom].Weight =
                                            Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                                    }
                                    //else we are just skipping this.  Let the previous part of the program take of this, look at 1492500AB
                                }
                            }
                        }

                        foreach (Excel.Range cell in facility.get_Range("B4:B6", "B8:B10"))
                        {
                            if (cell.Value != null)
                            {
                                double newVal;
                                if (Double.TryParse(cell.Value.ToString(), out newVal)) //has a decimal to be rounded to
                                {
                                    var dec = newVal - Math.Truncate(newVal);
                                    dec = SetSigFigs(dec, 2);
                                    if (dec.ToString().Length > 0 && dec != 0)
                                    {
                                        string formatString;
                                        if (newVal < 1)
                                        {
                                            formatString = "0.";
                                        }
                                        else
                                        {
                                            formatString = "#.";
                                        }
                                        string sdec = dec.ToString("0.################");
                                        for (int fc = 0; fc < dec.ToString("0.################").Length - 2; fc++)
                                            //0. is two, so subtract from this
                                        {
                                            formatString += "#";
                                        }
                                        cell.NumberFormat = formatString;
                                    }
                                }
                            }

                        }
                        string[] wsNamesArr = {"Bill of Materials", "Resources"};
                        foreach (Excel.Worksheet sh in newWorkbook.Sheets)
                        {
                            if (wsNamesArr.Contains(sh.Name))
                            {
                                sh.Select();
                                sh.Application.ActiveWindow.SplitColumn = 2;
                                sh.Application.ActiveWindow.FreezePanes = true;
                                while (sh.Application.ActiveWindow.FreezePanes == false ||
                                       sh.Application.ActiveWindow.SplitColumn != 2)
                                {
                                    sh.Application.ActiveWindow.SplitColumn = 2;
                                    sh.Application.ActiveWindow.FreezePanes = true;
                                }
                            }
                        }

                        foreach (Excel.Worksheet csh in newWorkbook.Sheets) //Hide all comments on all sheets
                        {
                            for (int j = 1; j <= csh.Comments.Count; j++)
                            {
                                csh.Comments[j].Visible = false;
                            }
                        }
                        xl.DisplayCommentIndicator = Excel.XlCommentDisplayMode.xlCommentIndicatorOnly;
                            //turn off all comments, I am less than confident that the above loop works

                        wbs.Columns[3].HorizontalAlignment = Excel.Constants.xlCenter; //Center the values in column C

                        facility.Range["B4:B6"].Style = "Normal";
                        facility.Range["B8:B10"].Style = "Normal";

                        facility.Range["B4:B6"].HorizontalAlignment = Excel.Constants.xlLeft;
                        facility.Range["B8:B10"].HorizontalAlignment = Excel.Constants.xlLeft;

                        foreach (Excel.Range cell in facility.Range["B4:B10"].Cells)
                        {
                            if (cell.Value > 0)
                            {
                                cell.NumberFormat = "#.##";
                            }
                        }


                        facility.Columns["C:E"].AutoFit(); //fit the column width to the contents

                        bom.Columns["C:D"].HorizontalAlignment = Excel.Constants.xlCenter;

                        res.Columns["C:D"].HorizontalAlignment = Excel.Constants.xlCenter;

                        for (int g = 3; g <= bom.UsedRange.Rows.Count; g++)
                        {
                            if (bom.Range["G" + g.ToString()].Value != null)
                            {
                                if (bom.Range["G" + g.ToString()].Value.ToString().Contains("ZZ"))
                                {
                                    bom.Range["G" + g.ToString()].Style = "Potential Problem/Missing Data";
                                    bom.Range["G" + g.ToString()].AddComment("ZZ NSN");
                                }
                            }
                        }
                        //Add column validation to bom and res
                        addValidation2(bom, "D");
                        addValidation(bom, "F");
                        addValidation(bom, "H:L");
                        addValidation(bom, "N:T");
                        addValidation2(res, "D");
                        addValidation(res, "F");
                        addValidation(res, "G");
                        addValidation(res, "I:P");
                        addValidation(wbs, "E");

                        fileName = fileName.Replace("/", "-");
                        string baseDirPath = Path.Combine(@"C:\Create_Workbooks\New Workbooks",
                            fileName.ToString().Replace(".xlsx", "").Replace(".xlsm", "").Replace("/", "-"));
                        Directory.CreateDirectory(baseDirPath);
                        Directory.CreateDirectory(Path.Combine(baseDirPath, "TOGS"));
                        Directory.CreateDirectory(Path.Combine(baseDirPath, "TOGS", "DOC"));
                        Directory.CreateDirectory(Path.Combine(baseDirPath, "SupportFiles"));
                        Directory.CreateDirectory(Path.Combine(baseDirPath, "SupportFiles", "DA"));
                        Directory.CreateDirectory(Path.Combine(baseDirPath, "SupportFiles", "Schedule"));
                        Directory.CreateDirectory(Path.Combine(baseDirPath, "SupportFiles", "ProductData"));
                        Directory.CreateDirectory(Path.Combine(baseDirPath, "Drawings"));
                        Directory.CreateDirectory(Path.Combine(baseDirPath, "Drawings", "PDF"));
                        int finalDirLength = Path.Combine(baseDirPath, fileName + ".xlsx").Length;
                        newWorkbook.SaveAs(Path.Combine(baseDirPath, fileName + ".xlsx"));
                        newWorkbook.Close(true);
                        exWb.Close(false);
                    }
                }
            }
            xl.Quit();
            Console.ReadLine();
        }

        public static string CheckNull(object v1)
        {

            if (v1 == null)
            {
                return "null";
            }
            else
            {
                return v1.ToString();
            }
        }
        public static double SetSigFigs(double d, int digits)
        {
            if (d == 0)
                return 0;

            decimal scale = (decimal) Math.Pow(10, Math.Floor(Math.Log10(Math.Abs(d))) + 1);

            return (double) (scale * Math.Round((decimal) d / scale, digits));
        }

        public static void RemoveEmptyRows(int topOfKey, Excel.Worksheet ws)
        {
            ws.Activate();
            ws.Rows[topOfKey].Select();
            ws.Rows[topOfKey + 2].Delete();
            ws.Rows[topOfKey + 2].Delete();
            for (int rn = topOfKey - 1; rn > 1; rn--)
            {
                double numCells = xl.WorksheetFunction.CountA(ws.Rows[rn]);
                if ((int) xl.WorksheetFunction.CountIf(ws.Rows[rn], "> \"\"") == 0)
                {
                    ws.Rows[rn].Delete();
                }
                else
                {
                    break;
                }
            }
        }

        public static void addValidation(Excel.Worksheet ws, string columnLetter)
        {
//"IsText(" + xl.ActiveCell.Value.ToString() + ")"
            var val = new Random();
            var rnCells = ws.Columns[columnLetter];
            rnCells.Validation.Delete();
            rnCells.Validation.Add(
                Microsoft.Office.Interop.Excel.XlDVType.xlValidateCustom,
                Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertInformation,
                Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlNotEqual, "-1");

            //(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, "5","5");
            //ws.Columns[columnLetter].Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, 5);
            rnCells.Validation.InputTitle = "Warning";
            rnCells.Validation.ErrorTitle = "Warning";
            rnCells.Validation.InputMessage = "Do not edit this column";
            rnCells.Validation.ErrorMessage = "Do not edit this cell!";
        }

        public static void addValidation2(Excel.Worksheet ws, string columnLetter)
        {
//"IsText(" + xl.ActiveCell.Value.ToString() + ")"
            var val = new Random();
            var rnCells = ws.Columns[columnLetter];
            rnCells.Validation.Delete();
            rnCells.Validation.Add(
                Microsoft.Office.Interop.Excel.XlDVType.xlValidateCustom,
                Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertInformation,
                Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlNotEqual, "-1");

            //(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, "5","5");
            //ws.Columns[columnLetter].Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, 5);
            rnCells.Validation.InputTitle = "Information";
            rnCells.Validation.ErrorTitle = "Information";
            rnCells.Validation.InputMessage = "Edit the formula, not the value in this column";
            rnCells.Validation.ErrorMessage = "Edit the formula, not the value in this column";
        }
    }
}
    


