using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace IDisburse
{
    public partial class ExcelUpload : System.Web.UI.Page
    {
        string connectionString = "";// ConfigurationManager.ConnectionStrings["Ginie"].ConnectionString;
        private object worksheet;

        protected void Page_Load(object sender, EventArgs e)
        {

        }


        protected void btnExclUpload_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;

            if (fileExcel.HasFile)
            {
                string FileExtension = System.IO.Path.GetExtension(fileExcel.FileName);

                if (FileExtension == ".xlsx" || FileExtension == ".xls")
                {
                    string strFileName = DateTime.Now.Day.ToString() + '_' + DateTime.Now.Month.ToString() + '_' + DateTime.Now.Year.ToString() + '_' + DateTime.Now.Hour.ToString() + '_' +
                                         DateTime.Now.Minute.ToString() + '_' + fileExcel.FileName.ToString();

                     filePath = Server.MapPath("~/Upload/" + strFileName);
                    fileExcel.SaveAs(filePath);


                    if (File.Exists(filePath))
                    {
                        DataTable dt = ReadExcelFile(filePath);
                        // Use the data (e.g., bind to GridView)
                    }
                }
            }
        }


        private DataTable ReadExcelFile(string filePath)
        {
            string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';";

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                // Get the sheet name (you could also hardcode it if known)
                DataTable schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = schema.Rows[0]["TABLE_NAME"].ToString();

                string query = $"SELECT * FROM [{sheetName}]";

                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    return dt;
                }
            }
        }
      

    }
}
