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
            if (fileExcel.HasFile)
            {
                string FileExtension = System.IO.Path.GetExtension(fileExcel.FileName);

                if (FileExtension == ".xlsx" || FileExtension == ".xls")
                {
                    string strFileName = DateTime.Now.Day.ToString() + '_' + DateTime.Now.Month.ToString() + '_' + DateTime.Now.Year.ToString() + '_' + DateTime.Now.Hour.ToString() + '_' +
                                         DateTime.Now.Minute.ToString() + '_' + fileExcel.FileName.ToString();

                    string filePath = Server.MapPath("~/Upload/" + strFileName);
                    fileExcel.SaveAs(filePath);

                    try
                    {
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                        {
                            // Licence for Non-Commercial applications
                            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


                            // Assuming the data is in the first worksheet
                            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                            // Access data from the worksheet
                            int rowCount = worksheet.Dimension.Rows;
                            int colCount = worksheet.Dimension.Columns;

                            DataTable dt = new DataTable();

                            // Assuming the first row contains column headers
                            for (int col = 1; col <= colCount; col++)
                            {
                                dt.Columns.Add(worksheet.Cells[1, col].Text);
                            }

                            // Starting from the second row to skip headers
                            for (int row = 2; row <= rowCount; row++)
                            {
                                DataRow dataRow = dt.NewRow();

                                for (int col = 1; col <= colCount; col++)
                                {
                                    dataRow[col - 1] = worksheet.Cells[row, col].Text;
                                }

                                dt.Rows.Add(dataRow);
                            }

                            // Checking column names present in excel sheet or not
                            if (dt.Columns[0].ColumnName.Trim() == "EmpName" && dt.Columns[1].ColumnName.Trim() == "EmpDesignation" && dt.Columns[2].ColumnName.Trim() == "EmpSalary")
                            {
                                if (dt.Rows.Count > 0)
                                {
                                    // Method 1: delete data from Temp Table
                                    // Method 2: inserting data into temp table if needed

                                    //InsertExcelDataToMainTable(dt);

                                    // Method 3: To display the inserted data from Temp or Main Table
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
        }

    }
}