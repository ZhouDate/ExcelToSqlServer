using System;
using System.Web.Configuration;
using System.IO;
using System.Data;

namespace Web_ExcelToSqlServer_Demo
{
    public partial class Web_Demo : System.Web.UI.Page
    {
        #region Private variables

        // Note:
        // if you use SQL Server, it will be better
        // if you use SQL Server , it will be better 
        // I use SQL Server 2008 R2, 10,000 compared to 10,000 , 3s
        // I use SqlExpress , 10,000 compared to 10,000 , 6s
        // But I can't guarantee it
        private string ConnectionString = WebConfigurationManager.ConnectionStrings["SqlServerHelper"].ToString();
        private OpenXmlHelper xmlhelper;
        private SqlBulkCopyHelper sqlcopyhelper;
        private CompareHelper comparehelper;

        #endregion

        #region Event

        #region Page_Load
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        #endregion

        #region Button1_Click
        protected void Button1_Click(object sender, EventArgs e)
        {
            // exist file
            if (FileUpload1.HasFile)
            {
                string[] name = this.FileUpload1.FileName.Split('.');

                // if you use NOPI ,it can read xls
                if (name.Length < 2 || name[1] != "xlsx")
                {
                    Response.Write("========================================================================= <br />");
                    Response.Write("file is not xlsx !<br />");
                    Response.Write("========================================================================= <br />");
                    return;
                }

                // begin
                Response.Write("begin time:" + DateTime.Now.ToString() + "<br />");

                // read stream
                using (Stream stream = FileUpload1.PostedFile.InputStream)
                {
                    // -------- you can use NOPI -------- //
                    xmlhelper = new OpenXmlHelper(stream);

                    // read worksheet "sheet1"
                    DataTable dt = xmlhelper.ReadExcelToDataTable("Sheet1");

                    // loop display 
                    for (int i = xmlhelper.ConstraintList.Count; i > 0; i--)
                    {
                        // row in the Excel exist in the Excel
                        Response.Write("row number " + xmlhelper.ConstraintList[i] + "in the Excel exist in the Excel<br />");
                    }

                    // set this table name equal table name in database
                    dt.TableName = "TableDemo";

                    // write to database
                    sqlcopyhelper = new SqlBulkCopyHelper(ConnectionString);
                    sqlcopyhelper.WriteData(dt);

                    // compare in database
                    comparehelper = new CompareHelper(ConnectionString);
                    DataSet ds = comparehelper.CompareDBData();

                    if (ds != null && ds.Tables.Count != 0)
                    {
                        DataTable dt2 = ds.Tables[0];

                        // loop display 
                        for (int i = 0; i < dt2.Rows.Count; i++)
                        {
                            Response.Write("row number " + dt2.Rows[i]["ROWNUM"] + " in the Excel exist in the database<br />");
                        }
                    }
                }

                // end
                Response.Write("end time:" + DateTime.Now.ToString() + "<br />");
            }
        }
        #endregion

        #endregion
    }
}