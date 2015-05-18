//Name：Web_ExcelToSqlServer_Demo
//Summary：aspx demo
//Maker：zzf
//Make time：201505281422
//Other：
//--------------------------------------------
//No     Editor     UpdateTime       Note
//1      zhou       20150518         Add NOPI and Simplify the process 

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
        // if you use NOPI , it will be better 
        // I use NOPI , 10,000 compared to 10,000 , 2s
        // I use OpenXml , 10,000 compared to 10,000 , 5s
        // because name is GetCellValue method have question  in class of OpenXmlHelper 
        // But I can't guarantee it
        private string ConnectionString = WebConfigurationManager.ConnectionStrings["SqlServerHelper"].ToString();
        private OpenXmlHelper xmlhelper;
        private NOPIExcelHelper nopihelper;
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
                if (name.Length < 2 || (name[1] != "xlsx" && name[1] != "xls"))
                {
                    Response.Write("========================================================================= <br />");
                    Response.Write("file is not xlsx or xls !<br />");
                    Response.Write("========================================================================= <br />");
                    return;
                }

                // begin
                Response.Write("begin time:" + DateTime.Now.ToString() + "<br />");

                //file name 
                string filename = FileUpload1.PostedFile.FileName;

                // read stream
                using (Stream stream = FileUpload1.PostedFile.InputStream)
                {
                    //you cam use OpenXml or NOPI

                    #region OpenXml

                    //// -------- you can use NOPI -------- //
                    //xmlhelper = new OpenXmlHelper(stream);

                    //// read worksheet "sheet1"
                    //DataTable dt = xmlhelper.ReadExcelToDataTable("Sheet1");

                    //// loop display 
                    //for (int i = xmlhelper.ConstraintList.Count; i > 0; i--)
                    //{
                    //    // row in the Excel exist in the Excel
                    //    Response.Write("row number " + xmlhelper.ConstraintList[i] + "in the Excel exist in the Excel<br />");
                    //}

                    #endregion

                    #region NOPI

                    // -------- this is NOPI -------- //
                    nopihelper = new NOPIExcelHelper(stream, filename);

                    // read worksheet "sheet1"
                    DataTable dt = nopihelper.ReadExcelToDataTable("Sheet1");

                    // loop display 
                    for (int i = nopihelper.ConstraintList.Count; i > 0; i--)
                    {
                        // row in the Excel exist in the Excel
                        Response.Write("row number " + nopihelper.ConstraintList[i] + "in the Excel exist in the Excel<br />");
                    }

                    #endregion

                    // set this table name equal table name in database
                    dt.TableName = "TableDemo";

                    // compare in database
                    comparehelper = new CompareHelper(ConnectionString);

                    // write to database
                    sqlcopyhelper = new SqlBulkCopyHelper(comparehelper.BeginTransaction());
                    sqlcopyhelper.WriteData(dt);

                    // compare repeat data
                    DataSet ds = comparehelper.CompareDBData();

                    #region write date
                    if (ds != null && ds.Tables.Count != 0)
                    {
                        DataTable dt2 = ds.Tables[0];

                        int flag = 0;

                        if (dt2.Rows.Count > 0)
                        {
                            //rollback
                            comparehelper.RollbackTransaction();

                            DataTable dt3 = ds.Tables[1];
                            flag = int.Parse(dt3.Rows[0][0].ToString());

                            //flag - dt.Rows.Count = min row num ,because row begin in 2,so - 1 
                            flag = flag - dt.Rows.Count - 1;

                            // loop display 
                            for (int i = 0; i < dt2.Rows.Count; i++)
                            {
                                Response.Write(string.Format("row number {0} in the Excel exist in the database<br />", Convert.ToInt32(dt2.Rows[i]["ID"]) - flag));
                            }

                            // end
                            Response.Write("end time:" + DateTime.Now.ToString() + "<br />");
                            return;
                        }

                        //commint
                        comparehelper.CommintTransaction();
                    }
                    #endregion
                }

                // end
                Response.Write("end time:" + DateTime.Now.ToString() + "<br />");
            }
        }
        #endregion

        #endregion
    }
}