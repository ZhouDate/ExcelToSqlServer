//Name：CompareHelper
//Summary：select repeat data in database
//Maker：zzf
//Make time：201504201354
//Other：
//--------------------------------------------
//No     Editor     UpdateTime       Note

using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace Web_ExcelToSqlServer_Demo
{
    public class CompareHelper
    {
        #region Private variables

        private string connectionstring;
        private SqlConnection conn;
        private SqlDataAdapter adapter;
        private DataSet ds;
        private SqlCommand comm;

        #endregion

        #region Constructor

        #region (string ConnectionString)
        /// <summary>
        /// CompareHelper
        /// </summary>
        /// <param name="ConnectionString">ConnectionString</param>
        public CompareHelper(string ConnectionString)
        {
            this.connectionstring = ConnectionString;
        }
        #endregion

        #endregion

        #region Public method

        #region SetDBData
        /// <summary>
        /// SetDBData
        /// </summary>
        /// <returns></returns>
        public void SetDBData()
        {
            using (conn = new SqlConnection(this.connectionstring))
            {
                // open database
                conn.Open();

                // sql text
                StringBuilder sb = new StringBuilder();

                // because "SELECT MAX(ROWNUM)", So it's possible that [ROWNUM] in database greater than [ROWNUM] in excel
                sb.Append(" UPDATE TABLEDEMO                                ");
                sb.Append(" SET ROWNUM = 0                                  ");

                comm = conn.CreateCommand();
                comm.CommandText = sb.ToString();
                comm.CommandType = CommandType.Text;

                // exceute sql
                comm.ExecuteNonQuery();
            }
        }
        #endregion

        #region CompareDBData
        /// <summary>
        /// CompareDBData
        /// </summary>
        /// <returns></returns>
        public DataSet CompareDBData()
        {
            using (conn = new SqlConnection(this.connectionstring))
            {
                // open database
                conn.Open();

                // sql text
                StringBuilder sb = new StringBuilder();

                // create temp table "temp1"
                sb.Append(" CREATE TABLE #TEMP1                             ");
                sb.Append(" (                                               ");
                sb.Append("     [ROWNUM] [INT] NULL,                        ");
                sb.Append("  	[INSERTTIME] [DATETIME] NULL,               ");
                sb.Append(" 	[COLUMN1] [VARCHAR](50) NULL,               ");
                sb.Append(" 	[COLUMN2] [VARCHAR](50) NULL,               ");
                sb.Append(" 	[COLUMN3] [VARCHAR](50) NULL,               ");
                sb.Append(" 	[COLUMN4] [VARCHAR](50) NULL,               ");
                sb.Append(" 	[COLUMN5] [VARCHAR](50) NULL                ");
                sb.Append(" );                                              ");

                // select repeat data and insert into "temp1"
                sb.Append(" INSERT INTO #TEMP1                              ");
                sb.Append(" (ROWNUM,INSERTTIME,COLUMN1,COLUMN2,             ");
                sb.Append(" COLUMN3,COLUMN4,COLUMN5)                        ");
                sb.Append(" SELECT MAX(ROWNUM),MAX(INSERTTIME),             ");
                sb.Append(" COLUMN1,COLUMN2,COLUMN3,COLUMN4,COLUMN5         ");
                sb.Append(" FROM TABLEDEMO                                  ");
                sb.Append(" GROUP BY COLUMN1,COLUMN2,COLUMN3,COLUMN4,COLUMN5");
                sb.Append(" HAVING COUNT(*) > 1 ;                           ");

                // delete all data in batch of this
                sb.Append(" DELETE                                          ");
                sb.Append(" FROM TABLEDEMO                                  ");
                sb.Append(" WHERE INSERTTIME IN (SELECT                     ");
                sb.Append(" INSERTTIME FROM #TEMP1);                        ");
                                                                            
                // return all repeat data                                   
                sb.Append(" SELECT ROWNUM,INSERTTIME,COLUMN1,               ");
                sb.Append(" COLUMN2,COLUMN3,COLUMN4,COLUMN5                 ");
                sb.Append(" FROM #TEMP1 ORDER BY ROWNUM DESC;               ");

                // exceute sql
                adapter = new SqlDataAdapter(sb.ToString(), conn);

                ds = new DataSet();

                // fill 
                adapter.Fill(ds);

                return ds;
            }
        }
        #endregion

        #endregion
    }
}