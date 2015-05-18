//Name：CompareHelper
//Summary：select repeat data in database
//Maker：zzf
//Make time：201504201354
//Other：
//--------------------------------------------
//No     Editor     UpdateTime       Note
//1      zhou       20150518         Simplify the process 

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
        private SqlCommand comm;
        private SqlDataAdapter adapter;
        private SqlTransaction tran;
        private DataSet ds;

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

        #region BeginTransaction
        /// <summary>
        /// BeginTransaction
        /// </summary>
        /// <returns></returns>
        public SqlTransaction BeginTransaction()
        {
            conn = new SqlConnection(this.connectionstring);
            conn.Open();
            comm = conn.CreateCommand();
            tran = conn.BeginTransaction();
            comm.Transaction = tran;

            return tran;
        }
        #endregion

        #region CommintTransaction
        /// <summary>
        /// CommintTransaction
        /// </summary>
        /// <returns></returns>
        public void CommintTransaction()
        {
            if (tran != null)
                tran.Commit();

            if (conn != null && conn.State == ConnectionState.Open)
                conn.Close();
        }
        #endregion

        #region RollbackTransaction
        /// <summary>
        /// RollbackTransaction
        /// </summary>
        /// <returns></returns>
        public void RollbackTransaction()
        {
            if (tran != null)
                tran.Rollback();

            if (conn != null && conn.State == ConnectionState.Open)
                conn.Close();
        }
        #endregion

        #region CompareDBData
        /// <summary>
        /// CompareDBData
        /// </summary>
        /// <returns></returns>
        public DataSet CompareDBData()
        {

            // sql text
            StringBuilder sb = new StringBuilder();

            // select repeat data
            sb.Append(" SELECT T1.ID                   ");
            sb.Append(" ,T1.[COLUMN1]                  ");
            sb.Append(" ,T1.[COLUMN2]                  ");
            sb.Append(" ,T1.[COLUMN3]                  ");
            sb.Append(" ,T1.[COLUMN4]                  ");
            sb.Append(" ,T1.[COLUMN5]                  ");
            sb.Append(" FROM (                         ");
            sb.Append("  SELECT MAX([ID]) AS ID        ");
            sb.Append("  ,[COLUMN1]                    ");
            sb.Append("  ,[COLUMN2]                    ");
            sb.Append("  ,[COLUMN3]                    ");
            sb.Append("  ,[COLUMN4]                    ");
            sb.Append("  ,[COLUMN5]                    ");
            sb.Append(" FROM [TABLEDEMO]               ");
            sb.Append(" GROUP BY [COLUMN1],[COLUMN2],  ");
            sb.Append(" [COLUMN3],[COLUMN4],[COLUMN5]  ");
            sb.Append(" HAVING COUNT(*) > 1            ");
            sb.Append(" ) T1                           ");
            sb.Append(" ORDER BY T1.ID DESC;           ");

            // select MAX ID
            sb.Append(" SELECT ISNULL(MAX(ID),0) AS ID FROM [TABLEDEMO];");

            // set Command
            comm.CommandType = CommandType.Text;
            comm.CommandText = sb.ToString();

            // exceute sql
            adapter = new SqlDataAdapter(comm);

            ds = new DataSet();

            // fill 
            adapter.Fill(ds);

            return ds;

        }
        #endregion

        #endregion
    }
}