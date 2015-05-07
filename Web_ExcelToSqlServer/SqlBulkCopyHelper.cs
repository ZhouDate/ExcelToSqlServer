//Name：SqlBulkCopyHelper
//Summary：big data insert into sqlserver
//Maker：zzf
//Make time：201504201354
//Other：
//--------------------------------------------
//No     Editor     UpdateTime       Note

using System;
using System.Data;
using System.Data.SqlClient;

namespace Web_ExcelToSqlServer_Demo
{
    public class SqlBulkCopyHelper
    {
        #region Private variables

        // if you use oracle
        // oracle need oracle11g
        // oracle11g --> Oracle.DataAccess.dll --> OracleBulkCopy
        private SqlBulkCopy sqlCpoy;

        #endregion

        #region Constructor

        #region (SqlConnection sql)
        /// <summary>
        /// SqlBulkCopyHelper
        /// </summary>
        public SqlBulkCopyHelper(SqlConnection sql)
        {
            if (sql == null)
            {
                throw new Exception("param SqlConnection is null!");
            }

            sqlCpoy = new SqlBulkCopy(sql);

        }
        #endregion

        #region (string connectionString)
        /// <summary>
        /// SqlBulkCopyHelper
        /// </summary>
        public SqlBulkCopyHelper(string connectionString)
        {
            if (string.IsNullOrEmpty(connectionString))
            {
                throw new Exception("param connectionString is null or empty!");
            }

            sqlCpoy = new SqlBulkCopy(connectionString);
        }
        #endregion

        #region (SqlTransaction externalTransaction)
        /// <summary>
        /// SqlBulkCopyHelper
        /// </summary>
        public SqlBulkCopyHelper(SqlTransaction externalTransaction)
        {
            if (externalTransaction == null)
            {
                throw new Exception("param SqlTransaction is null!");
            }

            SqlConnection sql = externalTransaction.Connection;

            sqlCpoy = new SqlBulkCopy(sql, SqlBulkCopyOptions.Default, externalTransaction);
        }
        #endregion

        #endregion

        #region Public method

        #region AddColumnMapping
        /// <summary>
        /// AddColumnMapping
        /// </summary>
        /// <param name="sourceColumn">SourceColumnName</param>
        /// <param name="destinationColumn">DestinationColumnName</param>
        public void AddColumnMapping(string sourceColumn, string destinationColumn)
        {
            if (string.IsNullOrEmpty(sourceColumn) || string.IsNullOrEmpty(destinationColumn))
            {
                throw new Exception("param is null or empty!");
            }

            // AddColumnMapping
            sqlCpoy.ColumnMappings.Add(sourceColumn, destinationColumn);
        }
        #endregion

        #region insert into database (DataTable dt)
        /// <summary>
        /// insert into database
        /// </summary>
        /// <param name="dt">DataTable</param>
        public void WriteData(DataTable dt)
        {
            WriteData(dt, null);
        }
        #endregion

        #region insert into database (DataTable dt,string DBTableName)
        /// <summary>
        ///  insert into database 
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="DBTableName">TableName</param>
        public void WriteData(DataTable dt, string DBTableName)
        {
            // param DataTable is null
            if (dt == null)
            {
                throw new Exception("param DataTable is null!");
            }

            // sqlCpoy is null
            if (sqlCpoy == null)
            {
                throw new Exception("sqlCpoy is null!");
            }

            try
            {
                // set table name
                sqlCpoy.DestinationTableName = string.IsNullOrEmpty(DBTableName) ? dt.TableName : DBTableName;

                // if ColumnMappings not have data
                if (sqlCpoy.ColumnMappings.Count == 0)
                {
                    // if database column name equal datatable column name
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        // add ColumnMapping
                        sqlCpoy.ColumnMappings.Add(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName);
                    }
                }

                // WriteToServer
                sqlCpoy.WriteToServer(dt);
            }
            catch (Exception e)
            {
                // goto
                throw e;
            }
            finally
            {
                //close
                sqlCpoy.Close();
            }
        }
        #endregion

        #endregion
    }
}