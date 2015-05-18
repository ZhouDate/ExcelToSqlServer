//Name：NOPIExcelHelper
//Summary：open and read Excel
//Maker：zzf
//Make time：201505281422
//Other：
//--------------------------------------------
//No     Editor     UpdateTime       Note

using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using System.Data;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace Web_ExcelToSqlServer_Demo
{
    public class NOPIExcelHelper : IDisposable
    {
        #region Private variables

        private string fileName = null;
        private Stream stream = null;
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private bool disposed;
        private List<int> _ConstraintList;

        #endregion

        #region Constructor

        #region NOPIExcelHelper(string fileName)
        /// <summary>
        /// NOPIExcelHelper
        /// </summary>
        /// <param name="fileName"></param>
        public NOPIExcelHelper(string fileName)
        {
            _ConstraintList = new List<int>();
            this.fileName = fileName;
            disposed = false;
        }
        #endregion

        #region NOPIExcelHelper(Stream stream, string fileName)
        /// <summary>
        /// NOPIExcelHelper
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="fileName"></param>
        public NOPIExcelHelper(Stream stream, string fileName)
        {
            _ConstraintList = new List<int>();
            this.stream = stream;
            this.fileName = fileName;
            disposed = false;
        }
        #endregion

        #endregion

        #region Property

        #region ConstraintList
        /// <summary>
        /// ConstraintList
        /// </summary>
        public List<int> ConstraintList
        {
            get { return _ConstraintList; }
            set { _ConstraintList = value; }
        }
        #endregion

        #endregion

        #region Public Method

        #region ReadExcelToDataTable
        /// <summary>
        /// ReadExcelToDataTable
        /// </summary>
        /// <param name="sheetName">sheetName</param>
        /// <returns>DataTable</returns>
        public DataTable ReadExcelToDataTable(string sheetName)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                if (stream == null)
                {
                    //read IO
                    fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                    stream = fs;
                }

                //File Extension name
                if (fileName.IndexOf(".xlsx") > 0) // 2007
                    workbook = new XSSFWorkbook(stream);
                else if (fileName.IndexOf(".xls") > 0) // 2003
                    workbook = new HSSFWorkbook(stream);

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //get sheets[0]
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    //get sheets[0]
                    sheet = workbook.GetSheetAt(0);
                }

                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //ColumnCount

                    //add column name
                    //column list
                    List<DataColumn> collist = new List<DataColumn>();
                    //add Column
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        ICell cell = firstRow.GetCell(i);
                        if (cell != null)
                        {
                            //get string value in cell
                            string cellValue = cell.StringCellValue;
                            if (cellValue != null)
                            {
                                DataColumn column = new DataColumn(cellValue);

                                //add Columns
                                data.Columns.Add(column);

                                //add list
                                collist.Add(column);
                            }
                        }
                    }
                    //set PrimaryKey for all
                    data.PrimaryKey = collist.ToArray();
                    //rownum+1
                    startRow = sheet.FirstRowNum + 1;

                    //add row
                    //LastRowNum
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        //get row
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;

                        //create new datarow
                        DataRow dataRow = data.NewRow();

                        //add cell value
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null)
                                dataRow[j] = row.GetCell(j).ToString();
                        }

                        try
                        {
                            //add row in data
                            data.Rows.Add(dataRow);
                        }
                        catch (ConstraintException e)
                        {
                            //this row repeat in excel
                            //add list
                            _ConstraintList.Add(i);
                            continue;
                        }
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }
        #endregion

        #endregion

        #region Dispose

        #region Dispose
        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

        #region Dispose(bool disposing)
        /// <summary>
        /// Dispose
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fs != null)
                        fs.Close();
                }

                fs = null;
                disposed = true;
            }
        }
        #endregion

        #endregion

    }
}
