//Name：OpenXmlHelper
//Summary：OpenXml SDK open and read Excel,need Open XML SDK 2.0 for Microsoft Office,source link http://www.microsoft.com/en-us/download/details.aspx?id=5124 
//Maker：zzf
//Make time：201504201354
//Other：
//--------------------------------------------
//No     Editor     UpdateTime       Note

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Web_ExcelToSqlServer_Demo
{
    public class OpenXmlHelper
    {
        #region Private variables

        //document
        private SpreadsheetDocument doucumentopen;
        //string
        private SharedStringTablePart sharedstring;
        //list of ConstraintException
        private List<int> _ConstraintList = new List<int>();

        #endregion

        #region Property

        #region ConstraintList
        /// <summary>
        /// ConstraintList
        /// </summary>
        public List<int> ConstraintList
        {
            get { return _ConstraintList; }
        }
        #endregion

        #endregion

        #region Constructor

        #region (string file)
        /// <summary>
        /// OpenXmlHelper
        /// </summary>
        /// <param name="file"></param>
        public OpenXmlHelper(string file)
        {
            //null or exist
            if (file == null || !File.Exists(file))
            {
                throw new Exception("file is null or not exist!");
            }

            //open document
            doucumentopen = SpreadsheetDocument.Open(file, false);
            //get stringsheet
            sharedstring = doucumentopen.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

        }
        #endregion

        #region (Stream stream)
        /// <summary>
        /// OpenXmlHelper
        /// </summary>
        /// <param name="stream"></param>
        public OpenXmlHelper(Stream stream)
        {
            //null or exist
            if (stream == null)
            {
                throw new Exception("stream is null");
            }

            //open document
            doucumentopen = SpreadsheetDocument.Open(stream, false);
            //get stringsheet
            sharedstring = doucumentopen.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

        }
        #endregion

        #endregion

        #region Private Method

        #region GetCellValue
        /// <summary>
        /// get string value of cell
        /// </summary>
        /// <param name="c">cell</param>
        /// <returns></returns>
        private string GetCellValue(Cell c)
        {
            //get InnerText
            string value = c.CellValue.InnerText;

            if (c.DataType != null)
            {
                //data type
                switch (c.DataType.Value)
                {
                    case CellValues.SharedString:
                        //get value
                        value = sharedstring.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                        break;
                }
            }

            return value;
        }
        #endregion

        #region ReadExcelToDataTable
        /// <summary>
        /// ReadExcelToDataTable
        /// </summary>
        /// <param name="worksheetpart">WorksheetPart</param>
        /// <param name="time">batch time flag</param>
        /// <returns></returns>
        private DataTable ReadExcelToDataTable(WorksheetPart worksheetpart, DateTime time)
        {
            DataTable dt = null;

            if (worksheetpart != null)
            {
                dt = new DataTable();
                bool HasColumn = true;

                //row index
                //if row number is 1 in excel is column name , RowIndex is 2
                int RowIndex = 2;

                //loop row
                foreach (var row in worksheetpart.Worksheet.Descendants<Row>())
                {
                    //if row number is 1 in excel is column name
                    if (HasColumn)
                    {
                        //column list
                        List<DataColumn> collist = new List<DataColumn>();

                        //data column
                        DataColumn col;

                        //loop add column
                        foreach (var cell in row.Descendants<Cell>())
                        {
                            //get string value in cell
                            col = new DataColumn(GetCellValue(cell), typeof(string));

                            //add Columns
                            dt.Columns.Add(col);

                            //add list
                            collist.Add(col);
                        }

                        //add rownum
                        //it is row number in excel 
                        dt.Columns.Add("ROWNUM");
                        //add inserttime
                        //it is batch flag of insert into table
                        dt.Columns.Add("INSERTTIME", typeof(DateTime));

                        //set PrimaryKey,without "ROWNUM" and "INSERTTIME"
                        dt.PrimaryKey = collist.ToArray();

                        //close add
                        HasColumn = false;
                        continue;
                    }

                    //create new row
                    DataRow dr = dt.NewRow();
                    //column index
                    int ColumnIndex = 0;

                    //set rownum
                    dr["ROWNUM"] = RowIndex;
                    //set inserttime
                    dr["INSERTTIME"] = time;

                    //loop cell
                    foreach (var cell in row.Descendants<Cell>())
                    {
                        //add cell
                        dr[ColumnIndex] = GetCellValue(cell);

                        //Column Index++
                        ColumnIndex++;
                    }

                    try
                    {
                        //add row in dt
                        dt.Rows.Add(dr);
                    }
                    catch (ConstraintException e)
                    {
                        //this row repeat in excel
                        //add list
                        _ConstraintList.Add(RowIndex);
                        continue;
                    }
                    catch (Exception e)
                    {
                        throw e;
                    }
                    finally
                    {
                        RowIndex++;
                    }
                }
            }

            //CloseExcel
            CloseExcel();

            return dt;
        }
        #endregion

        #region Close Excel
        /// <summary>
        /// CloseExcel
        /// </summary>
        private void CloseExcel()
        {
            if (doucumentopen == null)
            {
                throw new Exception("doucumentopen is close or null!");
            }

            //CloseExcel
            doucumentopen.Close();
        }
        #endregion

        #endregion

        #region Public Method

        #region ReadExcelToDataTable (string SheetName,bool columnnameflag)
        /// <summary>
        /// ReadExcelToDataTable
        /// </summary>
        /// <param name="SheetName">SheetName</param>
        /// <returns></returns>
        public DataTable ReadExcelToDataTable(string SheetName)
        {
            // null or empty
            if (string.IsNullOrEmpty(SheetName))
            {
                throw new Exception("SheetName is null or empty!");
            }

            // flag time
            DateTime time = DateTime.Now;

            // get worksheet of name is SheetName
            IEnumerable<Sheet> s = doucumentopen.WorkbookPart.Workbook.Descendants<Sheet>().Where(x => x.Name == SheetName);

            // select WorksheetPart
            WorksheetPart worksheetpart = doucumentopen.WorkbookPart.GetPartById(s.First().Id) as WorksheetPart;

            return ReadExcelToDataTable(worksheetpart, time);
        }
        #endregion

        #region ReadExcelToDataTable (int SheetIndex,bool columnnameflag)
        /// <summary>
        /// ReadExcelToDataTable
        /// </summary>
        /// <param name="SheetIndex">SheetIndex</param>
        /// <returns></returns>
        public DataTable ReadExcelToDataTable(int SheetIndex)
        {
            // SheetIndex < 0
            if (SheetIndex < 0)
            {
                throw new Exception("SheetIndex < 0!");
            }

            // flag time
            DateTime time = DateTime.Now;

            // get number SheetIndex worksheet 
            Sheet sheet = doucumentopen.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(SheetIndex);

            // select WorksheetPart
            WorksheetPart worksheetpart = doucumentopen.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;

            return ReadExcelToDataTable(worksheetpart, time);
        }
        #endregion

        #endregion

    }
}