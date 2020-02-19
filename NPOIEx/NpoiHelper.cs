using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace NPOIEx
{
    public class NpoiHelper
    {

        #region DataSet/DataTable To Xls
        /// <summary>
        /// DataSet To Xls
        /// </summary>
        /// <param name="path"></param>
        /// <param name="ds"></param>
        /// <returns></returns>
        public static Tuple<bool,string> DataSet2Xls(string path, DataSet ds)
        {
            try
            {
                HSSFWorkbook wk = new HSSFWorkbook();//创建工作簿
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    CreateASheet(wk, ds.Tables[i], ds.Tables[i].TableName);
                }
                Save(path, wk);
                return new Tuple<bool, string>(true, string.Empty);
            }
            catch (Exception ex)
            {
                return new Tuple<bool, string>(false, ex.ToString());
            }
        }
        /// <summary>
        /// DataTable To Xls
        /// </summary>
        /// <param name="path"></param>
        /// <param name="dt"></param>
        /// <param name="sheetName"></param>
        public static Tuple<bool, string> DataTable2Xls(string path, DataTable dt, string sheetName = null)
        {
            try
            {
                HSSFWorkbook wk = new HSSFWorkbook();//创建工作簿
                CreateASheet(wk, dt, sheetName);
                Save(path, wk);
                return new Tuple<bool, string>(true, string.Empty);
            }
            catch (Exception ex)
            {
                return new Tuple<bool, string>(false, ex.ToString());
            }
        }

        private static void CreateASheet(HSSFWorkbook wk, DataTable dt, string sheetName = null)
        {
            if (wk == null || dt == null)
            {
                return;
            }
            if (sheetName == null)
            {
                sheetName = wk.NumberOfSheets == 0 ? "Sheet1" : "Sheet" + (wk.NumberOfSheets + 1);
            }
            ISheet tb = wk.CreateSheet(sheetName);
            int rowsCount = dt.Rows.Count;
            int colCount = dt.Columns.Count;
            //第一行为列头
            IRow row = tb.CreateRow(0);
            for (int i = 0; i < colCount; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }
            //数据记录
            for (int i = 0; i < rowsCount; i++)
            {
                row = tb.CreateRow(i + 1);
                for (int j = 0; j < colCount; j++)
                {

                    ICell cell = row.CreateCell(j);
                    tb.SetColumnWidth(j, 20 * 256);

                    string value = dt.Rows[i][j] is DBNull ? string.Empty : dt.Rows[i][j].ToString();
                    switch (dt.Rows[i][j].GetType().ToString())
                    {
                        case "System.String"://字符串类型 
                            cell.SetCellValue(value);
                            break;
                        case "System.DateTime"://日期类型
                            cell.SetCellValue(DateTime.Parse(value));
                            IDataFormat dataFormat= wk.CreateDataFormat();
                            ICellStyle cellStyle = wk.CreateCellStyle();
                            cellStyle.DataFormat = dataFormat.GetFormat("yyyy-MM-dd HH:mm:ss");
                            cell.CellStyle = cellStyle;
                            break;
                        case "System.Int16"://整型
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            cell.SetCellValue(Int64.Parse(value));
                            break;
                        case "System.Decimal"://浮点型
                        case "System.Double":
                            cell.SetCellValue(double.Parse(value));
                            break;
                    }
                }
            }
        }

        private static void Save(string path, HSSFWorkbook wk)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Write
                     ))
                {
                    wk.Write(fs);
                }
            }
            catch (IOException ex)
            {
                
            }
        }
        #endregion

        #region Xls To DataSet/DataTable 
        /// <summary>
        /// 根据sheet名称返回DataTable对象
        /// </summary>
        /// <param name="path"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static DataTable Xls2DataTable(string path,string sheetName, int[] skipColIndex = null)
        {
            DataSet ds = Xls2DataSet(path, skipColIndex);
            foreach (DataTable dt in ds.Tables)
            {
                if (dt.TableName==sheetName)
                {
                    return dt;
                }
            }
            return null;
        }
       
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static DataSet Xls2DataSet(string path, int[] skipColIndex = null)
        {
            DataSet ds = new DataSet();
            HSSFWorkbook wk = GetWorkbook(path);

            for (int i = 0; i < wk.NumberOfSheets; i++)
            {
                ISheet sheet = wk.GetSheetAt(i);//读取当前表数据
                DataTable dt = new DataTable(sheet.SheetName);

                //第一行数据
                object[] obj = GetRowData(sheet, 0, skipColIndex);
                //设置DataTable列头
                for (int col = 0; col < obj.Count(); col++)
                {
                    dt.Columns.Add(obj[col].ToString());
                }

                for (int j = 1; j <= sheet.LastRowNum; j++)//LastRowNum当前行数据
                {
                    dt.Rows.Add(GetRowData(sheet, j, skipColIndex));
                }
                ds.Tables.Add(dt);
            }
            return ds;
        }
        /// <summary>
        /// 获取一个工作簿
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private static HSSFWorkbook GetWorkbook(string path)
        {
            using (FileStream fs = File.OpenRead(path))
            {
                HSSFWorkbook wk = new HSSFWorkbook(fs);//将xls文件中数据加载到wk
                return wk;
            }
        }
        /// <summary>
        /// 获取某一行数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="skipColIndex">跳过指定列索引数据</param>
        /// <returns></returns>
        private static object[] GetRowData(ISheet sheet,int row,int[] skipColIndex=null)
        {
            IRow r = sheet.GetRow(row);//当前行数据
            List<object> lst = new List<object>();
            if (r != null)
            {
                for (int k = 0; k < r.LastCellNum; k++)//LastCellNum当前行的总列数
                {
                    if (skipColIndex!=null&&skipColIndex.Contains(k))
                    {
                        continue;
                    }
                    ICell cell = r.GetCell(k);//当前表格
                    lst.Add(cell);
                }
            }
            return lst.ToArray();
        }
        #endregion



    }
}
