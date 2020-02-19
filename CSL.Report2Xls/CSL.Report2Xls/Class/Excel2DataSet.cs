using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSL.Report2Xls.Class
{
    public class Excel2DataSet
    {
         /// <summary>  
        ///  读取excel文件内容并存放在DataSet中.  
        /// </summary>  
        /// <returns>返回DataSet对象</returns>  
        public static DataSet ExcelToDS(string path, out string err)
        {
            DataSet ds = null;
            err = "";
            try
            {
                string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + @path + ";Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'";
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string strExcel = "select * from [sheet1$]";
                OleDbDataAdapter myCommand = new OleDbDataAdapter(strExcel, strConn);
                DataTable table1 = new DataTable();
                ds = new DataSet();
                myCommand.Fill(table1);
                myCommand.Fill(ds);
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }
            return ds;
        }
    }
}
