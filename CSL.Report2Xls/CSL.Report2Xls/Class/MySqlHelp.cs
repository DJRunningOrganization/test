using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSL.Report2Xls.Class
{
    /// <summary>
    /// 最后需要手动关闭连接
    /// </summary>
    public class MySqlHelp
    {
        public static string connStr = String.Empty;
        private static object _lock = new object();
        private static MySqlConnection con = null;
        private static MySqlHelp ins;
        private TLConfiguration.DelayerLogger logger = new TLConfiguration.DelayerLogger();
        /// <summary>
        /// 日志记录器
        /// </summary>
        public TLConfiguration.DelayerLogger Logger
        {
            get { return logger; }
            set { logger = value; }
        }

        public static MySqlHelp Instance
        {
            get
            {
                if (ins == null)
                {
                    lock (_lock)
                    {
                        if (ins == null)
                        {
                            ins = new MySqlHelp();
                        }
                    }
                }
                return ins;
            }
        }
        public static void Open()
        {
            if (con == null)
            {
                con = new MySqlConnection(connStr);
                con.Open();
            }
            else if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            else if (con.State == ConnectionState.Broken)
            {
                con.Close();
                con.Open();
            }
        }
        /// <summary>
        /// 关闭数据库连接
        /// </summary>
        public void Close()
        {
            if (con != null && con.State == ConnectionState.Open)
            {
                con.Close();
            }
        }
        public int SaveToDb(DataSet ds)
        {
            StringBuilder sb = null;
            int result = -1;
            if (ds.Tables.Count > 0)
            {
                foreach (DataTable dt in ds.Tables)
                {
                    sb = new StringBuilder();
                    sb.Append(string.Format("insert into {0} values", dt.TableName));
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sb.Append("(NULL,");
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (dt.Rows[i][j] is DBNull||string.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                            {
                                sb.Append("NULL,");

                            }
                            else
                            {
                                sb.Append("'" + dt.Rows[i][j] + "'" + ",");
                            }
                        }
                        sb.Remove(sb.ToString().LastIndexOf(','), 1);
                        sb.Append("),");
                    }
                    sb.Remove(sb.ToString().LastIndexOf(','), 1);
                    sb.Append(";");
                    result = ExecuteNonquery(sb.ToString(), null);
                    if (result==-1)
                    {
                        return result;
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 需要获得多个结果集的时候用该方法，返回DataSet对象
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="paras"></param>
        /// <returns></returns>
        public DataSet ExecuteDataSet(string sql, params MySqlParameter[] paras)
        {
            DataSet ds = new DataSet();
            try
            {
                Open();
                using (MySqlDataAdapter da = new MySqlDataAdapter(sql, con))
                {
                    if (paras != null)
                    {
                        da.SelectCommand.Parameters.AddRange(paras);
                    }
                    da.Fill(ds);
                }
            }
            catch (Exception ex)
            {
                if (logger != null)
                    logger.Write2Logger(TLConfiguration.ELogLevel.Error, "DB", "ExecuteDataSet:获得多个结果集," + ex.ToString());
            }
            finally
            {
                Close();
            }
            return ds;
        }
        /// <summary>
        /// 获得单个结果集时使用该方法，返回DataTable对象。
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="paras"></param>
        /// <returns></returns>
        public DataTable ExecuteDataTable(string sql, params MySqlParameter[] paras)
        {
            DataTable dt = new DataTable();
            try
            {
                Open();
                using (MySqlDataAdapter da = new MySqlDataAdapter(sql, con))
                {
                    if (paras != null)
                    {
                        da.SelectCommand.Parameters.AddRange(paras);
                    }
                    da.Fill(dt);
                }
                return dt;
            }
            catch (Exception ex)
            {
                if (logger != null)
                    logger.Write2Logger(TLConfiguration.ELogLevel.Error, "DB", "ExecuteDataTable:获得单个结果集," + ex.ToString());
                return dt;
            }
            finally
            {
                Close();
            }
        }
        /// <summary>
        /// MySqlDataReader读取表数据
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public MySqlDataReader ExecuteReader(string sql)
        {
            MySqlDataReader reader = null;
            try
            {
                Open();
                using (MySqlCommand cmd = new MySqlCommand(sql, con))
                {
                    reader = cmd.ExecuteReader();
                    return reader;
                }
            }
            catch (Exception ex)
            {
                if (logger != null)
                    logger.Write2Logger(TLConfiguration.ELogLevel.Error, "DB", "ExeReader:获得单个结果集," + ex.ToString());
                return reader;
            }
            finally
            {

            }
        }
        private static object _lockDataReader = new object();
        /// <summary>
        /// 获取第一行的第一列
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="paras"></param>
        /// <returns></returns>
        public object ExecuteScalar(string sql, params MySqlParameter[] paras)
        {
            object obj = null;
            try
            {
                Open();
                using (MySqlCommand cmd = new MySqlCommand(sql, con))
                {
                    if (paras != null)
                    {
                        cmd.Parameters.AddRange(paras);
                    }
                    lock (_lockDataReader)
                    {
                        obj = cmd.ExecuteScalar();
                    }
                    if ((object.Equals(obj, null)) || Object.Equals(obj, System.DBNull.Value))
                    {
                        return null;
                    }
                    else
                    {
                        return obj;
                    }
                }
            }
            catch (Exception ex)
            {
                if (logger != null)
                    logger.Write2Logger(TLConfiguration.ELogLevel.Error, "DB", "ExecuteScalar:获取第一行的第一列," + ex.ToString());
                return obj;
            }
            finally
            {
                Close();
            }
        }

        /// <summary>
        /// 获取BindingSource
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        //public BindingSource ExecuteBindingSource(string sql)
        //{
        //    BindingSource bs = new BindingSource();
        //    try
        //    {
        //        Open();
        //        using (MySqlCommand cmd = new MySqlCommand(sql, con))
        //        {
        //            using (MySqlDataReader dr = cmd.ExecuteReader())
        //            {
        //                bs.DataSource = dr;
        //            }
        //        }
        //        return bs;
        //    }
        //    catch (Exception ex)
        //    {
        //        LogHelp.Instance.PublicException("ExecuteBindingSource:获取BindingSource", ex);
        //        return bs;
        //    }
        //    finally
        //    {
        //        Close();
        //    }
        //}

        /// <summary>
        /// 执行Update,Delete,Insert操作
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public int ExecuteNonquery(string sql, ref long keyID, params MySqlParameter[] paras)
        {
            int result = -1;
            try
            {
                Open();
                using (MySqlCommand cmd = new MySqlCommand(sql, con))
                {
                    if (paras != null)
                    {
                        cmd.Parameters.AddRange(paras);
                    }
                    result = cmd.ExecuteNonQuery();
                    keyID = cmd.LastInsertedId;
                }
                return result;
            }
            catch (Exception ex)
            {
                if (logger != null)
                    logger.Write2Logger(TLConfiguration.ELogLevel.Error, "DB", "ExecuteNonquery:执行Update,Delete,Insert操作," + ex.ToString());
                return result;
            }
            finally
            {
                Close();
            }
        }
        /// <summary>
        /// 执行Update,Delete,Insert操作
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public int ExecuteNonquery(string sql, params MySqlParameter[] paras)
        {
            int result = -1;
            try
            {
                Open();
                using (MySqlCommand cmd = new MySqlCommand(sql, con))
                {
                    if (paras != null)
                    {
                        cmd.Parameters.AddRange(paras);
                    }
                    result = cmd.ExecuteNonQuery();
                }
                return result;
            }
            catch (Exception ex)
            {
                if (logger != null)
                    logger.Write2Logger(TLConfiguration.ELogLevel.Error, "DB", "ExecuteNonquery:执行Update,Delete,Insert操作," + ex.ToString());
                return result;
            }
            finally
            {
                Close();
            }
        }
        /// <summary>
        /// 存储过程 返回Datatable
        /// </summary>
        /// <param name="proName"></param>
        /// <param name="paras"></param>
        /// <returns></returns>
        public DataTable ExecuteProcQuery(string proName, params MySqlParameter[] paras)
        {
            DataTable dt = new DataTable();
            try
            {
                Open();
                using (MySqlCommand cmd = new MySqlCommand(proName, con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    MySqlDataAdapter da = new MySqlDataAdapter(cmd);
                    if (paras != null)
                    {
                        da.SelectCommand.Parameters.AddRange(paras);
                    }
                    da.Fill(dt);
                }
                return dt;
            }
            catch (Exception ex)
            {
                if (logger != null)
                    logger.Write2Logger(TLConfiguration.ELogLevel.Error, "DB", "ExecuteProcQuery:存储过程," + ex.ToString());
                return dt;
            }
            finally
            {
                Close();
            }
        }
        /// <summary>
        /// 多语句的事物管理
        /// </summary>
        /// <param name="cmds"></param>
        /// <returns></returns>
        public bool ExecuteCommandByTram(params MySqlCommand[] cmds)
        {
            try
            {
                Open();
                using (MySqlTransaction tran = con.BeginTransaction())
                {
                    foreach (var item in cmds)
                    {
                        item.Connection = con;
                        item.Transaction = tran;
                        item.ExecuteNonQuery();
                    }
                    tran.Commit();
                }
                return true;
            }
            catch (Exception ex)
            {
                if (logger != null)
                    logger.Write2Logger(TLConfiguration.ELogLevel.Error, "DB", "ExecuteCommandByTram:多语句的事物管理," + ex.ToString());
                return false;
            }
            finally
            {
                Close();
            }

        }

    }
}
