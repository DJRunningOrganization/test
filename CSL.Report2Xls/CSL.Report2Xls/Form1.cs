using CSL.Report2Xls.Class;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSL.Report2Xls
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public static TLConfiguration.ConfigurationMan configM = null;
        private void button1_Click(object sender, EventArgs e)
        {
            var dt = NPOIEx.NpoiHelper.Xls2DataSet(@"D:/zc.xls",skipColIndex:new int[]{0});
            int ddd=MySqlHelp.Instance.SaveToDb(dt);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            string t_match = "select * from t_match";
            DataTable dt = MySqlHelp.Instance.ExecuteDataTable(t_match, null);
            dt.TableName = "t_match";
            ds.Tables.Add(dt);
            string t_player = "select * from t_player";
            dt = MySqlHelp.Instance.ExecuteDataTable(t_player, null);
            dt.TableName = "t_player";
             ds.Tables.Add(dt);

            string t_team_base = "select * from t_team_base";
            dt = MySqlHelp.Instance.ExecuteDataTable(t_team_base, null);
            dt.TableName = "t_team_base";
             ds.Tables.Add(dt);

            string t_team_person = "select * from t_team_person";
            dt = MySqlHelp.Instance.ExecuteDataTable(t_team_person, null);
            dt.TableName = "t_team_person";
             ds.Tables.Add(dt);



             Tuple<bool, string> tup = NPOIEx.NpoiHelper.DataSet2Xls(@"D:/zc.xls", ds);
            //long id = 0;
            //string sql = "SELECT * from reporttest INTO OUTFILE 'E:/test.xls'";
            //MySqlHelp.Instance.ExecuteNonquery(sql,ref id);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            configM = new TLConfiguration.ConfigurationMan(AppDomain.CurrentDomain.BaseDirectory + "DbConfig.xml");
            CSL.Report2Xls.Class.MySqlHelp.connStr =configM.Read<string>("/Configuration/DbConStr", string.Empty);
        }


        
    }
}
