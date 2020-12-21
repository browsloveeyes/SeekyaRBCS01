using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;
using System.Data.OleDb;
using System.Data;

namespace Seekya
{
    /// <summary>
    /// pwdInsertForm.xaml 的交互逻辑
    /// </summary>
    public partial class pwdInsertForm : Window
    {
        private dbManager f2 = null;

        public pwdInsertForm()
        {
            InitializeComponent();
        }
        public pwdInsertForm(dbManager f1)
        {
            InitializeComponent();
            f2 = f1;

        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+System.AppDomain.CurrentDomain.BaseDirectory+"Data\\checkDb.mdb");
            string strSql = "Select * from 1".ToString();
            string passwd = tBoxPwd.Text;
            DataSet ds = new DataSet();
            MD5_16 myEncryption = new MD5_16();

            passwd = myEncryption.MD5Encrypt16(passwd);       

            try
            {
                
                aConnection.Open();        
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                adapter.SelectCommand = new OleDbCommand(strSql, aConnection);
                
                adapter.Fill(ds);

                //MessageBox.Show(ds.Tables[0].Rows[0]["password"].ToString());

                if (String.Compare(passwd,ds.Tables[0].Rows[0]["密码"].ToString()) == 0)//密码匹配
                {
                    f2.EnableAdmin();

                }
                else
                {
                    f2.DeleteYes();

                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR21:" + ex.Message);

            }
            finally
            {
                
                if (aConnection != null)
                    aConnection.Close();
                this.DialogResult = true;


            }
          
          

        }
    }
}
