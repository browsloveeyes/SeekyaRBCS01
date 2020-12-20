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
using System.Data;
using System.Data.OleDb;

namespace Seekya
{
    /// <summary>
    /// pwdModifyForm2.xaml 的交互逻辑
    /// </summary>
    public partial class pwdModifyForm2 : Window
    {
        private string pwd;

        public pwdModifyForm2()
        {
            InitializeComponent();
        }
        public pwdModifyForm2(string passwd)
        {
            InitializeComponent();
            pwd = passwd;
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            string password;

            MD5_16 myEncryption = new MD5_16();
            password = myEncryption.MD5Encrypt16(pwd);

            if (textBox1.Text == pwd)
            {
                string strSqlDelete = "Delete from 1 where 编号='" + 1 + "'";
                string strSqlInsert = "Insert into 1 (编号,密码) values ('"+"1"+"','"+password+"')";
                //MessageBox.Show("修改密码成功！！");

                try
                {
                    aConnection.Open();
                    OleDbCommand myCmd = new OleDbCommand(strSqlDelete,aConnection);
                    myCmd.ExecuteNonQuery();
                    OleDbCommand myCmd1 = new OleDbCommand(strSqlInsert, aConnection);
                    myCmd1.ExecuteNonQuery();

                    MessageBox.Show("修改密码成功！！");
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERROR22:" + ex.Message);
                }
                finally
                {
                    if (aConnection != null)
                        aConnection.Close();
                    this.DialogResult = true;
                }


            }
            else
            {
                MessageBox.Show("两次输入密码不一致，修改密码失败");
                this.DialogResult = true;

            }
          
        }
    }
}
