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
using System.Data.OleDb;

namespace Seekya
{
    /// <summary>
    /// dbInsertForm.xaml 的交互逻辑
    /// </summary>
    public partial class dbInsertForm : Window
    {
        private dbManager f2 = null; //为了调用dbManager中的成员和方法

        public dbInsertForm()
        {
            InitializeComponent();
        }
        public dbInsertForm(dbManager f1)
        {
            InitializeComponent();
            f2 = f1;
        }

        private void OKInsert_Click(object sender, RoutedEventArgs e)
        {
             MessageBoxResult dr = MessageBox.Show("确认插入记录吗？", "提示", MessageBoxButton.OKCancel);

             if (dr == MessageBoxResult.OK)
             {
                 OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
                 string cell = f2.GetTableName();
                 string date = "";

                 date += cell.Substring(0, 4) + "/";
                 date += cell.Substring(4, 2) + "/";
                 date += cell.Substring(6, 2);

                 try
                 {
                     if (aConnection.State.ToString() == "Closed")//如果没连接数据库，连接
                     {
                         aConnection.Open();
                     }

                     //允许用户不输入部分数据
                     tBox1.Text = (tBox1.Text == "") ? " " : tBox1.Text;
                     tBox2.Text = (tBox2.Text == "") ? " " : tBox2.Text;
                     tBox3.Text = (tBox3.Text == "") ? " " : tBox3.Text;
                     tBox4.Text = (tBox4.Text == "") ? " " : tBox4.Text;
                     tBox5.Text = (tBox5.Text == "") ? " " : tBox5.Text;
                     tBox6.Text = (tBox6.Text == "") ? " " : tBox6.Text;
                     tBox7.Text = (tBox7.Text == "") ? " " : tBox7.Text;
                     tBox8.Text = (tBox8.Text == "") ? " " : tBox8.Text;
                     tBox9.Text = (tBox9.Text == "") ? " " : tBox9.Text;
                     tBox10.Text = (tBox10.Text == "") ? " " : tBox10.Text;
                     tBox11.Text = (tBox11.Text == "") ? " " : tBox11.Text;
                     tBox12.Text = (tBox12.Text == "") ? " " : tBox12.Text;
                     tBox13.Text = (tBox13.Text == "") ? " " : tBox13.Text;
                     tBox14.Text = (tBox14.Text == "") ? " " : tBox14.Text;
                     //复核医生和报告医生
                     tBox16.Text = (tBox16.Text == "") ? " " : tBox16.Text;
                     tBox17.Text = (tBox17.Text == "") ? " " : tBox17.Text;
                     //备注1和备注2
                     tBox18.Text = (tBox18.Text == "") ? " " : tBox18.Text;
                     tBox19.Text = (tBox19.Text == "") ? " " : tBox19.Text;
                     tBox20.Text = (tBox19.Text == "") ? " " : tBox20.Text;

                    string insertSql = "Insert into " + cell + " (医院名称,科室名称,仪器型号,姓名,性别,年龄,住院号,CO,CO2,红细胞寿命,血红蛋白浓度,送检医生,复核医生,报告医生,初步诊断,样本类型,时间,日期,备注1,备注2) values ('" + tBox1.Text + "','" + tBox2.Text + "','" + tBox3.Text + "','" + tBox4.Text + "','" + tBox5.Text + "','" + tBox6.Text + "','" + tBox7.Text + "','" + tBox8.Text + "','" + tBox9.Text + "','" + tBox10.Text + "','" + tBox11.Text + "','" + tBox12.Text + "','" + tBox16.Text + "','" + tBox17.Text + "','" + tBox14.Text + "','" +tBox20.Text+"','" + tBox13.Text + "','" + date + "','" + tBox18.Text + "','" + tBox19.Text + "')";
                     OleDbCommand myCmd = new OleDbCommand(insertSql, aConnection);
                     myCmd.ExecuteNonQuery();
                     MessageBox.Show("插入成功");

                 }
                 catch (Exception ex)
                 {
                     MessageBox.Show("ERROR14:" + ex.Message);

                 }
                 finally
                 {
                     if (aConnection != null)
                         aConnection.Close();

                     //插入数据后，刷新DataGridView控件数据的操作，调用了Form1的函数
                     f2.DataGridViewTableDisplay(cell);

                     //关闭插入数据窗口
                     this.DialogResult = true;

                 }
             
             }
             else if (dr == MessageBoxResult.Cancel)
             { 
                //用户选择取消操作
             }
           
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }
    }
}
