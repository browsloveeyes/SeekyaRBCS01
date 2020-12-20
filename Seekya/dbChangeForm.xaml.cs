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

namespace Seekya
{
    /// <summary>
    /// dbChangeForm.xaml 的交互逻辑
    /// </summary>
    public partial class dbChangeForm : Window
    {

        private dbManager f1 = null;

        public dbChangeForm()
        {
            InitializeComponent();
        }
        public dbChangeForm(dbManager f2)
        {
            InitializeComponent();
            f1 = f2;
        }

        private void OKChange_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult dr = MessageBox.Show("确认更改记录吗？", "提示", MessageBoxButton.OKCancel);

            if (dr == MessageBoxResult.OK)
            {
                DbOperate test = new DbOperate();
                string time = f1.GetTime();
                string date = f1.GetDate();
                bool success;
                string date1 = "";

                int row = f1.dataGridViewTable.CurrentCell.RowIndex;

                date1 += date.Substring(0, 4) + "/";
                date1 += date.Substring(4, 2) + "/";
                date1 += date.Substring(6, 2);

                //允许用户不输入部分数据
                tBox1.Text = (tBox1.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[0].Value.ToString() : tBox1.Text;
                tBox2.Text = (tBox2.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[1].Value.ToString() : tBox2.Text;
                tBox3.Text = (tBox3.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[2].Value.ToString() : tBox3.Text;
                tBox4.Text = (tBox4.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[3].Value.ToString() : tBox4.Text;
                tBox5.Text = (tBox5.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[4].Value.ToString() : tBox5.Text;
                tBox6.Text = (tBox6.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[5].Value.ToString() : tBox6.Text;
                tBox7.Text = (tBox7.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[6].Value.ToString() : tBox7.Text;
                tBox8.Text = (tBox8.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[7].Value.ToString() : tBox8.Text;
                tBox9.Text = (tBox9.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[8].Value.ToString() : tBox9.Text;
                tBox10.Text = (tBox10.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[9].Value.ToString() : tBox10.Text;
                tBox11.Text = (tBox11.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[10].Value.ToString() : tBox11.Text;
                tBox12.Text = (tBox12.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[11].Value.ToString() : tBox12.Text;
                tBox13.Text = (tBox13.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[16].Value.ToString() : tBox13.Text;
                tBox14.Text = (tBox14.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[14].Value.ToString() : tBox14.Text;
                //复核医生和报告医生
                tBox16.Text = (tBox16.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[12].Value.ToString() : tBox16.Text;
                tBox17.Text = (tBox17.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[13].Value.ToString() : tBox17.Text;
                //备注1和备注2
                tBox18.Text = (tBox18.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[18].Value.ToString() : tBox18.Text;
                tBox19.Text = (tBox19.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[19].Value.ToString() : tBox19.Text;
                tBox20.Text= (tBox20.Text == "") ? f1.dataGridViewTable.Rows[row].Cells[15].Value.ToString() : tBox20.Text;
                success = test.ModifyRecord(date, time, tBox1.Text, tBox2.Text, tBox3.Text, tBox4.Text, tBox5.Text, tBox6.Text, tBox7.Text, tBox8.Text, tBox9.Text, tBox10.Text, tBox11.Text, tBox12.Text, tBox16.Text, tBox17.Text, tBox14.Text,tBox20.Text, tBox13.Text, date1, tBox18.Text, tBox19.Text);

                if (success == true)
                    MessageBox.Show("更改记录成功！！");
                else
                    MessageBox.Show("更改记录失败！！");

                f1.DataGridViewTableDisplay(date);

                this.DialogResult = true;
            }
            else if (dr == MessageBoxResult.Cancel)
            {
                //用户选择取消的操作
            }

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            int row = f1.dataGridViewTable.CurrentCell.RowIndex;

            //允许用户不输入部分数据
            tBox1.Text = f1.dataGridViewTable.Rows[row].Cells[0].Value.ToString();
            tBox2.Text = f1.dataGridViewTable.Rows[row].Cells[1].Value.ToString();
            tBox3.Text = f1.dataGridViewTable.Rows[row].Cells[2].Value.ToString();
            tBox4.Text = f1.dataGridViewTable.Rows[row].Cells[3].Value.ToString();
            tBox5.Text = f1.dataGridViewTable.Rows[row].Cells[4].Value.ToString();
            tBox6.Text = f1.dataGridViewTable.Rows[row].Cells[5].Value.ToString();
            tBox7.Text = f1.dataGridViewTable.Rows[row].Cells[6].Value.ToString();
            tBox8.Text = f1.dataGridViewTable.Rows[row].Cells[7].Value.ToString();
            tBox9.Text = f1.dataGridViewTable.Rows[row].Cells[8].Value.ToString();
            tBox10.Text = f1.dataGridViewTable.Rows[row].Cells[9].Value.ToString();
            tBox11.Text = f1.dataGridViewTable.Rows[row].Cells[10].Value.ToString();
            tBox12.Text = f1.dataGridViewTable.Rows[row].Cells[11].Value.ToString();
            tBox13.Text = f1.dataGridViewTable.Rows[row].Cells[16].Value.ToString();
            tBox14.Text = f1.dataGridViewTable.Rows[row].Cells[14].Value.ToString();
            //复核医生和报告医生
            tBox16.Text = f1.dataGridViewTable.Rows[row].Cells[12].Value.ToString();
            tBox17.Text = f1.dataGridViewTable.Rows[row].Cells[13].Value.ToString();
            //备注1和备注2
            tBox18.Text = f1.dataGridViewTable.Rows[row].Cells[18].Value.ToString();
            tBox19.Text = f1.dataGridViewTable.Rows[row].Cells[19].Value.ToString();
            tBox20.Text = f1.dataGridViewTable.Rows[row].Cells[15].Value.ToString();
        }
        
    }
}
