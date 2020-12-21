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
using System.Threading;
using System.IO;
//DataSet
using System.Data;
//ArrayList
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;

namespace Seekya
{
    /// <summary>
    /// dbManager.xaml 的交互逻辑
    /// </summary>
    public partial class dbManager : Window
    {
        MainWindow f1 = null;

        public Excel.Application app;
        public Excel.Workbooks wbs;
        public Excel.Workbook wb;

        //存储血红蛋白浓度
        public string rbcon1 = null;

        //判断office是否可用
        public bool officeavailable = true;  //与mainwindow中一样

        public dbManager(MainWindow f)
        {
            InitializeComponent();

            f1 = f;

            //如果当天表不存在，则创建
            DateTime dt = System.DateTime.Now;
            string date = dt.ToString("yyyyMMdd");
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            string strSql = "Select * from " + date;
            string patientPathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\patientInfo.txt";
            string[] item=new string[6];

            try//判断表是否存在，程序不够严谨（只要判断打开数据库表时出现错误，就归结于表不存在，以后改进）!!
            {
                aConnection.Open();
                OleDbCommand myCmd = new OleDbCommand(strSql, aConnection);
                myCmd.ExecuteNonQuery();

            }
            catch (Exception ex)//表不存在，创建表
            {
                try
                {
                    int i;
                    FileStream fs1 = new FileStream(patientPathString, FileMode.Open, FileAccess.Read);
                    StreamReader sr1 = new StreamReader(fs1);

                    for (i = 1; i < 21; i++)//读取txt文件到21行
                    {
                        sr1.ReadLine();
                    }
                    for (; i < 32; i=i+2)
                    {
                        item[(i - 21) / 2] = sr1.ReadLine();
                        sr1.ReadLine();
                            
                    }
                    
                    sr1.Close();
                    fs1.Close();

                }
                catch (Exception e)
                {
                    //System.Windows.MessageBox.Show("Error:" + e.Message);
                }
                
                ArrayList headList = new ArrayList();
                DbOperate testDb = new DbOperate();

                headList.Add("医院名称"); headList.Add("科室名称"); headList.Add("仪器型号");
                headList.Add("姓名"); headList.Add("性别"); headList.Add("年龄"); headList.Add("住院号");
                headList.Add("CO"); headList.Add("CO2"); headList.Add("红细胞寿命"); headList.Add("血红蛋白浓度");
                headList.Add("送检医生"); headList.Add("复核医生"); headList.Add("报告医生");
                headList.Add("初步诊断");headList.Add("样本类型");
                headList.Add("时间"); headList.Add("日期"); headList.Add("备注1"); headList.Add("备注2");

                for (int i = 0; i < 6; i++)
                {
                    if (item[i] != "null")
                        headList.Add(item[i]);
                }

                testDb.CreateTable(System.AppDomain.CurrentDomain.BaseDirectory+"Data\\checkDb.mdb", date, headList);

            }
            finally
            {
                if (aConnection != null)
                    aConnection.Close();

            }
           
 
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Thread td = new Thread(new ThreadStart(LoadData));

            td.IsBackground = true;
            td.Start();

        }
        private void LoadData()
        {
            Thread.Sleep(10);

            this.Dispatcher.Invoke(new Action(() =>
            {
                DateTime dt = System.DateTime.Now;
                string date = dt.ToString("yyyyMMdd");

                DataGridViewTablesListDisplay();
                DataGridViewTableDisplay(date);

                //屏蔽管理员操作按钮
                DisableAdmin();
            }));
        
        }
        private void dbManager_Load_1(object sender, EventArgs e)
        {
            DateTime dt = System.DateTime.Now;
            string date = dt.ToString("yyyyMMdd");

            DataGridViewTablesListDisplay();
            DataGridViewTableDisplay(date);

            //屏蔽管理员操作按钮
            DisableAdmin();

        }
        //使只允许管理员使用的数据库操作按钮无效
        public void DisableAdmin()
        {
            btnModifyPwd.IsEnabled = false;
            btnDeleteTable.IsEnabled = false;
            btnDeleteRecord.IsEnabled = false;
            btnChangeRecord.IsEnabled = false;
            btnInsertRecord.IsEnabled = false;

        }
        //使能只允许管理员使用的数据库操作按钮
        public void EnableAdmin()
        {
            btnModifyPwd.IsEnabled = true;
            btnDeleteTable.IsEnabled = true;
            btnDeleteRecord.IsEnabled = true;
            btnChangeRecord.IsEnabled = true;
            btnInsertRecord.IsEnabled = true;

        }
        //把管理员的复选框的勾去掉
        public void DeleteYes()
        {
            chBoxVerifyAdmin.IsChecked = false;//取消勾选

        }
        //把数据库的表显示在DataGridViewTable控件上
        public void DataGridViewTableDisplay(string tableName)
        {
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            //MessageBox.Show("Select * from " + tableName);
            string querySql = ("Select * from " + tableName).ToString();

            try
            {
                aConnection.Open();
                OleDbDataAdapter dadapter = new OleDbDataAdapter();
                dadapter.SelectCommand = new OleDbCommand(querySql, aConnection);
                DataSet dSet = new DataSet();

                dadapter.Fill(dSet);

                //为使dataGridView容器，当行数不足以填满容器时，进行补行操作
                if (dSet.Tables[0].Rows.Count < 18)
                {
                    // MessageBox.Show("表中数据的行数为：" + dSet.Tables[0].Rows.Count);
                    int j = dSet.Tables[0].Rows.Count;

                    for (int i = 0; i < (18 - j); i++)
                    {
                        DataRow dr = dSet.Tables[0].NewRow();
                        //for (int x = 0; x < 13; x++)
                        //{
                            //dr[x] = "";//新行的单元格装入空值
                        //}
                        dSet.Tables[0].Rows.Add(dr);

                    }

                }
                dataGridViewTable.DataSource = dSet.Tables[0];

                //for (int i = 0; i < 13; i++)    //解除表头（每列头字段）的点中以及排序模式
                    //dataGridViewTable.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error:" + ex.Message);

            }
            finally
            {
                if (aConnection != null)
                {
                    aConnection.Close();

                }

            }

        }
        //把数据库的表显示在DataGridViewTablesListDisplay控件上
        public void DataGridViewTablesListDisplay()
        {
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            try
            {
                aConnection.Open();
                DataSet dSet1 = new DataSet();
                DataTable shemaTable = aConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                DataTable tablesList = new DataTable();
                DataColumn tablesName = new DataColumn("表名");

                int tableCount = shemaTable.Rows.Count - 1;
                string[] strTmp=new string[tableCount];
                int c = 0;

                tablesList.Columns.Add(tablesName);

                foreach (DataRow dr in shemaTable.Rows)
                {
                    //DataRow row = tablesList.NewRow();
                    //row[0] = dr["TABLE_NAME"].ToString();

                    if (String.Compare(dr["TABLE_NAME"].ToString(), "1") != 0)
                    {
                        strTmp[c++] = dr["TABLE_NAME"].ToString();

                    }
                        //tablesList.Rows.Add(row);
                    
                }
                for (; c > 0; c--)
                {
                    DataRow row = tablesList.NewRow();
                    row[0] = strTmp[c-1];

                    tablesList.Rows.Add(row);
                
                }


                dSet1.Tables.Add(tablesList);

                //为使dataGridView容器，当行数不足以填满容器时，进行补行操作
                if (dSet1.Tables[0].Rows.Count < 18)
                {
                    // MessageBox.Show("表中数据的行数为：" + dSet.Tables[0].Rows.Count);
                    int j = dSet1.Tables[0].Rows.Count;

                    for (int i = 0; i < (18 - j); i++)
                    {
                        DataRow dr = dSet1.Tables[0].NewRow();
                        for (int x = 0; x < 1; x++)
                        {
                            dr[x] = "";//新行的单元格装入空值
                        }
                        dSet1.Tables[0].Rows.Add(dr);

                    }

                }
                dataGridViewTablesList.DataSource = dSet1.Tables[0];

                //for (int i = 0; i < 1; i++)    //解除表头（每列头字段）的点中以及排序模式
                    //dataGridViewTablesList.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error in handling:" + ex.Message);

            }
            finally
            {
                if (aConnection != null)
                {
                    aConnection.Close();

                }

            }

        }
        //返回dataGridView控件中当前选中行的“时间”列的值
        public string GetTime()
        {
            return dataGridViewTable.Rows[dataGridViewTable.CurrentCell.RowIndex].Cells[15].Value.ToString();

        }
        //返回dataGridView控件中当前选中行的“日期”列的值
        public string GetDate()
        {
            return DateToString(dataGridViewTable.Rows[dataGridViewTable.CurrentCell.RowIndex].Cells[16].Value.ToString());

        }
        //获取当前显示的表名
        public string GetTableName()
        {
            string cell = dataGridViewTablesList.CurrentCell.Value.ToString();

            return cell;

        }
        //把“20170613”转为“2017/06/13”
        private string StringToDate(string str)
        {
            string date="";

            date += str.Substring(0, 4)+"/";
            date += str.Substring(4, 2)+"/";
            date += str.Substring(6, 2);

            return date;
        
        }
        //退出操作
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = true;//关闭当前窗口
        }
        //插入记录操作
        private void btnInsertRecord_Click(object sender, EventArgs e)
        {
            string cell = dataGridViewTablesList.CurrentCell.Value.ToString();

            if (String.Compare(cell, "") != 0)//选中的表存在
            {
                //dbInsertForm insertRecord = new dbInsertForm(this);
                //insertRecord.ShowDialog();
                //insertRecord.ShowDialog(this);//"this"必不可少（将窗口显示为具有指定拥有者：insertRecord的所有者为Form1类的当前对象），目的为了insertRecord可调用
            }
        }
        //修改记录操作
        private void btnChangeRecord_Click(object sender, EventArgs e)
        {
            //dbChangeForm changeRecord = new dbChangeForm(this);
            //changeRecord.ShowDialog();
        }

        //管理员复选框勾选事件
        private void chBoxVerifyAdmin_MouseClick(object sender, MouseEventArgs e)
        {
            //if (chBoxVerifyAdmin.Checked)//管理员的可选框被勾上
            //{
                //pwdInsertForm f1 = new pwdInsertForm(this);//输入密码对话框
                //f1.ShowDialog();

           // }
            //else
            //{
               // DisableAdmin();

            //}

        }
        //修改管理员密码操作
        private void btnModifyPwd_Click(object sender, EventArgs e)
        {
           // pwdModifyForm1 passwdModify = new pwdModifyForm1();
           // passwdModify.ShowDialog();
        }

        //表名显示列表的点击事件，为实现右边表记录更新显示
       /* private void dataGridViewTablesList_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string cell = dataGridViewTablesList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                //MessageBox.Show(cell);
                if (String.Compare(cell, "") != 0)//可用的表被点中
                {
                    DataGridViewTableDisplay(cell);
                }
            }
            catch (Exception ex)
            {
                return;

            }


        }*/
        //删除记录操作
        private void btnDeleteRecord_Click(object sender, EventArgs e)
        {
            DbOperate test = new DbOperate();
            string time = dataGridViewTable.Rows[dataGridViewTable.CurrentCell.RowIndex].Cells[11].Value.ToString();
            string date = dataGridViewTable.Rows[dataGridViewTable.CurrentCell.RowIndex].Cells[12].Value.ToString();
            bool success;

            success = test.DeleteRecord(date, time);

            if (success == true)
                MessageBox.Show("删除记录成功！！");
            else
                MessageBox.Show("删除记录失败！！");


            DataGridViewTableDisplay(date);


        }

        private void btnDeleteTable_Click(object sender, EventArgs e)
        {
            string frontTable;
            DbOperate test = new DbOperate();
            string tableName = dataGridViewTablesList.CurrentCell.Value.ToString();

            test.DeleteTable(tableName);
            DataGridViewTablesListDisplay();
            frontTable = dataGridViewTablesList.CurrentCell.Value.ToString();
            DataGridViewTableDisplay(frontTable);

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            DataSet dSet = new DataSet();
            string name = tBoxName.Text;
            string num = tBoxNumber.Text;

            try
            {
                string strSql;

                aConnection.Open();
                DataTable shemaTable = aConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });//读取数据库的表名

                foreach (DataRow dr in shemaTable.Rows)
                {
                    if ((name != "" && num != "") || (name == "" && num == ""))//姓名和住院号都有输入，或者姓名和住院号都没有输入
                        strSql = "select * from " + dr["TABLE_NAME"].ToString() + " where 姓名=" + "\'" + name + "\'" + " and 住院号=" + "\'" + num + "\'";
                    else if (name != "" && num == "")//只有姓名被输入
                        strSql = "select * from " + dr["TABLE_NAME"].ToString() + " where 姓名=" + "\'" + name + "\'";
                    else//只有住院号被输入
                        strSql = "select * from " + dr["TABLE_NAME"].ToString() + " where 住院号=" + "\'" + num + "\'";


                    OleDbDataAdapter dadapter = new OleDbDataAdapter();
                    dadapter.SelectCommand = new OleDbCommand(strSql, aConnection);
                    //dadapter.SelectCommand = new OleDbCommand(strSql1, aConnection);


                    dadapter.Fill(dSet);
                }

                //为使dataGridView容器，当行数不足以填满容器时，进行补行操作
                if (dSet.Tables[0].Rows.Count < 18)
                {
                    // MessageBox.Show("表中数据的行数为：" + dSet.Tables[0].Rows.Count);
                    int j = dSet.Tables[0].Rows.Count;

                    for (int i = 0; i < (18 - j); i++)
                    {
                        DataRow dr = dSet.Tables[0].NewRow();
                        /*for (int x = 0; x < 13; x++)
                        {
                            dr[x] = "";//新行的单元格装入空值
                        }*/
                        dSet.Tables[0].Rows.Add(dr);

                    }

                }
                dataGridViewTable.DataSource = dSet.Tables[0];

                //for (int i = 0; i < 13; i++)    //解除表头（每列头字段）的点中以及排序模式
                    //dataGridViewTable.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error:" + ex.Message);

            }
            finally
            {
                if (aConnection != null)
                {
                    aConnection.Close();

                }

            }
        }
        //打印检验报告
        private void btnPrint_Click(object sender, EventArgs e)
        {
            //PrintReport print = new PrintReport();
            int row = dataGridViewTable.CurrentCell.RowIndex;
            //string hospitalName = dataGridViewTable.Rows[row].Cells[0].ToString();
            string dept = dataGridViewTable.Rows[row].Cells[1].Value.ToString();
            string device = dataGridViewTable.Rows[row].Cells[3].Value.ToString();
            string name = dataGridViewTable.Rows[row].Cells[4].Value.ToString();
            string sex = dataGridViewTable.Rows[row].Cells[5].Value.ToString();
            string age = dataGridViewTable.Rows[row].Cells[6].Value.ToString();
            string id = dataGridViewTable.Rows[row].Cells[7].Value.ToString();
            string PCO = dataGridViewTable.Rows[row].Cells[8].Value.ToString();
            string CO2L = dataGridViewTable.Rows[row].Cells[9].Value.ToString();
            string RBC = dataGridViewTable.Rows[row].Cells[10].Value.ToString();
            string time = dataGridViewTable.Rows[row].Cells[11].Value.ToString();
            string date = dataGridViewTable.Rows[row].Cells[12].Value.ToString();
            MessageBox.Show(dept + device + name + sex + age + id + PCO + CO2L + RBC + time + date);

           // print.ReportPrint(dept, device, name, sex, age, id, PCO, CO2L, RBC, time, date);
        }
        //查找记录
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            DataSet dSet = new DataSet();
            string name = tBoxName.Text;
            string num = tBoxNumber.Text;

            try
            {
                string strSql;

                aConnection.Open();
                DataTable shemaTable = aConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });//读取数据库的表名

                foreach (DataRow dr in shemaTable.Rows)
                {
                    if (String.Compare(dr["TABLE_NAME"].ToString(), "1") != 0)
                    {

                        if ((name != "" && num != "") || (name == "" && num == ""))//姓名和住院号都有输入，或者姓名和编号都没有输入
                            strSql = "select * from " + dr["TABLE_NAME"].ToString() + " where 姓名=" + "\'" + name + "\'" + " and 住院号=" + "\'" + num + "\'";
                        else if (name != "" && num == "")//只有姓名被输入
                            strSql = "select * from " + dr["TABLE_NAME"].ToString() + " where 姓名=" + "\'" + name + "\'";
                        else//只有住院号被输入
                            strSql = "select * from " + dr["TABLE_NAME"].ToString() + " where 住院号=" + "\'" + num + "\'";

                        OleDbDataAdapter dadapter = new OleDbDataAdapter();
                        dadapter.SelectCommand = new OleDbCommand(strSql, aConnection);
                        //dadapter.SelectCommand = new OleDbCommand(strSql1, aConnection);

                        dadapter.Fill(dSet);
                    }
                }

                //为使dataGridView容器，当行数不足以填满容器时，进行补行操作
                if (dSet.Tables[0].Rows.Count < 18)
                {
                    // MessageBox.Show("表中数据的行数为：" + dSet.Tables[0].Rows.Count);
                    int j = dSet.Tables[0].Rows.Count;

                    for (int i = 0; i < (18 - j); i++)
                    {
                        DataRow dr = dSet.Tables[0].NewRow();
                       /* for (int x = 0; x < 13; x++)
                        {
                            dr[x] = "";//新行的单元格装入空值
                        }*/
                        dSet.Tables[0].Rows.Add(dr);

                    }

                }

                dataGridViewTable.DataSource = dSet.Tables[0];

                //for (int i = 0; i < 13; i++)    //解除表头（每列头字段）的点中以及排序模式
                    //dataGridViewTable.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error:" + ex.Message);

            }
            finally
            {
                if (aConnection != null)
                {
                    aConnection.Close();

                }

            }
        }
        
        //管理员复选框勾选事件
        private void chBoxVerifyAdmin_Click(object sender, RoutedEventArgs e)
        {
            if (chBoxVerifyAdmin.IsChecked == true)//管理员的可选框被勾上
            {
                pwdInsertForm f1 = new pwdInsertForm(this);//输入密码对话框
                f1.ShowDialog();

            }
            else
            {
                DisableAdmin();

            }
        }

        private void btnDeleteTable_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult dr = MessageBox.Show("确定删除表吗？","提示",MessageBoxButton.OKCancel);

            if(dr==MessageBoxResult.OK)
            {
                string frontTable;
                DbOperate test = new DbOperate();
                string tableName = dataGridViewTablesList.CurrentCell.Value.ToString();

                test.DeleteTable(tableName);
                DataGridViewTablesListDisplay();

                frontTable = dataGridViewTablesList.CurrentCell.Value.ToString();
                DataGridViewTableDisplay(frontTable);
            }
            else if(dr==MessageBoxResult.Cancel)
            {
                //用户选择取消的操作
            }

        }

        private void btnDeleteRecord_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult dr = MessageBox.Show("确认删除记录吗？", "提示", MessageBoxButton.OKCancel);

            if (dr == MessageBoxResult.OK)
            {
                //用户选择确认的操作
                DbOperate test = new DbOperate();
                string time = dataGridViewTable.Rows[dataGridViewTable.CurrentCell.RowIndex].Cells[15].Value.ToString();
                string date = DateToString(dataGridViewTable.Rows[dataGridViewTable.CurrentCell.RowIndex].Cells[16].Value.ToString());
                bool success;

                success = test.DeleteRecord(date, time);

                if (success == true)
                    MessageBox.Show("删除记录成功");
                else
                    MessageBox.Show("删除记录失败");

                DataGridViewTableDisplay(date);
            }
            else if (dr == MessageBoxResult.Cancel)
            {
                //用户选择取消的操作

            }

            //删除记录后更新主界面的数据库
            f1.UpdateTodayReport();
            
        }
        //把日期“2017/06/09”转为“20170609”
        private string DateToString(string date)
        {
            string str="";
            string[] s = date.Split('/');

            foreach (string t in s)
                str += t;

            return str;
        
        }
        //修改管理员密码操作
        private void btnModifyPwd_Click(object sender, RoutedEventArgs e)
        {
            pwdModifyForm1 passwdModify = new pwdModifyForm1();
            passwdModify.ShowDialog();
        }
        //插入记录操作
        private void btnInsertRecord_Click(object sender, RoutedEventArgs e)
        {         
            string cell = dataGridViewTablesList.CurrentCell.Value.ToString();

            if (String.Compare(cell, "") != 0)//选中的表存在
            {
                dbInsertForm insertRecord = new dbInsertForm(this);
                insertRecord.ShowDialog();
                //insertRecord.ShowDialog(this);//"this"必不可少（将窗口显示为具有指定拥有者：insertRecord的所有者为Form1类的当前对象），目的为了insertRecord可调用
            }     
           
        }

        private void btnChangeRecord_Click(object sender, RoutedEventArgs e)
        {
                dbChangeForm changeRecord = new dbChangeForm(this);
                changeRecord.ShowDialog();
            
        }

        private void dataGridViewTablesList_CellClick(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            try
            {
                string cell = dataGridViewTablesList.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                //MessageBox.Show(cell);
                if (String.Compare(cell, "") != 0)//可用的表被点中
                {
                    DataGridViewTableDisplay(cell);
                }
            }
            catch (Exception ex)
            {
                return;

            }
        }

        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            PrintReport print = new PrintReport();
            int row = dataGridViewTable.CurrentCell.RowIndex;
            string[] userDefine = { "", "", "", "", "", "" };
            int i;
            string date = GetDate();

            //先获取CO浓度和血红蛋白浓度
            string co = dataGridViewTable.Rows[row].Cells[7].Value.ToString().Trim();
            string hb = dataGridViewTable.Rows[row].Cells[10].Value.ToString().Trim();
            string rbc = dataGridViewTable.Rows[row].Cells[9].Value.ToString();
            bool sign = false;
            string hospital = dataGridViewTable.Rows[row].Cells[0].Value.ToString();
            string department = dataGridViewTable.Rows[row].Cells[1].Value.ToString();
            string instrumentType = dataGridViewTable.Rows[row].Cells[2].Value.ToString();
            string name = dataGridViewTable.Rows[row].Cells[3].Value.ToString();
            string gender = dataGridViewTable.Rows[row].Cells[4].Value.ToString();
            string age = dataGridViewTable.Rows[row].Cells[5].Value.ToString();
            string id = dataGridViewTable.Rows[row].Cells[6].Value.ToString();
            string co2 = dataGridViewTable.Rows[row].Cells[8].Value.ToString();
            string submitDoctor = dataGridViewTable.Rows[row].Cells[11].Value.ToString();
            string checkDoctor = dataGridViewTable.Rows[row].Cells[12].Value.ToString();
            string reportDoctor = dataGridViewTable.Rows[row].Cells[13].Value.ToString();
            string firstVisit = dataGridViewTable.Rows[row].Cells[14].Value.ToString();
            string ty = dataGridViewTable.Rows[row].Cells[15].Value.ToString();
            string reportTime = dataGridViewTable.Rows[row].Cells[16].Value.ToString();
            string testDateLine = dataGridViewTable.Rows[row].Cells[17].Value.ToString();
            string remark1 = dataGridViewTable.Rows[row].Cells[18].Value.ToString();
            string remark2 = dataGridViewTable.Rows[row].Cells[19].Value.ToString();

            try
            {
                for (i = 18; i < 24; i++)
                {
                    userDefine[i - 18] = dataGridViewTable.Rows[row].Cells[i].Value.ToString();

                }

            }
            catch { }

            //判断血红蛋白浓度是否有效
            if (int.Parse(hb) == 0)
            {
                hbInputDbManager t = new hbInputDbManager();

                t.Owner = this;

                t.ShowDialog();

                hb = rbcon1;

                //红细胞寿命换算
                rbc = ((int)(1.38 * int.Parse(hb) / float.Parse(co))).ToString();
                sign = true;

            }
            if (sign == true)
            {
                //修改数据库未存有血红蛋白浓度的检验数据
                DbOperate test = new DbOperate();
                //更改记录
                test.ModifyRecord(testDateLine.Substring(0, 4) + testDateLine.Substring(5, 2) + testDateLine.Substring(8, 2), reportTime, hospital, department, instrumentType, name, gender, age, id, co, co2, rbc, hb, submitDoctor, checkDoctor, reportDoctor, firstVisit, ty,reportTime, testDateLine, remark1, remark2);

                DataGridViewTableDisplay(date);
            }

            string date1 = testDateLine.Substring(0, 4) + testDateLine.Substring(5, 2) + testDateLine.Substring(8, 2);
            string datetime2 = reportTime.Substring(0, 2) + reportTime.Substring(3, 2) + reportTime.Substring(6, 2);

            //添加数据到excel表格中，并创建患者检测报告
            string str = null;//System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template.xls";
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\print.txt";

            //读打印模板名
            try
            {
                StreamReader sr = new StreamReader(pathString, Encoding.GetEncoding("gb2312"));

                sr.ReadLine();
                str = sr.ReadLine();

                sr.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }
            //try
            //{
            //    Open(str);
            //}
            //catch (Exception)
            //{
            //    officeavailable = false;
            //}
            string officepath = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\office.txt";

            try
            {
                StreamReader sroffice = new StreamReader(officepath, Encoding.GetEncoding("gb2312"));
                string TorF = sroffice.ReadLine();
                if (TorF == "True")
                {
                    officeavailable = true;
                }
                else
                {
                    officeavailable = false;
                }
                sroffice.Close();
            }
            catch (Exception)
            {

                throw;
            }
            try
            {
                StreamReader sroffice = new StreamReader(officepath, Encoding.GetEncoding("gb2312"));
                string TorF = sroffice.ReadLine();
                if (TorF == "True")
                {
                    officeavailable = true;
                }
                else
                {
                    officeavailable = false;
                }
                sroffice.Close();
            }
            catch (Exception)
            {

                throw;
            }

            if (officeavailable)
            {
                Open(str);
                Excel.Worksheet ws = (Excel.Worksheet)app.ActiveSheet;

                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("name", typeof(string));
                dataTable.Columns.Add("age", typeof(string));
                dataTable.Columns.Add("zyh", typeof(string));
                dataTable.Columns.Add("sex", typeof(string));
                dataTable.Columns.Add("yqxh", typeof(string));
                dataTable.Columns.Add("cbzd", typeof(string));
                dataTable.Columns.Add("sjys", typeof(string));
                dataTable.Columns.Add("hb", typeof(string));
                dataTable.Columns.Add("yymc", typeof(string));
                dataTable.Columns.Add("rbc", typeof(string));
                dataTable.Columns.Add("CO", typeof(string));
                dataTable.Columns.Add("eyht", typeof(string));
                dataTable.Columns.Add("jyrq", typeof(string));
                dataTable.Columns.Add("ksmc", typeof(string));
                dataTable.Columns.Add("dyyi", typeof(string));
                dataTable.Columns.Add("dyer", typeof(string));
                dataTable.Columns.Add("dysan", typeof(string));
                dataTable.Columns.Add("dysi", typeof(string));
                dataTable.Columns.Add("dywu", typeof(string));
                dataTable.Columns.Add("dyliu", typeof(string));
                dataTable.Columns.Add("fhys", typeof(string));
                dataTable.Columns.Add("bgys", typeof(string));
                dataTable.Columns.Add("bgsj", typeof(string));
                dataTable.Columns.Add("ldgd", typeof(string));
                dataTable.Columns.Add("eyhtgd", typeof(string));
                dataTable.Columns.Add("yblx", typeof(string));

                dataTable.Columns.Add("htime", typeof(string));
                dataTable.Columns.Add("bnum", typeof(string));
                dataTable.Columns.Add("advice", typeof(string));
                dataTable.Columns.Add("ptype", typeof(string));
                dataTable.Columns.Add("height", typeof(string));
                dataTable.Columns.Add("weight", typeof(string));
                dataTable.Columns.Add("nation", typeof(string));
                dataTable.Columns.Add("nplace", typeof(string));
                dataTable.Columns.Add("tel", typeof(string));
                dataTable.Columns.Add("address", typeof(string));
                dataTable.Columns.Add("pay", typeof(string));
                dataTable.Columns.Add("mstate", typeof(string));

                DataRow dr = dataTable.NewRow();
                dr["name"] = name;
                dr["age"] = age;
                dr["zyh"] = id;
                dr["sex"] = gender;
                dr["yqxh"] = instrumentType;
                dr["cbzd"] = firstVisit;
                dr["sjys"] = submitDoctor;
                dr["hb"] = hb;
                dr["yymc"] = hospital;
                dr["rbc"] = rbc;
                dr["CO"] = co;
                dr["eyht"] = co2;
                dr["jyrq"] = testDateLine;
                dr["ksmc"] = department;
                dr["dyyi"] = null;
                dr["dyer"] = null;
                dr["dysan"] = null;
                dr["dysi"] = null;
                dr["dywu"] = null;
                dr["dyliu"] = null;
                dr["fhys"] = checkDoctor;
                dr["bgys"] = reportDoctor;
                dr["bgsj"] = reportTime;
                dr["ldgd"] = remark1;
                dr["eyhtgd"] = remark2;
                dr["yblx"] = ty;
                //for (int h = 0; h < 12; h++)
                //{
                //    if (values[h] != null)
                //    {
                //        dr[propts[h]] = values[h];
                //    }
                //} //暂注释，意味者配置里自定义的属性在数据管理界面的打印属性之外，后续需改进
                dataTable.Rows.Add(dr);
                int nameCellCount = app.ActiveWorkbook.Names.Count - 1;//获得命名单元格的总数
                int[] nameCellRow = new int[nameCellCount];//某个命名单元格的行
                int[] nameCellColumn = new int[nameCellCount];//某个命名单元格的列
                string[] nameCellName = new string[nameCellCount];//某个命名单元格的自定义名称，比如 工资
                string strName;
                string tmp;
                int nameCellIdx = 0;
                for (int j = 0; j < nameCellCount + 1; j++)
                {
                    strName = app.ActiveWorkbook.Names.Item(j + 1).Name;
                    if (strName != "Sheet1!Print_Area")
                    {
                        app.Goto(strName);
                        nameCellColumn[nameCellIdx] = app.ActiveCell.Column;
                        nameCellRow[nameCellIdx] = app.ActiveCell.Row;
                        nameCellName[nameCellIdx] = strName;
                        nameCellIdx++;//真实的循环的命名单元格序号
                    }
                }
                for (int index = 0; index < nameCellCount; index++)
                {
                    tmp = dataTable.Rows[0][nameCellName[index]].ToString();
                    ws.Cells[nameCellRow[index], nameCellColumn[index]] = tmp;
                }
                app.Goto(ws.Range["A1"], true);

                //忘了为啥这里有个加m的代码
                //int nameCellCountm = app.ActiveWorkbook.Names.Count;//获得命名单元格的总数
                //int[] nameCellRowm = new int[nameCellCount];//某个命名单元格的行
                //int[] nameCellColumnm = new int[nameCellCount];//某个命名单元格的列
                //string[] nameCellNamem = new string[nameCellCount];//某个命名单元格的自定义名称，比如 工资
                //string strNamem;
                //string tmpm;
                //int nameCellIdxm = 0;
                //for (int j = 0; j < nameCellCount; j++)
                //{
                //    strNamem = app.ActiveWorkbook.Names.Item(j + 1).Name;
                //    app.Goto(strNamem);
                //    nameCellColumnm[nameCellIdxm] = app.ActiveCell.Column;
                //    nameCellRowm[nameCellIdxm] = app.ActiveCell.Row;
                //    nameCellNamem[nameCellIdxm] = strNamem;
                //    nameCellIdxm++;//真实的循环的命名单元格序号
                //}
                //for (int index = 0; index < nameCellCount; index++)
                //{
                //    tmpm = dataTable.Rows[0][nameCellNamem[index]].ToString();
                //    ws.Cells[nameCellRowm[index], nameCellColumnm[index]] = tmpm;
                //}

                //检索数据
                OleDbConnection Connec = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
                DataSet dset = new DataSet();
                try
                {
                    Connec.Open();
                    DataTable shemaTable = Connec.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });//读取数据库的表名
                    string strsql;
                    foreach (DataRow dtrw in shemaTable.Rows)
                    {
                        if (String.Compare(dtrw["TABLE_NAME"].ToString(), "1") != 0)
                        {
                            //strSql = "select * from " + dr["TABLE_NAME"].ToString() + " where 姓名=" + "\'" + name + "\'";
                            string xy = id;
                            if (id!=null&&id!=" ")
                            {
                                strsql = "select 红细胞寿命,日期,时间 from " + dtrw["TABLE_NAME"].ToString() + " where 住院号=" + "\'" + xy + "\'";
                                //strsql = "select 红细胞寿命,日期,时间 from " + dtrw["TABLE_NAME"].ToString() + " where 姓名='吕布'";
                                //tBoxName.Text = strsql;
                                OleDbDataAdapter dpter = new OleDbDataAdapter();
                                dpter.SelectCommand = new OleDbCommand(strsql, Connec);
                                //dadapter.SelectCommand = new OleDbCommand(strSql1, aConnection);
                                dpter.Fill(dset);
                            }

                        }
                    }
                }
                catch (Exception eex)
                {
                    System.Windows.MessageBox.Show("ERROR:" + eex.Message);
                }
                finally
                {
                    if (Connec != null)
                    {
                        Connec.Close();

                    }
                }
                int[] hxbsm = new int[20];
                int num = 0;
                if (id!=null&&id!=" ")
                {
                    num = dset.Tables[0].Rows.Count;
                }
                string[] time = new string[20];
                try
                {
                    for (int t = 0; t < num; t++)
                    {
                        if (dset.Tables[0].Rows[t]["红细胞寿命"].ToString() == ">250")
                        {
                            hxbsm[t] = 250;
                        }
                        else
                        {
                            hxbsm[t] = Convert.ToInt32(dset.Tables[0].Rows[t]["红细胞寿命"]);
                        }
                        time[t] = string.Concat(dset.Tables[0].Rows[t]["日期"].ToString(), dset.Tables[0].Rows[t]["时间"].ToString().Substring(0, 5));
                    }
                }
                catch (Exception et)
                {
                    System.Windows.MessageBox.Show("ERRORt:" + et.Message);
                }

                if (num > 1)
                {
                    for (int w = 0; w < num; w++)
                    {
                        ws.Cells[30 + w, 19] = time[w];
                        //ws.Cells[30 + w, 18] = w;
                        ws.Cells[30 + w, 20] = hxbsm[w];
                    }
                    Excel.Range oResizeRange;
                    Excel.Series oSeries;
                    if (ws.Shapes.Count < 1)
                    {
                        wb.Charts.Add(Type.Missing, Type.Missing, 1, Type.Missing);
                        wb.ActiveChart.ChartType = Excel.XlChartType.xlLine;//设置图形
                        wb.ActiveChart.SetSourceData(ws.get_Range("T30", "T" + (num + 29).ToString()), Excel.XlRowCol.xlColumns);
                        wb.ActiveChart.Location(Excel.XlChartLocation.xlLocationAsObject, ws.Name);
                        oResizeRange = (Excel.Range)ws.Rows.get_Item(24, Type.Missing);
                        ws.Shapes.Item("图表 1").Top = (float)(double)oResizeRange.Top;
                        oResizeRange = (Excel.Range)ws.Columns.get_Item(1, Type.Missing); //调图表的位置左边距
                        ws.Shapes.Item("图表 1").Left = (float)(double)oResizeRange.Left;
                        ws.Shapes.Item("图表 1").Width = 443;
                        ws.Shapes.Item("图表 1").Height = 200;
                        //wb.ActiveChart.PlotArea.Interior.ColorIndex = 19; //设置绘图区的背景色
                        //wb.ActiveChart.PlotArea.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;//设置绘图区边框线条
                        wb.ActiveChart.PlotArea.Width = 443; //设置绘图区宽度
                        wb.ActiveChart.HasLegend = false;
                        //设置Y轴的显示
                        Excel.Axis yAxis = (Excel.Axis)wb.ActiveChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                        yAxis.MajorGridlines.Border.LineStyle = Excel.XlLineStyle.xlDot;
                        yAxis.MajorGridlines.Border.ColorIndex = 1;//gridLine横向线条的颜色
                        yAxis.HasTitle = true;
                        //xAxis.MinimumScale = 1500;
                        //xAxis.MaximumScale = 6000;
                        yAxis.TickLabels.Font.Name = "宋体";
                        //yAxis.TickLabels.Font.Size = 9;
                        yAxis.AxisTitle.Text = "红细胞寿命/天";
                        //设置X轴的显示
                        Excel.Axis xAxis = (Excel.Axis)wb.ActiveChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                        xAxis.CategoryNames = ws.get_Range("S30", "S" + (num + 29).ToString());
                        xAxis.HasTitle = true;
                        xAxis.AxisTitle.Text = "时间";
                        xAxis.AxisTitle.Left = 480;
                        xAxis.TickLabels.Orientation = Excel.XlTickLabelOrientation.xlTickLabelOrientationHorizontal;//X轴显示的方向,是水平还是垂直等
                        xAxis.TickLabels.Font.Size = 6;
                        //以下是设置标题
                        wb.ActiveChart.HasTitle = true;
                        wb.ActiveChart.ChartTitle.Text = "红细胞寿命变化示意图";
                        //在图线上显示数据点
                        wb.ActiveChart.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowValue, false, true
                    , false, false, false, true, false, false, false);
                        //    wb.ActiveChart.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowPercent, false, false
                        //, false, false, false, true, true, false, false);
                    }
                }
                testDateLine = testDateLine.Substring(0, 4) + testDateLine.Substring(5, 2) + testDateLine.Substring(8, 2);
                reportTime = reportTime.Substring(0, 2) + reportTime.Substring(3, 2) + reportTime.Substring(6, 2);
                string excelName = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\Template\\" + name + "(" + testDateLine + reportTime + ")" + ".xlsx";
                try
                {

                    int postn = excelName.LastIndexOf(".");
                    int k = 1;
                    while (System.IO.File.Exists(excelName))
                    {
                        excelName = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\Template\\" + name + "(" + testDateLine + reportTime + ")" + ".xlsx";

                        excelName = excelName.Insert(postn, "(" + k + ")");
                        //excelName = string.Format(excelName + i);
                        k++;
                    }
                    wb.SaveAs(excelName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                }
                catch (Exception eee)
                {
                    System.Windows.MessageBox.Show("ERROR26m:" + eee.Message);
                }

                //DataTable dataTable = new DataTable();
                //dataTable.Columns.Add("hb", typeof(string));
                //dataTable.Columns.Add("rbc", typeof(string));
                //DataRow dr = dataTable.NewRow();
                //dr["hb"] = hb;
                //dr["rbc"] = rbc;
                //dataTable.Rows.Add(dr);
                //app.Goto("rbc");
                //app.ActiveCell.FormulaR1C1 = rbc;
                //app.Goto("hb");
                //app.ActiveCell.FormulaR1C1 = hb;
                //wb.SaveCopyAs(excelName);

                //wb.Close(Type.Missing, Type.Missing, Type.Missing);      //(W)注意这里有注释掉
                //wbs.Close();                                             //(W)注意这里有注释掉
                //app.Quit();
                wb = null;
                wbs = null;
                //app = null;
                GC.Collect();
                PublicMethod.Kill(app);

                print.ReportPrintHand(name, gender, age, id, instrumentType, submitDoctor, firstVisit, hb, hospital, rbc, co, co2, testDateLine, department, userDefine[0], userDefine[1], userDefine[2], userDefine[3], userDefine[4], userDefine[5], checkDoctor, reportDoctor, reportTime, remark1, remark2, ty,excelName);
            }
            else
            {
                //检索数据
                OleDbConnection Connec = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
                DataSet dset = new DataSet();
                try
                {
                    Connec.Open();
                    DataTable shemaTable = Connec.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });//读取数据库的表名
                    string strsql;
                    foreach (DataRow dtrw in shemaTable.Rows)
                    {
                        if (String.Compare(dtrw["TABLE_NAME"].ToString(), "1") != 0)
                        {
                            if (id != null && id != " ")
                            {
                                //strSql = "select * from " + dr["TABLE_NAME"].ToString() + " where 姓名=" + "\'" + name + "\'";
                                string xy = id;
                                strsql = "select 红细胞寿命,日期,时间 from " + dtrw["TABLE_NAME"].ToString() + " where 住院号=" + "\'" + xy + "\'";
                                //strsql = "select 红细胞寿命,日期,时间 from " + dtrw["TABLE_NAME"].ToString() + " where 姓名='吕布'";
                                //tBoxName.Text = strsql;
                                OleDbDataAdapter dpter = new OleDbDataAdapter();
                                dpter.SelectCommand = new OleDbCommand(strsql, Connec);
                                //dadapter.SelectCommand = new OleDbCommand(strSql1, aConnection);
                                dpter.Fill(dset);
                            }
                            
                        }
                    }
                }
                catch (Exception eex)
                {
                    System.Windows.MessageBox.Show("ERROR:" + eex.Message);
                }
                finally
                {
                    if (Connec != null)
                    {
                        Connec.Close();

                    }
                }
                int[] hxbsm = new int[20];
                int num = 0;
                if (id!=null&&id!=" ")
                {
                    num = dset.Tables[0].Rows.Count;
                }
                string[] time = new string[20];
                try
                {
                    for (int t = 0; t < num; t++)
                    {
                        if (dset.Tables[0].Rows[t]["红细胞寿命"].ToString() == ">250")
                        {
                            hxbsm[t] = 250;
                        }
                        else
                        {
                            hxbsm[t] = Convert.ToInt32(dset.Tables[0].Rows[t]["红细胞寿命"]);
                        }
                        time[t] = string.Concat(dset.Tables[0].Rows[t]["日期"].ToString(), dset.Tables[0].Rows[t]["时间"].ToString().Substring(0, 5));
                    }
                }
                catch (Exception et)
                {
                    System.Windows.MessageBox.Show("ERRORt:" + et.Message);
                }

                try
                {
                    print.ReportPrintHandold(name, gender, age, id, instrumentType, submitDoctor, firstVisit, hb, hospital, rbc, co, co2, testDateLine, department, userDefine[0], userDefine[1], userDefine[2], userDefine[3], userDefine[4], userDefine[5], checkDoctor, reportDoctor, reportTime, remark1, remark2, ty,hxbsm, time, num);
                }
                catch (Exception e25)
                {
                    System.Windows.MessageBox.Show("ERROR25:" + e25.Message);
                }
            }
            

            
        }

        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            f1.setText();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            f1.dbOpen = false;
        }

        private void Open(string FileName)
        {
            app = new Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(FileName);
        }

    }
}
