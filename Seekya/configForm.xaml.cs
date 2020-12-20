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
using System.IO.Ports;//SerialPort
using System.Threading;
//Process
using System.Diagnostics;
//ObservableCollection
using System.Collections.ObjectModel;
//
using System.Data;

namespace Seekya
{
    /// <summary>
    /// configForm.xaml 的交互逻辑
    /// </summary>
    public partial class configForm : Window
    {
        private MainWindow f1 = null;
        ObservableCollection<doctor> doctorList = new ObservableCollection<doctor>();

        public configForm(MainWindow f)
        {
            InitializeComponent();
            f1 = f;

            DisplayDoctor();

            ((this.FindName("dataGrid1")) as DataGrid).ItemsSource = doctorList;
        }

        public void DisplayDoctor()
        {
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\doctor.txt";
            string dcName;
            //string doctorName;

            try
            {
                //FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.ReadWrite);
                StreamReader sr = new StreamReader(pathString, Encoding.GetEncoding("gb2312"));

                while ((dcName = sr.ReadLine()) != null)
                {
                    doctorList.Add(new doctor()
                    {
                        //num="0",
                        name = dcName

                    });

                }

                sr.Close();
                //fs1.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }

        }
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\hosipitalInfo.txt";
            string hosipitalName = hosipitalN.Text;
            string roomName = roomN.Text;
            string deviceNum = deviceN.Text;
            //string typeName = typeN.Text;

            try
            {
                FileStream fs1 = new FileStream(pathString, FileMode.Create, FileAccess.Write);
                StreamWriter sw1 = new StreamWriter(fs1);

                sw1.WriteLine(hosipitalName);
                sw1.WriteLine(roomName);
                sw1.WriteLine(deviceNum);
                //sw1.WriteLine(typeName);

                sw1.Close();
                fs1.Close();

                //跨线程访问
                //f1.Dispatcher.Invoke(new Action(() => { f1.HosipitalInfoDisplay(); }));
                //f1.HosipitalInfoDisplay();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error9:" + ex.Message);
            }
            finally
            {
                MessageBox.Show("医院信息配置成功");
                //this.DialogResult = true;

            }

            //刷新主界面的医院信息
            f1.RefresHosipitalInfo();

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            bool comExistence = false;//有可用串口标志位
            string t;
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\patientInfo.txt";

            try
            {
                FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.Read);
                StreamReader sr1 = new StreamReader(fs1);

                for (int i = 1; i <= 32; i++)
                {
                    ((TextBox)this.FindName("tBox" + i)).Text = ((t = sr1.ReadLine()) == "null") ? "" : t;

                }

                sr1.Close();
                fs1.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error10:" + ex.Message);
            }

            //typeN.Items.Add("吹气样本");
            //typeN.Items.Add("采气样本");
            //typeN.SelectedIndex = 1;
            serialNum.Items.Clear();//清除当前串口号中的所有串口名称
            for (int i = 0; i < 255; i++)
            {
                try
                {
                    SerialPort sp = new SerialPort("COM" + (i + 1).ToString());
                    sp.Open();
                    sp.Close();
                    serialNum.Items.Add("COM" + (i + 1).ToString());
                    comExistence = true;
                }
                catch (Exception)
                {
                    continue;

                }
            }
            if (comExistence)
            {
                serialNum.SelectedIndex = 0;//使ListBox显示第1个添加的索引
            }
            else
            {
                MessageBox.Show("没有找到可用串口", "错误提示");
            }

            //模板文件  
            string printPathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\print.txt";

            //写入打印模板路径
            try
            {
                StreamReader sr = new StreamReader(printPathString, Encoding.GetEncoding("gb2312"));

                sr.ReadLine();
                tBoxModelName.Text = sr.ReadLine();

                sr.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }

            //读取是否使用扫描枪
            string pathStringCom = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\scan.txt";

            try
            {
                //FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.ReadWrite);
                StreamReader sr = new StreamReader(pathStringCom, Encoding.GetEncoding("gb2312"));

                if (sr.ReadLine() == "0")//不使用扫描枪
                {
                    scanner.IsChecked = false;

                    //读入医院代码以及url
                    hosipitalCode.Text = sr.ReadLine();
                    url.Text = sr.ReadLine();

                    //把连接后台控件全部不启用
                    tbk1.IsEnabled = false;
                    tbk2.IsEnabled = false;
                    hosipitalCode.IsEnabled = false;
                    url.IsEnabled = false;
                    button3.IsEnabled = false;

                }
                else
                {
                    scanner.IsChecked = true;

                    //读入医院代码以及url
                    hosipitalCode.Text = sr.ReadLine();
                    url.Text = sr.ReadLine();

                    //把连接后台控件全部启用
                    tbk1.IsEnabled = true;
                    tbk2.IsEnabled = true;
                    hosipitalCode.IsEnabled = true;
                    url.IsEnabled = true;
                    button3.IsEnabled = true;

                }

                sr.Close();
                //fs1.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }

        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\com.txt";
            string com = serialNum.Text;

            try
            {

                FileStream fs1 = new FileStream(pathString, FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs1);
                sw.WriteLine(com);

                sw.Close();
                fs1.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR11:" + ex.Message);
            }
            finally
            {
                MessageBox.Show("串口号配置成功");
                //this.DialogResult = true;//关闭当前窗口
            }
        }
        //配置患者信息
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\patientInfo.txt";
            string itemName;
            int i;

            try
            {
                FileStream fs1 = new FileStream(pathString, FileMode.Create, FileAccess.Write);
                StreamWriter sw1 = new StreamWriter(fs1);

                for (i = 1; i <= 32; i++)
                {
                    itemName = (((TextBox)this.FindName("tBox" + i)).Text == "") ? "null" : ((TextBox)this.FindName("tBox" + i)).Text;
                    sw1.WriteLine(itemName);
                }

                sw1.Close();
                fs1.Close();
                MessageBox.Show("患者信息配置成功");
                this.DialogResult = true;


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error12:" + ex.Message);
            }

        }

        private void radioButton2_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void add_Click(object sender, RoutedEventArgs e)
        {
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\doctor.txt";
            //string doctorName;

            try
            {
                //FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.ReadWrite);
                StreamWriter sw = new StreamWriter(pathString, true, Encoding.GetEncoding("gb2312"));//true:尾部追加

                sw.WriteLine(tboxDoctor.Text);

                sw.Close();
                //fs1.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }

            doctorList.Add(new doctor()
            {
                //num="0",
                name = tboxDoctor.Text

            });

            //刷新主界面医生姓名
            f1.RefreshDoctor();

        }

        private void del_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string selectItem = ((Seekya.doctor)(this.dataGrid1.SelectedItem)).name as string;//((Seekya.doctor)(this.dataGrid1.SelectedItem)).name
                string[] name = new string[20];// 只可加入20个名字
                //System.Windows.MessageBox.Show(selectItem);
                //cell=dataGrid1.Items[dataGrid1.SelectedIndex].Cells[0].Text;

                string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\doctor.txt";
                int i = 0;
                //string doctorName;

                try
                {
                    //FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.ReadWrite);
                    StreamReader sr = new StreamReader(pathString, Encoding.GetEncoding("gb2312"));

                    while ((name[i] = sr.ReadLine()) != null)
                    {
                        if (String.Compare(name[i], selectItem) == 0)
                            i = i;
                        else
                            i++;
                    }

                    sr.Close();
                    //fs1.Close();

                }
                catch (Exception ex)
                {
                    // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

                }

                //清空
                doctorList.Clear();

                try
                {
                    //FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.ReadWrite);
                    StreamWriter sw = new StreamWriter(pathString, false, Encoding.GetEncoding("gb2312"));//true:尾部追加

                    for (int j = 0; j < i; j++)
                    {
                        sw.WriteLine(name[j]);

                        doctorList.Add(new doctor()
                        {
                            //num="0",
                            name = name[j]

                        });

                    }

                    sw.Close();
                    //fs1.Close();

                }
                catch (Exception ex)
                {
                    // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

                }

                //刷新主界面医生姓名
                f1.RefreshDoctor();
            }
            catch (Exception ex)
            { 
            
            }

        }

        private void clr_Click(object sender, RoutedEventArgs e)
        {

            tboxDoctor.Text = "";
        }

        //刷新
        private delegate void outputDelegate();

        private void button4_Click_1(object sender, RoutedEventArgs e)
        {
            this.serialNum.Dispatcher.Invoke(new outputDelegate(RefreshCom));
        }

        public void RefreshCom()
        {
            bool comExistence = false;//有可用串口标志位
            serialNum.Items.Clear();//清除当前串口号中的所有串口名称

            for (int i = 0; i < 255; i++)
            {
                try
                {
                    SerialPort sp = new SerialPort("COM" + (i + 1).ToString());
                    sp.Open();
                    sp.Close();
                    serialNum.Items.Add("COM" + (i + 1).ToString());
                    comExistence = true;
                }
                catch (Exception)
                {
                    continue;

                }
            }
            if (comExistence)
            {
                serialNum.SelectedIndex = 0;//使ListBox显示第1个添加的索引
            }
            else
            {
                MessageBox.Show("没有找到可用串口", "错误提示");
            }

        }

        private void button5_Click(object sender, RoutedEventArgs e)
        {
            string fileName = null;
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "所有文件(*.*)|*.*";
            string tmp = null;

            var result = openFileDialog.ShowDialog();
            if (result == true)
            {
                fileName = openFileDialog.FileName;
            }

            MessageBoxResult res = MessageBox.Show("确定选择该打印模板吗？", "提示", MessageBoxButton.OKCancel);

            if (res == MessageBoxResult.OK)
            {
                //把模板名写进txt文件中存储
                string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\print.txt";

                //读打印方式
                try
                {
                    StreamReader sr = new StreamReader(pathString, Encoding.GetEncoding("gb2312"));

                    tmp = sr.ReadLine();

                    sr.Close();

                }
                catch (Exception ex)
                {
                    // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

                }
                //写模板名
                try
                {
                    StreamWriter sw = new StreamWriter(pathString, false, Encoding.GetEncoding("gb2312"));//true:尾部追加

                    sw.WriteLine(tmp);
                    sw.WriteLine(fileName);

                    sw.Close();
                    //fs1.Close();

                }
                catch (Exception ex)
                {
                    // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

                }

                //把选择的模板路径，写进tBoxModelName控件中
                tBoxModelName.Text = fileName;

            }
        }

        

        private void disconnServer_Click(object sender, RoutedEventArgs e)
        {
            //Thread disconn = new Thread(my.DisconnectToServer);

            //disconn.Start();

            f1.DisconnectToServer();

        }

        private void scanner_Click(object sender, RoutedEventArgs e)
        {
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\scan.txt";

            if (scanner.IsChecked == true)//使用扫描枪
            {
                try
                {
                    FileStream fs1 = new FileStream(pathString, FileMode.Create, FileAccess.Write);
                    StreamWriter sw1 = new StreamWriter(fs1);

                    sw1.WriteLine("1");
                    sw1.WriteLine(hosipitalCode.Text.Trim());
                    sw1.WriteLine(url.Text.Trim());

                    sw1.Close();
                    fs1.Close();

                    //使能连接后台配置
                    tbk1.IsEnabled = true;
                    tbk2.IsEnabled = true;
                    hosipitalCode.IsEnabled = true;
                    url.IsEnabled = true;
                    button3.IsEnabled = true;

                    //使能主界面条形码确认按键
                    f1.enableBar();

                }
                catch (Exception ex)
                {

                }

            }
            else
            {
                try
                {
                    FileStream fs1 = new FileStream(pathString, FileMode.Create, FileAccess.Write);
                    StreamWriter sw1 = new StreamWriter(fs1);

                    sw1.WriteLine("0");
                    sw1.WriteLine(hosipitalCode.Text.Trim());
                    sw1.WriteLine(url.Text.Trim());

                    sw1.Close();
                    fs1.Close();

                    //不使能连接后台配置
                    tbk1.IsEnabled = false;
                    tbk2.IsEnabled = false;
                    hosipitalCode.IsEnabled = false;
                    url.IsEnabled = false;
                    button3.IsEnabled = false;

                    //使能主界面条形码确认按键
                    f1.unenableBar();
                }
                catch (Exception ex)
                {

                }

            }
        }

        //存储扫描功能所需的医院代码以及网址
        private void button3_Click(object sender, RoutedEventArgs e)
        {
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\scan.txt";

            try
            {
                StreamWriter sw = new StreamWriter(pathString, false, Encoding.GetEncoding("gb2312"));//true:尾部追加

                if (scanner.IsChecked == true)//使能扫描功能
                    sw.WriteLine("1");
                else
                    sw.WriteLine("0");

                sw.WriteLine(hosipitalCode.Text.Trim());
                sw.WriteLine(url.Text.Trim());

                sw.Close();
                //fs1.Close();

                MessageBox.Show("扫描枪支持配置成功");

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }

        }

        private void Window_Closed(object sender, EventArgs e)
        {
            f1.cfgOpen = false;
        }
    
    }
}
