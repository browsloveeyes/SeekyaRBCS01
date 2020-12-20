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
using System.Xml.Linq;

namespace Seekya
{
    /// <summary>
    /// QC.xaml 的交互逻辑
    /// </summary>
    public partial class QC : Window
    {
        MainWindow f1 = null;

        //public  bool sn = false;

        public QC(MainWindow f)
        {
            InitializeComponent();

            f1 = f;
        }

        private void GetElement(XElement root)
        {
            List<QCRecord> list = new List<QCRecord>();

            foreach (XElement element in root.Elements())
            {
                /*
                 //递归获取质控记录
                if (element.Elements().Count() > 0)
                {
                    GetElement(element);

                }
                else 
                {
                    textBox2.Text += element.Value + "\n";
                    
                }*/
             
                try
                {
                    string[] tp = new string[4];
                    tp = element.Value.Split(',');

                    list.Add(new QCRecord()
                    {
                        Time = tp[0],
                        Precision = tp[1],
                        Accuracy = tp[2],
                        Result = tp[3]

                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error23:" + ex.Message);

                }

            }

            List<QCRecord> tpList = new List<QCRecord>();

            for (int i = list.Count - 1; i >= 0; i--)
            {
                tpList.Add(list[i]);
            }

            dataGridQcRecord.AutoGenerateColumns = false;
            dataGridQcRecord.ItemsSource = tpList;

        
        }

        //追加记录
        public void QCSave(string qcRecord)
        {
            string path = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\qc.xml";
            XDocument document = XDocument.Load(path);

            XElement root = document.Root;  //获取根节点

            //创建一个子节点
            XElement xele = new XElement("record");
            root.Add(xele);
            //添加属性
            xele.SetValue(qcRecord);
            //xele.SetElementValue("precision", "1");
            //xele.SetElementValue("accuracy", "2");
            //xele.SetElementValue("result", "3");

            document.Save(path);
        }

        //建立质控历史记录的类
        public class QCRecord
        {
            public string Time { get; set; }
            public string Precision { get; set; }
            public string Accuracy { get; set; }
            public string Result { get; set; }
        
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (MainWindow.qcOpend == true)
            {
                button1.IsEnabled = false;
            }

        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            f1.sn = true;
            if (MainWindow.qcOpend==true)
            {
                MessageBox.Show("提示：当前使用的是仪器质控流程！");
                button1.IsEnabled = false;

            }
            //输入信息不为空，执行质控
            else if (co.Text.Trim().Length != 0 && co2.Text.Trim().Length != 0)
            {

                try
                {
                    Byte[] temp = new Byte[6];

                    temp[5] = 0X21;
                    temp[4] = 0X00;
                    temp[3] = 0X00;
                    temp[2] = 0X00;
                    temp[1] = 0X01;
                    temp[0] = 0X20;

                    f1.sp.Write(temp, 0, 6);

                    f1.WriteLog("【质控】 " + "200100000021");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERROR24:" + ex.Message);
                }

                button1.IsEnabled = false; 
                f1.qcSign = true;
            }
            else 
            {
                MessageBox.Show("请输全参数","提示");
            
            }

        }

        //获取测试气A的CO浓度差值*100
        public Int16 GetCO()
        {
            return (Int16)(Convert.ToDouble(co.Text.Trim())*100);
        
        }

        //获取测试气A的CO₂浓度值*100
        public Int16 GetCO2()
        {
            return (Int16)(Convert.ToDouble(co2.Text.Trim()) * 100);

        }

        private void tabItem2_MouseDown(object sender, MouseButtonEventArgs e)
        {

            string path = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\qc.xml";
            XDocument document = XDocument.Load(path);

            XElement root = document.Root;  //获取根节点

            //通过递归，获取所有下面的子元素
            GetElement(root);

        }

        //从qc.xml文件，加载质控的历史记录
        private void tabControl1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string path = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\qc.xml";
            XDocument document = XDocument.Load(path);

            XElement root = document.Root;  //获取根节点

            //通过递归，获取所有下面的子元素
            GetElement(root);

        }

        private void Window_Closed(object sender, EventArgs e)
        {
            f1.qcOpen = false;
            f1.softwareOperate = false;
        }
      
    }
}
