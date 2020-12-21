using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows;
//using System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO.Ports;
using System.Threading;
using System.IO;
//ArrayList
using System.Collections;//新
using Excel = Microsoft.Office.Interop.Excel;


namespace Seekya
{
    public partial class MainWindow : Window
    {
        public Excel.Application app;
        public Excel.Workbooks wbs;
        public Excel.Workbook wb;

        private List<byte> buffer = new List<byte>(4096);

        public SerialPort sp = null;//声明一个串口类
        bool l1 = true, l2 = true, l3 = true, l4 = true;//显示灯状态标志位，true表示灯亮，false表示灯暗

        Byte[] RecvData = new Byte[6];//创建接收字节数组    

        //联机标志位，联机成功值置为true
        Boolean firstConn = false;

        //当前连接的串口号
        string com1 = null;
        //打算连接的串口号
        string com2 = null;

        //串口打开标志位，判断是否处理串口通信事件，true，为处理，否则，不处理，目的防止串口关闭软件卡死问题
        bool spOpenSign = false;

        //零点过大的变量，初始值为
        int zeroOver = 0;
        //co备注栏
        string coSign = "";
        //co2过低，true：过低 false：正常
        string co2LowSign = "";

        //按下“测量”键，步骤标志位，默认为0
        int measureStep = 0;

        //质控标志位，false：不是质控阶段，true：质控阶段
        public bool qcSign = false;

        //质控进行到哪一步的标志位
        public Int16 qcStep = 0;

        //红细胞寿命
        Int32 RBCT = 0;


        //获取配置好的串口号
        private string GetCom()
        {
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory+"Data\\com.txt";
            string com3;          
            try
            {

                FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fs1);

                com3 = sr.ReadLine();

                sr.Close();
                fs1.Close();

                return com3;

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR1:" + ex.Message);
                return  "";

            }

        }

        private void SetPortProperty()//设置串口的属性
        {
            string tmp = GetCom();
            
            sp = new SerialPort();
            sp.PortName = tmp;//设置串口号

            sp.BaudRate = 9600;//设置串口的波特率为9600
            sp.StopBits = StopBits.One;//设置停止位为1位
            sp.DataBits = 8;//设置数据位为8位
            sp.Parity = Parity.None;//设置奇偶校验位为None

            sp.ReadTimeout = -1;//设置超时读取时间
            sp.RtsEnable = true; //定义DataReceived事件，当串口收到数据后触发事件
            sp.DataReceived += new SerialDataReceivedEventHandler(sp_DataReceived);
            //isHexDisplay=true;//16进制显示
            //isHexSend = true;//16进制发送
       

        }

        private Byte CheckSum(Byte[] arr)
        {
            Byte sum = 0;

            for (int i = 0; i < 5; i++)
                sum += arr[i];

            return sum;

        }

        //public void SerialOpen()
        //{
        //    DateTime dt = System.DateTime.Now;
        //    string date = dt.ToLocalTime().ToString();
        //    string time = dt.ToString("HH:mm:ss");

        //    SetPortProperty();//设置串口属性

        //    try//打开串口
        //    {
        //        if (string.Compare(com1, com2) != 0)
        //        {
        //            sp.Open();
        //            receiveInfo.Text += "[" + time + "]:" + "串口打开" + System.Environment.NewLine;

        //            com1 = com2;

        //            //给下位机发送DD
        //            Byte[] temp = new Byte[1];
        //            temp[0] = 0XDD;

        //            //写日志
        //            WriteLog("[" + date + "]" + ":" + "DD");

        //            //
        //            this.receiveInfo.ScrollToEnd();

        //            sp.Write(temp, 0, 1);//(temp, 0, 1);
        //        }

        //    }
        //    catch (Exception)
        //    { 
        //        //打开串口失败后，相应标志位取消
        //        MessageBox.Show("串口无效或已被占用，连接仪器失败", "错误提示");
        //    }

        //}

        public void SerialClose()
        {
            try
            {
                sp.Close();
                sp.Dispose();
            }
            catch (Exception)
            {
                //MessageBox.Show("断开仪器失败", "提示错误");

            }

        }
        //发送ASCII码的数据
        private void SendASCII(object SD)
        {
            string sd = SD as string;
            Encoding gb = System.Text.Encoding.GetEncoding("gb2312");
            Byte[] writeBytes = gb.GetBytes(sd);
            Byte[] head = {0X5A,0XA5,(Byte)writeBytes.Length };
            Byte[] info = CombineByteArray(head,writeBytes);

            Thread.Sleep(3000);

            try
            {
                sp.Write(info, 0, info.Length);
            }
            catch(Exception ex)
            {
                //出错不显示
            }
        }

        //合并两个字节数组
        private Byte[] CombineByteArray(Byte[] a, Byte[] b)
        {
            Byte[] c=new Byte[a.Length + b.Length];

            a.CopyTo(c,0);
            b.CopyTo(c,a.Length);

            return c;
        
        }
        //根据接收到的提示信息显示
        private void ShowTip(Byte[] ReceivedData)
        {
            //获取接收数据时的系统时间
            DateTime dt1 = System.DateTime.Now;
            string time1 = dt1.ToString("HH:mm:ss");

            //在C#当中通常以Image_Test.Source=new BitmapImage(new Uri(“图片路径”,UriKind. RelativeOrAbsolute))的方式来为Image控件指定Source属性。
            if (ReceivedData[0] == 0X80)//气袋状态
            {

                if (ReceivedData[3] == 0X00 && ReceivedData[4] == 0X00)
                {
                    switch (ReceivedData[1])
                    {
                        case 0X00: if (ReceivedData[2] == 0X00) { receiveInfo.Text += "[" + time1 + "]:" + "肺泡气袋空闲" + System.Environment.NewLine; light1.Source = new BitmapImage(new Uri("/Seekya;component/Images/lightOff1.jpg", UriKind.Relative)); l1 = false; if (l2 == false && l3 == false && l4 == false) Reflash();  } else if (ReceivedData[2] == 0X01) { receiveInfo.Text += "[" + time1 + "]:" + "肺泡气袋插入" + System.Environment.NewLine; light1.Source = new BitmapImage(new Uri("/Seekya;component/Images/lightOn1.jpg", UriKind.Relative)); l1 = true;  } break;
                        case 0X01: if (ReceivedData[2] == 0X00) { receiveInfo.Text += "[" + time1 + "]:" + "本底气袋空闲" + System.Environment.NewLine; light2.Source = new BitmapImage(new Uri("/Seekya;component/Images/lightOff1.jpg", UriKind.Relative)); l2 = false; if (l1 == false && l3 == false && l4 == false) Reflash();  } else if (ReceivedData[2] == 0X01) { receiveInfo.Text += "[" + time1 + "]:" + "本底气袋插入" + System.Environment.NewLine; light2.Source = new BitmapImage(new Uri("/Seekya;component/Images/lightOn1.jpg", UriKind.Relative)); l2 = true;  } break;
                        case 0X02: if (ReceivedData[2] == 0X00) { receiveInfo.Text += "[" + time1 + "]:" + "倒气袋1空闲" + System.Environment.NewLine; light3.Source = new BitmapImage(new Uri("/Seekya;component/Images/lightOff1.jpg", UriKind.Relative)); l3 = false; if (l1 == false && l2 == false && l4 == false) Reflash();  } else if (ReceivedData[2] == 0X01) { receiveInfo.Text += "[" + time1 + "]:" + "倒气袋1插入" + System.Environment.NewLine; light3.Source = new BitmapImage(new Uri("/Seekya;component/Images/lightOn1.jpg", UriKind.Relative)); l3 = true;  } break;
                        case 0X03: if (ReceivedData[2] == 0X00) { receiveInfo.Text += "[" + time1 + "]:" + "倒气袋2空闲" + System.Environment.NewLine; light4.Source = new BitmapImage(new Uri("/Seekya;component/Images/lightOff1.jpg", UriKind.Relative)); l4 = false; if (l1 == false && l2 == false && l3 == false) Reflash();  } else if (ReceivedData[2] == 0X01) { receiveInfo.Text += "[" + time1 + "]:" + "倒气袋2插入" + System.Environment.NewLine; light4.Source = new BitmapImage(new Uri("/Seekya;component/Images/lightOn1.jpg", UriKind.Relative)); l4 = true;  } break;
                        //default: MessageBox.Show("接收数据有误！！"); break;

                    }
                }

            }
            else if (ReceivedData[0] == 0X90)//预热状态
            {
                if (ReceivedData[1] == 0 && ReceivedData[3] == 0 && ReceivedData[4] == 0)
                {
                    switch (ReceivedData[2])
                    {
                        case 0X00: receiveInfo.Text += "[" + time1 + "]:" + "仪器初始化 ..." + System.Environment.NewLine;  break;
                        case 0X01: receiveInfo.Text += "[" + time1 + "]:" + "仪器初始化完成" + System.Environment.NewLine;  break;
                        case 0X02: receiveInfo.Text += "[" + time1 + "]:" + "仪器就绪" + System.Environment.NewLine;  break;
                        //default: MessageBox.Show("接收数据有误！！"); break;
                        
                    }
                }

            }
            else if (ReceivedData[0] == 0XA0)//
            {
                if (ReceivedData[3] == 0 && ReceivedData[4] == 0)
                {
                    switch (ReceivedData[1])
                    {
                        case 0X00: if (ReceivedData[2] == 0X00)
                            {
                                receiveInfo.Text += "[" + time1 + "]:" + "测量开始..." + System.Environment.NewLine;
                            }
                            else if (ReceivedData[2] == 0X01) { receiveInfo.Text += ("[" + time1 + "]:" + "测量完成" + System.Environment.NewLine);  }
                            else if (ReceivedData[2] == 0X02) { receiveInfo.Text += ("[" + time1 + "]:" + "测量出错" + System.Environment.NewLine);  }
                            break;
                        case 0X01: if (ReceivedData[2] == 0X00)
                            {
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
                                receiveInfo.Text += "[" + time1 + "]:" + "第一步进行中...." + System.Environment.NewLine;
                                string pathoffice = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\office.txt";
                                FileStream fsoffice = new FileStream(pathoffice, FileMode.Create, FileAccess.Write);
                                StreamWriter swoffice = new StreamWriter(fsoffice);
                                try
                                {                                 
                                    Open(str);
                                    swoffice.WriteLine("True");
                                }
                                catch (Exception err)
                                {
                                    swoffice.WriteLine("False");
                                }
                                swoffice.Close();
                                fsoffice.Close();
                            }
                            else if (ReceivedData[2] == 0X01)
                            {
                                receiveInfo.Text += "[" + time1 + "]:" + "第一步完成" + System.Environment.NewLine;
                            }
                            else if (ReceivedData[2] == 0X02)
                            {
                                receiveInfo.Text += "[" + time1 + "]:" + "第一步出错" + System.Environment.NewLine;
                            }
                            break;
                        case 0X02: if (ReceivedData[2] == 0X00) { receiveInfo.Text += "[" + time1 + "]:" + "第二步进行中...." + System.Environment.NewLine;  } else if (ReceivedData[2] == 0X01) { receiveInfo.Text += "[" + time1 + "]:" + "第二步完成" + System.Environment.NewLine;  } else if (ReceivedData[2] == 0X02) { receiveInfo.Text += "[" + time1 + "]:" + "第二步出错" + System.Environment.NewLine;  } break;
                        case 0X03: if (ReceivedData[2] == 0X00) { receiveInfo.Text += "[" + time1 + "]:" + "第三步进行中...." + System.Environment.NewLine;  } else if (ReceivedData[2] == 0X01) { receiveInfo.Text += "[" + time1 + "]:" + "第三步完成" + System.Environment.NewLine;  } else if (ReceivedData[2] == 0X02) { receiveInfo.Text += "[" + time1 + "]:" + "第三步出错" + System.Environment.NewLine;  } break;
                        case 0X04: if (ReceivedData[2] == 0X00) { receiveInfo.Text += "[" + time1 + "]:" + "第四步进行中...." + System.Environment.NewLine;  } else if (ReceivedData[2] == 0X01) { receiveInfo.Text += "[" + time1 + "]:" + "第四步完成" + System.Environment.NewLine;  } else if (ReceivedData[2] == 0X02) { receiveInfo.Text += "[" + time1 + "]:" + "第四步出错" + System.Environment.NewLine;  } break;

                        //default: MessageBox.Show("接收数据有误！！"); break;

                    }
                }

            }
            else if (ReceivedData[0] == 0XC0)//测量结果
            {           
                
                switch (ReceivedData[1])
                {   
                    //0X05:接收到零点数据
                    case 0X05: double zero = ReceivedData[2] * 16 * 16 * 16 * 16 + ReceivedData[3] * 16 * 16 + ReceivedData[4]; DateTime dt3 = System.DateTime.Now; string date2 = dt3.ToLocalTime().ToString(); WriteZero("[" + date2 + "]" + ":" + zero.ToString()); break;
                    //0X06:提示零点过大
                    case 0X06: if (ReceivedData[2] == 0X00) { zeroOver = ReceivedData[3] * 16 * 16 + ReceivedData[4]; Thread zeroOversize = new Thread(new ThreadStart(ShowZeroOversizeFault)); zeroOversize.IsBackground = true; zeroOversize.SetApartmentState(ApartmentState.STA); zeroOversize.Start(); } break;
                    case 0X07: if (ReceivedData[2] == 0X00) { Int32 pre = ReceivedData[3] * 16 * 16 + ReceivedData[4]; myQC.precision.Text = (100 - pre).ToString() + "～" + (100 + pre).ToString(); } break;
                    case 0X08: if (ReceivedData[2] == 0X00) { Int32 acc = ReceivedData[3] * 16 * 16 + ReceivedData[4]; myQC.accuracy.Text = (100 - acc).ToString() + "～" + (100 + acc).ToString(); } break;
                    case 0X00: if (ReceivedData[1] == 0X00 && ReceivedData[2] == 0) RBCT = ReceivedData[3] * 16 * 16 + ReceivedData[4]; break;
                    case 0X02: double PCO = (ReceivedData[2] * 16 * 16 * 16 * 16 + ReceivedData[3] * 16 * 16 + ReceivedData[4]) / 10000.0; receiveInfo.Text += ("[" + time1 + "]:" + "内源性CO浓度为：" + PCO.ToString("0.0000") + "ppm" + System.Environment.NewLine); CO.Text = PCO.ToString("0.0000"); break;
                    case 0X03: double CO2 = (ReceivedData[2] * 16 * 16 * 16 * 16 + ReceivedData[3] * 16 * 16 + ReceivedData[4]) / 100.0; receiveInfo.Text += ("[" + time1 + "]:" + "CO2浓度:" + CO2.ToString("0.00") + "%" + System.Environment.NewLine); PCO2.Text = CO2.ToString("0.00"); break;
                    case 0x04: if (ReceivedData[2] == 0) { Int32 r = ReceivedData[3] * 16 * 16 + ReceivedData[4]; rbConcentration.Text = r.ToString();
                    if (rbConcentration.Text.Trim().Length == 0 || String.Compare(rbConcentration.Text, "0") == 0)   //没输入血红蛋白浓度//显示红细胞寿命
                    {
                        receiveInfo.Text += ("[" + time1 + "]:" + "未输入血红蛋白浓度，红细胞寿命未知" + System.Environment.NewLine);
                        day.Text = "";

                    }
                    else
                    {
                        string strRBC = (RBCT > 250) ? ">250" : RBCT.ToString();
                        receiveInfo.Text += ("[" + time1 + "]:" + "红细胞寿命为：" + strRBC + "天" + System.Environment.NewLine);
                        day.Text = strRBC;
                    }
                    
                    }
                    //case 0X04: double rbC= ReceivedData[3] * 16 * 16 + ReceivedData[4]; rbConcentration.Text = rbC.ToString();

                        //建立后台多线程，发送数据给仪器（SD）
                        Thread t = new Thread(new ParameterizedThreadStart(SendASCII));
                        Encoding gb = System.Text.Encoding.GetEncoding("gb2312");
                        
                        t.IsBackground = true;//后台运作

                        //接收测量结果完成，把结果数据导入数据库
                        DateTime dt = System.DateTime.Now;
                        string date = dt.ToString("yyyy/MM/dd");
                        string date1 = dt.ToString("yyyyMMdd");
                        string time = dt.ToString("HH:mm:ss");

                        date = date.Substring(0, 4) + '/' + date.Substring(5, 2) + '/' + date.Substring(8,2);

                        string hsptName = (hosipitalName.Text == "") ? " " : hosipitalName.Text;
                        string rName = (roomName.Text == "") ? " " : roomName.Text;
                        string dNum = (deviceNum.Text == "") ? " " : deviceNum.Text;
                        string i = (id.Text == "") ? " " : id.Text;
                        string nm = (name.Text == "") ? " " : name.Text;
                        string ag = (age.Text == "") ? " " : age.Text;
                        string sx = (sex.Text == "") ? " " : sex.Text;
                        string dy = (day.Text == "") ? " " : day.Text;
                        string CO1 = (CO.Text == "") ? " " : CO.Text;
                        string CO21 = (PCO2.Text == "") ? " " : PCO2.Text;
                        string rb = (rbConcentration.Text == "") ? "0" : rbConcentration.Text;
                        string sDoctor = (sendDoctor.Text == "") ? " " : sendDoctor.Text;
                        string fCheck = (firstCheck.Text == "") ? " " : firstCheck.Text;

                        //加入报告医生和复核医生
                        string cDoctor = (checkDoctor.Text == "") ? " " : checkDoctor.Text;
                        string rDoctor = (reviewDoctor.Text == "") ? " " : reviewDoctor.Text;
                        //备注1，零点过大和备注2，co2浓度过低
                        string cork = (String.Compare(coSign.Trim(), "") == 0) ? " " : coSign;
                        string co2Low = (String.Compare(co2LowSign.Trim(), "") == 0) ? " " : co2LowSign;

                        string SDItem = "";//存储每次的检验样品的信息，以用于发送到仪器上的SD卡中存储

                        //如果当天表不存在，则创建
                        OleDbConnection aConnection1 = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
                        string strSql1 = "Select * from " + date1;
                        string patientPathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\patientInfo.txt";
                        string[] item = new string[6];

                        try//判断表是否存在，程序不够严谨（只要判断打开数据库表时出现错误，就归结于表不存在，以后改进）!!
                        {
                            aConnection1.Open();
                            OleDbCommand myCmd = new OleDbCommand(strSql1, aConnection1);
                            myCmd.ExecuteNonQuery();

                        }
                        catch (Exception ex)//表不存在，创建表
                        {
                            try
                            {
                                int j;
                                FileStream fs1 = new FileStream(patientPathString, FileMode.Open, FileAccess.Read);
                                StreamReader sr1 = new StreamReader(fs1);

                                for (j = 1; j < 21; j++)//读取txt文件到21行
                                {
                                    sr1.ReadLine();
                                }
                                for (; j < 32; j = j + 2)
                                {
                                    item[(j - 21) / 2] = sr1.ReadLine();
                                    sr1.ReadLine();

                                }

                                sr1.Close();
                                fs1.Close();

                            }
                            catch (Exception e)
                            {
                                System.Windows.MessageBox.Show("Error2:" + e.Message);
                            }

                            ArrayList headList = new ArrayList();
                            DbOperate testDb = new DbOperate();

                            headList.Add("医院名称"); headList.Add("科室名称"); headList.Add("仪器型号");
                            headList.Add("姓名"); headList.Add("性别"); headList.Add("年龄"); headList.Add("住院号");
                            headList.Add("CO"); headList.Add("CO2"); headList.Add("红细胞寿命"); headList.Add("血红蛋白浓度");
                            headList.Add("送检医生"); headList.Add("复核医生"); headList.Add("报告医生");
                            headList.Add("初步诊断"); 
                            headList.Add("时间"); headList.Add("日期"); headList.Add("备注1"); headList.Add("备注2");

                            for (int k = 0; k < 6; k++)
                            {
                                if (item[k] != "null")
                                    headList.Add(item[k]);
                            }

                            testDb.CreateTable(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb", date1, headList);

                        }
                        finally
                        {
                            if (aConnection1 != null)
                                aConnection1.Close();

                        }

                        OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
                        string strSql = "Insert into " + date1 + " (医院名称,科室名称,仪器型号,姓名,性别,年龄,住院号,CO,CO2,红细胞寿命,血红蛋白浓度,送检医生,复核医生,报告医生,初步诊断,时间,日期,备注1,备注2) values ('" + hsptName + "','" + rName + "','" + dNum + "','" + nm + "','" + sx + "','" + ag + "','" + i + "','" + CO1 + "','" + CO21 + "','" + dy + "','" + rb + "','" + sDoctor + "','" + rDoctor + "','" + cDoctor + "','" + fCheck + "','"  + time + "','" + date + "','" + cork + "','" + co2Low + "')";

                        //MessageBox.Show(hsptName + "," + rName + "," + dNum + "," + nm + "," + sx + "," + ag + "," + i + "," + CO1 + "," + CO21 + "," + dy + "," + rb + "," + sDoctor + ","  + fCheck + "," + rmk + "," + time + "," + date);
                        try
                        {
                            aConnection.Open();
                            OleDbCommand myCmd = new OleDbCommand(strSql, aConnection);
                            myCmd.ExecuteNonQuery();

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("ERROR3:" + ex.Message);

                        }
                        finally
                        {
                            if (aConnection != null)
                                aConnection.Close();

                        }

                        todayReportDisplay();

                        hsptName = (hosipitalName.Text.Trim() == "") ? "null" : hosipitalName.Text.Trim();
                        rName = (roomName.Text.Trim() == "") ? "null" : roomName.Text.Trim();
                        dNum = (deviceNum.Text.Trim() == "") ? "null" : deviceNum.Text.Trim();
                        i = (id.Text.Trim() == "") ? "null" : id.Text.Trim();
                        nm = (name.Text.Trim() == "") ? "null" : name.Text.Trim();
                        ag = (age.Text.Trim() == "") ? "null" : age.Text.Trim();
                        sx = (sex.Text.Trim() == "") ? "null" : sex.Text.Trim();
                        dy = (day.Text.Trim() == "") ? "null" : day.Text.Trim();
                        CO1 = (CO.Text.Trim() == "") ? "null" : CO.Text.Trim();
                        CO21 = (PCO2.Text.Trim() == "") ? "null" : PCO2.Text.Trim();
                        rb = (rbConcentration.Text.Trim() == "") ? "null" : rbConcentration.Text.Trim();
                        sDoctor = (sendDoctor.Text.Trim() == "") ? "null" : sendDoctor.Text.Trim();
                        fCheck = (firstCheck.Text.Trim() == "") ? "null" : firstCheck.Text.Trim();
                        //加入报告医生和复核医生
                        cDoctor = (checkDoctor.Text.Trim() == "") ? "null" : checkDoctor.Text.Trim();
                        rDoctor = (reviewDoctor.Text.Trim() == "") ? "null" : reviewDoctor.Text.Trim();

                        SDItem = hsptName + "@" + rName + "@" + dNum + "@" + nm + "@" + i + "@" + sx + "@" + ag + "@" + rb + "@" + dy + "@" + CO1 + "@" + CO21 + "@" + sDoctor + "@" + cDoctor + "@" + rDoctor + "@" + fCheck ;
                        
                        Byte[] writeBytes = gb.GetBytes(SDItem);

                        //把检验结果发送到下位机的SD卡
                        if (writeBytes.Length < 155)//样品信息的字节数不超过了155（限定传输的字节长度）
                        {
                            t.Start("["+date+"]:"+SDItem);
                        }
                        else//字节数超过了155，则省去备注
                        {
                            SDItem = hsptName + "@" + rName + "@" + dNum + "@" + nm + "@" + i + "@" + sx + "@" + ag + "@" + rb + "@" + dy + "@" + CO1 + "@" + CO21 + "@" + sDoctor + "@" + cDoctor + "@" + rDoctor + "@" + fCheck;

                            t.Start("[" + date + "]:" + SDItem);
                        }

                        //判断是否调用后台接口
                        if (scanBarOk.IsEnabled == true)
                        {
                            //向后台检验结果
                            string[] args = new string[3];
                            args[0] = tBoxScanBar.Text;
                            args[1] = CO1 + "|" + CO21 + "|" + dy ;  //CO|CO2|红细胞寿命
                            args[2] = "TNSHXBSM";
                            string url = null;

                            string pathStringCom = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\scan.txt";

                            try
                            {
                                //FileStream fs1 = new FileStream(pathString, FileMode.Open, FileAccess.ReadWrite);
                                StreamReader sr = new StreamReader(pathStringCom, Encoding.GetEncoding("gb2312"));

                                sr.ReadLine();
                                sr.ReadLine();

                                url = sr.ReadLine();

                                sr.Close();
                                //fs1.Close();

                            }
                            catch (Exception ex)
                            {
                                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

                            }

                            try
                            {
                                object result = WebServiceHelper.InvokeWebService(url, "DHCUpdateResult", args);

                                if (String.Compare(result.ToString(), "0") == 0)
                                {
                                    //接收成功

                                }
                                else
                                {
                                    

                                }
                            }
                            catch (Exception ex)
                            {
                                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

                            }

                        }

                        //把“测量”按键使能
                        measure.IsEnabled = true;

                        //把零点过大的数据置0
                        zeroOver = 0;
                        //把CO置为空
                        coSign = "";
                        //把CO2低置空
                        co2LowSign = "";
                        //把测量步骤置0
                        measureStep = 0;

                        break;

                }
            }
            else if (ReceivedData[0] == 0X0D)//
            {
                if (ReceivedData[1]==0X00 && ReceivedData[2] == 0 && ReceivedData[3] == 0 && ReceivedData[4] == 0)
                {
                    receiveInfo.Text += "[" + time1 + "]:" + "准备就绪" + System.Environment.NewLine;
                    //default: MessageBox.Show("接收数据有误！！"); break;

                }

            }
            else if (ReceivedData[0] == 0XD0)
            {
                if (ReceivedData[1] == 0X04 && ReceivedData[2] == 0X00 && ReceivedData[3] == 0X00 && ReceivedData[4] == 0X00)
                {
                    Thread co2Lower = new Thread(new ThreadStart(ShowCO2LowFault)); co2Lower.IsBackground = true; co2Lower.SetApartmentState(ApartmentState.STA); co2Lower.Start();
                    co2LowSign = "*";
                }
                else if (ReceivedData[1] == 0X04 && ReceivedData[2] == 0X01 && ReceivedData[3] == 0X00 && ReceivedData[4] == 0X00)
                {
                    Thread co2Lower = new Thread(new ThreadStart(ShowCO2LowFault)); co2Lower.IsBackground = true; co2Lower.SetApartmentState(ApartmentState.STA); co2Lower.Start();
                    co2LowSign = "**";
                
                }
                else if (ReceivedData[1] == 0X05 && ReceivedData[2] == 0X00)  //co备注
                {
                    int tp = ReceivedData[3] * 16 * 16 + ReceivedData[4];
                    coSign = "*(" + tp.ToString() + ")";
                
                }
                else if (ReceivedData[1] == 0X05 && ReceivedData[2] == 0X01)  //co备注
                {
                    int tp = ReceivedData[3] * 16 * 16 + ReceivedData[4];
                    coSign = "**(" + tp.ToString() + ")";

                }
                else if (ReceivedData[1] == 0X05 && ReceivedData[2] == 0X02)  //co备注
                {
                    coSign = "*";

                }
                else if (ReceivedData[1] == 0X06 && ReceivedData[2] == 0X00)    //质控CO2出错
                {
                    MessageBox.Show("质控未完成（请检查CO2测量系统），拔掉所有气袋，仪器返回待机界面", "提示");

                }
                else if (ReceivedData[1] == 0X06 && ReceivedData[2] == 0X01)     //质控零点错误
                {
                    MessageBox.Show("质控未完成（Zero Fault），拔掉所有气袋，仪器返回待机界面", "提示");

                }
                else if (ReceivedData[1] == 0X06 && ReceivedData[2] == 0X02)     //质控未通过
                {
                    myQC.result.Text = "未通过";
                    myQC.textBox1.Text += "[" + time1 + "]  " + "质控未通过" + System.Environment.NewLine;
                    string record = DateTime.Now.ToString() + "," + myQC.precision.Text.Trim() + "," + myQC.accuracy.Text.Trim() + "," + myQC.result.Text.Trim();
                    myQC.QCSave(record);

                    //进度显示最新信息
                    myQC.textBox1.ScrollToEnd();

                }
                else if(ReceivedData[1] == 0X01 && ReceivedData[2] == 0 && ReceivedData[3] == 0 && ReceivedData[4] == 0)
                {
                    receiveInfo.Text += "[" + time1 + "]:" + "零点错误" + System.Environment.NewLine; Thread zero = new Thread(new ThreadStart(ShowZeroFault)); zero.IsBackground = true; zero.SetApartmentState(ApartmentState.STA); zero.Start();
                }    
                else if(ReceivedData[1] == 0X02 && ReceivedData[2] == 0 && ReceivedData[3] == 0 && ReceivedData[4] == 0)   
                {
                    receiveInfo.Text += "[" + time1 + "]:" + "测试错误" + System.Environment.NewLine; Thread test = new Thread(new ThreadStart(ShowTestFault)); test.IsBackground = true; test.SetApartmentState(ApartmentState.STA); test.Start(); 
                }
                else if (ReceivedData[1] == 0X03 && ReceivedData[2] == 0 && ReceivedData[3] == 0 && ReceivedData[4] == 0)
                {
                    receiveInfo.Text += "[" + time1 + "]:" + "样本错误" + System.Environment.NewLine; Thread sample = new Thread(new ThreadStart(ShowSampleFault)); sample.IsBackground = true; sample.SetApartmentState(ApartmentState.STA); sample.Start(); 
                }
            }
            else if (ReceivedData[0] == 0XB0)
            {
                if (ReceivedData[1] == 0X00 && ReceivedData[2] == 0X00)
                {
                    myQC.textBox1.Text += "[" + time1 + "]  " + "质控开始" + System.Environment.NewLine;
                    myQC.textBox1.ScrollToEnd();
                }
                else if (ReceivedData[1] == 0X00 && ReceivedData[2] == 0X01)
                {
                    myQC.textBox1.Text += "[" + time1 + "]  " + "质控结束" + System.Environment.NewLine;
                    myQC.textBox1.ScrollToEnd();
                }
                else if (ReceivedData[1] == 0X01 && ReceivedData[2] == 0X00)
                {
                    myQC.textBox1.Text += "[" + time1 + "]  " + "第一阶段开始" + System.Environment.NewLine;
                    myQC.textBox1.ScrollToEnd();
                }
                else if (ReceivedData[1] == 0X01 && ReceivedData[2] == 0X01)
                {
                    myQC.textBox1.Text += "[" + time1 + "]  " + "第一阶段结束" + System.Environment.NewLine;
                    myQC.textBox1.ScrollToEnd();
                }
                else if (ReceivedData[1] == 0X02 && ReceivedData[2] == 0X00)
                {
                    myQC.textBox1.Text += "[" + time1 + "]  " + "第二阶段开始" + System.Environment.NewLine;
                    myQC.textBox1.ScrollToEnd();
                }
                else if (ReceivedData[1] == 0X02 && ReceivedData[2] == 0X01)
                {
                    myQC.textBox1.Text += "[" + time1 + "]  " + "第二阶段结束" + System.Environment.NewLine;
                    myQC.textBox1.ScrollToEnd();
                }
                else if (ReceivedData[1] == 0X03 && ReceivedData[2] == 0X00)
                {
                    myQC.textBox1.Text += "[" + time1 + "]  " + "第三阶段开始" + System.Environment.NewLine;
                    myQC.textBox1.ScrollToEnd();
                }
                else if (ReceivedData[1] == 0X03 && ReceivedData[2] == 0X01)
                {
                    myQC.textBox1.Text += "[" + time1 + "]  " + "第三阶段结束" + System.Environment.NewLine;
                    myQC.textBox1.ScrollToEnd();
                }
                else if (ReceivedData[1] == 0X04 && ReceivedData[2] == 0X00)
                {
                    myQC.result.Text = "通过";
                    string record = DateTime.Now.ToString() + "," + myQC.precision.Text.Trim() + "," + myQC.accuracy.Text.Trim() + "," + myQC.result.Text.Trim();
                    myQC.QCSave(record);

                    myQC.textBox1.Text += "[" + time1 + "]  " + "质控通过" + System.Environment.NewLine;
                    myQC.textBox1.ScrollToEnd();
                }
                else if (ReceivedData[1] == 0X04 && ReceivedData[2] == 0X01)
                {
                    myQC.result.Text = "通过*";
                    string record = DateTime.Now.ToString() + "," + myQC.precision.Text.Trim() + "," + myQC.accuracy.Text.Trim() + "," + myQC.result.Text.Trim();
                    myQC.QCSave(record);

                    myQC.textBox1.Text += "[" + time1 + "]  " + "质控通过*" + System.Environment.NewLine;
                    myQC.textBox1.ScrollToEnd();
                }
                else if (ReceivedData[1] == 0X05 && ReceivedData[2] == 0X00)
                {
                    myQC.textBox1.Text += "[" + time1 + "]  " + "质控返回待机" + System.Environment.NewLine;

                    myQC.textBox1.ScrollToEnd();

                    //一次质控完成，回复待机界面
                    QCReset();

                }
                
            
            }
            else if (ReceivedData[0] == 0XE0)
            {
                if (ReceivedData[1] == 0X02 && ReceivedData[2] == 0X01 && ReceivedData[3] == 0X00 && ReceivedData[4] == 0X00)
                    sex.Text = "女";
                else if (ReceivedData[1] == 0X02 && ReceivedData[2] == 0X01 && ReceivedData[3] == 0X00 && ReceivedData[4] == 0X01)
                    sex.Text = "男";
            }

            //把提示框拉倒最后一行
            this.receiveInfo.ScrollToEnd();
        
        }
        //质控返回初始界面
        private void QCReset()
        {
            myQC.co.Text = null; myQC.co2.Text = null;

            myQC.precision.Text = null; myQC.accuracy.Text = null; myQC.result.Text = null;

            myQC.textBox1.Text = null;

            myQC.button1.IsEnabled = true;

            qcSign = false;
            qcStep = 0;
            qcOpend = false;
            sn = false;
        }

        //对主界面当日报告进行刷新，把患者信息复位
        private void Reflash()
        {            
            name.Text = "";
            age.Text = "";
            sex.Text = "男";
            id.Text = "";
            rbConcentration.Text = "0"; //血红蛋白默认值为0
            sendDoctor.Text = "";
            firstCheck.Text = "";
            receiveInfo.Text = "";
            day.Text = "";
            CO.Text = "";
            PCO2.Text = "";
            measure.IsEnabled = true;
        
        }
        private void sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //System.Threading.Thread.Sleep(150);//延时100ms等待接收完数据

            //串口打开标志位为false，则不处理串口事件
            if (spOpenSign == false)
                return;

            //this.Invoke就是跨线程访问ui的方法
            this.Dispatcher.Invoke(new Action(() =>
            {   //委托操作GUI控件的部分

                int n = sp.BytesToRead;                       //buffer

                Byte[] ReceivedData = new Byte[6];//创建接收字节数组
                string RecvDataText=null;
                //sp.Read(ReceivedData, 0, ReceivedData.Length);//读取所接收到的数据                  //buffer


                //sp.DiscardInBuffer();//丢弃接收缓冲区数据                    //buffer
                //sp.DiscardOutBuffer();//清空发送缓冲区数据                  //buffer

                byte[] buf = new byte[n];                             //buffer
                sp.Read(buf, 0, n);                             //buffer
                buffer.AddRange(buf);                         //buffer


                while (buffer.Count>0) ///*buffer
                {                  
                    //receiveInfo.Text += buffer.Count + System.Environment.NewLine;
                    try
                    {
                        if (buffer[0] != 0X80 && buffer[0] != 0X90 && buffer[0] != 0XC0 && buffer[0] != 0X0D && buffer[0] != 0XD0 && buffer[0] != 0XB0 && buffer[0] != 0XE0 && buffer[0] != 0XA0 && buffer[0] != 0XCC && buffer[0] != 0XAA && buffer[0] != 0XFF && buffer[0] != 0X00)
                        {
                            buffer.RemoveRange(0, 1);
                            //break;
                            continue;
                        }
                        if (buffer[0] == 0X80 || buffer[0] == 0X90 || buffer[0] == 0XC0||buffer[0]==0X0D|| buffer[0] == 0XD0|| buffer[0] == 0XB0|| buffer[0] == 0XE0||buffer[0]==0XA0)
                        {
                            if (buffer.Count < 6)
                            {
                                //receiveInfo.Text += buffer.ToString();
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR4:" + ex.Message);
                    }
                    //获取接收数据时的系统时间
                    DateTime dt1 = System.DateTime.Now;
                    string date1 = dt1.ToLocalTime().ToString();
                    string time1 = dt1.ToString("HH:mm:ss");
                    //把接收到的数据写进日志中
                    string recv = null;
                    string buff = null;

                    for (int i = 0; i < buffer.Count; i++)
                        buff += (buffer[i].ToString("X2"));

                    WriteLog("[" + date1 + "]" + ":" + buff);

                    //if (String.Compare(buff, "00800001000081") == 0)
                    //{
                    //    Byte[] temp = new Byte[1];
                    //    string date3 = dt1.ToLocalTime().ToString();
                    //    temp[0] = 0X00;

                    //    receiveInfo.Text += "[" + time1 + "]:" + "肺泡气袋插入" + System.Environment.NewLine;
                    //    light1.Source = new BitmapImage(new Uri("/Seekya;component/Images/lightOn1.jpg", UriKind.Relative));
                    //    l1 = true;

                    //    //开始测量
                    //    measure.IsEnabled = false;

                    //    //写日志
                    //    WriteLog("[" + date3 + "]" + ":" + "00");

                    //    sp.Write(temp, 0, 1);
                    //    buffer.RemoveRange(0, n);


                    //}

                    if (String.Compare(buffer[0].ToString("X2"), "AA") == 0)//当上位机接收到仪器发送过来的0XAA，则返回0XBB,以表示同意接收
                    {
                        Byte[] temp = new Byte[1];
                        string date3 = dt1.ToLocalTime().ToString();
                        temp[0] = 0XBB;

                        //写日志
                        WriteLog("[" + date3 + "]" + ":" + "BB");

                        /*
                        if (firstConn == false)
                        {
                            receiveInfo.Text += "[" + time1 + "]:" + "联机成功" + System.Environment.NewLine;
                            firstConn = true;
                        }
                        */
                        //receiveInfo.Text += "hellowworld" + System.Environment.NewLine;
                        sp.Write(temp, 0, 1);
                        buffer.RemoveRange(0, 1);


                    }
                    else if (String.Compare(buffer[0].ToString("X2"), "CC") == 0)
                    {
                        receiveInfo.Text += "[" + time1 + "]:" + "联机成功" + System.Environment.NewLine;
                        buffer.RemoveRange(0, 1);


                    }
                    else if (String.Compare(buffer[0].ToString("X2"), "FF") == 0)//下位机接收失败
                    {

                        if (qcSign == true) //处于质控阶段
                        {
                            Byte[] temp = new Byte[6];

                            switch (qcStep)
                            {
                                case 0: temp[5] = 0X21; temp[4] = 0X00; temp[3] = 0X00; temp[2] = 0X00; temp[1] = 0X01; temp[0] = 0X20; sp.Write(temp, 0, 6);buffer.RemoveRange(0, 1); break;
                                case 1: temp[0] = 0XE0; temp[1] = 0X03; temp[2] = 0X00; temp[3] = (Byte)(myQC.GetCO() / 256); temp[4] = (Byte)(myQC.GetCO() % 256); temp[5] = (Byte)(temp[0] + temp[1] + temp[3] + temp[4]); sp.Write(temp, 0, 6);buffer.RemoveRange(0,1); break;
                                case 2: temp[0] = 0XE0; temp[1] = 0X04; temp[2] = 0X00; temp[3] = (Byte)(myQC.GetCO2() / 256); temp[4] = (Byte)(myQC.GetCO2() % 256); temp[5] = (Byte)(temp[0] + temp[1] + temp[3] + temp[4]); sp.Write(temp, 0, 6);buffer.RemoveRange(0, 1); break;

                            }

                        }
                        else
                        {
                            Byte[] temp = new Byte[6];
                            //获取接收数据时的系统时间
                            DateTime dt2 = System.DateTime.Now;
                            string date3 = dt1.ToLocalTime().ToString();

                            switch (measureStep)
                            {
                                case 0: temp[5] = 0X20; temp[4] = 0X00; temp[3] = 0X00; temp[2] = 0X00; temp[1] = 0X00; temp[0] = 0X20; WriteLog("[" + date3 + "]" + ":" + "200000000020"); sp.Write(temp, 0, 6);buffer.RemoveRange(0, 1); break;  //重发开始测量指令
                                case 1:
                                    temp[0] = 0XE0; temp[1] = 0X00; temp[2] = 0X00;  //重新发送血红蛋白浓度
                                    if (rbConcentration.Text.Trim().Length == 0)
                                    {

                                        temp[3] = 0; temp[4] = 0; temp[5] = 0XE0;

                                        WriteLog("[" + date3 + "]" + ":" + "E000000000E0");
                                        sp.Write(temp, 0, 6);
                                        buffer.RemoveRange(0, 1);

                                    }
                                    else
                                    {
                                        int rb = Convert.ToInt16(rbConcentration.Text.Trim());

                                        temp[3] = (Byte)(rb / 256); temp[4] = (Byte)(rb % 256); temp[5] = (Byte)(temp[0] + temp[3] + temp[4]);

                                        WriteLog("[" + date3 + "]" + ":" + Convert.ToString(temp));
                                        sp.Write(temp, 0, 6);
                                        buffer.RemoveRange(0, 1);


                                    }
                                    break;
                                case 2:
                                    temp[0] = 0XE0; temp[1] = 0X02; temp[2] = 0X01; temp[3] = 0X00;
                                    if (String.Compare(sex.Text.Trim(), "男") == 0)
                                        temp[4] = 0X01;
                                    else
                                        temp[4] = 0X00;

                                    temp[5] = (byte)(temp[0] + temp[1] + temp[2] + temp[4]);

                                    WriteLog("[" + date3 + "]" + ":" + Convert.ToString(temp));
                                    sp.Write(temp, 0, 6);
                                    buffer.RemoveRange(0, 1);

                                    break;

                            }
                        }
                    }
                    else if (String.Compare(buffer[0].ToString("X2"), "00") == 0)//下位机接收成功
                    {
                        //receiveInfo.Text += 00 + System.Environment.NewLine;

                        //质控时，接收到00
                        if (qcSign == true)
                        {
                            Byte[] temp = new Byte[6];

                            switch (qcStep)
                            {
                                case 0:buffer.RemoveRange(0, 1); break;
                                case 1: qcStep++; temp[0] = 0XE0; temp[1] = 0X04; temp[2] = 0X00; temp[3] = (Byte)(myQC.GetCO2() / 256); temp[4] = (Byte)(myQC.GetCO2() % 256); temp[5] = (Byte)(temp[0] + temp[1] + temp[3] + temp[4]); sp.Write(temp, 0, 6);buffer.RemoveRange(0, 1); break;
                                case 2: qcStep = 0;buffer.RemoveRange(0, 1); break;

                            }
                        }
                        else
                        {
                            Byte[] temp = new Byte[6];
                            //获取接收数据时的系统时间
                            DateTime dt2 = System.DateTime.Now;
                            string date3 = dt1.ToLocalTime().ToString();

                            switch (measureStep)
                            {
                                case 0: measure.IsEnabled = false;buffer.RemoveRange(0, 1); break;
                                case 1:
                                    measureStep++; temp[0] = 0XE0; temp[1] = 0X02; temp[2] = 0X01; temp[3] = 0X00;
                                    if (String.Compare(sex.Text.Trim(), "男") == 0)
                                    {
                                        temp[4] = 0X01;
                                        temp[5] = (byte)(temp[0] + temp[1] + temp[2] + temp[4]);

                                    }
                                    else
                                    {
                                        temp[4] = 0X00;
                                        temp[5] = (byte)(temp[0] + temp[1] + temp[2] + temp[4]);

                                    }
                                    WriteLog("[" + date3 + "]" + ":" + Convert.ToString(temp));
                                    sp.Write(temp, 0, 6);
                                    buffer.RemoveRange(0, 1);

                                    break;
                                case 2: measureStep = 0;buffer.RemoveRange(0, 1); break;

                            }
                        }


                    }
                    else //接收到协议中不同命令时的处理
                    {

                        //if (buffer.Count == 6)
                        if (buffer.Count>=6)
                        {
                            buffer.CopyTo(0, ReceivedData, 0, 6);
                            if (buffer[5] != CheckSum(ReceivedData))
                            {
                                string date2 = dt1.ToLocalTime().ToString();
                                Byte[] temp3 = new Byte[1];
                                temp3[0] = 0XFF;

                                //写日志
                                WriteLog("[" + date2 + "]" + " " + "FF");

                                sp.Write(temp3, 0, 1);//(temp, 0, 1);
                                buffer.RemoveRange(0, 6);
                                MessageBox.Show("数据包不正确！");
                                continue;
                            }
                            else
                            {
                                string date2 = dt1.ToLocalTime().ToString();

                                Byte[] temp3 = new Byte[1];
                                temp3[0] = 0X00;

                                sp.Write(temp3, 0, 1);//(temp, 0, 1);

                                //写日志
                                WriteLog("[" + date2 + "]" + ":" + "00");
                                buffer.RemoveRange(0, 6);

                                if (String.Compare(buff, "800400000084") == 0)    //气袋全部插入
                                {
                                    Byte[] temp = new Byte[6];
                                    //获取接收数据时的系统时间
                                    DateTime dt2 = System.DateTime.Now;
                                    string date3 = dt1.ToLocalTime().ToString();

                                    Thread.Sleep(500);    //休眠100ms     //.....500ms

                                    //if (qcSign == true)     //重新发送测试气A的CO浓度差值
                                    //{
                                    //    temp[0] = 0XE0; temp[1] = 0X03; temp[2] = 0X00; temp[3] = (Byte)(myQC.GetCO() / 256); temp[4] = (Byte)(myQC.GetCO() % 256); temp[5] = (Byte)(temp[0] + temp[1] + temp[3] + temp[4]);

                                    //    sp.Write(temp, 0, 6);

                                    //    qcStep++;
                                    //}
                                    //else
                                    {
                                        //使“测量键”无效
                                        measure.IsEnabled = false;

                                        temp[0] = 0XE0; temp[1] = 0X00; temp[2] = 0X00;  //重新发送血红蛋白浓度
                                        if (rbConcentration.Text.Trim().Length == 0)
                                        {

                                            temp[3] = 0X00; temp[4] = 0; temp[5] = 0XE0;

                                            WriteLog("[" + date3 + "]" + ":" + "E000000000E0");
                                            sp.Write(temp, 0, 6);

                                        }
                                        else
                                        {
                                            int rb = Convert.ToInt16(rbConcentration.Text.Trim());

                                            temp[3] = (Byte)(rb / 256); temp[4] = (Byte)(rb % 256); temp[5] = (Byte)(temp[0] + temp[3] + temp[4]);

                                            WriteLog("[" + date3 + "]" + ":" + Convert.ToString(temp));
                                            sp.Write(temp, 0, 6);

                                        }
                                        measureStep++;
                                    }
                                }
                                else if ((String.Compare(buff, "800401000085") == 0))
                                {
                                    MessageBox.Show("气袋未插到位", "提示");

                                }
                                else if (string.Compare(buff,"800402000086")==0)
                                {
                                    Byte[] temp = new Byte[6];
                                    //获取接收数据时的系统时间
                                    DateTime dt2 = System.DateTime.Now;
                                    string date3 = dt1.ToLocalTime().ToString();

                                    Thread.Sleep(500);
                                    //if (qcSign == true)     //重新发送测试气A的CO浓度差值
                                    //{
                                    //    temp[0] = 0XE0; temp[1] = 0X03; temp[2] = 0X00; temp[3] = (Byte)(myQC.GetCO() / 256); temp[4] = (Byte)(myQC.GetCO() % 256); temp[5] = (Byte)(temp[0] + temp[1] + temp[3] + temp[4]);

                                    //    sp.Write(temp, 0, 6);

                                    //    qcStep++;
                                    //}
                                    //if (softwareOperate==false)
                                    //{
                                    //    myQC = new QC(this);

                                    //}
                                    if (sn==false)
                                    {
                                        qcOpend = true;

                                        if (qcOpen == false)
                                        {
                                            qcDialogShow();
                                        }
                                        else if (myQC.WindowState == WindowState.Minimized)
                                        {
                                            myQC.WindowState = WindowState.Normal;
                                            //qcOpend = true;
                                        }
                                        else
                                        {
                                            //qcOpend = true;
                                        }

                                    }
                                    else
                                    {
                                        if (qcOpen==false)
                                        {
                                            qcDialogShow();
                                        }
                                        else if (myQC.WindowState == WindowState.Minimized)
                                        {
                                            myQC.WindowState = WindowState.Normal;
                                            temp[0] = 0XE0; temp[1] = 0X03; temp[2] = 0X00; temp[3] = (Byte)(myQC.GetCO() / 256); temp[4] = (Byte)(myQC.GetCO() % 256); temp[5] = (Byte)(temp[0] + temp[1] + temp[3] + temp[4]);

                                            sp.Write(temp, 0, 6);

                                            qcStep++;
                                        }
                                        else
                                        {
                                            temp[0] = 0XE0; temp[1] = 0X03; temp[2] = 0X00; temp[3] = (Byte)(myQC.GetCO() / 256); temp[4] = (Byte)(myQC.GetCO() % 256); temp[5] = (Byte)(temp[0] + temp[1] + temp[3] + temp[4]);

                                            sp.Write(temp, 0, 6);

                                            qcStep++;
                                        }


                                        myQC.Activate();


                                    }
                                }
                                else
                                    ShowTip(ReceivedData);
                            }


                            //checkSum = CheckSum(ReceivedData);//计算检验和
                            //string date2 = dt1.ToLocalTime().ToString();

                            
                           

                        }

                    }
                    //buffer.CopyTo(0, ReceivedData, 0, n);
                    //buffer.RemoveRange(0, n);

                }                                                                           //buffer*/ 

                //获取接收数据时的系统时间
                //DateTime dt1 = System.DateTime.Now;
                //string date1 = dt1.ToLocalTime().ToString();
                //string time1 = dt1.ToString("HH:mm:ss");

                //把接收到的数据写进日志中
                //string recv = null;

                //for (int i = 0; i < ReceivedData.Length; i++)
                //    recv += (ReceivedData[i].ToString("X2"));

                //WriteLog("[" + date1 + "]" + ":" + recv);

                
 

                    //接受到错误代码00800001000081，回复00，显示灯1亮
                    //if (String.Compare(RecvDataText, "00800001000081") == 0)
                    //{
                    //    Byte[] temp = new Byte[1];
                    //    string date3 = dt1.ToLocalTime().ToString();
                    //    temp[0] = 0X00;

                    //    receiveInfo.Text += "[" + time1 + "]:" + "肺泡气袋插入" + System.Environment.NewLine; 
                    //    light1.Source = new BitmapImage(new Uri("/Seekya;component/Images/lightOn1.jpg", UriKind.Relative)); 
                    //    l1 = true;

                    //    //开始测量
                    //    measure.IsEnabled = false;

                    //    //写日志
                    //    WriteLog("[" + date3 + "]" + ":" + "00");

                    //    sp.Write(temp, 0, 1);
 
                    //}
                    //else if (String.Compare(RecvDataText, "AA") == 0)//当上位机接收到仪器发送过来的0XAA，则返回0XBB,以表示同意接收
                    //{
                    //    Byte[] temp = new Byte[1];
                    //    string date3 = dt1.ToLocalTime().ToString();
                    //    temp[0] = 0XBB;

                    //    //写日志
                    //    WriteLog("[" + date3 + "]" + ":" + "BB");

                    //    /*
                    //    if (firstConn == false)
                    //    {
                    //        receiveInfo.Text += "[" + time1 + "]:" + "联机成功" + System.Environment.NewLine;
                    //        firstConn = true;
                    //    }
                    //    */

                    //    sp.Write(temp, 0, 1);

                    //}
                    //else if (String.Compare(RecvDataText, "CC") == 0)
                    //{
                    //    receiveInfo.Text += "[" + time1 + "]:" + "联机成功" + System.Environment.NewLine;

                    //}
                    //else if (String.Compare(RecvDataText, "FF") == 0)//下位机接收失败
                    //{

                    //    if (qcSign == true) //处于质控阶段
                    //    {
                    //        Byte[] temp = new Byte[6];

                    //        switch (qcStep)
                    //        {
                    //            case 0: temp[5] = 0X21; temp[4] = 0X00; temp[3] = 0X00; temp[2] = 0X00; temp[1] = 0X01; temp[0] = 0X20; sp.Write(temp, 0, 6); break;
                    //            case 1: temp[0] = 0XE0; temp[1] = 0X03; temp[2] = 0X00; temp[3] = (Byte)(myQC.GetCO() / 256); temp[4] = (Byte)(myQC.GetCO() % 256); temp[5] = (Byte)(temp[0] + temp[1] + temp[3] + temp[4]); sp.Write(temp, 0, 6); break;
                    //            case 2: temp[0] = 0XE0; temp[1] = 0X04; temp[2] = 0X00; temp[3] = (Byte)(myQC.GetCO2() / 256); temp[4] = (Byte)(myQC.GetCO2() % 256); temp[5] = (Byte)(temp[0] + temp[1] + temp[3] + temp[4]); sp.Write(temp, 0, 6); break;

                    //        }

                    //    }
                    //    else
                    //    {
                    //        Byte[] temp = new Byte[6];
                    //        //获取接收数据时的系统时间
                    //        DateTime dt2 = System.DateTime.Now;
                    //        string date3 = dt1.ToLocalTime().ToString();

                    //        switch (measureStep)
                    //        {
                    //            case 0: temp[5] = 0X20; temp[4] = 0X00; temp[3] = 0X00; temp[2] = 0X00; temp[1] = 0X00; temp[0] = 0X20; WriteLog("[" + date3 + "]" + ":" + "200000000020"); sp.Write(temp, 0, 6); break;  //重发开始测量指令
                    //            case 1: temp[0] = 0XE0; temp[1] = 0X00; temp[2] = 0X00;  //重新发送血红蛋白浓度
                    //                if (rbConcentration.Text.Trim().Length == 0)
                    //                {

                    //                    temp[3] = 0; temp[4] = 0; temp[5] = 0XE0;

                    //                    WriteLog("[" + date3 + "]" + ":" + "E000000000E0");
                    //                    sp.Write(temp, 0, 6);

                    //                }
                    //                else
                    //                {
                    //                    int rb = Convert.ToInt16(rbConcentration.Text.Trim());

                    //                    temp[3] = (Byte)(rb / 256); temp[4] = (Byte)(rb % 256); temp[5] = (Byte)(temp[0] + temp[3] + temp[4]);

                    //                    WriteLog("[" + date3 + "]" + ":" + Convert.ToString(temp));
                    //                    sp.Write(temp, 0, 6);


                    //                }
                    //                break;
                    //            case 2: temp[0] = 0XE0; temp[1] = 0X02; temp[2] = 0X01; temp[3] = 0X00;
                    //                if (String.Compare(sex.Text.Trim(), "男") == 0)
                    //                    temp[4] = 0X01;
                    //                else
                    //                    temp[4] = 0X00;

                    //                temp[5] = (byte)(temp[0] + temp[1] + temp[2] + temp[4]);

                    //                WriteLog("[" + date3 + "]" + ":" + Convert.ToString(temp));
                    //                sp.Write(temp, 0, 6);

                    //                break;

                    //        }
                    //    }
                    //}
                    //else if (String.Compare(RecvDataText, "00") == 0)//下位机接收成功
                    //{

                    //    //质控时，接收到00
                    //    if (qcSign == true)
                    //    {
                    //        Byte[] temp = new Byte[6];

                    //        switch (qcStep)
                    //        {
                    //            case 0: break;
                    //            case 1: qcStep++; temp[0] = 0XE0; temp[1] = 0X04; temp[2] = 0X00; temp[3] = (Byte)(myQC.GetCO2() / 256); temp[4] = (Byte)(myQC.GetCO2() % 256); temp[5] = (Byte)(temp[0] + temp[1] + temp[3] + temp[4]); sp.Write(temp, 0, 6); break;
                    //            case 2: qcStep = 0; break;

                    //        }
                    //    }
                    //    else
                    //    {
                    //        Byte[] temp = new Byte[6];
                    //        //获取接收数据时的系统时间
                    //        DateTime dt2 = System.DateTime.Now;
                    //        string date3 = dt1.ToLocalTime().ToString();

                    //        switch (measureStep)
                    //        {
                    //            case 0: measure.IsEnabled = false; break;
                    //            case 1: measureStep++; temp[0] = 0XE0; temp[1] = 0X02; temp[2] = 0X01; temp[3] = 0X00;
                    //                if (String.Compare(sex.Text.Trim(), "男") == 0)
                    //                {
                    //                    temp[4] = 0X01;
                    //                    temp[5] = (byte)(temp[0] + temp[1] + temp[2] + temp[4]);

                    //                }
                    //                else
                    //                {
                    //                    temp[4] = 0X00;
                    //                    temp[5] = (byte)(temp[0] + temp[1] + temp[2] + temp[4]);

                    //                }
                    //                WriteLog("[" + date3 + "]" + ":" + Convert.ToString(temp));
                    //                sp.Write(temp, 0, 6);

                    //                break;
                    //            case 2: measureStep = 0; break;

                    //        }
                    //    }


                    //}
                    //else //接收到协议中不同命令时的处理
                    //{
                    //    Byte checkSum = 0;

                    //    if (ReceivedData.Length == 6)
                    //    {
                    //        checkSum = CheckSum(ReceivedData);//计算检验和
                    //        string date2 = dt1.ToLocalTime().ToString();

                    //        if (ReceivedData[5] == checkSum)//检验和成功
                    //        {
                    //            Byte[] temp3 = new Byte[1];
                    //            temp3[0] = 0X00;

                    //            sp.Write(temp3, 0, 1);//(temp, 0, 1);

                    //            //写日志
                    //            WriteLog("[" + date2 + "]" + ":" + "00");

                    //            if (String.Compare(RecvDataText, "800400000084") == 0)    //气袋全部插入
                    //            {
                    //                Byte[] temp = new Byte[6];
                    //                //获取接收数据时的系统时间
                    //                DateTime dt2 = System.DateTime.Now;
                    //                string date3 = dt1.ToLocalTime().ToString();

                    //                Thread.Sleep(500);    //休眠100ms     //.....500ms

                    //                if (qcSign == true)     //重新发送测试气A的CO浓度差值
                    //                {
                    //                    temp[0] = 0XE0; temp[1] = 0X03; temp[2] = 0X00; temp[3] = (Byte)(myQC.GetCO() / 256); temp[4] = (Byte)(myQC.GetCO() % 256); temp[5] = (Byte)(temp[0] + temp[1] + temp[3] + temp[4]);

                    //                    sp.Write(temp, 0, 6);

                    //                    qcStep++;
                    //                }
                    //                else
                    //                {
                    //                    //使“测量键”无效
                    //                    measure.IsEnabled = false;

                    //                    temp[0] = 0XE0; temp[1] = 0X00; temp[2] = 0X00;  //重新发送血红蛋白浓度
                    //                    if (rbConcentration.Text.Trim().Length == 0)
                    //                    {

                    //                        temp[3] = 0X00; temp[4] = 0; temp[5] = 0XE0;

                    //                        WriteLog("[" + date3 + "]" + ":" + "E000000000E0");
                    //                        sp.Write(temp, 0, 6);

                    //                    }
                    //                    else
                    //                    {
                    //                        int rb = Convert.ToInt16(rbConcentration.Text.Trim());

                    //                        temp[3] = (Byte)(rb / 256); temp[4] = (Byte)(rb % 256); temp[5] = (Byte)(temp[0] + temp[3] + temp[4]);

                    //                        WriteLog("[" + date3 + "]" + ":" + Convert.ToString(temp));
                    //                        sp.Write(temp, 0, 6);

                    //                    }
                    //                    measureStep++;
                    //                }
                    //            }
                    //            else if ((String.Compare(RecvDataText, "800401000085") == 0))
                    //            {
                    //                MessageBox.Show("气袋未插到位", "提示");

                    //            }
                    //            else
                    //                ShowTip(ReceivedData);

                    //        }
                    //        else
                    //        {
                    //            Byte[] temp3 = new Byte[1];
                    //            temp3[0] = 0XFF;

                    //            //写日志
                    //            WriteLog("[" + date2 + "]" + " " + "FF");

                    //            sp.Write(temp3, 0, 1);//(temp, 0, 1);

                    //        }

                    //    }

                    //}
                //tBoxDataReceive.Text+=RecvDataText;
                //receiveInfo.Text+=System.Environment.NewLine;//Windows下换行用“\r\n”,Linux下换行用“\n”,“System.Environment.NewLine”都适用

            }));
         
        }
        private void ShowZeroFault()
        {
            MessageBox.Show("测量未完成(Zero Fault)，拔掉所有气袋，仪器返回待机界面。", "报错");

            System.Windows.Threading.Dispatcher.Run();//如果去掉这个，会发现启动的窗口显示出来以后会很快就关掉。
        
        }
        private void ShowTestFault()
        {
            MessageBox.Show("测量未完成(Test Fault)，拔掉所有气袋，仪器返回待机界面。", "报错");

            System.Windows.Threading.Dispatcher.Run();
        }
        private void ShowSampleFault()
        {
            MessageBox.Show("Sample Fault，拔掉所有气袋，仪器返回待机界面。", "报错");

            System.Windows.Threading.Dispatcher.Run();
        }
        private void ShowZeroOversizeFault()
        {
            MessageBox.Show("问题提示：测试过程受干扰，该测试结果可能存在异常风险，请将测试结果反馈给授权经销商或生产产家。", "提示");

            System.Windows.Threading.Dispatcher.Run();
        }
        private void ShowCO2LowFault()
        {
            MessageBox.Show("问题提示：样本采集过程中混入了较多的空气，请规范采样。", "提示");

            System.Windows.Threading.Dispatcher.Run();
        }

        //监听设备串口，并连上设备
        private void ListenCom()
        {
            bool discon = false;
            while (true)
            {
                if (IsSeekyaRBCSConn())
                {
                    DateTime dt = System.DateTime.Now;
                    string date = dt.ToLocalTime().ToString();

                    //没设置串口号，什么也不做
                    if ((com2 = GetCom()) == null)
                    {
                        //do nothing
                    }
                    else
                    {
                        if (String.Compare(com1, com2) != 0)
                        {
                            if (com1 != null)
                            {
                                try
                                {
                                    //断开串口com1
                                    sp.Close();
                                    sp.Dispose();

                                    spOpenSign = false;//把串口打开标志位设置为false

                                }
                                catch (Exception)
                                {

                                }

                            }
                            SetPortProperty();//设置串口属性

                            try//打开串口
                            {
                                if (!sp.IsOpen)
                                {
                                    sp.Open();
                                }
                                //给下位机发送DD
                                Byte[] temp = new Byte[1];
                                temp[0] = 0XDD;

                                //写日志
                                WriteLog("[" + date + "]" + ":" + "DD");

                                sp.Write(temp, 0, 1);//(temp, 0, 1);

                                spOpenSign = true;//把串口打开标志位设置为true
                                com1 = com2;
                                discon = false;

                            }
                            catch (Exception)
                            {
                                this.receiveInfo.Dispatcher.Invoke(new Action(() => {
                                    this.receiveInfo.Text += "串口无效或已被占用，连接失败！" + System.Environment.NewLine;
                                    this.receiveInfo.ScrollToEnd();
                                }));
                                //打开串口失败后，相应标志位取消
                                //MessageBox.Show("串口无效或已被占用，连接仪器失败", "错误提示");
                            }
                        }
                    }
                }
                else
                {
                    //当前有连接上串口，就断开串口，com1置null
                    if (com1 != null)
                    {
                        try
                        {
                            DateTime dt = System.DateTime.Now;
                            string time1 = dt.ToString("HH:mm:ss");

                            //把占用串口断开
                            sp.Close();
                            sp.Dispose();

                            com1 = null;

                            //在提示框，提示串口已断开
                            this.receiveInfo.Dispatcher.Invoke(new Action(()=>{this.receiveInfo.Text += "[" + time1 + "]:" + "串口断开" + System.Environment.NewLine; this.receiveInfo.ScrollToEnd(); }));
                            discon = true;

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("ERROR5:" + ex.Message);
                        
                        }
                    
                    }
                
                }
                try
                {
                    if (!sp.IsOpen) //串口关闭
                    {
                        //this.receiveInfo.Dispatcher.Invoke(new Action(() => {
                        //    this.receiveInfo.Text += "sp is not open" + System.Environment.NewLine;
                        //    this.receiveInfo.ScrollToEnd();
                        //}));
                        try
                        {
                            //sp.Close();//close可以执行，open不能
                            sp.Open();
                            if (sp.IsOpen)
                            {
                                //给下位机发送DD
                                Byte[] temp = new Byte[1];
                                temp[0] = 0XDD;

                                //写日志
                                DateTime dt = System.DateTime.Now;
                                string date = dt.ToLocalTime().ToString();
                                WriteLog("[" + date + "]" + ":" + "DD");
                                //con = true;
                                sp.Write(temp, 0, 1);//(temp, 0, 1);

                                spOpenSign = true;//把串口打开标志位设置为true
                                //com1 = com2;
                                discon = false;
                            }
                            //this.receiveInfo.Dispatcher.Invoke(new Action(() => {
                            //    this.receiveInfo.Text += "sp is open" + System.Environment.NewLine;
                            //    this.receiveInfo.ScrollToEnd();
                            //}));
                        }
                        catch (Exception)
                        {
                            if (discon == false)
                            {
                                DateTime dt = System.DateTime.Now;
                                string time1 = dt.ToString("HH:mm:ss");
                                this.receiveInfo.Dispatcher.Invoke(new Action(() => {
                                    this.receiveInfo.Text += "[" + time1 + "]:" + "串口断开" + System.Environment.NewLine;
                                    this.receiveInfo.ScrollToEnd();
                                }));
                                discon = true;
                            }
                        }
                    }
                }
                catch (Exception)
                {

                }

                //等待2秒钟,继续连串口
                Thread.Sleep(2000);

            }
        }

        //把接收到的数据写进日志中
        public void WriteLog(string str)
        {
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\log.txt";//读取日志的txt文件

            try
            {
                FileStream fs1 = new FileStream(pathString, FileMode.Append, FileAccess.Write);
                StreamWriter sw1 = new StreamWriter(fs1);

                sw1.WriteLine(str);

                sw1.Close();
                fs1.Close();

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error:" + ex.Message);
            }
        }
        //把零点写进日志中
        private void WriteZero(string str)
        {
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\zero.txt";//读取日志的txt文件

            try
            {
                FileStream fs1 = new FileStream(pathString, FileMode.Append, FileAccess.Write);
                StreamWriter sw1 = new StreamWriter(fs1);

                sw1.WriteLine(str);

                sw1.Close();
                fs1.Close();

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error:" + ex.Message);
            }
        }

        //判断仪器是否串口连接电脑，连接了，返回true，否则，返回false
        private bool IsSeekyaRBCSConn()
        {
            string[] ports = SerialPort.GetPortNames();
            string comTmp = GetCom();

            if (comTmp == null)
            {
                comTmp = "COM";

            }

            foreach (string port in ports)
            {
                if (String.Compare(comTmp, port) == 0)
                {
                    return true;
                }

            }

            return false;
        
        }

        public void Open(string FileName)
        {
            app = new Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(FileName);
            //wb = wbs.Open(FileName,  0, true, 5,"", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true,Type.Missing,Type.Missing);
            //wb = wbs.Open(FileName,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Excel.XlPlatform.xlWindows,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing);
        }

    }
}

