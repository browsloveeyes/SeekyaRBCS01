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
using System.Threading;

namespace Seekya
{
    /// <summary>
    /// HL7.xaml 的交互逻辑
    /// </summary>
    public partial class HL7 : Window
    {
        private MainWindow my = null;

        public HL7(MainWindow mainWin)
        {
            InitializeComponent();

            my = mainWin;

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnSend.Focus();
            
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            textBox1.Focus();

        }

        //条形码发送
        private void btnSend_Click(object sender, RoutedEventArgs e)
        {
            Thread qryAA = new Thread(my.QRYAA);
            
            my.SerialNumber=textBox1.Text;

            qryAA.Start();

        }

        //连接远程服务器
        private void connServer_Click(object sender, RoutedEventArgs e)
        {
            //Thread conn = new Thread(new ThreadStart(my.ConnectToServer));

            my.ServerIP = serverIP.Text;
            my.ServerPort = serverPort.Text;

            my.ConnectToServer();
            //conn.Start();

        }

        //断开远程服务器
        private void disconnServer_Click(object sender, RoutedEventArgs e)
        {
            //Thread disconn = new Thread(my.DisconnectToServer);

            //disconn.Start();

            my.DisconnectToServer();
        }

    }
}
