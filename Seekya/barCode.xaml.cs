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
    /// barCode.xaml 的交互逻辑
    /// </summary>
    public partial class barCode : Window
    {
        MainWindow f1 = null;

        public barCode(MainWindow f)
        {
            InitializeComponent();

            f1 = f;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            tBoxBarCode.Focus();

        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Thread qryAA = new Thread(f1.QRYAA);

            f1.SerialNumber = tBoxBarCode.Text;

            qryAA.Start();

            //关闭扫描条形码窗口
            this.DialogResult = true;

        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }
    }
}
