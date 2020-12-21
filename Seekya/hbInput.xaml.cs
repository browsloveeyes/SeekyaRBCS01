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
    /// hbInput.xaml 的交互逻辑
    /// </summary>
    public partial class hbInput : Window
    {
        public hbInput()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MainWindow f = (MainWindow)this.Owner;

            f.rbcon = tboxhb.Text.Trim();

            this.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }
    }
}
