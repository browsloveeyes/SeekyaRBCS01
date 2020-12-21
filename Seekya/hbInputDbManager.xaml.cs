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
    /// hbInputDbManager.xaml 的交互逻辑
    /// </summary>
    public partial class hbInputDbManager : Window
    {
        public hbInputDbManager()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            dbManager f = (dbManager)this.Owner;

            f.rbcon1 = tboxhb.Text.Trim();

            this.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }
    }
}
