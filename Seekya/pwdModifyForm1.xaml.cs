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
    /// pwdModifyForm1.xaml 的交互逻辑
    /// </summary>
    public partial class pwdModifyForm1 : Window
    {
        public pwdModifyForm1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            string passwd = textBox1.Text;
            pwdModifyForm2 f1 = new pwdModifyForm2(passwd);

            f1.ShowDialog();
            this.DialogResult = true;    
            
          
        }
    }
}
