using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;
using System.Threading;
using System.Diagnostics;//Stopwatch

namespace Seekya
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        private const int MINIMUM_SPLASH_TIME = 1500; // Miliseconds  

        protected override void OnStartup(StartupEventArgs e)
        {
            StartInterface start= new StartInterface();

            // Step 2 - Start a stop watch
            //Stopwatch timer = new Stopwatch();
            //timer.Start();

            // Step 3 - Load your windows but don't show it yet  
            base.OnStartup(e);
            MainWindow main = new MainWindow();

            //timer.Stop();

            //int remainingTimeToShowSplash = MINIMUM_SPLASH_TIME - (int)timer.ElapsedMilliseconds;
            //if (remainingTimeToShowSplash > 0)
                //Thread.Sleep(remainingTimeToShowSplash);

            start.Show();

            //显示进度条
            for (int i = 0; i <= 100; i++)
            {
                double value = i * 100.0 / 100;
                start.lbBar.Content = "加载中..      " + i + "%";
                start.pbBar.Dispatcher.Invoke(new Action<System.Windows.DependencyProperty, object>(start.pbBar.SetValue), System.Windows.Threading.DispatcherPriority.Background, System.Windows.Controls.ProgressBar.ValueProperty, value);
                Thread.Sleep(1);
            }

            start.Close();

        }

    }  
    
}
