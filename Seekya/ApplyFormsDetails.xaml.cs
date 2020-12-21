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
using System.Data;
using System.Data.OleDb;
using System.Xml;

namespace Seekya
{
    /// <summary>
    /// ApplyFormsDetails.xaml 的交互逻辑
    /// </summary>
    public partial class ApplyFormsDetails : Window
    {
        public MainWindow m1;
        public ApplyFormsDetails(MainWindow m2)
        {
            InitializeComponent();
            m1 = m2;
        }

        public void Window_Loaded(object sender, RoutedEventArgs e)
        {
            OleDbConnection apfmDb = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data/applyforms.mdb");
            string apfmStr = "Select * from " + m1.ApplyForms;
            try
            {
                apfmDb.Open();
                OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter();
                oleDbDataAdapter.SelectCommand=new OleDbCommand(apfmStr, apfmDb);
                DataSet dset = new DataSet();
                oleDbDataAdapter.Fill(dset);
                datagrid.ItemsSource = dset.Tables[0].DefaultView;
            }
            catch (Exception e202012081431)
            {
                MessageBox.Show("ERROR202012081431:" + e202012081431.Message);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //datagrid.SelectionUnit = DataGridSelectionUnit.FullRow;

            int apfmIndex = datagrid.SelectedIndex;            
            if (apfmIndex!=-1)
            {
                var x = (System.Data.DataRowView)(datagrid.Items.GetItemAt(apfmIndex));
                m1.apfm = x.Row.ItemArray[0].ToString();
                m1.ptID = x.Row.ItemArray[1].ToString();
                m1.ptnb = x.Row.ItemArray[2].ToString();
                m1.imcd = x.Row.ItemArray[3].ToString();
                m1.imnm = x.Row.ItemArray[4].ToString();
                m1.aptm = x.Row.ItemArray[5].ToString();

                string XmlFile = string.Empty;
                XmlFile += "        </DHCLISTOHXBSM></HXBSMCDYJCJG>";
                string[] args = new string[2];
                string msgHeader = string.Empty;
                msgHeader = @"<?xml version='1.0' encoding='utf-8'?>                                                   
                                                        <root>                                                         
                                                                   <serverName>" + "GetLisReports" + "</serverName><format>" + "XML" + "</format><callOperator>" + "" + "</callOperator><certificate>" + "NF6LprJJMrqt6ePCODNhQQ==" + "</certificate><orgCode>" + 01 + "</orgCode>  </root>";
                string msgBody = string.Empty;
                msgBody = @"<?xml version='1.0' encoding='utf-8'?>                                                   
                                                        <root>                                                         
                                                                   <PatientId>" + m1.ptID + "</PatientId><VisitNo>" + m1.ptnb + "</VisitNo></root>";
                args[0] = msgHeader;
                args[1] = msgBody;
                //string url = "http://168.2.5.28:1506/services/WSInterface?wsdl";  //FJ
                string url = "http://192.168.31.164/Webservice1.asmx?wsdl";
                try
                {
                    object result = WebServiceHelper.InvokeWebService(url, "CallInterface", args);
                    XmlDocument xdoc = new XmlDocument();
                    xdoc.LoadXml(result.ToString());
                    XmlElement root = xdoc.DocumentElement;
                    XmlNodeList xnl = null;
                    xnl = root.SelectNodes("/root/returnContents/returnContent/PatientId");
                    foreach (XmlNode node in xnl)
                    {
                        m1.id.Dispatcher.Invoke(new Action(() =>
                        {
                            m1.id.Text = node.InnerText;
                        }));
                    }
                    xnl = root.SelectNodes("/root/returnContents/returnContent/PatientName");
                    foreach (XmlNode node in xnl)
                    {
                        m1.name.Dispatcher.Invoke(new Action(() =>
                        {
                            m1.name.Text = node.InnerText;
                            m1.PatientName = m1.name.Text;
                        }));
                    }
                    xnl = root.SelectNodes("/root/returnContents/returnContent/Sex");
                    foreach (XmlNode node in xnl)
                    {
                        m1.sex.Dispatcher.Invoke(new Action(() =>
                        {
                            m1.sex.Text = node.InnerText;
                        }));
                    }
                    xnl = root.SelectNodes("/root/returnContents/returnContent/Age");
                    foreach (XmlNode node in xnl)
                    {
                        m1.age.Dispatcher.Invoke(new Action(() =>
                        {
                            m1.age.Text = node.InnerText;
                        }));
                    }
                    xnl = root.SelectNodes("/root/returnContents/returnContent/Sex");
                    foreach (XmlNode node in xnl)
                    {
                        m1.sex.Dispatcher.Invoke(new Action(() =>
                        {
                            m1.sex.Text = node.InnerText;
                        }));
                    }
                    xnl = root.SelectNodes("/root/returnContents/returnContent/ReportOperator");
                    foreach (XmlNode node in xnl)
                    {
                        m1.checkDoctor.Dispatcher.Invoke(new Action(() =>
                        {
                            m1.checkDoctor.Text = node.InnerText;
                            m1.ReportOperator = m1.checkDoctor.Text;
                        }));
                    }
                    xnl = root.SelectNodes("/root/returnContents/returnContent/Sex");
                    foreach (XmlNode node in xnl)
                    {
                        m1.sex.Dispatcher.Invoke(new Action(() =>
                        {
                            m1.sex.Text = node.InnerText;
                        }));
                    }
                    xnl = root.SelectNodes("/root/returnContents/returnContent/ItemResult");
                    foreach (XmlNode node in xnl)
                    {
                        m1.textboxhb.Dispatcher.Invoke(new Action(() =>
                        {
                            m1.textboxhb.Text = node.InnerText;
                        }));
                    }
                    m1.receiveInfo.Dispatcher.Invoke(new Action(() =>
                    {
                        m1.receiveInfo.Text += "Get Data Success!"+System.Environment.NewLine;
                    }));

                    //关闭窗口
                    this.Close();
                }
                catch (Exception e202012081640)
                {
                    MessageBox.Show("ERROR202012081640:" + e202012081640.Message+e202012081640.StackTrace);
                }
            }

        }

        public void Window_Closed(object sender, EventArgs e)
        {
            m1.apfmOpen = false;
        }
    }
}
