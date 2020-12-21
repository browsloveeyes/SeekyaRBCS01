//PrintReport.cs,把mdb的数据导入到EXCEL中显示
//添加net引用：Microsoft.Office.Interop.Excel

using System;
using System.Text;
using System.Diagnostics;
//File
using System.IO;
//NPOI
using NPOI.HSSF.UserModel;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using System.Runtime.InteropServices;

namespace Seekya
{
    class PrintReport
    {
        //直接打印
        public void ReportPrintDirectold(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2, string tyn, int[] RBC,string[] xtime,int number )
        {
            WriteCopyTemplateold(name, gender, age, id, instrumentType, submitDoctor, firstVisit, hb, hospital, rbc, co, co2, testDateLine, department, userDefine1, userDefine2, userDefine3, userDefine4, userDefine5, userDefine6, checkDoctor, reportDoctor, reportTime, remark1, remark2,tyn,RBC,xtime,number);
            //直接打印
            //Process.Start(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls");

        }
        //手动打印
        public void ReportPrintHandold(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2,string tyn,int[] RBC,string[] xtime,int number)
        {
            WriteCopyTemplateold(name, gender, age, id, instrumentType, submitDoctor, firstVisit, hb, hospital, rbc, co, co2, testDateLine, department, userDefine1, userDefine2, userDefine3, userDefine4, userDefine5, userDefine6, checkDoctor, reportDoctor, reportTime, remark1, remark2,tyn,RBC,xtime,number);
            //间接打印
            //Process.Start(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls");

        }

        //往临时报告中写数据
        public void WriteCopyTemplateold(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2,string tyn,int[] RBC,string[] xtime,int number)
        {

            //模板文件  
            string TempletFileName = null;//System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template.xls";
            string pathString = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\print.txt";

            //读打印模板名
            try
            {
                StreamReader sr = new StreamReader(pathString, Encoding.GetEncoding("gb2312"));

                sr.ReadLine();
                TempletFileName = sr.ReadLine();

                sr.Close();

            }
            catch (Exception ex)
            {
                // System.Windows.MessageBox.Show("ERROR:" + ex.Message);

            }

            TempletFileName = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\templatex.xlsx";
            //导出文件  
            //string ReportFileName = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls";
            FileStream file = new FileStream(TempletFileName, FileMode.Open, FileAccess.Read);
            IWorkbook hssfworkbook = new XSSFWorkbook(file);
            ISheet ws = hssfworkbook.GetSheet("Sheet1");
            //添加或修改WorkSheet里的数据  
            //System.Data.DataTable dt = new System.Data.DataTable();  
            //dt = DbHelperMySQLnew.Query("select * from t_jb_info where id='" + id + "'").Tables[0];  
            #region

            //姓名
            IRow row = ws.GetRow(1);
            ICell cell = row.GetCell(19);
            cell.SetCellValue(name);

            //性别
            row = ws.GetRow(2);
            cell = row.GetCell(19);
            cell.SetCellValue(gender);

            //年龄
            row = ws.GetRow(3);
            cell = row.GetCell(19);
            cell.SetCellValue(age);

            //住院号
            row = ws.GetRow(4);
            cell = row.GetCell(19);
            cell.SetCellValue(id);

            //仪器型号
            row = ws.GetRow(5);
            cell = row.GetCell(19);
            cell.SetCellValue(instrumentType);

            //送检医生
            row = ws.GetRow(6);
            cell = row.GetCell(19);
            cell.SetCellValue(submitDoctor);

            //初步诊断
            row = ws.GetRow(7);
            cell = row.GetCell(19);
            cell.SetCellValue(firstVisit);

            //血红蛋白浓度
            row = ws.GetRow(8);
            cell = row.GetCell(19);
            cell.SetCellValue(hb);

            //医院名称
            row = ws.GetRow(9);
            cell = row.GetCell(19);
            cell.SetCellValue(hospital);

            //红细胞寿命
            row = ws.GetRow(10);
            cell = row.GetCell(19);
            cell.SetCellValue(rbc);

            //一氧化碳浓度
            row = ws.GetRow(11);
            cell = row.GetCell(19);
            cell.SetCellValue(co);

            //二氧化碳浓度
            row = ws.GetRow(12);
            cell = row.GetCell(19);
            cell.SetCellValue(co2);

            //检验日期
            row = ws.GetRow(13);
            cell = row.GetCell(19);
            cell.SetCellValue(testDateLine);

            //科室名称
            row = ws.GetRow(14);
            cell = row.GetCell(19);
            cell.SetCellValue(department);

            //定义1
            row = ws.GetRow(15);
            cell = row.GetCell(19);
            cell.SetCellValue(userDefine1);

            //定义2
            row = ws.GetRow(16);
            cell = row.GetCell(19);
            cell.SetCellValue(userDefine2);

            //定义3
            row = ws.GetRow(17);
            cell = row.GetCell(19);
            cell.SetCellValue(userDefine3);

            //定义4
            row = ws.GetRow(18);
            cell = row.GetCell(19);
            cell.SetCellValue(userDefine4);

            //定义5
            row = ws.GetRow(19);
            cell = row.GetCell(19);
            cell.SetCellValue(userDefine5);

            //定义6
            row = ws.GetRow(20);
            cell = row.GetCell(19);
            cell.SetCellValue(userDefine6);

            //复核医生
            row = ws.GetRow(21);
            cell = row.GetCell(19);
            cell.SetCellValue(checkDoctor);

            //报告医生
            row = ws.GetRow(22);
            cell = row.GetCell(19);
            cell.SetCellValue(reportDoctor);

            //报告时间
            row = ws.GetRow(23);
            cell = row.GetCell(19);
            cell.SetCellValue(reportTime);

            //零点过大
            row = ws.GetRow(24);
            cell = row.GetCell(19);
            cell.SetCellValue(remark1);

            //CO2过低
            row = ws.GetRow(25);
            cell = row.GetCell(19);
            cell.SetCellValue(remark2);

            //样本类型
            row = ws.GetRow(26);
            cell = row.GetCell(19);
            cell.SetCellValue(tyn);
            #endregion

            if (number > 1)
            {
                for (int w = 0; w < number; w++)
                {
                    row = ws.GetRow(100);
                    cell = row.GetCell(w);
                    cell.SetCellValue(xtime[w]);
                    row = ws.GetRow(101);
                    cell = row.GetCell(w);
                    cell.SetCellValue(RBC[w]);
                }
                NPOI.SS.UserModel.IDrawing drawing = ws.CreateDrawingPatriarch();
                IClientAnchor anchor1 = drawing.CreateAnchor(0, 0, 0, 0, 0, 19, 9, 40);
                CreateChart(drawing, ws, anchor1, "红细胞寿命变化示意图", number);
            }

            ws.ForceFormulaRecalculation = true;

            //另存为以姓名+日期+时间+序号为文件名的文件 (3.23new)
            string date1 = testDateLine.Substring(0, 4) + testDateLine.Substring(5, 2) + testDateLine.Substring(8, 2);
            string datetime2 = reportTime.Substring(0, 2) + reportTime.Substring(3, 2) + reportTime.Substring(6, 2);
            string excelname = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\Template\\" + name + "(" + date1 + datetime2 + ")" + ".xlsx";
            int postn = excelname.LastIndexOf(".");
            int k = 1;
            while (System.IO.File.Exists(excelname))
            {
                excelname = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\Template\\" + name + "(" + date1 + datetime2 + ")" + ".xlsx";

                excelname = excelname.Insert(postn, "(" + k + ")");
                //excelName = string.Format(excelName + i);
                k++;
            }

            //using (FileStream filess = File.OpenWrite(TempletFileName))
            //{
            //    hssfworkbook.Write(filess);
            //}

            using (FileStream filess = File.Create(excelname))
            {
                hssfworkbook.Write(filess);
            }

            //Process.Start(TempletFileName);
            Process.Start(excelname);

            //using (FileStream filess = File.OpenWrite(TempletFileName))
            //{
            //    hssfworkbook.Write(filess);
            //}

            //Process.Start(TempletFileName);
        }

        public void CreateChart(NPOI.SS.UserModel.IDrawing drawing, ISheet sheet, IClientAnchor anchor, string serie1,/* string serie2,*/int number)
        {
            try
            {
                var chart = drawing.CreateChart(anchor) as XSSFChart;
                //生成图例
                var legend = chart.GetOrCreateLegend();
                //图例位置
                legend.Position = LegendPosition.TopRight;

                //图表
                var data = chart.ChartDataFactory.CreateLineChartData<double, double>(); //折线图
                //var data = chart.ChartDataFactory.CreateScatterChartData<double, double>(); //散点图

                // X轴.
                var bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
                bottomAxis.IsVisible = true; //默认为true 不显示  设置为fase 显示坐标轴(BUG?)

                //Y轴
                IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
                leftAxis.Crosses = (AxisCrosses.AutoZero);
                leftAxis.IsVisible = true; //设置显示坐标轴

                //数据源
                IChartDataSource<double> xs = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(100, 100, 0, number-1));

                IChartDataSource<double> ys1 = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(101, 101, 0, number-1));
                //IChartDataSource<double> ys2 = DataSources.FromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, number - 1));

                //数据系列
                var s1 = data.AddSeries(xs, ys1);
                s1.SetTitle(serie1);
                //s1.GetXValues();
                //var s2 = data.AddSeries(xs, ys2);
                //s2.SetTitle(serie2);


                chart.Plot(data, bottomAxis, leftAxis);
            }
            catch (Exception e35)
            {
                System.Windows.MessageBox.Show("ERROR35:" + e35.Message);
            }
            

        }

        //手动打印
        public void ReportPrintHand(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2,string tyn, string excelname)
        {
            WriteCopyTemplatemanual(name, gender, age, id, instrumentType, submitDoctor, firstVisit, hb, hospital, rbc, co, co2, testDateLine, department, userDefine1, userDefine2, userDefine3, userDefine4, userDefine5, userDefine6, checkDoctor, reportDoctor, reportTime, remark1, remark2,tyn, excelname);
            //间接打印
            //Process.Start(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls");

        }

        public void WriteCopyTemplatemanual(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2, string tyn,string excelname)
        {

            try
            {
                //Process.Start(filename);
                Process.Start(excelname);
            }
            catch (Exception e)
            {
                System.Windows.MessageBox.Show("ERROR28:" + e.Message);
            }
        }

    }

    public class PublicMethod
    {
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
        {
            IntPtr t = new IntPtr(excel.Hwnd);//得到这个句柄，具体作用是得到这块内存入口 

            int k = 0;
            GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
            p.Kill();     //关闭进程k
        }
    }
}
