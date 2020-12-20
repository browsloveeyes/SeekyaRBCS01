//PrintReport.cs,把mdb的数据导入到EXCEL中显示
//添加net引用：Microsoft.Office.Interop.Excel

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;//引用这个才能使用Missing字段
using System.Diagnostics;
using System.Windows;
//File
using System.IO;
//NPOI
using NPOI.HSSF.UserModel;

namespace Seekya
{
    class PrintReport
    {
        //直接打印
        public void ReportPrintDirect(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2)
        {
            WriteCopyTemplate(name, gender, age, id, instrumentType, submitDoctor, firstVisit, hb, hospital, rbc, co, co2, testDateLine, department, userDefine1, userDefine2, userDefine3, userDefine4, userDefine5, userDefine6, checkDoctor, reportDoctor, reportTime, remark1, remark2);
            //直接打印
            //Process.Start(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls");

        }
        //手动打印
        public void ReportPrintHand(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2)
        {
            WriteCopyTemplate(name, gender, age, id, instrumentType, submitDoctor, firstVisit, hb, hospital, rbc, co, co2, testDateLine, department, userDefine1, userDefine2, userDefine3, userDefine4, userDefine5, userDefine6, checkDoctor, reportDoctor, reportTime, remark1, remark2);
            //间接打印
            //Process.Start(System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls");

        }

        //往临时报告中写数据
        public void WriteCopyTemplate(string name, string gender, string age, string id, string instrumentType, string submitDoctor, string firstVisit, string hb, string hospital, string rbc, string co, string co2, string testDateLine, string department, string userDefine1, string userDefine2, string userDefine3, string userDefine4, string userDefine5, string userDefine6, string checkDoctor, string reportDoctor, string reportTime, string remark1, string remark2)
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

            TempletFileName = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls";
            //导出文件  
            //string ReportFileName = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\template\\template.xls";
            FileStream file = new FileStream(TempletFileName, FileMode.Open, FileAccess.Read);
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);
            HSSFSheet ws = hssfworkbook.GetSheet("Sheet1");
            //添加或修改WorkSheet里的数据  
            //System.Data.DataTable dt = new System.Data.DataTable();  
            //dt = DbHelperMySQLnew.Query("select * from t_jb_info where id='" + id + "'").Tables[0];  
            #region

            //姓名
            HSSFRow row = ws.GetRow(1);
            HSSFCell cell = row.GetCell(19);
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

            //ws.GetRow(1).GetCell(1).SetCellValue("5");  
            #endregion
            ws.ForceFormulaRecalculation = true;

            using (FileStream filess = File.OpenWrite(TempletFileName))
            {
                hssfworkbook.Write(filess);
            }

            Process.Start(TempletFileName);
        }

    }
}
