//添加两个com组件引用
//Microsoft ADO Ext. 2.8 for DDL and Security  //ADOX
//Microsoft ActiveX Data Objects 2.8 Library  //ADOX
using System;
using System.Collections;
using System.Windows;
using System.Collections.Generic;
using System.Data.OleDb;   //OleDbConnection
using System.Linq;
using System.Text;
using System.Data;
using ADOX;

namespace Seekya
{
    class DbOperate
    {
        //创建数据库
        public void CreateDb()
        {
            ADOX.CatalogClass m_Adox = new CatalogClass();    // 初始化CatalogClass对象
            try
            {

                m_Adox.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");//创建数据库

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR15:" + ex.Message);
            }
            finally
            {
                m_Adox = null;
                GC.Collect();  // 强制回收资源并解除LDB锁定

            }

        }

        //创建表
        //新建mdb的表,C#操作Access之创建表 

        //mdbHead是一个ArrayList，存储的是table表中的具体列名。  

        public void CreateTable(
        string mdbPath, string tableName, ArrayList mdbHead)
        {
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            try
            {
                ADOX.CatalogClass cat = new ADOX.CatalogClass();

                string sAccessConnection = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbPath;
                ADODB.Connection cn = new ADODB.Connection();

                cn.Open(sAccessConnection, null, null, -1);

                cat.ActiveConnection = cn;

                //新建一个表,C#操作Access之创建表
                ADOX.TableClass tbl = new ADOX.TableClass();
                tbl.ParentCatalog = cat;
                tbl.Name = tableName;

                int size = mdbHead.Count;
                for (int i = 0; i < size; i++)
                {
                    //增加一个文本字段
                    ADOX.ColumnClass col2 = new ADOX.ColumnClass();

                    col2.ParentCatalog = cat;
                    col2.Name = mdbHead[i].ToString();//列的名称

                    col2.Properties["Jet OLEDB:Allow Zero Length"].Value = false;

                    tbl.Columns.Append(col2, ADOX.DataTypeEnum.adVarWChar, 500);
                }
                cat.Tables.Append(tbl); //这句把表加入数据库(非常重要)  ,C#操作Access之创建表
                tbl = null;
                cat = null;
                cn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR16:" + ex.Message);

            }

            try
            {
                aConnection.Open();
                //创建姓名的索引
                string strSql = "Create index myIndex on " + tableName + "(姓名)";
                OleDbCommand myCmd = new OleDbCommand(strSql, aConnection);
                myCmd.ExecuteNonQuery();
                //创建住院号的索引
                strSql = "Create index myIndex1 on " + tableName + "(住院号)";
                OleDbCommand myCmd1 = new OleDbCommand(strSql, aConnection);
                myCmd1.ExecuteNonQuery();

                /*
                //修改“医院名称”的属性
                strSql = "Alter table " + tableName + " alter column 医院名称 Text(255) not null default 1";//设置允许“医院名称”字段为空命令
                OleDbCommand myCmd2 = new OleDbCommand(strSql, aConnection);
                myCmd2.ExecuteNonQuery();
                */

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR17:" + ex.Message);

            }
            finally
            {
                if (aConnection != null)
                    aConnection.Close();

            }
        }
        //删除表
        public void DeleteTable(string tableName)
        {
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");

            try
            {
                aConnection.Open();
                string strSql = "Drop table " + tableName;
                //strSql = strSql.Substring(0, 15) + "\\" + strSql.Substring(16, 2) + "\\" + strSql.Substring(19, 2);
                OleDbCommand myCmd = new OleDbCommand(strSql, aConnection);
                myCmd.ExecuteNonQuery();
                //MessageBox.Show("删除表操作成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR18:" + ex.Message);

            }
            finally
            {
                if (aConnection != null)
                    aConnection.Close();

            }



        }
        //删除记录
        public bool DeleteRecord(string date, string time)
        {
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            string strSql = "Delete from " + date + " where 时间=" + "\'" + time + "\'";

            try
            {
                aConnection.Open();
                OleDbCommand myCmd = new OleDbCommand(strSql, aConnection);
                myCmd.ExecuteNonQuery();

                if (aConnection != null)
                    aConnection.Close();
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR19:" + ex.Message);

                if (aConnection != null)
                    aConnection.Close();
                return false;

            }

        }

        //更改记录
        public bool ModifyRecord(string originDate, string originTime, string hospitalName, string roomName, string instrumentModel, string name, string sex, string age, string number, string CO, string CO2, string RBC, string textboxhb, string sendDoctor, string reviewDoctor, string checkDoctor, string firstCheck, string time, string date, string remark1, string remark2)
        {
            DeleteRecord(originDate, originTime);//删除原始的记录
            OleDbConnection aConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + System.AppDomain.CurrentDomain.BaseDirectory + "Data\\checkDb.mdb");
            string strSql = "Insert into " + originDate + " (医院名称,科室名称,仪器型号,姓名,性别,年龄,住院号,CO,CO2,红细胞寿命,血红蛋白浓度,送检医生,复核医生,报告医生,初步诊断,时间,日期,备注1,备注2) values ('" + hospitalName + "','" + roomName + "','" + instrumentModel + "','" + name + "','" + sex + "','" + age + "','" + number + "','" + CO + "','" + CO2 + "','" + RBC + "','" + textboxhb + "','" + sendDoctor + "','" + reviewDoctor + "','" + checkDoctor + "','" + firstCheck + "','" + time + "','" + date + "','" + remark1 + "','" + remark2 + "')";

            try
            {
                aConnection.Open();
                OleDbCommand myCmd = new OleDbCommand(strSql, aConnection);
                myCmd.ExecuteNonQuery();

                if (aConnection != null)
                    aConnection.Close();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR20:" + ex.Message);

                if (aConnection != null)
                    aConnection.Close();
                return false;

            }


        }
    }
}
