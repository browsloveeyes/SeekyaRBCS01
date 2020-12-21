using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.IO;
using System.Xml;

namespace SeekyaWS
{
    /// <summary>
    /// WebService1 的摘要说明
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消注释以下行。 
    // [System.Web.Script.Services.ScriptService]
    public class WebService1 : System.Web.Services.WebService
    {

        [WebMethod]
        public string HelloWorld()
        {
            return "Hello World";
        }

        [WebMethod]
        public string AppointLfeApply(string msgHeader, string msgBody)
        {
            string path = System.AppDomain.CurrentDomain.BaseDirectory + "Data\\yyxx.txt";
            FileStream fileStream = new FileStream(path, FileMode.Append);
            StreamWriter streamWriter = new StreamWriter(fileStream);
            DateTime time = System.DateTime.Now;
            string time1 = time.ToString("MMddHHmmss");
            var x = time1.Substring(4, 4).ToCharArray();
            Array.Reverse(x);
            string number = new string(x) + time1.Substring(1, 3) + time1[0] + time1[9] + time1[8];
            NHapi.Base.Parser.PipeParser Parser = new NHapi.Base.Parser.PipeParser();
            NHapi.Base.Model.IMessage m = Parser.Parse(msgBody);
            NHapi.Model.V24.Message.ORM_O01 orm001 = m as NHapi.Model.V24.Message.ORM_O01;

            string yysj = orm001.GetORDER(0).ORC.OrderEffectiveDateTime.TimeOfAnEvent.Value;
            string xmlfile = @"<?xml version='1.0' encoding='utf-8'?><ORC><yyh>" + number + "</yyh><yysj>" + yysj + "</yysj><ORC>";
            streamWriter.Write(xmlfile);
            streamWriter.Close();
            fileStream.Close();
            //            MSH |^ ~\&| HIS || LFE || 消息发送时间 || ORM ^ O01 | 消息GUID | P | 2.4
            //PID||| 患者唯一标识ID ^^^^ 标识类型（字典）~患者唯一标识ID ^^^^ 标识类型~身份证号 ^^^^ PN || 患者姓名 ^ 姓名拼音 || 出生日期 | 性别(字典) ||| 患者住址 || 联系电话 || 婚姻状况(字典)
            //PV1 || 患者类别(字典) | 住院科室 ^ 房间号(非必填) ^ 病床号 |||| 主管医生ID ^ 姓名 ||||||||||| 患者类型(字典) | 就诊号（门诊号或住院号）||||||||||||||||||||||||| 就诊时间（入院）
            //ORC | NW | 申请单号 |||||||||| 申请医生ID ^ 姓名 |^^^^^^^^ 开单科室名称 || 申请时间
            //OBR || 申请单号 || 检查项目代码 ^ 检查项目名称 ||||||||||| 检查部位代码 & 检查部位 | 开单医生ID ^ 姓名 ||||||| 单据总费用 ^ 扣费科室ID & 扣费科室名称 | 检查科室代码（可作为检查类别使用）||||||| 检查原因ID ^ 检查原因 ||||||||||||||| 1 ^ 主诉 ~2 ^ 病史及体检状况~3 ^ 辅助检查~4 ^ 特殊要求~5 ^ 旧检查号
            //NTE | 顺序号 | 说明的来源 | 单据类型代码 | 单据类型代码 ^ 单据类型名称
            //DG1 | 顺序号 || 诊断ID ^ 诊断名称 || 诊断时间 | 诊断类型 ||||||||| 诊断优先级 | 诊断医生ID ^ 姓名 ||
            
            //            MSH |^ ~\&| 消息发送方 || 消息接收方 || 消息发送时间 || ACK ^ varies | 消息GUID | P | 2.4
            //MSA | 控制码 | 信息GUID | 错误信息文本
            //PID ||| 患者唯一标识ID ^^^^ 标识类型~患者唯一标识ID ^^^^ 标识类型 || 患者姓名 ^ 姓名拼音 || 出生日期 | 性别 ||| 患者住址 || 联系电话 ||| 婚姻状况
            //ORC | 控制码 | 申请单号 | 报告单号 | 申请单组号 | 申请单状态 |||||| 审核医生ID ^ 姓名 | 开申请单医生ID ^ 姓名 | 开单科室
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(msgHeader.ToString());
            string GUID = doc.SelectSingleNode("/root/msgNo").InnerText;

            string rt = "MSH|^~\\&|P01||HIS||" + time + "||ACK^ varies|" + GUID + "|P|2.4\\.br\\."
                + "MSA|AE|" + GUID + "|\\.br\\."
                + "PID|||" + orm001.PATIENT.PID.GetPatientIdentifierList(0).ID.Value + "||" + orm001.PATIENT.PID.GetPatientName(0).GivenName.Value + "||" + orm001.PATIENT.PID.DateTimeOfBirth.TimeOfAnEvent.Value + "||" + orm001.PATIENT.PID.AdministrativeSex.Value + "|||" + orm001.PATIENT.PID.GetPatientAddress(0).OtherGeographicDesignation.Value + "||" + orm001.PATIENT.PID.GetPhoneNumberHome(0).Get9999999X99999CAnyText.Value + "|||" + orm001.PATIENT.PID.MaritalStatus.Components[0] + "\\.br\\."
                + "ORC|SC|" + orm001.GetORDER(0).ORC.PlacerOrderNumber.EntityIdentifier.Value + "|" + orm001.GetORDER(0).ORC.FillerOrderNumber.EntityIdentifier.Value + "|" + orm001.GetORDER(0).ORC.PlacerGroupNumber.EntityIdentifier.Value + "|" + orm001.GetORDER(0).ORC.OrderStatus.Value + "||||||" + orm001.GetORDER(0).ORC.GetVerifiedBy(0).IDNumber.Value + "|" + orm001.GetORDER(0).ORC.GetOrderingProvider(0).IDNumber.Value + "|" + orm001.GetORDER(0).ORC.EntererSLocation.Room.Value + "\\.br\\.";
            return rt;
        }
    }
}
