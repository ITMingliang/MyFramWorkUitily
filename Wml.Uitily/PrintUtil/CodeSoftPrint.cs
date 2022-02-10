using LabelManager2;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wml.Uitily.PrintUtil
{
    public class CodeSoftPrint
    {
        #region 私有方法
        /// <summary>
        /// 打印功能 CodeSoft
        /// </summary>
        /// <param name="PrintParam1"></param>
        /// <param name="PrintParam2"></param>
        /// <param name="PrintParam3"></param>
        /// <param name="PrintParam4"></param>
        /// <returns></returns>
        public bool BOE_SoftCodePrint(string PrintParam1 = "", string PrintParam2 = "", string PrintParam3 = "", string PrintParam4 = "")
        {
            bool result = false;
            int printNum = 2;
            try
            {
                string text = string.Empty;
                ApplicationClass labApp = null;
                Document doc = null;
                string labFileName = AppDomain.CurrentDomain.BaseDirectory + "Template\\" + "Test.Lab";
                if (!File.Exists(labFileName))
                {
                    throw new Exception("沒有找到标签模版");
                }

                for (int i = 0; i < printNum; i++)
                {
                    labApp = new ApplicationClass();
                    labApp.Documents.Open(labFileName, false);// 调用设计好的label文件
                    doc = labApp.ActiveDocument;

                    //可通过配置档进行配置打印信息
                    if (!string.IsNullOrEmpty("SN"))
                        doc.Variables.FreeVariables.Item("SN").Value = PrintParam1;
                    if (!string.IsNullOrEmpty("MSN"))
                        doc.Variables.FreeVariables.Item("MSN").Value = PrintParam2;
                    if (!string.IsNullOrEmpty("PSN"))
                        doc.Variables.FreeVariables.Item("PSN").Value = PrintParam3;
                    if (!string.IsNullOrEmpty("IMEI"))
                        doc.Variables.FreeVariables.Item("IMEI").Value = PrintParam4;
                    doc.PrintDocument(1);
                }

                labApp.Quit();
                result = true;
            }
            catch (Exception ex)
            {
               
            }
            return result;

        }
        #endregion
    }
}
