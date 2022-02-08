using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wml.Uitily.Logs
{
    public class WriteExcuteLog
    {
        #region 日志记录
        public static void WriteLog(string strLog)
        {
            string text = Environment.CurrentDirectory + "\\LOG\\" + DateTime.Now.ToString("yyyyMMdd");
            string text2 = "Execute.log";
            text2 = text + "\\" + text2;
            if (!Directory.Exists(text))
            {
                Directory.CreateDirectory(text);
            }
            FileStream fileStream = null;
            StreamWriter streamWriter = null;
            try
            {
                if (File.Exists(text2))
                {
                    fileStream = new FileStream(text2, FileMode.Append, FileAccess.Write);
                }
                else
                {
                    fileStream = new FileStream(text2, FileMode.Create, FileAccess.Write);
                }
                streamWriter = new StreamWriter(fileStream);
                streamWriter.WriteLine("【" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "】" + strLog);
            }
            finally
            {
                streamWriter.Close();
                fileStream.Close();
            }
        }
        #endregion
    }
}
