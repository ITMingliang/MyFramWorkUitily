using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wml.Uitily.CommonHelper
{
    public class PDFWaterMark
    {
        /// <summary>
        /// 添加普通偏转角度文字水印
        /// </summary>
        public void SetWatermark(string filePath, string text)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            string tempPath = Path.GetDirectoryName(filePath) + Path.GetFileNameWithoutExtension(filePath) + "_temp.pdf";

            try
            {
                pdfReader = new PdfReader(filePath);
                pdfStamper = new PdfStamper(pdfReader, new FileStream(tempPath, FileMode.Create));
                int total = pdfReader.NumberOfPages + 1;
                Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont.CreateFont(@"C:\WINDOWS\Fonts\SIMFANG.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                PdfGState gs = new PdfGState();
                for (int i = 1; i < total; i++)
                {
                    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                    //content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                    //透明度
                    gs.FillOpacity = 0.3f;
                    content.SetGState(gs);
                    //content.SetGrayFill(0.3f);
                    //开始写入文本
                    content.BeginText();
                    content.SetColorFill(BaseColor.GRAY);
                    content.SetFontAndSize(font, 30);
                    content.SetTextMatrix(0, 0);
                    content.ShowTextAligned(Element.ALIGN_CENTER, text, width - 80, height - 80, -45);//设置水印角度位置
                    //content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, 0, 0, 0);
                    content.EndText();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
                 File.Copy(tempPath, filePath, true);
                 File.Delete(tempPath);
            }
        }
    }
}
