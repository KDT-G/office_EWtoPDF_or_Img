using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office_ToDLL.Funtion
{
    /// <summary>
    /// PDF转图片
    /// </summary>
    public class PdfTo_Img
    {
        /// <summary>
        /// PDF转图片
        /// </summary>
        /// <param name="Inputpath">PDF文件路径</param>
        /// <param name="Outpath">图片输出路径</param>
        /// <param name="imageFormat">图片格式</param>
        /// <param name="FileName">文件名称</param>
        public  static void PdfToImg(string Inputpath, string Outpath, ImageFormat imageFormat, string FileName)
        {
            PdfDocument doc = new PdfDocument();
            try
            {
               
                //读取pdf路径
                doc.LoadFromFile(Inputpath);
                for (int i = 0; i < doc.Pages.Count; i++)
                {
                    //设置图片dpi
                    Image bmp = doc.SaveAsImage(i, 300, 300);
                    //保存图片
                    bmp.Save(Outpath + FileName + "-" + (i + 1) + "." + imageFormat);
                }

                doc.Close();
                doc.Dispose();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

        }
    }
}
