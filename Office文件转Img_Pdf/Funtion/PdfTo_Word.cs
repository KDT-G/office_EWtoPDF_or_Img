using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Pdf;

namespace Office_ToDLL.Funtion
{
    /// <summary>
    /// PDF转Word
    /// </summary>
    public class PdfTo_Word
    {
        /// <summary>
        /// PDF转Word
        /// </summary>
        /// <param name="Inputpath">导入地址</param>
        /// <param name="Outpath">导出地址</param>
        /// <param name="FileName">文件名称</param>
        /// <param name="FileOutType">文件后缀名</param>
        public static void PdfToWrod(string Inputpath, string Outpath, string FileName,string FileOutType=null)
        {
            PdfDocument doc = new PdfDocument();

            try
            {
                //读取PDF
                doc.LoadFromFile(Inputpath);

                //设置导出地址
                string WordOutpath = FileOutType == null ? Outpath + FileName +".docx" : Outpath + FileName + FileOutType;
                //导出Word
                doc.SaveToFile(WordOutpath,FileFormat.DOCX);

                doc.Close();
                doc.Dispose();
            }
            catch (Exception e)
            {

                Console.WriteLine(e);
            }
           

            //打开文件
           // System.Diagnostics.Process.Start("图文版丽江旅游攻略大全.doc");
        }
    }
}
