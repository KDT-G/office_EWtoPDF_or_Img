using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace Office文件转Img_Pdf
{
    /// <summary>
    ///  Word转PDF
    /// </summary>
    class WordTo_Pdf
    {
        /// <summary>
        /// Word转PDF
        /// </summary>
        /// <param name="Inputpath">Word文件路径</param>
        /// <param name="Outpath">Word保存路径</param>
        /// <param name="FileName">Word文件名</param>
        /// <returns></returns>
        public static bool WordToPDF(string Inputpath, string Outpath, string FileName)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application App = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document Doc = new Microsoft.Office.Interop.Word.Document();
                object lobjMissing = System.Reflection.Missing.Value;
                //判断是否为打开后产生的临时文档
                if (Path.GetFileNameWithoutExtension(FileName)[0] == '~' || Path.GetFileNameWithoutExtension(FileName)[1] == '$')
                {
                    return false;
                }
                else//如果不是打开后产生的临时文档，则执行Word转PDF
                {
                    //读取Word文档
                    Doc = App.Documents.Open(Inputpath, false, false, lobjMissing, lobjMissing, lobjMissing, false,
                          lobjMissing, lobjMissing, lobjMissing, lobjMissing, lobjMissing, false, lobjMissing, lobjMissing);
                    //保存为PDF
                    Doc.ExportAsFixedFormat(Outpath + Path.GetFileNameWithoutExtension(FileName) + ".pdf", Word.WdExportFormat.wdExportFormatPDF);
                    Doc.Close();
                    App.Quit();
                    //生成没有错误返回True
                    return true;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                //生成错误返回Flase
                return false;
            }

        }
    }
}
