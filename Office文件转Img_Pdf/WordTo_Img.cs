using Spire.Pdf;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace Office文件转Img_Pdf
{
    /// <summary>
    /// Word转图片
    /// </summary>
    class WordTo_Img
    {
        /// <summary>
        /// Word转图片
        /// </summary>
        /// <param name="WordInputpath">Word文件夹路径</param>
        /// <param name="Outpath">图片保存路径</param>
        /// <param name="imageFormat">图片格式</param>
        public static void WordToImg(string WordInputpath, string Outpath, ImageFormat imageFormat)
        {
            try
            {
                String path = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\"+"临时";
                Directory.CreateDirectory(path);
               
                if (WordToPDF(WordInputpath, path + "\\", Path.GetFileName(WordInputpath)))//判读Word转pdf是否成功
                {
                    DirectoryInfo PDFinfo = new DirectoryInfo(path+"\\");//获取pdf文件路径
                    FileInfo[] PdfFileList = PDFinfo.GetFiles();
                    foreach (FileInfo PDFfile in PdfFileList)
                    {
                        if (PDFfile.Extension == ".pdf")//当前遍历到的文件是否是pdf文件
                        {
                            //PDF转图片
                            PdfToImg(path+"\\" + PDFfile.Name, Outpath, imageFormat, Path.GetFileNameWithoutExtension(PDFfile.Name));
                            PDFfile.Delete();//删除
                            Directory.Delete(path);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
        /// <summary>
        /// Word转PDF
        /// </summary>
        /// <param name="Inputpath">Word文件路径</param>
        /// <param name="Outpath">Word保存路径</param>
        /// <param name="FileName">Word文件名</param>
        /// <returns></returns>
        private static bool WordToPDF(string Inputpath, string Outpath, string FileName)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application App = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document Doc = new Microsoft.Office.Interop.Word.Document();
                object lobjMissing = System.Reflection.Missing.Value;
                //判断是否为打开后产生的临时文档
                if (Path.GetFileNameWithoutExtension(FileName)[0] == '~' && Path.GetFileNameWithoutExtension(FileName)[1] == '$')
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
        /// <summary>
        /// PDF转图片
        /// </summary>
        /// <param name="Inputpath">PDF文件路径</param>
        /// <param name="Outpath">图片输出路径</param>
        /// <param name="imageFormat">图片格式</param>
        private static void PdfToImg(string Inputpath, string Outpath, ImageFormat imageFormat, string FileName)
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
