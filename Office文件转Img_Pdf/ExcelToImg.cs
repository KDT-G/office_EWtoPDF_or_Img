using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office文件转Img_Pdf
{
    class ExcelToImg
    {
        /// <summary>
        /// Excel转图片
        /// </summary>
        /// <param name="ExcelPath">Excel地址</param>
        /// <param name="ImgOutPath">图片输出地址</param>
        /// <param name="ImgType">图片格式</param>
        public static void ExcelToImg_(string ExcelPath, string ImgOutPath, ImageFormat ImgType,string NowFolder_Path)
        {


            Workbook book = new Workbook(ExcelPath);
            var list = book.Worksheets;
            foreach (var item in list)
            {
                item.PageSetup.LeftMargin = 0.3;
                item.PageSetup.RightMargin = 0;
                item.PageSetup.BottomMargin = 0;
                item.PageSetup.TopMargin = 0.3;


                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

                imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                //设置清晰度
                imgOptions.HorizontalResolution = 200;
                imgOptions.VerticalResolution = 200;
                //设置是否全表内容显示在一页中
                imgOptions.OnePagePerSheet = true;

                imgOptions.PrintingPage = PrintingPageType.Default;
                //读取工作薄和设置
                SheetRender sr = new SheetRender(item, imgOptions);
                try
                {
                    //设置图片输出地址和名称
                    string Outpath = $@"{ImgOutPath}{NowFolder_Path}-{Path.GetFileNameWithoutExtension(ExcelPath)}-{item.Name}.{ImgType}";
                    //获取Excel转换的图片
                    Bitmap s = sr.ToImage(0);
                    //生成新图片
                    Bitmap k = new Bitmap(s.Width + 20, s.Height + 20);
                    //获取生成的新图片进行绘图
                    System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(k);
                    //背景设置为白色
                    g.Clear(Color.White);
                    //将Excel转换的图片绘制在新生成的图片上
                    g.DrawImage(s, 0, 0, s.Width, s.Height);
                    //保存绘制完成的图片
                    k.Save(Outpath);

                    
                }
                catch (Exception)
                {
                    //如果检测到空白Excel跳过，进行下一张的转换
                    continue;
                }
            }
            //结束进程
            book.Dispose();
        }

    }
}
