using Office_ToDLL.Enum;
using Office_ToDLL.Funtion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace Office文件转换.Common
{
    #region 辅助类
    public class Help_Class
    {
        /// <summary>
        /// 文件转换类
        /// </summary>
        /// <param name="FuntionIndex">转换方法选择</param>
        /// <param name="InputPath">导入地址</param>
        /// <param name="OutputPath">导出地址</param>
        /// <param name="FileName">文件名称</param>
        /// <param name="imageFormat">图片类型</param>
        /// <param name="FileOutType">文件导出类型</param>
        public void Class_Main(int FuntionIndex, string InputPath, string OutputPath, string FileName, ImageFormat imageFormat, string FileOutType = null)
        {
            //默认图片类型
            if (imageFormat == null)
            {
                imageFormat = ImageFormat.Png;
            }
            switch (FuntionIndex)
            {
                case ((int)Funtion_Index.Excel_to_img):

                    ExcelTo_Img.ExcelToImg(InputPath, OutputPath, imageFormat);//Excel转图片
                    break;
                case ((int)Funtion_Index.Excel_to_pdf):

                    ExcelTo_PDF.ExcelToPDF(InputPath, OutputPath, FileOutType);//Excel转PDF
                    break;
                case ((int)Funtion_Index.Pdf_to_img):

                    PdfTo_Img.PdfToImg(InputPath, OutputPath, imageFormat, FileName);//PDF转图片
                    break;
                case ((int)Funtion_Index.Pdf_to_word):
                    PdfTo_Word.PdfToWrod(InputPath, OutputPath, FileName, FileOutType);//PDF转Word
                    break;
                case ((int)Funtion_Index.Word_to_img):
                    WordTo_Img.WordToImg(InputPath, OutputPath, imageFormat);//Word转图片
                    break;
                case ((int)Funtion_Index.Word_to_pdf):
                    WordTo_Pdf.WordToPDF(InputPath, OutputPath, FileName);//Word转PDF
                    break;
            }

        }

        /// <summary>
        /// 把毫秒解析成小时、分钟、秒、毫秒
        /// </summary>
        /// <param name="time">获取到得毫秒数</param>
        /// <returns></returns>
        public int[] Tiem_num(double time)
        {
            int[] Time = new int[4];//声明列表存储

            Time[0] = Convert.ToInt32(time / 1000 / 60 / 60);//小时
            Time[1] = Convert.ToInt32(time / 1000 / 60 % 60);//分钟
            Time[2] = Convert.ToInt32(time / 1000 % 60);//秒
            Time[3] = Convert.ToInt32(time % 1000);//毫秒

            return Time;
        }
        /// <summary>
        /// 检查路径或后缀名是否正确
        /// </summary>
        ///  <param name="InputPath_Text">导入框文本</param>
        ///  <param name="OutPath_Text">导出框文本</param>
        ///  <param name="InputCheckBox_Checked">是否选中文件夹</param>
        ///  <param name="message">提示消息</param>
        public bool Check_FilePath(string InputPath_Text,string OutPath_Text,bool InputCheckBox_Checked, out string message)
        {

            //判断文件夹或者文件是否存在          
            if (InputPath_Text == "选择文件" || OutPath_Text == "选择文件夹")
            {
                message = "文件或文件夹路径不能为空！";
                return false;
            }
            else
            {
                //如果没有选中选择文件夹，则判断输入文件和存储文件夹路径是否存在
                if (!InputCheckBox_Checked)
                {
                    if (!File.Exists(InputPath_Text))
                    {

                        message = "输入路径文件不存在！";
                        return false;
                    }

                    if (!System.IO.Directory.Exists(OutPath_Text))
                    {
                        message = "输出路径文件夹不存在！";
                        return false;
                    }
                }
                else              //反之判断输入文件夹和存储文件夹是否存在
                {
                    if (!System.IO.Directory.Exists(InputPath_Text))
                    {
                        message = "输入路径文件夹不存在！";
                        return false;
                    }
                    if (!System.IO.Directory.Exists(OutPath_Text))
                    {
                        message = "输出路径文件夹不存在！";
                        return false;
                    }
                }

            }
            message = null;
            return true;
        }
        /// <summary>
        /// 判断文件类型是否匹配
        /// </summary>
        /// <param name="path">文件地址</param>
        /// <param name="extName">文件类型</param>
        /// <returns></returns>
        public bool File_Type_Or_Right(string path, string[] extName)
        {
            FileInfo file = new FileInfo(path);
            foreach (string item in extName)
            {
                if (file.Extension.ToLower() == item.ToLower())
                {
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// 获取当前目录以及子目录中对应类型的文件
        /// </summary>
        /// <param name="path">导入路径</param>
        /// <param name="extName">文件类型数组</param>
        /// <param name="lst">文件存储列表</param>
        /// <returns></returns>
        public List<FileInfo> File_Name(string path, string[] extName, List<FileInfo> lst)
        {
            try
            {
                //读取文件夹路径
                DirectoryInfo info = new DirectoryInfo(path);
                //获取子目录地址
                string[] dir = Directory.GetDirectories(path);

                FileInfo[] file = info.GetFiles();
                if (file.Length != 0 || dir.Length != 0)
                {

                    foreach (FileInfo File in file)
                    {
                        if (Path.GetFileNameWithoutExtension(File.Name)[0] == '~' && Path.GetFileNameWithoutExtension(File.Name)[1] == '$')
                        {
                            continue;
                        }
                        else
                        {
                            for (int i = 0; i < extName.Length; i++)
                            {
                                //if (extName[i].ToLower().IndexOf(File.Extension.ToLower()) >= 0)
                                //{
                                //    lst.Add(File);
                                //    break;
                                //}
                                if (extName[i].ToLower() == File.Extension.ToLower())
                                {
                                    lst.Add(File);
                                    break;
                                }
                            }
                        }

                    }


                    foreach (string item in dir)
                    {
                        File_Name(item + "\\", extName, lst);//递归，不断查找子目录中的文件
                    }
                }
                return lst;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return null;
            }
        }
        /// <summary>
        /// 获取图片文件大小
        /// </summary>
        /// <param name="path">图片路径</param>
        /// <returns></returns>
        public double GetFileSize(string path)
        {
            System.IO.FileInfo fileInfo = null;
            try
            {
                fileInfo = new System.IO.FileInfo(path);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return 0;
            }
            if (fileInfo != null && fileInfo.Exists)
            {
                //所得文件MB大小
                double Length = (System.Math.Floor(fileInfo.Length / 1024.0 / 1024.0) + System.Math.Floor(fileInfo.Length / 1024.0 % 1024.0 / 10) / 100);
                if (Length > 1000)
                {
                    return 0;
                }
                return Length;
            }
            else
            {
                return 0;
            }
        }
        /// <summary>
        /// 无损压缩图片
        /// </summary>
        /// <param name="sFile">原图片地址</param>
        /// <param name="dFile">保存压缩后图片路径</param>
        /// <param name="resolution_Max_H">最大分辨率高</param>
        /// <param name="resolution_Max_W">最大分辨率宽</param>
        /// <param name="resolution_Min_Multiple">分辨率缩小倍数</param>
        /// <param name="flag">压缩质量（数字越小压缩率越高）1-100</param>
        /// <param name="size">压缩后图片的最大大小</param>
        /// <returns></returns>
        public  bool CompressImage(string sFile, string dFile, int resolution_Max_H = 0, int resolution_Max_W = 0, double resolution_Min_Multiple = 1, int flag = 90, int size = 300)
        {
            Image iSource = Image.FromFile(sFile);
            ImageFormat tFormat = iSource.RawFormat;

            Size tem_size = new Size(iSource.Width, iSource.Height);
            //设置图片高宽
            int dHeight = Convert.ToInt32(iSource.Height / resolution_Min_Multiple);//分辨率高度
            int dWidth = Convert.ToInt32(iSource.Width / resolution_Min_Multiple);   //分辨率宽度


            //判断是否为手动设置最大分辨率高宽
            //判断处理后的分辨率高度是否大于手动设置的分辨率高度,如果大于则设置为指定高度
            if (resolution_Max_H != 0 && iSource.Height / resolution_Min_Multiple > resolution_Max_H)
            {
                //赋值为手动设置的分辨率的高
                dHeight = resolution_Max_H;
            }
            //判断处理后的分辨率宽度是否大于手动设置的分辨率宽度，如果大于则设置为指定宽度
            if (resolution_Max_W != 0 && iSource.Width / resolution_Min_Multiple > resolution_Max_W)
            {
                //赋值为手动设置的分辨率的宽
                dWidth = resolution_Max_W;
            }
            if (resolution_Max_H == 0 && resolution_Max_W == 0)//如果为空则继续赋值为原处理后的分辨率高宽
            {
                dHeight = Convert.ToInt32(iSource.Height / resolution_Min_Multiple);//分辨率高度
                dWidth = Convert.ToInt32(iSource.Width / resolution_Min_Multiple);   //分辨率宽度
            }

            Bitmap ob = new Bitmap(dWidth, dHeight);
            Graphics g = Graphics.FromImage(ob);

            g.Clear(Color.WhiteSmoke);
            g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            //绘制图片
            g.DrawImage(iSource, new Rectangle(0, 0, dWidth, dHeight), 0, 0, iSource.Width, iSource.Height, GraphicsUnit.Pixel);

            g.Dispose();

            //以下代码为保存图片时，设置压缩质量
            EncoderParameters ep = new EncoderParameters();
            long[] qy = new long[1];
            qy[0] = flag;//设置压缩的比例1-100
            EncoderParameter eParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, qy);
            ep.Param[0] = eParam;

            try
            {
                ImageCodecInfo[] arrayICI = ImageCodecInfo.GetImageEncoders();
                ImageCodecInfo jpegICIinfo = null;
                for (int x = 0; x < arrayICI.Length; x++)
                {
                    if (arrayICI[x].FormatDescription.Equals("JPEG"))
                    {
                        jpegICIinfo = arrayICI[x];
                        break;
                    }
                }
                if (jpegICIinfo != null)
                {
                    ob.Save(dFile, jpegICIinfo, ep);//dFile是压缩后的新路径
                    FileInfo fi = new FileInfo(dFile);
                    if (fi.Length > 1024 * size)
                    {
                        flag = flag - 5;
                        CompressImage(sFile, dFile, resolution_Max_H, resolution_Max_W, resolution_Min_Multiple, flag, size);
                    }
                }
                else
                {
                    ob.Save(dFile, tFormat);
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ",堆栈:" + ex.StackTrace + "InnerException:" + ex.InnerException);
                return false;
            }
            finally
            {
                iSource.Dispose();
                ob.Dispose();
            }
        }
   
    }
    #endregion
}