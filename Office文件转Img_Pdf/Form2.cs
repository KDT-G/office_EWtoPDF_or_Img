using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Office文件转Img_Pdf
{
    public partial class Compress_Demo : Form
    {
        /// <summary>
        /// 进程锁
        /// </summary>
        private static readonly object TestLock_Lock_0 = new object();
        /// <summary>
        /// 程序入口
        /// </summary>
        public Compress_Demo()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 页面关闭时执行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Compress_Demo_FormClosing(object sender, FormClosingEventArgs e)
        {
            //页面关闭时直接关闭应用进程
            System.Environment.Exit(0);
        }
        /// <summary>
        /// 返回主页面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Back_up_Click(object sender, EventArgs e)
        {
            //主页面
            Office_Demo form1 = new Office_Demo();
            //显示主页面
            form1.Show();
            //当前页面隐藏
            this.Hide();
        }

        /// <summary>
        /// 压缩图片
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Compress_Img_Click(object sender, EventArgs e)
        {

            //初始化进度条
            ProgressBar_.Value = 0;
            ProgressBar_.Show();
            
            //判断单个文件还是文件夹
            bool One = InputCheckBox.Checked;
            //判断文件夹或者文件是否存在          
            if (InputPath.Text == "选择文件" || OutPath.Text == "选择文件夹")
            {
                MessageBox.Show("文件或文件夹路径不能为空！");
                return;
            }
            else
            {
                if (One == false)
                {
                    if (!File.Exists(InputPath.Text))
                    {
                        MessageBox.Show("输入路径文件不存在！");
                        return;
                    }
                    if (!System.IO.Directory.Exists(OutPath.Text))
                    {
                        MessageBox.Show("输出路径文件夹不存在！");
                        return;
                    }
                }
                else              //反之判断输入文件夹和存储文件夹是否存在
                {
                    if (!System.IO.Directory.Exists(InputPath.Text))
                    {
                        MessageBox.Show("输入路径文件夹不存在！");
                        return;
                    }
                    if (!System.IO.Directory.Exists(OutPath.Text))
                    {
                        MessageBox.Show("输出路径文件夹不存在！");
                        return;
                    }
                }
            }
            //原图片路径
            string OldImgPath = One == false ? InputPath.Text : Path.GetFullPath(InputPath.Text) + "\\";
            //保存压缩后图片路径
            string NewImgPath = OutPath.Text + "\\";
            //压缩质量（数字越小压缩率越高）1-100
            int ImgQuality = 90;

            try
            {
                //压缩图片最大大小
                int ImgSize = Img_Size.Text == "" ? 300 : Convert.ToInt32(Img_Size.Text);
                //最大分辨率高 默认是0
                int resolution_Max_H = Img_H.Text == "" ? 0 : Convert.ToInt32(Img_H.Text);
                //最大分辨率宽 默认是0
                int resolution_Max_W = Img_W.Text == "" ? 0 : Convert.ToInt32(Img_W.Text);
                try
                {
                    if (!One)
                    {
                        FileInfo File = new FileInfo(OldImgPath);
                        if (File.Extension == ".Png" || File.Extension == ".png" || File.Extension == ".Jpg" || File.Extension == ".jpg" || File.Extension == ".JPEG" || File.Extension == ".jpeg")
                        {

                            //调用压缩图片方法
                            Image_Compression_Main(OldImgPath, NewImgPath, ImgSize, ImgQuality, resolution_Max_H, resolution_Max_W, One);
                        }
                        else
                        {
                            MessageBox.Show("文件类型暂不支持，请重新选择");
                        }
                    }
                    else
                    {

                        //调用压缩图片方法
                        Image_Compression_Main(OldImgPath, NewImgPath, ImgSize, ImgQuality, resolution_Max_H, resolution_Max_W, One);
                    }


                }
                catch (Exception)
                {
                    if (MessageBox.Show("压缩失败") == DialogResult.OK)
                    {
                        ProgressBar_.Hide();
                        //恢复界面
                        Button_True();
                        return;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("请输入正确的数字格式!");
                //恢复界面
                Button_True();
                return;
            }
        }

        /// <summary>
        /// 输入图片路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InputPathButton_Click(object sender, EventArgs e)
        {
            //如果选中选择文件夹则直接选择文件夹，反之选择文件
            if (InputCheckBox.Checked)
            {
                FolderBrowserDialog folder = new FolderBrowserDialog();
                folder.Description = "选择保存的文件夹";

                //如果点击取消选择，则重新赋值路径显示部分
                if (folder.ShowDialog() == DialogResult.OK)
                {
                    //获取目录信息
                    InputPath.Text = folder.SelectedPath;
                }
            }
            else
            {
                //打开文件界面
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Title = "请选择文件";
                ofd.Filter = "所有文件(*.*)|*.*";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    //获取目录信息
                    InputPath.Text = ofd.FileName;
                }

            }
        }

        /// <summary>
        /// 保存图片路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OutPathButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            folder.Description = "选择保存的文件夹";

            //如果点击取消选择，则重新赋值路径显示部分
            if (folder.ShowDialog() == DialogResult.OK)
            {
                //获取目录信息
                OutPath.Text = folder.SelectedPath;
            }
        }
        /// <summary>
        /// 压缩图片
        /// </summary>
        /// <param name="OldImgPath">原图片路径</param>
        /// <param name="NewImgPath">保存压缩后图片路径</param>
        /// <param name="ImgSize">压缩图片最大大小</param>
        /// <param name="ImgQuality">压缩质量（数字越小压缩率越高）1-100</param>
        /// <param name="resolution_Max_H">最大分辨率高 默认是0</param>
        /// <param name="resolution_Max_W">最大分辨率宽 默认是0</param>
        /// <param name="One">是否是单个文件</param>
        public void Image_Compression_Main(string OldImgPath, string NewImgPath, int ImgSize = 300, int ImgQuality = 90, int resolution_Max_H = 0, int resolution_Max_W = 0, bool One = false)
        {
            //执行中隐藏
            Button_False();
            MessageBox.Show("压缩开始执行。。。");
            //用于检测是否完成Task任务
            TaskFactory taskFactory = new TaskFactory();
            if (One == false)
            {
                //定义计时器
                Stopwatch st1 = new Stopwatch();
                //创建Task列表
                List<Task> taskList = new List<Task>();
                st1.Start();//开始计时
                            //添加Task
                taskList.Add(
                Task.Run(() =>
                {
                    //加锁
                    lock (TestLock_Lock_0)
                    {
                        FileInfo File = new FileInfo(OldImgPath);
                        //规定输入文件格式
                        if (File.Extension == ".Png" || File.Extension == ".png" || File.Extension == ".Jpg" || File.Extension == ".jpg" || File.Extension == ".JPEG" || File.Extension == ".jpeg")
                        {
                            ProgressBar_.Value = 50;
                            //分辨率处理  
                            double resolution_Min_Multiple = 1.1;//默认会缩小一点分辨率和大小，避免出现压缩错误问题
                                                                 //拼接具体的文件路径
                            string FilePath = OldImgPath;
                            //文件大小
                            double FileSize = GetFileSize(FilePath);
                            if (1 - FileSize < 0)
                            {
                                if ((1 - (FileSize / 100)) - 1 < -10)//1000MB之外
                                {
                                    resolution_Min_Multiple = 64;
                                }
                                else if ((1 - (FileSize / 100)) <= 0 && (1 - (FileSize / 100)) - 1 > -10)//1000MB之内
                                {
                                    resolution_Min_Multiple = 16;
                                }
                                else if ((1 - (FileSize / 10)) <= 0 && (1 - (FileSize / 10)) - 1 > -10)//100以内
                                {
                                    resolution_Min_Multiple = 8;
                                }
                                else if ((1 - FileSize) <= 0 && (1 - FileSize) - 1 < -1)//10以内
                                {
                                    resolution_Min_Multiple = 4;
                                }
                            }
                            else if (FileSize < ImgSize / 1024)//如果太小则不压缩直接传递
                            {
                                File.CopyTo(NewImgPath + Path.GetFileNameWithoutExtension(File.Name) + ".jpg", true);
                                //ProgressBar_.Value = 100;

                            }
                            //判断文件大小，确定是否压缩，如果大小满足则直接保存到新路径，如果不满足则进行压缩操作
                            if (CompressImage(FilePath, NewImgPath + Path.GetFileNameWithoutExtension(File.Name) + ".jpg", resolution_Max_H, resolution_Max_W, resolution_Min_Multiple, ImgQuality, ImgSize))//压缩成功
                            {
                                //ProgressBar_.Value = 100;

                            }
                        }
                    }
                }));
                //等着全部任务完成后，启动一个新的task来完成后续动作
                taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                {
                    st1.Stop();//结束计时

                    //进度条增加
                    ProgressBar_.Value = 50;
                    Thread.Sleep(1000);
                    ProgressBar_.Value = 100;//结束后进度条拉满

                    double time = st1.ElapsedMilliseconds;//总用时

                    //解析出具体时间
                    int[] Time = Tiem_num(time);
                    //提示完成并告知具体运行时间
                    if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                    {
                        ProgressBar_.Hide();
                        //恢复界面
                        Button_True();
                    }
                });
            }
            else//多文件，文件夹
            {
                //定义计时器
                Stopwatch st1 = new Stopwatch();
                //创建Task列表
                List<Task> taskList = new List<Task>();
                st1.Start();//开始计时
                            //添加Task
                taskList.Add(
                Task.Run(() =>
                {
                    //加锁
                    lock (TestLock_Lock_0)
                    {
                        //获取图片文件路径
                        DirectoryInfo path = new DirectoryInfo(OldImgPath);
                        //获取文件数
                        int count = path.GetFiles().Length;
                        //遍历指定路径目录下文件
                        foreach (FileInfo File in path.GetFiles())
                        {

                            //规定输入文件格式
                            if (File.Extension == ".Png" || File.Extension == ".png" || File.Extension == ".Jpg" || File.Extension == ".jpg" || File.Extension == ".JPEG" || File.Extension == ".jpeg")
                            {
                                //分辨率处理  
                                double resolution_Min_Multiple = 1.1;//默认会缩小一点分辨率和大小，避免出现压缩错误问题
                                                                     //拼接具体的文件路径
                                string FilePath = OldImgPath + File.Name;
                                //文件大小
                                double FileSize = GetFileSize(FilePath);
                                if (1 - FileSize < 0)
                                {
                                    if ((1 - (FileSize / 100)) - 1 < -10)//1000MB之外
                                    {

                                        resolution_Min_Multiple = 64;
                                    }
                                    else if ((1 - (FileSize / 100)) <= 0 && (1 - (FileSize / 100)) - 1 > -10)//1000MB之内
                                    {
                                        resolution_Min_Multiple = 16;
                                    }
                                    else if ((1 - (FileSize / 10)) <= 0 && (1 - (FileSize / 10)) - 1 > -10)//100以内
                                    {
                                        resolution_Min_Multiple = 8;
                                    }
                                    else if ((1 - FileSize) <= 0 && (1 - FileSize) - 1 < -1)//10以内
                                    {
                                        resolution_Min_Multiple = 4;
                                    }
                                }
                                else if (FileSize < ImgSize / 1024)//如果太小则不压缩直接传递
                                {
                                    File.CopyTo(NewImgPath + Path.GetFileNameWithoutExtension(File.Name) + ".jpg", true);

                                    ProgressBar_.Value = ProgressBar_.Value + Convert.ToInt32(count / 100) >= 98 ? 98 : ProgressBar_.Value + Convert.ToInt32(count / 100);
                                    continue;
                                }
                                //判断文件大小，确定是否压缩，如果大小满足则直接保存到新路径，如果不满足则进行压缩操作
                                if (CompressImage(FilePath, NewImgPath + Path.GetFileNameWithoutExtension(File.Name) + ".jpg", resolution_Max_H, resolution_Max_W, resolution_Min_Multiple, ImgQuality, ImgSize))//压缩成功
                                {
                                    ProgressBar_.Value = ProgressBar_.Value + Convert.ToInt32(count / 100) >= 98 ? 98 : ProgressBar_.Value + Convert.ToInt32(count / 100);
                                }


                            }
                        }

                    }
                }));
                //等着全部任务完成后，启动一个新的task来完成后续动作
                taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                {
                    st1.Stop();//结束计时

                    //进度条增加
                    ProgressBar_.Value = 100;//结束后进度条拉满

                    double time = st1.ElapsedMilliseconds;//总用时

                    //解析出具体时间
                    int[] Time = Tiem_num(time);
                    //提示完成并告知具体运行时间
                    if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                    {
                        ProgressBar_.Hide();
                        //恢复界面
                        Button_True();
                    }
                });

            }
        }
        /// <summary>
        /// 获取图片文件大小
        /// </summary>
        /// <param name="path">图片路径</param>
        /// <returns></returns>
        private static double GetFileSize(string path)
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
        private static bool CompressImage(string sFile, string dFile, int resolution_Max_H = 0, int resolution_Max_W = 0, double resolution_Min_Multiple = 1, int flag = 90, int size = 300)
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




        /******************************************附加功能区******************************************/

        /// <summary>
        /// 执行中的操作界面
        /// </summary>
        private void Button_False()
        {
            //执行按钮
            Compress_Img.Text = "执行中";
            Compress_Img.Enabled = false;
            Compress_Img.Cursor = Cursors.No;
            //输入地址按钮
            InputPathButton.Enabled = false;
            InputPathButton.Cursor = Cursors.No;
            //输出地址按钮
            OutPathButton.Enabled = false;
            OutPathButton.Cursor = Cursors.No;
            //返回主菜单按钮
            Back_up.Enabled = false;
            Back_up.Cursor = Cursors.No;
            //选择单文件或者多文件按钮
            InputCheckBox.Enabled = false;
            InputCheckBox.Cursor = Cursors.No;
            //输入路径输入框
            InputPath.Enabled = false;
            InputPath.Cursor = Cursors.No;
            //输出路径输入框
            OutPath.Enabled = false;
            OutPath.Cursor = Cursors.No;
            //图片指定高度
            Img_H.Enabled = false;
            Img_H.Cursor = Cursors.No;
            //图片指定宽度
            Img_W.Enabled = false;
            Img_W.Cursor = Cursors.No;
            //图片指定大小
            Img_Size.Enabled = false;
            Img_Size.Cursor = Cursors.No;

        }
        /// <summary>
        /// 恢复操作界面
        /// </summary>
        private void Button_True()
        {
            //执行按钮
            Compress_Img.Text = "开始执行";
            Compress_Img.Cursor = Cursors.Hand;
            Compress_Img.Enabled = true;
            //输入地址按钮
            InputPathButton.Enabled = true;
            InputPathButton.Cursor = Cursors.Hand;
            //输出地址按钮
            OutPathButton.Enabled = true;
            OutPathButton.Cursor = Cursors.Hand;
            //返回主菜单按钮
            Back_up.Enabled = true;
            Back_up.Cursor = Cursors.Hand;
            //选择单文件或者多文件按钮
            InputCheckBox.Enabled = true;
            InputCheckBox.Cursor = Cursors.Hand;
            //输入路径输入框
            InputPath.Enabled = true;
            InputPath.Cursor = Cursors.Hand;
            //输出路径输入框
            OutPath.Enabled = true;
            OutPath.Cursor = Cursors.Hand;
            //图片指定高度
            Img_H.Enabled = true;
            Img_H.Cursor = Cursors.Hand;
            //图片指定宽度
            Img_W.Enabled = true;
            Img_W.Cursor = Cursors.Hand;
            //图片指定大小
            Img_Size.Enabled = true;
            Img_Size.Cursor = Cursors.Hand;

        }


        /// <summary>
        /// 把毫秒解析成小时、分钟、秒、毫秒
        /// </summary>
        /// <param name="time">获取到得毫秒数</param>
        /// <returns></returns>
        private int[] Tiem_num(double time)
        {
            int[] Time = new int[4];//声明列表存储

            Time[0] = Convert.ToInt32(time / 1000 / 60 / 60);//小时
            Time[1] = Convert.ToInt32(time / 1000 / 60 % 60);//分钟
            Time[2] = Convert.ToInt32(time / 1000 % 60);//秒
            Time[3] = Convert.ToInt32(time % 1000);//毫秒

            return Time;
        }

        #region 拖拽功能
        /*拖拽功能*/
        //输入路径拖拽
        private void InputPath_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();       //获得路径
            InputPath.Text = path;
        }

        private void InputPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;//重要代码：表明是所有类型的数据，比如文件路径

            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        //输出路径拖拽
        private void OutPath_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();       //获得路径
            OutPath.Text = path;
        }

        private void OutPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;//重要代码：表明是所有类型的数据，比如文件路径

            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        #endregion

    }

}

