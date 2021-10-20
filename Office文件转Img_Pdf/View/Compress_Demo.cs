using Office文件转换;
using Office文件转换.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Office文件转换
{
    public partial class Compress_Demo : Form
    {
        /**************************************初始化**************************************/
        /// <summary>
        /// 实例化辅助类
        /// </summary>
        Help_Class help_Class = new Help_Class();
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

        /************************************组件功能区************************************/

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
        /// 压缩图片执行按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Compress_Img_Click(object sender, EventArgs e)
        {
            //原图片路径
            string OldImgPath = InputCheckBox.Checked == false ? InputPath.Text : Path.GetFullPath(InputPath.Text) + "\\";
            //保存压缩后图片路径
            string NewImgPath = OutPath.Text + "\\";          
            try
            {
                //判断文件夹或者文件是否存在
                string message = "";
                if (!help_Class.Check_FilePath(InputPath.Text, OutPath.Text, InputCheckBox.Checked, out message))
                {
                    MessageBox.Show(message, "提示");
                    return;
                }
                else
                {
                    //初始化进度条
                    ProgressBar_.Value = 0;
                    ProgressBar_.Show();
                    //压缩质量（数字越小压缩率越高）1-100
                    int ImgQuality = 90;
                    //压缩图片最大大小
                    int ImgSize = Img_Size.Text == "" ? 300 : Convert.ToInt32(Img_Size.Text);
                    //最大分辨率高 默认是0
                    int resolution_Max_H = Img_H.Text == "" ? 0 : Convert.ToInt32(Img_H.Text);
                    //最大分辨率宽 默认是0
                    int resolution_Max_W = Img_W.Text == "" ? 0 : Convert.ToInt32(Img_W.Text);
                    try
                    {
                        string[] FileType = { ".Png", ".png", ".Jpg", ".jpg", ".JPEG", ".jpeg" };
                        if (!InputCheckBox.Checked)
                        {                        
                            if (help_Class.File_Type_Or_Right(OldImgPath, FileType))
                            {
                                //调用压缩图片方法
                                Image_Compression_Main(FileType,OldImgPath, NewImgPath, ImgSize, ImgQuality, resolution_Max_H, resolution_Max_W);
                            }
                            else
                            {
                                MessageBox.Show("文件类型暂不支持，请重新选择");
                            }                      
                        }
                        else
                        {
                            //调用压缩图片方法
                            Image_Compression_Main(FileType,OldImgPath, NewImgPath, ImgSize, ImgQuality, resolution_Max_H, resolution_Max_W,true);
                        }
                    }
                    catch (Exception)
                    {
                        if (MessageBox.Show("压缩失败") == DialogResult.OK)
                        {
                            ProgressBar_.Hide();
                            //恢复界面
                            Execution_Enabled(true, "开始压缩");
                            return;
                        }
                    }
                }                           
            }
            catch (Exception)
            {
                MessageBox.Show("请输入正确的数字格式!");
                //恢复界面
                Execution_Enabled(true, "开始压缩");
                return;
            }
        }
        #region 拖拽功能
        /*拖拽功能*/
        //输入路径拖拽     
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
        private void InputPath_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();       //获得路径
            InputPath.Text = path;
        }
        //输出路径拖拽
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
        private void OutPath_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();       //获得路径
            OutPath.Text = path;
        }
        #endregion

        /************************************附加功能区************************************/

        /// <summary>
        /// 压缩图片方法
        /// </summary>
        /// <param name="FileType">文件类型</param>
        /// <param name="OldImgPath">原图片路径</param>
        /// <param name="NewImgPath">保存压缩后图片路径</param>
        /// <param name="ImgSize">压缩图片最大大小</param>
        /// <param name="ImgQuality">压缩质量（数字越小压缩率越高）1-100</param>
        /// <param name="resolution_Max_H">最大分辨率高 默认是0</param>
        /// <param name="resolution_Max_W">最大分辨率宽 默认是0</param>
        /// <param name="One">是否是单个文件</param>
        public void Image_Compression_Main(string[] FileType, string OldImgPath, string NewImgPath, int ImgSize = 300, int ImgQuality = 90, int resolution_Max_H = 0, int resolution_Max_W = 0, bool One = false)
        {
            try
            {
                //用于检测是否完成Task任务
                TaskFactory taskFactory = new TaskFactory();
                //定义计时器
                Stopwatch st1 = new Stopwatch();
                //创建Task列表
                List<Task> taskList = new List<Task>();
                //创建文件列表，用于获取有效文件
                List<FileInfo> files = new List<FileInfo>();
                if (One == false)
                {
                    if (help_Class.File_Type_Or_Right(OldImgPath, FileType))
                    {
                        files.Add(new FileInfo(OldImgPath));
                    }
                    else
                    {
                        MessageBox.Show("文件类型不正确,请检查！", "提示");
                        return;
                    }
                }
                else
                {
                    //检测的文件类型
                    //执行检测方法，获取FileInfo对象列表
                    files = help_Class.File_Name(OldImgPath, FileType, files);
                    //判断文件夹中是否有相关文件
                    if (files.Count < 1)
                    {
                        MessageBox.Show("未找到相关类型文件,请检查！", "提示");
                        return;
                    }
                }
                st1.Start();//开始计时
                            //添加Task
                taskList.Add(
                Task.Run(() =>
                {
                    //加锁
                    lock (TestLock_Lock_0)
                    {
                        for (int count = 0; count < files.Count; count++)
                        {
                            //加载过滤完的文件夹
                            FileInfo file = files[count] as FileInfo;
                            //获取当前文件所在的文件夹
                            Directory.SetCurrentDirectory(Directory.GetParent(file.FullName).FullName);
                            //获取到文件夹名称
                            string NowFolder_Path = Path.GetFileNameWithoutExtension(Directory.GetCurrentDirectory());
                            //文件名称
                            string FileName = Path.GetFileNameWithoutExtension(file.Name);
                            //分辨率处理  
                            double resolution_Min_Multiple = 1.1;//默认会缩小一点分辨率和大小，避免出现压缩错误问题
                                                                 //拼接具体的文件路径
                                                                 //文件路径
                            string FilePath = file.FullName;
                            //文件大小
                            double FileSize = help_Class.GetFileSize(FilePath);
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
                                file.CopyTo(NewImgPath + FileName + ".jpg", true);
                                continue;
                            }
                            //判断文件大小，确定是否压缩，如果大小满足则直接保存到新路径，如果不满足则进行压缩操作
                            if (help_Class.CompressImage(file.FullName, NewImgPath + FileName + ".jpg", resolution_Max_H, resolution_Max_W, resolution_Min_Multiple, ImgQuality, ImgSize))//压缩成功
                            {
                                continue;
                            }
                        }
                    }
                }));
                //进度条假加载                
                Thread Bar_Thread = new Thread(new ThreadStart(() =>
                {
                    //初始化进度条
                    ProgressBar_.Value = 0;
                    ProgressBar_.Show();
                    //执行中
                    Execution_Enabled(false, "执行中");
                    while (ProgressBar_.Value < 98)
                    {
                        ProgressBar_.Value += 1;

                        Thread.CurrentThread.Join(1000);
                    }
                }));
                //进度条后台执行
                Bar_Thread.IsBackground = true;
                //进度条开始执行
                Bar_Thread.Start();
                //等着全部任务完成后，启动一个新的task来完成后续动作
                taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                {
                    Bar_Thread.Abort();//结束进度条增加
                    st1.Stop();//结束计时

                    ProgressBar_.Value = 100;//结束后进度条拉满

                    double time = st1.ElapsedMilliseconds;//总用时
                    //解析出具体时间
                    int[] Time = help_Class.Tiem_num(time);
                    //提示完成并告知具体运行时间
                    if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                    {
                        ProgressBar_.Hide();
                        //恢复界面
                        Execution_Enabled(true, "开始压缩");
                    }
                });
            }
            catch (Exception)//压缩失败
            {
                MessageBox.Show("压缩失败", "提示");
                Execution_Enabled(true, "开始压缩");
            }
        }
        /// <summary>
        /// 执行中的操作界面
        /// </summary>
        /// <param name="enabled">是否禁用全部控件</param>
        /// <param name="Button_text">执行按钮的文本</param>
        private void Execution_Enabled(bool enabled, string Button_text)
        {
            //是否禁用全部控件
            // this.Enabled = enabled;
            #region 禁用的控件
            //执行按钮
            Compress_Img.Text = Button_text;
            Compress_Img.Enabled = enabled;
            //输入地址按钮
            InputPathButton.Enabled = enabled;
            //输出地址按钮
            OutPathButton.Enabled = enabled;
            //返回主菜单按钮
            Back_up.Enabled = enabled;
            //选择单文件或者多文件按钮
            InputCheckBox.Enabled = enabled;
            //输入路径输入框
            InputPath.Enabled = enabled;
            //输出路径输入框
            OutPath.Enabled = enabled;
            //图片指定高度
            Img_H.Enabled = enabled;
            //图片指定宽度
            Img_W.Enabled = enabled;
            //图片指定大小
            Img_Size.Enabled = enabled;
            #endregion
        }
    }
}