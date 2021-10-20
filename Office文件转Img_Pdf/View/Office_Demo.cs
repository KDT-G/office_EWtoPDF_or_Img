using Office_ToDLL;
using Office文件转换.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Office文件转换
{
    public partial class Office_Demo : Form
    {
        /************************************初始化************************************/

        /// <summary>
        /// 实例化辅助类
        /// </summary>
        Help_Class help_Class = new Help_Class();
        /// <summary>
        /// 保存默认的文件图片
        /// </summary>
        public Image Defult_img;
        /// <summary>
        /// 选择方法的下标
        /// </summary>
        internal static int FuntionIndex_num;
        /// <summary>
        /// 程序入口
        /// </summary>
        public Office_Demo()
        {
            InitializeComponent();          
        }

        /************************************组件功能区************************************/

        /// <summary>
        /// 界面加载  获取默认图片
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Office_Demo_Load(object sender, EventArgs e)
        {
            //获取默认图片
            Defult_img = File_ico.Image;
        }
        /// <summary>
        /// 选择输入文件
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
                    //显示文件图片
                    File_ico.Image = System.Drawing.Icon.ExtractAssociatedIcon(InputPath.Text).ToBitmap();
                }
            }
        }
        /// <summary>
        /// 选择输出文件夹
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
        /// 恢复默认格式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Default_Button_Click(object sender, EventArgs e)
        {
            //恢复输入路径文本默认格式
            InputPath.Text = "选择文件";
            //恢复保存路径文本默认格式
            OutPath.Text = "选择文件夹";
            //恢复文件图片默认格式
            File_ico.Image = Defult_img;
            //取消选择当前文件夹内全部文件的按钮
            InputCheckBox.Checked = false;
            //功能区恢复恢复默认状态
            Pdf_Img_radio.Checked = true;
            //进度条
            ProgressBar.Value = 0;
        }
        /// <summary>
        /// 打开压缩图片窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Compress_Img_Click(object sender, EventArgs e)
        {
            Compress_Demo form2 = new Compress_Demo();
            form2.Show();
            this.Hide();
        }
        /// <summary>
        /// 点击×退出进程
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Office_Demo_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.Environment.Exit(0);          
        }       
        /// <summary>
        /// 执行按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Read_Button_Click(object sender, EventArgs e)
        {
            //输入路径
            string FileInputPath= InputCheckBox.Checked == false ? InputPath.Text : InputPath.Text + "\\";
            //保存路径
            string FileOutPath = OutPath.Text + "\\";
            //判断文件夹或者文件是否存在
            string message = "";
            if (!help_Class.Check_FilePath(InputPath.Text, OutPath.Text, InputCheckBox.Checked,out  message))
            {
                MessageBox.Show(message,"提示");
                return;
            }
            else
            {
                #region 功能区
                //判断选择了哪个功能
                if (Pdf_Img_radio.Checked)//PDF转图片
                {
                    Select_Check(FileInputPath, FileOutPath, Pdf_Img_radio);
                }
                else if (Pdf_Word_radio.Checked)//PDF转图片
                {
                    Select_Check(FileInputPath, FileOutPath, Pdf_Word_radio);
                }
                else if (Excel_Img_radio.Checked)//Excel转图片
                {
                    Select_Check(FileInputPath, FileOutPath, Excel_Img_radio);
                }
                else if (Word_Img_radio.Checked)//Word转图片
                {
                    Select_Check(FileInputPath, FileOutPath, Word_Img_radio);
                }
                else if (Word_Pdf_radio.Checked)//Word转PDF
                {
                    Select_Check(FileInputPath, FileOutPath, Word_Pdf_radio);
                }
                else if (Excel_Pdf_radio.Checked)//Excel转PDF
                {
                    Select_Check(FileInputPath, FileOutPath, Excel_Pdf_radio);
                }
                #endregion               
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
            File_ico.Image = System.Drawing.Icon.ExtractAssociatedIcon(InputPath.Text).ToBitmap();
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
            File_ico.Image = System.Drawing.Icon.ExtractAssociatedIcon(InputPath.Text).ToBitmap();
        }
        #endregion

        /************************************功能方法区************************************/

        /// <summary>
        /// 线程锁
        /// </summary>
        private static readonly object TestLock_Lock = new object();
        /// <summary>
        /// 设置转换类型
        /// </summary>
        /// <param name="FileInputPath">导入地址</param>
        /// <param name="FileOutPath">导出地址</param>
        private void Select_Check(string FileInputPath, string FileOutPath, RadioButton radioButton)
        {
            //文件类型列表
            List<string> FileType =new List<string>();
            //判断功能选择
            switch (radioButton.Name)
            {
                case "Excel_Img_radio":
                    FileType.AddRange(new string[] { ".xls", ".xlsx" });
                    FuntionIndex_num = 0;
                    break;
                case "Excel_Pdf_radio":
                    FileType.AddRange(new string[] { ".xls", ".xlsx" });
                    FuntionIndex_num = 1;
                    break;
                case "Pdf_Img_radio":
                    FileType.AddRange(new string[] { ".pdf" });
                    FuntionIndex_num = 2;
                    break;
                case "Pdf_Word_radio":
                    FileType.AddRange(new string[] { ".pdf" });
                    FuntionIndex_num = 3;
                    break;            
                case "Word_Img_radio":
                    FileType.AddRange(new string[] { ".docx", ".doc", ".dot", ".docm", ".dotm", ".dotx" });
                    FuntionIndex_num = 4;
                    break;
                case "Word_Pdf_radio":
                    FileType.AddRange(new string[] { ".docx", ".doc", ".dot", ".docm", ".dotm", ".dotx" });
                    FuntionIndex_num = 5;
                    break;
                case "Null":
                    
                    break;               
            }
           
            Select_Check_Main(FileInputPath, FileOutPath, FileType.ToArray(), FuntionIndex_num);
        }
        /// <summary>
        /// 调用DLL方法执行工作
        /// </summary>
        /// <param name="FileInputPath">导入地址</param>
        /// <param name="FileOutPath">导出地址</param>
        /// <param name="Full_File">是否为文件夹</param>
        /// <param name="files1">文件列表</param>
        /// <param name="FileTpye">文件类型</param>
        /// <param name="image">图片类型</param>
        private void Select_Check_Main(string FileInputPath, string FileOutPath, string[] FileType , int FuntionIndex, Image image = null)
        {
            try
            {
                //用于检测是否完成Task任务
                TaskFactory taskFactory = new TaskFactory();

                //创建文件列表，用于获取有效文件
                List<FileInfo> files = new List<FileInfo>();
                //判断是单个文件还是整个文件夹文件
                if (InputCheckBox.Checked)
                {                    
                    //检测的文件类型
                    //执行检测方法，获取FileInfo对象列表
                    files = help_Class.File_Name(FileInputPath, FileType, files);
                    //判断文件夹中是否有相关文件
                    if (files.Count < 1)
                    {
                        MessageBox.Show("未找到相关类型文件,请检查！", "提示");
                        return;
                    }
                }
                else
                {
                    if (help_Class.File_Type_Or_Right(FileInputPath, FileType))
                    {
                        files.Add(new FileInfo(FileInputPath));
                    }
                    else
                    {
                        MessageBox.Show("文件类型不正确,请检查！", "提示");
                        return;
                    }
                }
                #region 测试代码
                //StringBuilder sb = new StringBuilder();
                //foreach (var item in files)
                //{
                //    sb.Append(item.DirectoryName + "-" + item.FullName + "-" + item.Name + "-" + item.Exists + "\n"+ FileInputPath);

                //}
                //MessageBox.Show(sb.ToString());
                //Execution_Enabled(true, "开始执行");
                //return;
                #endregion
                //定义计时器
                Stopwatch st1 = new Stopwatch();
                //创建Task列表
                List<Task> taskList = new List<Task>();
                //开始计时             
                st1.Start();
                //添加Task          
                taskList.Add(
                Task.Run(() =>
                {
                    //加锁
                    lock (TestLock_Lock)
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
                            //执行对应的函数(导入地址，导出地址，文件名称,图片类型,导出文件后缀[.类型])
                            help_Class.Class_Main(FuntionIndex, file.FullName, FileOutPath,FileName,null,null);
                        }
                    }
                }));
                //进度条假加载                
                Thread Bar_Thread = new Thread(new ThreadStart(() =>
                {
                    //初始化进度条
                    ProgressBar.Value = 0;
                    ProgressBar.Show();
                    //执行中
                    Execution_Enabled(false, "执行中");
                    while (ProgressBar.Value < 98)
                    {
                        ProgressBar.Value += 1;
                        
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

                    ProgressBar.Value = 100;//结束后进度条拉满
                                                               
                    double time = st1.ElapsedMilliseconds;//总用时

                    //解析出具体时间
                    int[] Time = help_Class.Tiem_num(time);
                    //提示完成并告知具体运行时间
                    if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                    {
                        ProgressBar.Hide(); //隐藏进度条
                        Execution_Enabled(true, "开始执行"); //恢复界面
                    }
                });
               
            }
            catch (Exception e)
            {
                MessageBox.Show("转换失败!错误代码:"+e.ToString()+"", "警告");
                return;
            }
        }
        /// <summary>
        /// 执行中的操作界面
        /// </summary>
        /// <param name="enabled">是否禁用全部控件</param>
        /// <param name="Button_text">执行按钮的文本</param>
        private void Execution_Enabled(bool enabled,string Button_text)
        {
            //是否禁用全部控件
           // this.Enabled = enabled;
            //执行按钮文本变化
            Read_Button.Text = Button_text;
            #region 禁用的控件
            //执行按钮
            Read_Button.Enabled = enabled;

            //压缩图片按钮
            Compress_Img.Enabled = enabled;

            //输入路径按钮
            InputPathButton.Enabled = enabled;

            //输出路径按钮
            OutPathButton.Enabled = enabled;

            //恢复默认按钮
            Default_Button.Enabled = enabled;

            //选择单文件和多文件按钮
            InputCheckBox.Enabled = enabled;

            //输入路径输入框
            InputPath.Enabled = enabled;
            //输出路径输入框
            OutPath.Enabled = enabled;
            //功能按钮按钮
            Pdf_Img_radio.Enabled = enabled;
            Word_Pdf_radio.Enabled = enabled;
            Excel_Img_radio.Enabled = enabled;
            Excel_Pdf_radio.Enabled = enabled;
            Word_Img_radio.Enabled = enabled;
            Pdf_Word_radio.Enabled = enabled;
            #endregion
        }    
 
    }
}