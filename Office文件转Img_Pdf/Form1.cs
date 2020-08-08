using Spire.Pdf;
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
    public partial class Office_Demo : Form
    {
        /// <summary>
        /// 保存默认的文件图片
        /// </summary>
        public Image Defult_img;
        /// <summary>
        /// 程序入口
        /// </summary>
        public Office_Demo()
        {
            InitializeComponent();
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
        /// 线程锁
        /// </summary>
        private static readonly object TestLock_Lock = new object();
        /// <summary>
        /// 执行按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Read_Button_Click(object sender, EventArgs e)
        {
            /******************************************初始区分界线***************************************/
          
            #region 初始化区
            //判断文件夹或者文件是否存在          
            if (InputPath.Text == "选择文件" || OutPath.Text == "选择文件夹")
            {
                MessageBox.Show("文件或文件夹路径不能为空！");
                return;
            }
            else
            {
                //如果没有选中选择文件夹，则判断输入文件和存储文件夹路径是否存在
                if (!InputCheckBox.Checked)
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
            //初始化进度条
            ProgressBar.Value = 0;
            ProgressBar.Show();
            //执行中
            Button_False();
            #endregion
            /******************************************功能区分界线***************************************/

            #region 功能区

            /*------------------------------------------PDF转图片------------------------------------------*/
            #region PDF转图片
            //判断选择了哪个功能
            if (Pdf_Img_radio.Checked)//PDF转图片
            {
                //MessageBox.Show("PDF转图片开始执行。。。");
                Pdf_Img_radio_Fun();
                
            }
            #endregion
            /*------------------------------------------Excel转图片------------------------------------------*/
            #region Excel转图片
            else if (Excel_Img_radio.Checked)//Excel转图片
            {
                //MessageBox.Show("Excel转图片开始执行。。。");
                Excel_Img_radio_Fun();
            }
            #endregion
            /*------------------------------------------Word转图片------------------------------------------*/
            #region Word转图片
            else if (Word_Img_radio.Checked)//Word转图片
            {
                //MessageBox.Show("Word转图片开始执行。。。");
                Word_Img_radio_Fun();
            }
            #endregion
            /*------------------------------------------Word转PDF------------------------------------------*/
            #region Word转PDF
            else if (Word_Pdf_radio.Checked)//Word转PDF
            {
                //MessageBox.Show("Word转PDF开始执行。。。");
                Word_Pdf_radio_Fun();


            }
            #endregion
            /*------------------------------------------Excel转PDF------------------------------------------*/
            #region Excel转PDF
            else if (Excel_Pdf_radio.Checked)//Excel转PDF
            {
                //MessageBox.Show("Excel转PDF开始执行。。。");
                Excel_Pdf_radio_Fun();
            }
            #endregion

            #endregion
            
        }      
        /// <summary>
        /// PDF转图片
        /// </summary>
        private void Pdf_Img_radio_Fun()
        {
            //PDF输入路径
            string pdfInputPath = InputCheckBox.Checked == false ? InputPath.Text : InputPath.Text + "\\";
            //图片保存路径
            string imgOutPath = OutPath.Text + "\\";
            //声明函数
            PdfDocument doc = new PdfDocument();
            //用于检测是否完成Task任务
            TaskFactory taskFactory = new TaskFactory();
            //判断是单个文件还是整个文件夹文件
            if (InputCheckBox.Checked)//整个文件夹
            {
                try
                {
                    //获取有效文件数
                    List<FileInfo> files = new List<FileInfo>();                   
                    //检测的文件类型
                    string [] es={ ".pdf" };
                    //执行检测方法，获取FileInfo对象列表
                    files = File_Name(pdfInputPath, es, files);
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
                        lock (TestLock_Lock)
                        {
                            for (int count = 0; count < files.Count; count++)
                            {
                                //加载过滤完的文件夹
                                FileInfo file = files[count] as FileInfo;
                                //获取当前文件所在的文件夹
                                Directory.SetCurrentDirectory(Directory.GetParent(file.FullName).FullName);
                                string NowFolder_Path = Path.GetFileNameWithoutExtension(Directory.GetCurrentDirectory());
                                //读取PDF
                                doc.LoadFromFile(file.FullName);
                                //遍历PDF页数
                                for (int i = 0; i < doc.Pages.Count; i++)
                                {
                                    //按页保存
                                    Image bmp = doc.SaveAsImage(i, 300, 300);
                                    //保存到指定地址  
                                    bmp.Save(imgOutPath + NowFolder_Path + "-" + Path.GetFileNameWithoutExtension(file.Name) + "-" + (i + 1) + ".png");
                                }
                                if (count <= Convert.ToInt32(Math.Floor(Convert.ToDecimal(files.Count / 2))))
                                {
                                    //如果处理的文件不到一半，那就定在50%
                                    ProgressBar.Value = count >= 50 ? 50 : count;
                                }
                                else
                                {
                                    ProgressBar.Value = count >= 98 ? 98 : count;
                                }
                                //关闭文件夹
                                doc.Close();
                                doc.Dispose();
                            }
                        }
                    }));

                    //等着全部任务完成后，启动一个新的task来完成后续动作
                    taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                    {
                        st1.Stop();//结束计时

                        ProgressBar.Value = 100;//结束后进度条拉满

                        double time = st1.ElapsedMilliseconds;//总用时

                        //解析出具体时间
                        int[] Time = Tiem_num(time);
                        //提示完成并告知具体运行时间
                        if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                        {
                            ProgressBar.Hide();
                            //恢复界面
                            Button_True();
                        }
                    });
                    
                }
                catch (Exception)
                {
                    MessageBox.Show("转换失败!");
                    //恢复界面
                    Button_True();
                    return;
                }

            }
            else//单个文件
            {
                try
                {
                    FileInfo file = new FileInfo(pdfInputPath);
                    if (file.Extension == ".pdf" || file.Extension == ".Pdf" || file.Extension == ".PDF" || file.Extension == ".PDf")
                    {
                        Stopwatch st1 = new Stopwatch();
                        List<Task> taskList = new List<Task>();
                        st1.Start();//开始计时
                        taskList.Add(
                        Task.Run(() =>
                        {
                            lock (TestLock_Lock)
                            {
                                    //读取PDF
                                    doc.LoadFromFile(pdfInputPath);
                                for (int i = 0; i < doc.Pages.Count; i++)//遍历PDF页数
                                {
                                        //按页保存
                                        Image bmp = doc.SaveAsImage(i, 300, 300);
                                        //保存到指定地址
                                        bmp.Save(imgOutPath + Path.GetFileNameWithoutExtension(file.Name) + (i + 1) + ".png");
                                }
                                doc.Close();
                                doc.Dispose();
                            }
                        }));
                        //等着全部任务完成后，启动一个新的task来完成后续动作
                        taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                        {
                           
                            st1.Stop();//结束计时
                                       //进度条增加
                            ProgressBar.Value = 50;
                            Thread.Sleep(1000);
                            ProgressBar.Value = 100;//结束后进度条拉满

                            double time = st1.ElapsedMilliseconds;//总用时

                            //解析出具体时间
                           int[] Time = Tiem_num(time);
                           //提示完成并告知具体运行时间
                           if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                           {
                                ProgressBar.Hide();
                                //恢复界面
                                Button_True();
                            }
                        });
                    }
                    else
                    {
                        MessageBox.Show("文件类型不正确!");
                        //恢复界面
                        Button_True();
                        return;
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("转换失败!");
                    //恢复界面
                    Button_True();
                    return;
                }
            }
        }

        /// <summary>
        /// Excel转图片
        /// </summary>
        private void Excel_Img_radio_Fun()
        {
            //可以破解附带的Aspose.Cells 16.12.0.0版本
            LicenseHelper.ModifyInMemory.ActivateMemoryPatching();
            //Excel地址
            string ExcelInputPath = InputCheckBox.Checked == false ? InputPath.Text : InputPath.Text + "\\";
            //图片输出地址
            string ImgOutPath = OutPath.Text + "\\";
            //用于检测是否完成Task任务
            TaskFactory taskFactory = new TaskFactory();
            //图片格式
            ImageFormat ImgType = ImageFormat.Png;
            if (InputCheckBox.Checked)//整个文件夹
            {                
                try
                {
                    //获取有效文件数
                    List<FileInfo> files = new List<FileInfo>();
                    //检测的文件类型
                    string[] es = { ".xls", ".xlsx" };
                    //执行检测方法，获取FileInfo对象列表
                    files = File_Name(ExcelInputPath, es, files);
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
                        lock (TestLock_Lock)
                        {
                            for (int count = 0; count < files.Count; count++)
                            {
                                //加载过滤完的文件夹
                                FileInfo file = files[count] as FileInfo;
                                //获取当前文件所在的文件夹
                                Directory.SetCurrentDirectory(Directory.GetParent(file.FullName).FullName);
                                string NowFolder_Path = Path.GetFileNameWithoutExtension(Directory.GetCurrentDirectory());
                                //Excel转图片（Excel地址+Excel文件名，图片输出地址，图片格式）
                                ExcelToImg.ExcelToImg_(ExcelInputPath + file.Name, ImgOutPath, ImgType, NowFolder_Path);
                                //进度条增进
                                if (count <= Convert.ToInt32(Math.Floor(Convert.ToDecimal(files.Count / 2))))
                                {
                                    //如果处理的文件不到一半，那就定在50%
                                    ProgressBar.Value = count >= 50 ? 50 : count;
                                }
                                else
                                {
                                    ProgressBar.Value = count >= 98 ? 98 : count;
                                }
                            }
                        }
                    }));

                    //等着全部任务完成后，启动一个新的task来完成后续动作
                    taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                    {
                        st1.Stop();//结束计时

                        ProgressBar.Value = 100;//结束后进度条拉满

                        double time = st1.ElapsedMilliseconds;//总用时

                        //解析出具体时间
                        int[] Time = Tiem_num(time);
                        //提示完成并告知具体运行时间
                        if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                        {
                            ProgressBar.Hide();
                            //恢复界面
                            Button_True();
                        }
                    });                
                }
                catch (Exception)
                {

                    MessageBox.Show("转换失败！");
                    //恢复界面
                    Button_True();
                    return;
                }

            }
            else//单个文件
            {
                try
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
                        lock (TestLock_Lock)
                        {
                            //判断格式
                            FileInfo file = new FileInfo(ExcelInputPath);
                            if (file.Extension == ".xls" || file.Extension == ".xlsx" || file.Extension == ".XLS" || file.Extension == ".XLSX")
                            {

                                //获取当前文件所在的文件夹
                                Directory.SetCurrentDirectory(Directory.GetParent(file.FullName).FullName);
                                string NowFolder_Path = Path.GetFileNameWithoutExtension(Directory.GetCurrentDirectory());
                                //读取Excel路径
                                //Excel转图片（Excel地址+Excel文件名，图片输出地址，图片格式）
                                ExcelToImg.ExcelToImg_(ExcelInputPath, ImgOutPath, ImgType, NowFolder_Path);
                            }
                            else
                            {
                                MessageBox.Show("类型错误！");
                                //恢复界面
                                Button_True();
                                return;
                            }
                        }
                    }));
                    //等着全部任务完成后，启动一个新的task来完成后续动作
                    taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                    {
                        st1.Stop();//结束计时

                        //进度条增加
                        ProgressBar.Value = 50;
                        Thread.Sleep(1000);
                        ProgressBar.Value = 100;//结束后进度条拉满

                        double time = st1.ElapsedMilliseconds;//总用时

                        //解析出具体时间
                        int[] Time = Tiem_num(time);
                        //提示完成并告知具体运行时间
                        if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                        {
                              ProgressBar.Hide();
                            //恢复界面
                            Button_True();
                        }
                    });

                }
                catch (Exception)
                {

                    MessageBox.Show("转换失败");
                    //恢复界面
                    Button_True();
                    return;
                }
            }
        }

        /// <summary>
        /// Word转图片
        /// </summary>
        private void Word_Img_radio_Fun()
        {
            //Word文件路径
            string WordInputPath = InputCheckBox.Checked == false ? InputPath.Text : InputPath.Text + "\\";
            //图片文件输出路径
            string PdfOutputPath = OutPath.Text + "\\";
            //用于检测是否完成Task任务
            TaskFactory taskFactory = new TaskFactory();
            if (InputCheckBox.Checked)//整个文件夹
            {          
                 try
                 {
                    //获取有效文件数
                    List<FileInfo> files = new List<FileInfo>();
                    //检测的文件类型
                    string[] es = { ".docx", ".doc", ".dot", ".docm" , ".dotm" , ".dotx" };
                    //执行检测方法，获取FileInfo对象列表
                    files = File_Name(WordInputPath, es, files);
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
                        lock (TestLock_Lock)
                        {
                         
                            for (int count = 0; count < files.Count; count++)
                            {
                                //加载过滤完的文件夹
                                FileInfo file = files[count] as FileInfo;
                                //Word转Img
                                WordTo_Img.WordToImg(file.FullName, PdfOutputPath, ImageFormat.Jpeg);
                                //进度条增进
                                if (count <= Convert.ToInt32(Math.Floor(Convert.ToDecimal(files.Count / 2))))
                                {
                                    //如果处理的文件不到一半，那就定在50%
                                    ProgressBar.Value = count >= 50 ? 50 : count;
                                }
                                else
                                {
                                    ProgressBar.Value = count >= 98 ? 98 : count;
                                }
                            }
                        }
                    }));
                    //等着全部任务完成后，启动一个新的task来完成后续动作
                    taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                    {
                        st1.Stop();//结束计时

                        ProgressBar.Value = 100;//结束后进度条拉满

                        double time = st1.ElapsedMilliseconds;//总用时

                        //解析出具体时间
                        int[] Time = Tiem_num(time);
                        //提示完成并告知具体运行时间
                        if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                        {
                            ProgressBar.Hide();
                            //恢复界面
                            Button_True();
                        }
                    });
                 }
                 catch (Exception)
                 {

                      MessageBox.Show("转换失败！");
                    //恢复界面
                    Button_True();
                    return;
                }
            }
            else       //单个文件
            {
                try
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
                        lock (TestLock_Lock)
                        {
                            //检查文件类型
                            FileInfo file = new FileInfo(WordInputPath);
                            //文件类型                          
                            string fileExs = file.Extension.ToLower();
                            if (fileExs == ".docx" || fileExs == ".doc" || fileExs == ".dot" || fileExs == ".docm" || fileExs == ".dotm" || fileExs == ".dotx")
                            {
                                WordTo_Img.WordToImg(WordInputPath, PdfOutputPath, ImageFormat.Jpeg);
                            }
                            else
                            {
                                MessageBox.Show("文件类型错误！");
                                //恢复界面
                                Button_True();
                                return;
                            }
                        }
                    }));
                    //等着全部任务完成后，启动一个新的task来完成后续动作
                    taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                    {
                        st1.Stop();//结束计时

                        //进度条增加
                        ProgressBar.Value = 50;
                        Thread.Sleep(1000);
                        ProgressBar.Value = 100;//结束后进度条拉满

                        double time = st1.ElapsedMilliseconds;//总用时

                        //解析出具体时间
                        int[] Time = Tiem_num(time);
                        //提示完成并告知具体运行时间
                        if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                        {
                            ProgressBar.Hide();
                        }
                    });
                }
                catch (Exception)
                {

                    MessageBox.Show("转换失败");
                    //恢复界面
                    Button_True();
                    return;
                }
            }
        }

        /// <summary>
        /// Word转PDF
        /// </summary>
        private void Word_Pdf_radio_Fun()
        {
            //Word文件路径
            string WordInputPath = InputCheckBox.Checked == false ? InputPath.Text : InputPath.Text + "\\";
            //PDF文件输出路径
            string PdfOutputPath = OutPath.Text + "\\";
            //用于检测是否完成Task任务
            TaskFactory taskFactory = new TaskFactory();
            if (InputCheckBox.Checked)//整个文件夹
            {
                try
                {
                    //获取有效文件数
                    List<FileInfo> files = new List<FileInfo>();
                    //检测的文件类型
                    string[] es = { ".docx", ".doc", ".dot", ".docm", ".dotm", ".dotx" };
                    //执行检测方法，获取FileInfo对象列表
                    files = File_Name(WordInputPath, es, files);
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
                        lock (TestLock_Lock)
                        {

                            for (int count = 0; count < files.Count; count++)
                            {
                                //加载过滤完的文件夹
                                FileInfo file = files[count] as FileInfo;
                                if (WordTo_Pdf.WordToPDF(file.FullName, PdfOutputPath, file.Name))
                                {
                                    //进度条增进
                                    if (count <= Convert.ToInt32(Math.Floor(Convert.ToDecimal(files.Count / 2))))
                                    {
                                        //如果处理的文件不到一半，那就定在50%
                                        ProgressBar.Value = count >= 50 ? 50 : count;
                                    }
                                    else
                                    {
                                        ProgressBar.Value = count >= 98 ? 98 : count;
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("转换失败！");
                                    //恢复界面
                                    Button_True();
                                }

                            }
                        }
                    }));
                    //等着全部任务完成后，启动一个新的task来完成后续动作
                    taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                    {
                        st1.Stop();//结束计时

                        ProgressBar.Value = 100;//结束后进度条拉满

                        double time = st1.ElapsedMilliseconds;//总用时

                        //解析出具体时间
                        int[] Time = Tiem_num(time);
                        //提示完成并告知具体运行时间
                        if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                        {
                            ProgressBar.Hide();
                            //恢复界面
                            Button_True();
                        }
                    });
                }
                catch (Exception)
                {
                    MessageBox.Show("转换失败！");
                    //恢复界面
                    Button_True();
                    return;
                }
            }
            else       //单个文件
            {
                try
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
                        lock (TestLock_Lock)
                        {
                            //检查文件类型
                            FileInfo file = new FileInfo(WordInputPath);
                            //文件类型
                            string fileExs = file.Extension.ToLower();
                            if (fileExs == ".docx" || fileExs == ".doc" || fileExs == ".dot" || fileExs == ".docm" || fileExs == ".dotm" || fileExs == ".dotx")
                            {
                                WordTo_Pdf.WordToPDF(file.FullName, PdfOutputPath, file.Name);
                            }
                            else
                            {
                                MessageBox.Show("文件类型错误！");
                                //恢复界面
                                Button_True();
                                return;
                            }
                        }
                    }));
                    //等着全部任务完成后，启动一个新的task来完成后续动作
                    taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                    {
                        st1.Stop();//结束计时

                        //进度条增加
                        ProgressBar.Value = 50;
                        Thread.Sleep(1000);
                        ProgressBar.Value = 100;//结束后进度条拉满

                        double time = st1.ElapsedMilliseconds;//总用时

                        //解析出具体时间
                        int[] Time = Tiem_num(time);
                        //提示完成并告知具体运行时间
                        if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                        {
                            ProgressBar.Hide();
                            //恢复界面
                            Button_True();
                        }
                    });
                }
                catch (Exception)
                {
                    MessageBox.Show("转换失败");
                    //恢复界面
                    Button_True();
                    return;
                }
             
            }
        }

        /// <summary>
        /// Excel转PDF
        /// </summary>
        private void Excel_Pdf_radio_Fun()
        {
            //Excel文件路径
            string ExcelInputPath = InputCheckBox.Checked == false ? InputPath.Text : InputPath.Text + "\\";
            //PDF文件输出路径
            string PdfOutputPath = OutPath.Text + "\\";
            //用于检测是否完成Task任务
            TaskFactory taskFactory = new TaskFactory();
            if (InputCheckBox.Checked)//整个文件夹
            {
                try
                {
                    //获取有效文件数
                    List<FileInfo> files = new List<FileInfo>();
                    //检测的文件类型
                    string[] es = { ".xls", ".xlsx"};
                    //执行检测方法，获取FileInfo对象列表
                    files = File_Name(ExcelInputPath, es, files);
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
                        lock (TestLock_Lock)
                        {
                            for (int count = 0; count < files.Count; count++)
                            {
                                //加载过滤完的文件夹
                                FileInfo file = files[count] as FileInfo;
                                //获取当前遍历到的文件的无后缀文件名
                                string FileName = Path.GetFileNameWithoutExtension(file.Name);
                                //Excel转PDF文件，并将PDF文件转换为图片文件 ExcelToPDF(Excel文件路径，图片输出路径，PDF输出路径，图片名称)
                                ExcelTo_PDF.ExcelToPDF(file.FullName, PdfOutputPath, FileName);
                                //进度条增进
                                if (count <= Convert.ToInt32(Math.Floor(Convert.ToDecimal(files.Count / 2))))
                                {
                                    //如果处理的文件不到一半，那就定在50%
                                    ProgressBar.Value = count >= 50 ? 50 : count;
                                }
                                else
                                {
                                    ProgressBar.Value = count >= 98 ? 98 : count;
                                }
                            }
                        }
                    }));
                    //等着全部任务完成后，启动一个新的task来完成后续动作
                    taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                    {
                        st1.Stop();//结束计时

                        ProgressBar.Value = 100;//结束后进度条拉满

                        double time = st1.ElapsedMilliseconds;//总用时

                        //解析出具体时间
                        int[] Time = Tiem_num(time);
                        //提示完成并告知具体运行时间
                        if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                        {
                            ProgressBar.Hide();
                            //恢复界面
                            Button_True();
                        }
                    });
                }
                catch (Exception)
                {
                    MessageBox.Show("转换失败！");
                    //恢复界面
                    Button_True();
                    return;
                }
            }
            else       //单个文件夹
            {
                try
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
                        lock (TestLock_Lock)
                        {
                            //检查文件类型
                            FileInfo file = new FileInfo(ExcelInputPath);

                            if (file.Extension == ".xls" || file.Extension == ".xlsx" || file.Extension == ".XLS" || file.Extension == ".XLSX")
                            {
                                //获取当前遍历到的文件的无后缀文件名
                                string FileName = Path.GetFileNameWithoutExtension(file.Name);
                                //Excel转PDF文件，并将PDF文件转换为图片文件 ExcelToPDF(Excel文件路径，图片输出路径，PDF输出路径，图片名称)
                                ExcelTo_PDF.ExcelToPDF(file.FullName, PdfOutputPath, FileName);
                            }
                            else
                            {
                                MessageBox.Show("文件类型错误！");
                                //恢复界面
                                Button_True();
                                return;
                            }
                        }
                    }));
                    //等着全部任务完成后，启动一个新的task来完成后续动作
                    taskFactory.ContinueWhenAll(taskList.ToArray(), tArray =>
                    {
                        st1.Stop();//结束计时

                        //进度条增加
                        ProgressBar.Value = 50;
                        Thread.Sleep(1000);
                        ProgressBar.Value = 100;//结束后进度条拉满

                        double time = st1.ElapsedMilliseconds;//总用时

                        //解析出具体时间
                        int[] Time = Tiem_num(time);
                        //提示完成并告知具体运行时间
                        if (MessageBox.Show($" \t用时:\n\n  {Time[0]}小时{Time[1]}分钟{Time[2]}秒{Time[3]}毫秒", "文件转换完成") == DialogResult.OK)
                        {
                            ProgressBar.Hide();
                            //恢复界面
                            Button_True();
                        }
                    });
                }
                catch (Exception)
                {
                    MessageBox.Show("转换失败");
                    //恢复界面
                    Button_True();
                    return;
                }          
            }
        }



        /******************************************附加功能区******************************************/
        /// <summary>
        /// 执行中的操作界面
        /// </summary>
        private void Button_False()
        {
            //执行按钮
            Read_Button.Text = "执行中";
            Read_Button.Enabled = false;
            Read_Button.Cursor = Cursors.No;


            //压缩图片按钮
            Compress_Img.Enabled = false;
            Compress_Img.Cursor = Cursors.No;

            //输入路径按钮
            InputPathButton.Enabled = false;
            InputPathButton.Cursor = Cursors.No;

            //输出路径按钮
            OutPathButton.Enabled = false;
            OutPathButton.Cursor = Cursors.No;

            //恢复默认按钮
            Default_Button.Enabled = false;
            Default_Button.Cursor = Cursors.No;

            //选择单文件和多文件按钮
            InputCheckBox.Enabled = false;
            InputCheckBox.Cursor = Cursors.No;

            //输入路径输入框
            InputPath.Enabled = false;
            InputPath.Cursor = Cursors.No;

            //输出路径输入框
            OutPath.Enabled = false;
            OutPath.Cursor = Cursors.No;

            //功能按钮按钮
            Pdf_Img_radio.Enabled = false;
            Word_Pdf_radio.Enabled = false;
            Excel_Img_radio.Enabled = false;
            Excel_Pdf_radio.Enabled = false;
            Word_Img_radio.Enabled = false;

        }
        /// <summary>
        /// 恢复操作界面
        /// </summary>
        private void Button_True()
        {
            //执行按钮
            Read_Button.Text = "开始执行";
            Read_Button.Enabled = true;
            Read_Button.Cursor = Cursors.Hand;
            

            //压缩图片按钮
            Compress_Img.Enabled = true;
            Compress_Img.Cursor = Cursors.Hand;

            //输入路径按钮
            InputPathButton.Enabled = true;
            InputPathButton.Cursor = Cursors.Hand;

            //输出路径按钮
            OutPathButton.Enabled = true;
            OutPathButton.Cursor = Cursors.Hand;

            //恢复默认按钮
            Default_Button.Enabled = true;
            Default_Button.Cursor = Cursors.Hand;

            //选择单文件和多文件按钮
            InputCheckBox.Enabled = true;
            InputCheckBox.Cursor = Cursors.Hand;

            //输入路径输入框
            InputPath.Enabled = true;
            InputPath.Cursor = Cursors.Hand;
            //输出路径输入框
            OutPath.Enabled = true;
            OutPath.Cursor = Cursors.Hand;
            //功能按钮按钮
            Pdf_Img_radio.Enabled = true;
            Word_Pdf_radio.Enabled = true;
            Excel_Img_radio.Enabled = true;
            Excel_Pdf_radio.Enabled = true;
            Word_Img_radio.Enabled = true;

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

        /// <summary>
        /// 获取当前目录以及子目录中对应类型的文件
        /// </summary>
        /// <param name="path"></param>
        /// <param name="extName"></param>
        /// <param name="lst"></param>
        /// <returns></returns>
        private List<FileInfo> File_Name(string path, string[] extName, List<FileInfo> lst)
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
                                if (extName[i].ToLower().IndexOf(File.Extension.ToLower()) >= 0)
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

                return null;
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

    }
}
