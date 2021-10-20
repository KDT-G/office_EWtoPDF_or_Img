namespace Office文件转换
{
    partial class Office_Demo
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Office_Demo));
            this.ProgressBar = new System.Windows.Forms.ProgressBar();
            this.InputPathButton = new System.Windows.Forms.Button();
            this.InputPath = new System.Windows.Forms.TextBox();
            this.OutPath = new System.Windows.Forms.TextBox();
            this.OutPathButton = new System.Windows.Forms.Button();
            this.InputPath_Txt = new System.Windows.Forms.Label();
            this.OutPath_Txt = new System.Windows.Forms.Label();
            this.Pdf_Img_radio = new System.Windows.Forms.RadioButton();
            this.Word_Img_radio = new System.Windows.Forms.RadioButton();
            this.Excel_Img_radio = new System.Windows.Forms.RadioButton();
            this.Word_Pdf_radio = new System.Windows.Forms.RadioButton();
            this.Excel_Pdf_radio = new System.Windows.Forms.RadioButton();
            this.label3 = new System.Windows.Forms.Label();
            this.Fileico = new System.Windows.Forms.Label();
            this.InputCheckBox = new System.Windows.Forms.CheckBox();
            this.Default_Button = new System.Windows.Forms.Button();
            this.Compress_Img = new System.Windows.Forms.Button();
            this.Read_Button = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.File_ico = new System.Windows.Forms.PictureBox();
            this.Pdf_Word_radio = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.File_ico)).BeginInit();
            this.SuspendLayout();
            // 
            // ProgressBar
            // 
            this.ProgressBar.Enabled = false;
            this.ProgressBar.Location = new System.Drawing.Point(0, 436);
            this.ProgressBar.Name = "ProgressBar";
            this.ProgressBar.Size = new System.Drawing.Size(800, 15);
            this.ProgressBar.TabIndex = 0;
            // 
            // InputPathButton
            // 
            this.InputPathButton.BackColor = System.Drawing.Color.LightGray;
            this.InputPathButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.InputPathButton.FlatAppearance.BorderColor = System.Drawing.Color.LightGray;
            this.InputPathButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.InputPathButton.Location = new System.Drawing.Point(417, 49);
            this.InputPathButton.Name = "InputPathButton";
            this.InputPathButton.Size = new System.Drawing.Size(75, 24);
            this.InputPathButton.TabIndex = 1;
            this.InputPathButton.Tag = "";
            this.InputPathButton.Text = "选择文件";
            this.InputPathButton.UseVisualStyleBackColor = false;
            this.InputPathButton.Click += new System.EventHandler(this.InputPathButton_Click);
            // 
            // InputPath
            // 
            this.InputPath.AllowDrop = true;
            this.InputPath.CausesValidation = false;
            this.InputPath.Cursor = System.Windows.Forms.Cursors.Default;
            this.InputPath.Location = new System.Drawing.Point(130, 49);
            this.InputPath.Multiline = true;
            this.InputPath.Name = "InputPath";
            this.InputPath.Size = new System.Drawing.Size(267, 24);
            this.InputPath.TabIndex = 2;
            this.InputPath.TabStop = false;
            this.InputPath.Text = "选择文件";
            this.InputPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.InputPath.WordWrap = false;
            this.InputPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.InputPath_DragDrop);
            this.InputPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.InputPath_DragEnter);
            // 
            // OutPath
            // 
            this.OutPath.AllowDrop = true;
            this.OutPath.CausesValidation = false;
            this.OutPath.Cursor = System.Windows.Forms.Cursors.Default;
            this.OutPath.Location = new System.Drawing.Point(128, 129);
            this.OutPath.Multiline = true;
            this.OutPath.Name = "OutPath";
            this.OutPath.Size = new System.Drawing.Size(267, 24);
            this.OutPath.TabIndex = 4;
            this.OutPath.TabStop = false;
            this.OutPath.Text = "选择文件夹";
            this.OutPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.OutPath.WordWrap = false;
            this.OutPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.OutPath_DragDrop);
            this.OutPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.OutPath_DragEnter);
            // 
            // OutPathButton
            // 
            this.OutPathButton.BackColor = System.Drawing.Color.LightGray;
            this.OutPathButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.OutPathButton.FlatAppearance.BorderColor = System.Drawing.Color.LightGray;
            this.OutPathButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OutPathButton.Location = new System.Drawing.Point(415, 129);
            this.OutPathButton.Name = "OutPathButton";
            this.OutPathButton.Size = new System.Drawing.Size(75, 24);
            this.OutPathButton.TabIndex = 3;
            this.OutPathButton.Tag = "";
            this.OutPathButton.Text = "选择文件夹";
            this.OutPathButton.UseVisualStyleBackColor = false;
            this.OutPathButton.Click += new System.EventHandler(this.OutPathButton_Click);
            // 
            // InputPath_Txt
            // 
            this.InputPath_Txt.AutoSize = true;
            this.InputPath_Txt.Location = new System.Drawing.Point(128, 34);
            this.InputPath_Txt.Name = "InputPath_Txt";
            this.InputPath_Txt.Size = new System.Drawing.Size(65, 12);
            this.InputPath_Txt.TabIndex = 5;
            this.InputPath_Txt.Text = "输入路径：";
            // 
            // OutPath_Txt
            // 
            this.OutPath_Txt.AutoSize = true;
            this.OutPath_Txt.Location = new System.Drawing.Point(128, 111);
            this.OutPath_Txt.Name = "OutPath_Txt";
            this.OutPath_Txt.Size = new System.Drawing.Size(65, 12);
            this.OutPath_Txt.TabIndex = 6;
            this.OutPath_Txt.Text = "输出路径：";
            // 
            // Pdf_Img_radio
            // 
            this.Pdf_Img_radio.AutoSize = true;
            this.Pdf_Img_radio.Checked = true;
            this.Pdf_Img_radio.Cursor = System.Windows.Forms.Cursors.Default;
            this.Pdf_Img_radio.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Pdf_Img_radio.Location = new System.Drawing.Point(147, 226);
            this.Pdf_Img_radio.Name = "Pdf_Img_radio";
            this.Pdf_Img_radio.Size = new System.Drawing.Size(85, 21);
            this.Pdf_Img_radio.TabIndex = 7;
            this.Pdf_Img_radio.TabStop = true;
            this.Pdf_Img_radio.Tag = "";
            this.Pdf_Img_radio.Text = "PDF转图片";
            this.Pdf_Img_radio.UseCompatibleTextRendering = true;
            this.Pdf_Img_radio.UseVisualStyleBackColor = true;
            // 
            // Word_Img_radio
            // 
            this.Word_Img_radio.AutoSize = true;
            this.Word_Img_radio.Cursor = System.Windows.Forms.Cursors.Default;
            this.Word_Img_radio.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Word_Img_radio.Location = new System.Drawing.Point(147, 270);
            this.Word_Img_radio.Name = "Word_Img_radio";
            this.Word_Img_radio.Size = new System.Drawing.Size(95, 18);
            this.Word_Img_radio.TabIndex = 8;
            this.Word_Img_radio.Tag = "";
            this.Word_Img_radio.Text = "Word转图片";
            this.Word_Img_radio.UseVisualStyleBackColor = true;
            // 
            // Excel_Img_radio
            // 
            this.Excel_Img_radio.AutoSize = true;
            this.Excel_Img_radio.Cursor = System.Windows.Forms.Cursors.Default;
            this.Excel_Img_radio.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Excel_Img_radio.Location = new System.Drawing.Point(147, 248);
            this.Excel_Img_radio.Name = "Excel_Img_radio";
            this.Excel_Img_radio.Size = new System.Drawing.Size(102, 18);
            this.Excel_Img_radio.TabIndex = 9;
            this.Excel_Img_radio.Tag = "";
            this.Excel_Img_radio.Text = "Excel转图片";
            this.Excel_Img_radio.UseVisualStyleBackColor = true;
            // 
            // Word_Pdf_radio
            // 
            this.Word_Pdf_radio.AutoSize = true;
            this.Word_Pdf_radio.Cursor = System.Windows.Forms.Cursors.Default;
            this.Word_Pdf_radio.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Word_Pdf_radio.Location = new System.Drawing.Point(300, 226);
            this.Word_Pdf_radio.Name = "Word_Pdf_radio";
            this.Word_Pdf_radio.Size = new System.Drawing.Size(88, 18);
            this.Word_Pdf_radio.TabIndex = 10;
            this.Word_Pdf_radio.Tag = "";
            this.Word_Pdf_radio.Text = "Word转PDF";
            this.Word_Pdf_radio.UseVisualStyleBackColor = true;
            // 
            // Excel_Pdf_radio
            // 
            this.Excel_Pdf_radio.AutoSize = true;
            this.Excel_Pdf_radio.Cursor = System.Windows.Forms.Cursors.Default;
            this.Excel_Pdf_radio.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Excel_Pdf_radio.Location = new System.Drawing.Point(300, 248);
            this.Excel_Pdf_radio.Name = "Excel_Pdf_radio";
            this.Excel_Pdf_radio.Size = new System.Drawing.Size(95, 18);
            this.Excel_Pdf_radio.TabIndex = 11;
            this.Excel_Pdf_radio.Tag = "";
            this.Excel_Pdf_radio.Text = "Excel转PDF";
            this.Excel_Pdf_radio.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(130, 193);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 16);
            this.label3.TabIndex = 12;
            this.label3.Text = "功能：";
            // 
            // Fileico
            // 
            this.Fileico.AutoSize = true;
            this.Fileico.Location = new System.Drawing.Point(610, 34);
            this.Fileico.Name = "Fileico";
            this.Fileico.Size = new System.Drawing.Size(53, 12);
            this.Fileico.TabIndex = 15;
            this.Fileico.Text = "文件图标";
            // 
            // InputCheckBox
            // 
            this.InputCheckBox.AutoSize = true;
            this.InputCheckBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.InputCheckBox.Location = new System.Drawing.Point(130, 79);
            this.InputCheckBox.Name = "InputCheckBox";
            this.InputCheckBox.Size = new System.Drawing.Size(192, 16);
            this.InputCheckBox.TabIndex = 17;
            this.InputCheckBox.Text = "是否选择当前文件夹内全部文件";
            this.InputCheckBox.UseVisualStyleBackColor = true;
            // 
            // Default_Button
            // 
            this.Default_Button.BackColor = System.Drawing.Color.LightGray;
            this.Default_Button.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Default_Button.FlatAppearance.BorderColor = System.Drawing.Color.LightGray;
            this.Default_Button.FlatAppearance.BorderSize = 0;
            this.Default_Button.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Default_Button.Location = new System.Drawing.Point(675, 375);
            this.Default_Button.Name = "Default_Button";
            this.Default_Button.Size = new System.Drawing.Size(92, 33);
            this.Default_Button.TabIndex = 18;
            this.Default_Button.Tag = "";
            this.Default_Button.Text = "恢复默认格式";
            this.Default_Button.UseVisualStyleBackColor = false;
            this.Default_Button.Click += new System.EventHandler(this.Default_Button_Click);
            // 
            // Compress_Img
            // 
            this.Compress_Img.BackColor = System.Drawing.Color.LightGray;
            this.Compress_Img.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Compress_Img.FlatAppearance.BorderColor = System.Drawing.Color.LightGray;
            this.Compress_Img.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Compress_Img.Location = new System.Drawing.Point(307, 299);
            this.Compress_Img.Name = "Compress_Img";
            this.Compress_Img.Size = new System.Drawing.Size(88, 23);
            this.Compress_Img.TabIndex = 19;
            this.Compress_Img.Tag = "";
            this.Compress_Img.Text = "图片压缩";
            this.Compress_Img.UseVisualStyleBackColor = false;
            this.Compress_Img.Click += new System.EventHandler(this.Compress_Img_Click);
            // 
            // Read_Button
            // 
            this.Read_Button.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Read_Button.Font = new System.Drawing.Font("宋体", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Read_Button.Location = new System.Drawing.Point(516, 226);
            this.Read_Button.Name = "Read_Button";
            this.Read_Button.Size = new System.Drawing.Size(147, 69);
            this.Read_Button.TabIndex = 20;
            this.Read_Button.Tag = "";
            this.Read_Button.Text = "开始执行";
            this.Read_Button.UseVisualStyleBackColor = true;
            this.Read_Button.Click += new System.EventHandler(this.Read_Button_Click);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Control;
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(133, 328);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ShortcutsEnabled = false;
            this.textBox1.Size = new System.Drawing.Size(385, 102);
            this.textBox1.TabIndex = 21;
            this.textBox1.Text = "提示：\r\n\r\n1.根据电脑性能不同和插件自身情况，存在一定几率失败，\r\n不保证百分百成功，如果执行失败,属于正常情况，对于Excel和Word一些复杂样式，失败几" +
    "率极大，谨慎选择\r\n\r\n2.如果对此有疑问，或者建议，亦或者想一块交流，\r\n微信搜索订阅号：菜鸟的温暖小窝   后台留言";
            // 
            // File_ico
            // 
            this.File_ico.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.File_ico.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.File_ico.Cursor = System.Windows.Forms.Cursors.No;
            this.File_ico.Enabled = false;
            this.File_ico.Image = global::Office文件转换.Properties.Resources.图标;
            this.File_ico.Location = new System.Drawing.Point(586, 49);
            this.File_ico.Name = "File_ico";
            this.File_ico.Size = new System.Drawing.Size(100, 100);
            this.File_ico.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.File_ico.TabIndex = 14;
            this.File_ico.TabStop = false;
            // 
            // Pdf_Word_radio
            // 
            this.Pdf_Word_radio.AutoSize = true;
            this.Pdf_Word_radio.Cursor = System.Windows.Forms.Cursors.Default;
            this.Pdf_Word_radio.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Pdf_Word_radio.Location = new System.Drawing.Point(300, 270);
            this.Pdf_Word_radio.Name = "Pdf_Word_radio";
            this.Pdf_Word_radio.Size = new System.Drawing.Size(88, 18);
            this.Pdf_Word_radio.TabIndex = 22;
            this.Pdf_Word_radio.Tag = "";
            this.Pdf_Word_radio.Text = "PDF转Wrod";
            this.Pdf_Word_radio.UseVisualStyleBackColor = true;
            // 
            // Office_Demo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.Pdf_Word_radio);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.Read_Button);
            this.Controls.Add(this.Compress_Img);
            this.Controls.Add(this.Default_Button);
            this.Controls.Add(this.InputCheckBox);
            this.Controls.Add(this.Fileico);
            this.Controls.Add(this.File_ico);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Excel_Pdf_radio);
            this.Controls.Add(this.Word_Pdf_radio);
            this.Controls.Add(this.Excel_Img_radio);
            this.Controls.Add(this.Word_Img_radio);
            this.Controls.Add(this.Pdf_Img_radio);
            this.Controls.Add(this.OutPath_Txt);
            this.Controls.Add(this.InputPath_Txt);
            this.Controls.Add(this.OutPath);
            this.Controls.Add(this.OutPathButton);
            this.Controls.Add(this.InputPath);
            this.Controls.Add(this.InputPathButton);
            this.Controls.Add(this.ProgressBar);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Office_Demo";
            this.Opacity = 0.95D;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Tag = "";
            this.Text = "Office文件转换——小菜鸟出品";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Office_Demo_FormClosing);
            this.Load += new System.EventHandler(this.Office_Demo_Load);
            ((System.ComponentModel.ISupportInitialize)(this.File_ico)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar ProgressBar;
        private System.Windows.Forms.Button InputPathButton;
        private System.Windows.Forms.Button OutPathButton;
        private System.Windows.Forms.Label InputPath_Txt;
        private System.Windows.Forms.Label OutPath_Txt;
        private System.Windows.Forms.RadioButton Pdf_Img_radio;
        private System.Windows.Forms.RadioButton Word_Img_radio;
        private System.Windows.Forms.RadioButton Excel_Img_radio;
        private System.Windows.Forms.RadioButton Word_Pdf_radio;
        private System.Windows.Forms.RadioButton Excel_Pdf_radio;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.PictureBox File_ico;
        private System.Windows.Forms.Label Fileico;
        private System.Windows.Forms.Button Default_Button;
        private System.Windows.Forms.Button Compress_Img;
        private System.Windows.Forms.Button Read_Button;
        private System.Windows.Forms.TextBox textBox1;
        internal System.Windows.Forms.TextBox OutPath;
        internal System.Windows.Forms.CheckBox InputCheckBox;
        public System.Windows.Forms.TextBox InputPath;
        private System.Windows.Forms.RadioButton Pdf_Word_radio;
    }
}

