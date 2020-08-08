namespace Office文件转Img_Pdf
{
    partial class Compress_Demo
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Compress_Demo));
            this.InputCheckBox = new System.Windows.Forms.CheckBox();
            this.OutPath_Txt = new System.Windows.Forms.Label();
            this.InputPath_Txt = new System.Windows.Forms.Label();
            this.OutPath = new System.Windows.Forms.TextBox();
            this.OutPathButton = new System.Windows.Forms.Button();
            this.InputPath = new System.Windows.Forms.TextBox();
            this.InputPathButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.Compress_Img = new System.Windows.Forms.Button();
            this.ProgressBar_ = new System.Windows.Forms.ProgressBar();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.Back_up = new System.Windows.Forms.Button();
            this.Img_H = new System.Windows.Forms.TextBox();
            this.Img_W = new System.Windows.Forms.TextBox();
            this.Img_Size = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // InputCheckBox
            // 
            this.InputCheckBox.AutoSize = true;
            this.InputCheckBox.Location = new System.Drawing.Point(176, 107);
            this.InputCheckBox.Name = "InputCheckBox";
            this.InputCheckBox.Size = new System.Drawing.Size(192, 16);
            this.InputCheckBox.TabIndex = 24;
            this.InputCheckBox.Text = "是否选择当前文件夹内全部文件";
            this.InputCheckBox.UseVisualStyleBackColor = true;
            // 
            // OutPath_Txt
            // 
            this.OutPath_Txt.AutoSize = true;
            this.OutPath_Txt.Location = new System.Drawing.Point(174, 139);
            this.OutPath_Txt.Name = "OutPath_Txt";
            this.OutPath_Txt.Size = new System.Drawing.Size(65, 12);
            this.OutPath_Txt.TabIndex = 23;
            this.OutPath_Txt.Text = "输出路径：";
            // 
            // InputPath_Txt
            // 
            this.InputPath_Txt.AutoSize = true;
            this.InputPath_Txt.Location = new System.Drawing.Point(174, 62);
            this.InputPath_Txt.Name = "InputPath_Txt";
            this.InputPath_Txt.Size = new System.Drawing.Size(65, 12);
            this.InputPath_Txt.TabIndex = 22;
            this.InputPath_Txt.Text = "输入路径：";
            // 
            // OutPath
            // 
            this.OutPath.AllowDrop = true;
            this.OutPath.CausesValidation = false;
            this.OutPath.Cursor = System.Windows.Forms.Cursors.Default;
            this.OutPath.Location = new System.Drawing.Point(174, 157);
            this.OutPath.Multiline = true;
            this.OutPath.Name = "OutPath";
            this.OutPath.Size = new System.Drawing.Size(267, 24);
            this.OutPath.TabIndex = 21;
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
            this.OutPathButton.Location = new System.Drawing.Point(461, 157);
            this.OutPathButton.Name = "OutPathButton";
            this.OutPathButton.Size = new System.Drawing.Size(75, 24);
            this.OutPathButton.TabIndex = 20;
            this.OutPathButton.Text = "选择文件夹";
            this.OutPathButton.UseVisualStyleBackColor = false;
            this.OutPathButton.Click += new System.EventHandler(this.OutPathButton_Click);
            // 
            // InputPath
            // 
            this.InputPath.AllowDrop = true;
            this.InputPath.CausesValidation = false;
            this.InputPath.Cursor = System.Windows.Forms.Cursors.Default;
            this.InputPath.Location = new System.Drawing.Point(176, 77);
            this.InputPath.Multiline = true;
            this.InputPath.Name = "InputPath";
            this.InputPath.Size = new System.Drawing.Size(267, 24);
            this.InputPath.TabIndex = 19;
            this.InputPath.TabStop = false;
            this.InputPath.Text = "选择文件";
            this.InputPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.InputPath.WordWrap = false;
            this.InputPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.InputPath_DragDrop);
            this.InputPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.InputPath_DragEnter);
            // 
            // InputPathButton
            // 
            this.InputPathButton.BackColor = System.Drawing.Color.LightGray;
            this.InputPathButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.InputPathButton.FlatAppearance.BorderColor = System.Drawing.Color.LightGray;
            this.InputPathButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.InputPathButton.Location = new System.Drawing.Point(463, 77);
            this.InputPathButton.Name = "InputPathButton";
            this.InputPathButton.Size = new System.Drawing.Size(75, 24);
            this.InputPathButton.TabIndex = 18;
            this.InputPathButton.Text = "选择文件";
            this.InputPathButton.UseVisualStyleBackColor = false;
            this.InputPathButton.Click += new System.EventHandler(this.InputPathButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(38, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 27);
            this.label1.TabIndex = 25;
            this.label1.Text = "图片压缩";
            // 
            // Compress_Img
            // 
            this.Compress_Img.BackColor = System.Drawing.Color.LightGray;
            this.Compress_Img.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Compress_Img.FlatAppearance.BorderColor = System.Drawing.Color.LightGray;
            this.Compress_Img.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Compress_Img.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Compress_Img.Location = new System.Drawing.Point(249, 325);
            this.Compress_Img.Name = "Compress_Img";
            this.Compress_Img.Size = new System.Drawing.Size(119, 39);
            this.Compress_Img.TabIndex = 26;
            this.Compress_Img.Text = "开始压缩";
            this.Compress_Img.UseVisualStyleBackColor = false;
            this.Compress_Img.Click += new System.EventHandler(this.Compress_Img_Click);
            // 
            // ProgressBar_
            // 
            this.ProgressBar_.Enabled = false;
            this.ProgressBar_.Location = new System.Drawing.Point(-1, 435);
            this.ProgressBar_.Name = "ProgressBar_";
            this.ProgressBar_.Size = new System.Drawing.Size(801, 14);
            this.ProgressBar_.TabIndex = 27;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(174, 209);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 12);
            this.label2.TabIndex = 31;
            this.label2.Text = "指定图片高度：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(174, 239);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 12);
            this.label3.TabIndex = 32;
            this.label3.Text = "指定图片宽度：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(159, 278);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(113, 12);
            this.label4.TabIndex = 35;
            this.label4.Text = "指定图片最大大小：";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(395, 278);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 12);
            this.label5.TabIndex = 37;
            this.label5.Text = "(单位：KB)";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(395, 239);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 12);
            this.label6.TabIndex = 38;
            this.label6.Text = "(单位：像素)";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(395, 209);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(77, 12);
            this.label7.TabIndex = 39;
            this.label7.Text = "(单位：像素)";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Control;
            this.textBox1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(523, 206);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(250, 204);
            this.textBox1.TabIndex = 40;
            this.textBox1.Text = "提示：\r\n\r\n图片会一定几率压缩图片的宽高，介意勿用！\r\n\r\n高度，宽度，图片大小可不填，有默认值\r\n\r\n高宽默认值取决于原图片大小\r\n\r\n图片最大大小默认值为3" +
    "00KB\r\n\r\n\r\n图片大小单位换算：\r\n1 MB = 1024 KB\r\n1 GB = 1024 MB\r\n";
            // 
            // Back_up
            // 
            this.Back_up.BackColor = System.Drawing.Color.LightGray;
            this.Back_up.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Back_up.FlatAppearance.BorderColor = System.Drawing.Color.LightGray;
            this.Back_up.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Back_up.Location = new System.Drawing.Point(698, 23);
            this.Back_up.Name = "Back_up";
            this.Back_up.Size = new System.Drawing.Size(75, 23);
            this.Back_up.TabIndex = 41;
            this.Back_up.Text = "返回主页面";
            this.Back_up.UseVisualStyleBackColor = false;
            this.Back_up.Click += new System.EventHandler(this.Back_up_Click);
            // 
            // Img_H
            // 
            this.Img_H.Location = new System.Drawing.Point(269, 206);
            this.Img_H.Name = "Img_H";
            this.Img_H.Size = new System.Drawing.Size(120, 21);
            this.Img_H.TabIndex = 42;
            // 
            // Img_W
            // 
            this.Img_W.Location = new System.Drawing.Point(269, 236);
            this.Img_W.Name = "Img_W";
            this.Img_W.Size = new System.Drawing.Size(120, 21);
            this.Img_W.TabIndex = 43;
            // 
            // Img_Size
            // 
            this.Img_Size.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.Img_Size.Location = new System.Drawing.Point(269, 275);
            this.Img_Size.Name = "Img_Size";
            this.Img_Size.Size = new System.Drawing.Size(120, 21);
            this.Img_Size.TabIndex = 44;
            // 
            // Compress_Demo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.Img_Size);
            this.Controls.Add(this.Img_W);
            this.Controls.Add(this.Img_H);
            this.Controls.Add(this.Back_up);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ProgressBar_);
            this.Controls.Add(this.Compress_Img);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.InputCheckBox);
            this.Controls.Add(this.OutPath_Txt);
            this.Controls.Add(this.InputPath_Txt);
            this.Controls.Add(this.OutPath);
            this.Controls.Add(this.OutPathButton);
            this.Controls.Add(this.InputPath);
            this.Controls.Add(this.InputPathButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Compress_Demo";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "图片压缩";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Compress_Demo_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox InputCheckBox;
        private System.Windows.Forms.Label OutPath_Txt;
        private System.Windows.Forms.Label InputPath_Txt;
        private System.Windows.Forms.TextBox OutPath;
        private System.Windows.Forms.Button OutPathButton;
        private System.Windows.Forms.TextBox InputPath;
        private System.Windows.Forms.Button InputPathButton;
        private System.Windows.Forms.Button Compress_Img;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button Back_up;
        private System.Windows.Forms.TextBox Img_H;
        private System.Windows.Forms.TextBox Img_W;
        private System.Windows.Forms.TextBox Img_Size;
        public System.Windows.Forms.ProgressBar ProgressBar_;
        public System.Windows.Forms.Label label1;
    }
}