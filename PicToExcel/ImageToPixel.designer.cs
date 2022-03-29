namespace PicToExcel
{
    partial class ImageToPixel
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ImageToPixel));
            this.pOriginal = new System.Windows.Forms.PictureBox();
            this.txtImgPath = new System.Windows.Forms.TextBox();
            this.btnOpenImg = new System.Windows.Forms.Button();
            this.pCurrent = new System.Windows.Forms.Panel();
            this.btnConvert = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblX = new System.Windows.Forms.Label();
            this.lblY = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblW = new System.Windows.Forms.Label();
            this.lblH = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.txtExcelPath = new System.Windows.Forms.TextBox();
            this.btnOpenExcel = new System.Windows.Forms.Button();
            this.bar = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.pOriginal)).BeginInit();
            this.SuspendLayout();
            // 
            // pOriginal
            // 
            this.pOriginal.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(25)))), ((int)(((byte)(25)))), ((int)(((byte)(25)))));
            this.pOriginal.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pOriginal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pOriginal.Location = new System.Drawing.Point(12, 87);
            this.pOriginal.Name = "pOriginal";
            this.pOriginal.Size = new System.Drawing.Size(295, 397);
            this.pOriginal.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pOriginal.TabIndex = 0;
            this.pOriginal.TabStop = false;
            // 
            // txtImgPath
            // 
            this.txtImgPath.BackColor = System.Drawing.SystemColors.InfoText;
            this.txtImgPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtImgPath.ForeColor = System.Drawing.SystemColors.Info;
            this.txtImgPath.Location = new System.Drawing.Point(12, 22);
            this.txtImgPath.Name = "txtImgPath";
            this.txtImgPath.Size = new System.Drawing.Size(354, 21);
            this.txtImgPath.TabIndex = 1;
            // 
            // btnOpenImg
            // 
            this.btnOpenImg.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.btnOpenImg.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnOpenImg.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Crimson;
            this.btnOpenImg.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Yellow;
            this.btnOpenImg.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOpenImg.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btnOpenImg.Location = new System.Drawing.Point(372, 21);
            this.btnOpenImg.Name = "btnOpenImg";
            this.btnOpenImg.Size = new System.Drawing.Size(75, 23);
            this.btnOpenImg.TabIndex = 2;
            this.btnOpenImg.Text = "打开图片";
            this.btnOpenImg.UseVisualStyleBackColor = false;
            this.btnOpenImg.Click += new System.EventHandler(this.btnOpenImg_Click);
            // 
            // pCurrent
            // 
            this.pCurrent.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(25)))), ((int)(((byte)(25)))), ((int)(((byte)(25)))));
            this.pCurrent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pCurrent.Location = new System.Drawing.Point(442, 87);
            this.pCurrent.Name = "pCurrent";
            this.pCurrent.Size = new System.Drawing.Size(295, 397);
            this.pCurrent.TabIndex = 3;
            // 
            // btnConvert
            // 
            this.btnConvert.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.btnConvert.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnConvert.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Crimson;
            this.btnConvert.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Yellow;
            this.btnConvert.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnConvert.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btnConvert.Location = new System.Drawing.Point(631, 21);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(91, 23);
            this.btnConvert.TabIndex = 2;
            this.btnConvert.Text = "测试专用";
            this.btnConvert.UseVisualStyleBackColor = false;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.btnSave.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnSave.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Crimson;
            this.btnSave.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Yellow;
            this.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSave.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btnSave.Location = new System.Drawing.Point(453, 21);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(91, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "保存像素文件";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(358, 87);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "X:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(358, 117);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(17, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "Y:";
            // 
            // lblX
            // 
            this.lblX.AutoSize = true;
            this.lblX.BackColor = System.Drawing.Color.Transparent;
            this.lblX.ForeColor = System.Drawing.Color.White;
            this.lblX.Location = new System.Drawing.Point(381, 87);
            this.lblX.Name = "lblX";
            this.lblX.Size = new System.Drawing.Size(11, 12);
            this.lblX.TabIndex = 4;
            this.lblX.Text = "0";
            // 
            // lblY
            // 
            this.lblY.AutoSize = true;
            this.lblY.BackColor = System.Drawing.Color.Transparent;
            this.lblY.ForeColor = System.Drawing.Color.White;
            this.lblY.Location = new System.Drawing.Point(381, 117);
            this.lblY.Name = "lblY";
            this.lblY.Size = new System.Drawing.Size(11, 12);
            this.lblY.TabIndex = 4;
            this.lblY.Text = "0";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(334, 147);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 5;
            this.label3.Text = "Width:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(328, 177);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(47, 12);
            this.label4.TabIndex = 5;
            this.label4.Text = "Height:";
            // 
            // lblW
            // 
            this.lblW.AutoSize = true;
            this.lblW.BackColor = System.Drawing.Color.Transparent;
            this.lblW.ForeColor = System.Drawing.Color.White;
            this.lblW.Location = new System.Drawing.Point(381, 147);
            this.lblW.Name = "lblW";
            this.lblW.Size = new System.Drawing.Size(11, 12);
            this.lblW.TabIndex = 6;
            this.lblW.Text = "0";
            // 
            // lblH
            // 
            this.lblH.AutoSize = true;
            this.lblH.BackColor = System.Drawing.Color.Transparent;
            this.lblH.ForeColor = System.Drawing.Color.White;
            this.lblH.Location = new System.Drawing.Point(381, 177);
            this.lblH.Name = "lblH";
            this.lblH.Size = new System.Drawing.Size(11, 12);
            this.lblH.TabIndex = 6;
            this.lblH.Text = "0";
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.btnCancel.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnCancel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Crimson;
            this.btnCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Yellow;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btnCancel.Location = new System.Drawing.Point(550, 21);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "中断并保存";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // txtExcelPath
            // 
            this.txtExcelPath.BackColor = System.Drawing.SystemColors.InfoText;
            this.txtExcelPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtExcelPath.ForeColor = System.Drawing.SystemColors.Info;
            this.txtExcelPath.Location = new System.Drawing.Point(12, 49);
            this.txtExcelPath.Name = "txtExcelPath";
            this.txtExcelPath.Size = new System.Drawing.Size(354, 21);
            this.txtExcelPath.TabIndex = 1;
            // 
            // btnOpenExcel
            // 
            this.btnOpenExcel.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.btnOpenExcel.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnOpenExcel.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Crimson;
            this.btnOpenExcel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Yellow;
            this.btnOpenExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOpenExcel.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btnOpenExcel.Location = new System.Drawing.Point(372, 48);
            this.btnOpenExcel.Name = "btnOpenExcel";
            this.btnOpenExcel.Size = new System.Drawing.Size(75, 23);
            this.btnOpenExcel.TabIndex = 2;
            this.btnOpenExcel.Text = "打开Excel";
            this.btnOpenExcel.UseVisualStyleBackColor = false;
            this.btnOpenExcel.Click += new System.EventHandler(this.btnOpenExcel_Click);
            // 
            // bar
            // 
            this.bar.BackColor = System.Drawing.SystemColors.ControlText;
            this.bar.Dock = System.Windows.Forms.DockStyle.Top;
            this.bar.Location = new System.Drawing.Point(0, 0);
            this.bar.Name = "bar";
            this.bar.Size = new System.Drawing.Size(749, 10);
            this.bar.Step = 1;
            this.bar.TabIndex = 8;
            // 
            // ImageToPixel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(749, 496);
            this.Controls.Add(this.bar);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblH);
            this.Controls.Add(this.lblW);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblY);
            this.Controls.Add(this.lblX);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pCurrent);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnConvert);
            this.Controls.Add(this.btnOpenExcel);
            this.Controls.Add(this.btnOpenImg);
            this.Controls.Add(this.txtExcelPath);
            this.Controls.Add(this.txtImgPath);
            this.Controls.Add(this.pOriginal);
            this.MaximizeBox = false;
            this.Name = "ImageToPixel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "像素分割";
            this.Load += new System.EventHandler(this.ImageToPixel_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pOriginal)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pOriginal;
        private System.Windows.Forms.TextBox txtImgPath;
        private System.Windows.Forms.Button btnOpenImg;
        private System.Windows.Forms.Panel pCurrent;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblX;
        private System.Windows.Forms.Label lblY;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblW;
        private System.Windows.Forms.Label lblH;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox txtExcelPath;
        private System.Windows.Forms.Button btnOpenExcel;
        private System.Windows.Forms.ProgressBar bar;
    }
}

