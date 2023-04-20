namespace GetKeywords
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            this.btnStart = new System.Windows.Forms.Button();
            this.tmrPlan01 = new System.Windows.Forms.Timer(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtVolMax = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtSpeed = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lstStatus = new System.Windows.Forms.ListBox();
            this.dgrListKeywords = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clnChecked = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cboPlan = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnPause = new System.Windows.Forms.Button();
            this.btnAuto = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtKeywords = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.quickExportExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.clearDataGridToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.trợGiúpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tmrPlan02 = new System.Windows.Forms.Timer(this.components);
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.txttaikhoan = new System.Windows.Forms.TextBox();
            this.txtmatkhau = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tmrPlan03 = new System.Windows.Forms.Timer(this.components);
            this.label6 = new System.Windows.Forms.Label();
            this.txtdiachi = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tmrPlan04 = new System.Windows.Forms.Timer(this.components);
            this.txtCur = new System.Windows.Forms.TextBox();
            this.txtTotal = new System.Windows.Forms.TextBox();
            this.tmrPlan05 = new System.Windows.Forms.Timer(this.components);
            this.txtSuggestKey = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.càiĐặtToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ghiFileCàiĐặtToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.đọcFileCàiĐặtToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.đọcFileKịchBảnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgrListKeywords)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(91, 50);
            this.btnStart.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 23);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // tmrPlan01
            // 
            this.tmrPlan01.Interval = 2000;
            this.tmrPlan01.Tick += new System.EventHandler(this.tmrPlanGetKeyword);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtVolMax);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.txtSpeed);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(24, 25);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox1.Size = new System.Drawing.Size(165, 68);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            // 
            // txtVolMax
            // 
            this.txtVolMax.Location = new System.Drawing.Point(53, 38);
            this.txtVolMax.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtVolMax.Name = "txtVolMax";
            this.txtVolMax.Size = new System.Drawing.Size(57, 22);
            this.txtVolMax.TabIndex = 3;
            this.txtVolMax.Text = "1000";
            this.txtVolMax.TextChanged += new System.EventHandler(this.txtVolMax_TextChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(3, 41);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(52, 16);
            this.label8.TabIndex = 2;
            this.label8.Text = "VolMax";
            // 
            // txtSpeed
            // 
            this.txtSpeed.Location = new System.Drawing.Point(53, 11);
            this.txtSpeed.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtSpeed.Name = "txtSpeed";
            this.txtSpeed.Size = new System.Drawing.Size(57, 22);
            this.txtSpeed.TabIndex = 1;
            this.txtSpeed.Text = "1000";
            this.txtSpeed.TextChanged += new System.EventHandler(this.txtSpeed_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 14);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(48, 16);
            this.label2.TabIndex = 0;
            this.label2.Text = "Speed";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(203, 25);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 16);
            this.label3.TabIndex = 2;
            this.label3.Text = "Status";
            // 
            // lstStatus
            // 
            this.lstStatus.FormattingEnabled = true;
            this.lstStatus.ItemHeight = 16;
            this.lstStatus.Location = new System.Drawing.Point(205, 43);
            this.lstStatus.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.lstStatus.Name = "lstStatus";
            this.lstStatus.Size = new System.Drawing.Size(85, 52);
            this.lstStatus.TabIndex = 6;
            // 
            // dgrListKeywords
            // 
            this.dgrListKeywords.AllowUserToAddRows = false;
            this.dgrListKeywords.AllowUserToDeleteRows = false;
            this.dgrListKeywords.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dgrListKeywords.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgrListKeywords.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.clnChecked});
            this.dgrListKeywords.Cursor = System.Windows.Forms.Cursors.Default;
            this.dgrListKeywords.EnableHeadersVisualStyles = false;
            this.dgrListKeywords.Location = new System.Drawing.Point(15, 234);
            this.dgrListKeywords.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dgrListKeywords.Name = "dgrListKeywords";
            this.dgrListKeywords.ReadOnly = true;
            this.dgrListKeywords.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.dgrListKeywords.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.dgrListKeywords.RowHeadersVisible = false;
            this.dgrListKeywords.RowHeadersWidth = 51;
            this.dgrListKeywords.RowTemplate.Height = 24;
            this.dgrListKeywords.Size = new System.Drawing.Size(277, 388);
            this.dgrListKeywords.TabIndex = 7;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Keywords";
            this.Column1.MinimumWidth = 6;
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            this.Column1.Width = 125;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Search Volume (Average)";
            this.Column2.MinimumWidth = 6;
            this.Column2.Name = "Column2";
            this.Column2.ReadOnly = true;
            this.Column2.Width = 50;
            // 
            // clnChecked
            // 
            this.clnChecked.HeaderText = "Checked";
            this.clnChecked.MinimumWidth = 6;
            this.clnChecked.Name = "clnChecked";
            this.clnChecked.ReadOnly = true;
            this.clnChecked.Width = 30;
            // 
            // cboPlan
            // 
            this.cboPlan.FormattingEnabled = true;
            this.cboPlan.Location = new System.Drawing.Point(101, 21);
            this.cboPlan.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.cboPlan.Name = "cboPlan";
            this.cboPlan.Size = new System.Drawing.Size(159, 24);
            this.cboPlan.TabIndex = 8;
            this.cboPlan.Text = "[Lựa chọn kịch bản]";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnPause);
            this.groupBox2.Controls.Add(this.btnAuto);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.cboPlan);
            this.groupBox2.Controls.Add(this.btnStart);
            this.groupBox2.Location = new System.Drawing.Point(24, 146);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox2.Size = new System.Drawing.Size(267, 82);
            this.groupBox2.TabIndex = 9;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Thực hiện";
            // 
            // btnPause
            // 
            this.btnPause.Location = new System.Drawing.Point(9, 49);
            this.btnPause.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnPause.Name = "btnPause";
            this.btnPause.Size = new System.Drawing.Size(75, 27);
            this.btnPause.TabIndex = 11;
            this.btnPause.Text = "Pause";
            this.btnPause.UseVisualStyleBackColor = true;
            this.btnPause.Click += new System.EventHandler(this.btnPause_Click);
            // 
            // btnAuto
            // 
            this.btnAuto.Location = new System.Drawing.Point(171, 50);
            this.btnAuto.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnAuto.Name = "btnAuto";
            this.btnAuto.Size = new System.Drawing.Size(75, 23);
            this.btnAuto.TabIndex = 10;
            this.btnAuto.Text = "Auto";
            this.btnAuto.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(37, 25);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(58, 16);
            this.label4.TabIndex = 9;
            this.label4.Text = "Kịch bản";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(27, 122);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 16);
            this.label5.TabIndex = 0;
            this.label5.Text = "Keywords";
            // 
            // txtKeywords
            // 
            this.txtKeywords.Location = new System.Drawing.Point(97, 119);
            this.txtKeywords.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtKeywords.Name = "txtKeywords";
            this.txtKeywords.Size = new System.Drawing.Size(193, 22);
            this.txtKeywords.TabIndex = 1;
            this.txtKeywords.Text = "Hình nền windows đẹp";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(24, 97);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(267, 12);
            this.progressBar1.TabIndex = 10;
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.trợGiúpToolStripMenuItem,
            this.càiĐặtToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(5, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(311, 28);
            this.menuStrip1.TabIndex = 11;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openExcelToolStripMenuItem,
            this.saveExcelToolStripMenuItem,
            this.quickExportExcelToolStripMenuItem,
            this.toolStripMenuItem1,
            this.clearDataGridToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(46, 24);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // openExcelToolStripMenuItem
            // 
            this.openExcelToolStripMenuItem.Name = "openExcelToolStripMenuItem";
            this.openExcelToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.openExcelToolStripMenuItem.Text = "Import Excel";
            this.openExcelToolStripMenuItem.Click += new System.EventHandler(this.openExcelToolStripMenuItem_Click);
            // 
            // saveExcelToolStripMenuItem
            // 
            this.saveExcelToolStripMenuItem.Name = "saveExcelToolStripMenuItem";
            this.saveExcelToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.saveExcelToolStripMenuItem.Text = "Export Excel";
            this.saveExcelToolStripMenuItem.Click += new System.EventHandler(this.saveExcelToolStripMenuItem_Click);
            // 
            // quickExportExcelToolStripMenuItem
            // 
            this.quickExportExcelToolStripMenuItem.Name = "quickExportExcelToolStripMenuItem";
            this.quickExportExcelToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.quickExportExcelToolStripMenuItem.Text = "Quick Export Excel";
            this.quickExportExcelToolStripMenuItem.Click += new System.EventHandler(this.quickExportExcelToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(221, 6);
            // 
            // clearDataGridToolStripMenuItem
            // 
            this.clearDataGridToolStripMenuItem.Name = "clearDataGridToolStripMenuItem";
            this.clearDataGridToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.clearDataGridToolStripMenuItem.Text = "Clear DataGrid";
            this.clearDataGridToolStripMenuItem.Click += new System.EventHandler(this.clearDataGridToolStripMenuItem_Click);
            // 
            // trợGiúpToolStripMenuItem
            // 
            this.trợGiúpToolStripMenuItem.Name = "trợGiúpToolStripMenuItem";
            this.trợGiúpToolStripMenuItem.Size = new System.Drawing.Size(78, 24);
            this.trợGiúpToolStripMenuItem.Text = "Trợ giúp";
            // 
            // tmrPlan02
            // 
            this.tmrPlan02.Interval = 2000;
            this.tmrPlan02.Tick += new System.EventHandler(this.tmrPlanLogin);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // txttaikhoan
            // 
            this.txttaikhoan.Location = new System.Drawing.Point(100, 705);
            this.txttaikhoan.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txttaikhoan.Name = "txttaikhoan";
            this.txttaikhoan.Size = new System.Drawing.Size(191, 22);
            this.txttaikhoan.TabIndex = 12;
            this.txttaikhoan.Text = "chuminhtue10667@gmail.com";
            // 
            // txtmatkhau
            // 
            this.txtmatkhau.Location = new System.Drawing.Point(100, 742);
            this.txtmatkhau.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtmatkhau.Name = "txtmatkhau";
            this.txtmatkhau.Size = new System.Drawing.Size(88, 22);
            this.txtmatkhau.TabIndex = 13;
            this.txtmatkhau.Text = "guihAaat";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 746);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 16);
            this.label1.TabIndex = 14;
            this.label1.Text = "Password";
            // 
            // tmrPlan03
            // 
            this.tmrPlan03.Interval = 2000;
            this.tmrPlan03.Tick += new System.EventHandler(this.tmrPlanDowloadKeywordtieptheo);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(11, 709);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(41, 16);
            this.label6.TabIndex = 15;
            this.label6.Text = "Email";
            // 
            // txtdiachi
            // 
            this.txtdiachi.Location = new System.Drawing.Point(100, 668);
            this.txtdiachi.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtdiachi.Name = "txtdiachi";
            this.txtdiachi.Size = new System.Drawing.Size(192, 22);
            this.txtdiachi.TabIndex = 16;
            this.txtdiachi.Text = "https://keywordtool.io/";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(11, 672);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(64, 16);
            this.label7.TabIndex = 17;
            this.label7.Text = "Websites";
            // 
            // tmrPlan04
            // 
            this.tmrPlan04.Interval = 2000;
            this.tmrPlan04.Tick += new System.EventHandler(this.tmrPlanLoadLaiExcel);
            // 
            // txtCur
            // 
            this.txtCur.Location = new System.Drawing.Point(197, 742);
            this.txtCur.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtCur.Name = "txtCur";
            this.txtCur.Size = new System.Drawing.Size(44, 22);
            this.txtCur.TabIndex = 18;
            // 
            // txtTotal
            // 
            this.txtTotal.Location = new System.Drawing.Point(251, 742);
            this.txtTotal.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtTotal.Name = "txtTotal";
            this.txtTotal.Size = new System.Drawing.Size(43, 22);
            this.txtTotal.TabIndex = 18;
            this.txtTotal.TextChanged += new System.EventHandler(this.txtTotal_TextChanged);
            // 
            // tmrPlan05
            // 
            this.tmrPlan05.Tick += new System.EventHandler(this.tmrPlanXen);
            // 
            // txtSuggestKey
            // 
            this.txtSuggestKey.Location = new System.Drawing.Point(100, 636);
            this.txtSuggestKey.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtSuggestKey.Name = "txtSuggestKey";
            this.txtSuggestKey.Size = new System.Drawing.Size(192, 22);
            this.txtSuggestKey.TabIndex = 16;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(11, 640);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(80, 16);
            this.label9.TabIndex = 17;
            this.label9.Text = "SuggestKey";
            // 
            // càiĐặtToolStripMenuItem
            // 
            this.càiĐặtToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ghiFileCàiĐặtToolStripMenuItem,
            this.đọcFileCàiĐặtToolStripMenuItem,
            this.toolStripMenuItem2,
            this.đọcFileKịchBảnToolStripMenuItem});
            this.càiĐặtToolStripMenuItem.Name = "càiĐặtToolStripMenuItem";
            this.càiĐặtToolStripMenuItem.Size = new System.Drawing.Size(70, 24);
            this.càiĐặtToolStripMenuItem.Text = "Cài đặt";
            // 
            // ghiFileCàiĐặtToolStripMenuItem
            // 
            this.ghiFileCàiĐặtToolStripMenuItem.Name = "ghiFileCàiĐặtToolStripMenuItem";
            this.ghiFileCàiĐặtToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.ghiFileCàiĐặtToolStripMenuItem.Text = "Ghi file cài đặt";
            // 
            // đọcFileCàiĐặtToolStripMenuItem
            // 
            this.đọcFileCàiĐặtToolStripMenuItem.Name = "đọcFileCàiĐặtToolStripMenuItem";
            this.đọcFileCàiĐặtToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.đọcFileCàiĐặtToolStripMenuItem.Text = "Đọc file cài đặt";
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(221, 6);
            // 
            // đọcFileKịchBảnToolStripMenuItem
            // 
            this.đọcFileKịchBảnToolStripMenuItem.Name = "đọcFileKịchBảnToolStripMenuItem";
            this.đọcFileKịchBảnToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.đọcFileKịchBảnToolStripMenuItem.Text = "Đọc file Kịch bản";
            this.đọcFileKịchBảnToolStripMenuItem.Click += new System.EventHandler(this.đọcFileKịchBảnToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(311, 783);
            this.Controls.Add(this.txtTotal);
            this.Controls.Add(this.txtCur);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtSuggestKey);
            this.Controls.Add(this.txtdiachi);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtmatkhau);
            this.Controls.Add(this.txttaikhoan);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.txtKeywords);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.dgrListKeywords);
            this.Controls.Add(this.lstStatus);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Location = new System.Drawing.Point(1050, 300);
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgrListKeywords)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Timer tmrPlan01;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtSpeed;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox lstStatus;
        private System.Windows.Forms.DataGridView dgrListKeywords;
        private System.Windows.Forms.ComboBox cboPlan;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtKeywords;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btnAuto;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem trợGiúpToolStripMenuItem;
        private System.Windows.Forms.Timer tmrPlan02;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.TextBox txttaikhoan;
        private System.Windows.Forms.TextBox txtmatkhau;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Timer tmrPlan03;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnPause;
        private System.Windows.Forms.TextBox txtdiachi;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Timer tmrPlan04;
        private System.Windows.Forms.TextBox txtCur;
        private System.Windows.Forms.TextBox txtTotal;
        private System.Windows.Forms.TextBox txtVolMax;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Timer tmrPlan05;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn clnChecked;
        private System.Windows.Forms.ToolStripMenuItem quickExportExcelToolStripMenuItem;
        private System.Windows.Forms.TextBox txtSuggestKey;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem clearDataGridToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem càiĐặtToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ghiFileCàiĐặtToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem đọcFileCàiĐặtToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem đọcFileKịchBảnToolStripMenuItem;
    }
}

