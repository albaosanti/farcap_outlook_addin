namespace DragDrapWatcher_AddIn
{
    partial class frmMailCounter
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
      System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
      this.statusStrip1 = new System.Windows.Forms.StatusStrip();
      this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
      this.lblStatus = new System.Windows.Forms.ToolStripStatusLabel();
      this.dgvList = new System.Windows.Forms.DataGridView();
      this.label2 = new System.Windows.Forms.Label();
      this.cmbPeriod = new System.Windows.Forms.ComboBox();
      this.numScan = new System.Windows.Forms.NumericUpDown();
      this.label1 = new System.Windows.Forms.Label();
      this.lblFolderName = new System.Windows.Forms.Label();
      this.label4 = new System.Windows.Forms.Label();
      this.chkAll = new System.Windows.Forms.CheckBox();
      this.label3 = new System.Windows.Forms.Label();
      this.btnProcess = new System.Windows.Forms.Button();
      this.lblSender = new System.Windows.Forms.Label();
      this.label6 = new System.Windows.Forms.Label();
      this.progressBar1 = new System.Windows.Forms.ProgressBar();
      this.bgProcess = new System.ComponentModel.BackgroundWorker();
      this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.statusStrip1.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.dgvList)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.numScan)).BeginInit();
      this.SuspendLayout();
      // 
      // statusStrip1
      // 
      this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.lblStatus});
      this.statusStrip1.Location = new System.Drawing.Point(0, 339);
      this.statusStrip1.Name = "statusStrip1";
      this.statusStrip1.Size = new System.Drawing.Size(452, 24);
      this.statusStrip1.TabIndex = 3;
      this.statusStrip1.Text = "statusStrip1";
      // 
      // toolStripStatusLabel1
      // 
      this.toolStripStatusLabel1.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Right;
      this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
      this.toolStripStatusLabel1.Size = new System.Drawing.Size(43, 19);
      this.toolStripStatusLabel1.Text = "Status";
      // 
      // lblStatus
      // 
      this.lblStatus.ForeColor = System.Drawing.Color.Blue;
      this.lblStatus.Name = "lblStatus";
      this.lblStatus.Size = new System.Drawing.Size(55, 19);
      this.lblStatus.Text = "<Ready>";
      // 
      // dgvList
      // 
      this.dgvList.AllowUserToAddRows = false;
      this.dgvList.AllowUserToDeleteRows = false;
      dataGridViewCellStyle6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
      this.dgvList.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle6;
      this.dgvList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
      this.dgvList.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleHorizontal;
      this.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dgvList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3});
      this.dgvList.GridColor = System.Drawing.Color.White;
      this.dgvList.Location = new System.Drawing.Point(12, 93);
      this.dgvList.Name = "dgvList";
      this.dgvList.ReadOnly = true;
      this.dgvList.RowHeadersVisible = false;
      this.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
      this.dgvList.Size = new System.Drawing.Size(429, 187);
      this.dgvList.TabIndex = 14;
      this.dgvList.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvList_CellContentClick);
      // 
      // label2
      // 
      this.label2.AutoSize = true;
      this.label2.Location = new System.Drawing.Point(183, 68);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(86, 13);
      this.label2.TabIndex = 17;
      this.label2.Text = "Scan mail up to :";
      // 
      // cmbPeriod
      // 
      this.cmbPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
      this.cmbPeriod.FormattingEnabled = true;
      this.cmbPeriod.Items.AddRange(new object[] {
            "Months",
            "Weeks",
            "Days"});
      this.cmbPeriod.Location = new System.Drawing.Point(336, 64);
      this.cmbPeriod.Name = "cmbPeriod";
      this.cmbPeriod.Size = new System.Drawing.Size(105, 21);
      this.cmbPeriod.TabIndex = 19;
      // 
      // numScan
      // 
      this.numScan.Location = new System.Drawing.Point(275, 64);
      this.numScan.Maximum = new decimal(new int[] {
            99,
            0,
            0,
            0});
      this.numScan.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
      this.numScan.Name = "numScan";
      this.numScan.Size = new System.Drawing.Size(56, 21);
      this.numScan.TabIndex = 18;
      this.numScan.Value = new decimal(new int[] {
            3,
            0,
            0,
            0});
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(12, 22);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(85, 13);
      this.label1.TabIndex = 20;
      this.label1.Text = "Folder To Scan: ";
      // 
      // lblFolderName
      // 
      this.lblFolderName.AutoSize = true;
      this.lblFolderName.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.lblFolderName.ForeColor = System.Drawing.Color.Blue;
      this.lblFolderName.Location = new System.Drawing.Point(94, 22);
      this.lblFolderName.Name = "lblFolderName";
      this.lblFolderName.Size = new System.Drawing.Size(66, 13);
      this.lblFolderName.TabIndex = 21;
      this.lblFolderName.Text = "<Not Set>";
      // 
      // label4
      // 
      this.label4.BackColor = System.Drawing.Color.Silver;
      this.label4.Location = new System.Drawing.Point(133, 56);
      this.label4.Name = "label4";
      this.label4.Size = new System.Drawing.Size(1, 32);
      this.label4.TabIndex = 28;
      // 
      // chkAll
      // 
      this.chkAll.AutoSize = true;
      this.chkAll.Location = new System.Drawing.Point(16, 64);
      this.chkAll.Name = "chkAll";
      this.chkAll.Size = new System.Drawing.Size(63, 17);
      this.chkAll.TabIndex = 27;
      this.chkAll.Text = "Scan All";
      this.chkAll.UseVisualStyleBackColor = true;
      this.chkAll.CheckedChanged += new System.EventHandler(this.chkAll_CheckedChanged);
      // 
      // label3
      // 
      this.label3.BackColor = System.Drawing.Color.DarkGray;
      this.label3.Location = new System.Drawing.Point(13, 51);
      this.label3.Name = "label3";
      this.label3.Size = new System.Drawing.Size(426, 1);
      this.label3.TabIndex = 29;
      // 
      // btnProcess
      // 
      this.btnProcess.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.btnProcess.Location = new System.Drawing.Point(351, 13);
      this.btnProcess.Name = "btnProcess";
      this.btnProcess.Size = new System.Drawing.Size(90, 30);
      this.btnProcess.TabIndex = 30;
      this.btnProcess.Text = "Start";
      this.btnProcess.UseVisualStyleBackColor = true;
      this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
      // 
      // lblSender
      // 
      this.lblSender.AutoSize = true;
      this.lblSender.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.lblSender.ForeColor = System.Drawing.Color.Blue;
      this.lblSender.Location = new System.Drawing.Point(94, 293);
      this.lblSender.Name = "lblSender";
      this.lblSender.Size = new System.Drawing.Size(32, 13);
      this.lblSender.TabIndex = 32;
      this.lblSender.Text = "<0>";
      // 
      // label6
      // 
      this.label6.AutoSize = true;
      this.label6.Location = new System.Drawing.Point(9, 293);
      this.label6.Name = "label6";
      this.label6.Size = new System.Drawing.Size(79, 13);
      this.label6.TabIndex = 31;
      this.label6.Text = "Sender found :";
      // 
      // progressBar1
      // 
      this.progressBar1.Location = new System.Drawing.Point(11, 314);
      this.progressBar1.Name = "progressBar1";
      this.progressBar1.Size = new System.Drawing.Size(430, 16);
      this.progressBar1.TabIndex = 33;
      // 
      // bgProcess
      // 
      this.bgProcess.WorkerReportsProgress = true;
      this.bgProcess.WorkerSupportsCancellation = true;
      this.bgProcess.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgProcess_DoWork);
      this.bgProcess.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bgProcess_ProgressChanged);
      this.bgProcess.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgProcess_RunWorkerCompleted);
      // 
      // Column1
      // 
      this.Column1.HeaderText = "Name";
      this.Column1.Name = "Column1";
      this.Column1.ReadOnly = true;
      this.Column1.Width = 130;
      // 
      // Column2
      // 
      this.Column2.HeaderText = "Email";
      this.Column2.Name = "Column2";
      this.Column2.ReadOnly = true;
      this.Column2.Width = 220;
      // 
      // Column3
      // 
      this.Column3.HeaderText = "Counts";
      this.Column3.Name = "Column3";
      this.Column3.ReadOnly = true;
      this.Column3.Width = 70;
      // 
      // frmMailCounter
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(452, 363);
      this.Controls.Add(this.progressBar1);
      this.Controls.Add(this.lblSender);
      this.Controls.Add(this.label6);
      this.Controls.Add(this.btnProcess);
      this.Controls.Add(this.label3);
      this.Controls.Add(this.label4);
      this.Controls.Add(this.chkAll);
      this.Controls.Add(this.lblFolderName);
      this.Controls.Add(this.label1);
      this.Controls.Add(this.label2);
      this.Controls.Add(this.cmbPeriod);
      this.Controls.Add(this.numScan);
      this.Controls.Add(this.dgvList);
      this.Controls.Add(this.statusStrip1);
      this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
      this.MaximizeBox = false;
      this.Name = "frmMailCounter";
      this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
      this.Text = "Mail Counter per Sender";
      this.Load += new System.EventHandler(this.frmMailCounter_Load);
      this.statusStrip1.ResumeLayout(false);
      this.statusStrip1.PerformLayout();
      ((System.ComponentModel.ISupportInitialize)(this.dgvList)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.numScan)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;
        private System.Windows.Forms.DataGridView dgvList;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.ComboBox cmbPeriod;
    private System.Windows.Forms.NumericUpDown numScan;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.Label lblFolderName;
    private System.Windows.Forms.Label label4;
    private System.Windows.Forms.CheckBox chkAll;
    private System.Windows.Forms.Label label3;
    private System.Windows.Forms.Button btnProcess;
    private System.Windows.Forms.Label lblSender;
    private System.Windows.Forms.Label label6;
    private System.Windows.Forms.ProgressBar progressBar1;
    private System.ComponentModel.BackgroundWorker bgProcess;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
    }
}