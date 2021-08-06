namespace DragDrapWatcher_AddIn
{
    partial class frmManager
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
      System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
      this.label1 = new System.Windows.Forms.Label();
      this.textBox1 = new System.Windows.Forms.TextBox();
      this.statusStrip1 = new System.Windows.Forms.StatusStrip();
      this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
      this.lblStatus = new System.Windows.Forms.ToolStripStatusLabel();
      this.btnRefresh = new System.Windows.Forms.Button();
      this.btnDelete = new System.Windows.Forms.Button();
      this.btnSearch = new System.Windows.Forms.Button();
      this.btnEdit = new System.Windows.Forms.Button();
      this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
      this.label2 = new System.Windows.Forms.Label();
      this.linkLabel1 = new System.Windows.Forms.LinkLabel();
      this.panel1 = new System.Windows.Forms.Panel();
      this.numMaxRecipient = new System.Windows.Forms.NumericUpDown();
      this.label8 = new System.Windows.Forms.Label();
      this.txtCategoryPrefix = new System.Windows.Forms.TextBox();
      this.label7 = new System.Windows.Forms.Label();
      this.linkLabel2 = new System.Windows.Forms.LinkLabel();
      this.txtRecipient = new System.Windows.Forms.TextBox();
      this.label6 = new System.Windows.Forms.Label();
      this.button1 = new System.Windows.Forms.Button();
      this.txtFolder = new System.Windows.Forms.TextBox();
      this.txtRuleName = new System.Windows.Forms.TextBox();
      this.label5 = new System.Windows.Forms.Label();
      this.label4 = new System.Windows.Forms.Label();
      this.label3 = new System.Windows.Forms.Label();
      this.dgvList = new System.Windows.Forms.DataGridView();
      this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
      this.statusStrip1.SuspendLayout();
      this.panel1.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.numMaxRecipient)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.dgvList)).BeginInit();
      this.SuspendLayout();
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(9, 8);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(85, 13);
      this.label1.TabIndex = 1;
      this.label1.Text = "Search Keyword";
      // 
      // textBox1
      // 
      this.textBox1.Location = new System.Drawing.Point(12, 26);
      this.textBox1.Name = "textBox1";
      this.textBox1.Size = new System.Drawing.Size(241, 21);
      this.textBox1.TabIndex = 2;
      this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
      // 
      // statusStrip1
      // 
      this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.lblStatus});
      this.statusStrip1.Location = new System.Drawing.Point(0, 335);
      this.statusStrip1.Name = "statusStrip1";
      this.statusStrip1.Size = new System.Drawing.Size(734, 24);
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
      // btnRefresh
      // 
      this.btnRefresh.Location = new System.Drawing.Point(453, 297);
      this.btnRefresh.Name = "btnRefresh";
      this.btnRefresh.Size = new System.Drawing.Size(62, 24);
      this.btnRefresh.TabIndex = 4;
      this.btnRefresh.Text = "Refresh";
      this.btnRefresh.UseVisualStyleBackColor = true;
      this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
      // 
      // btnDelete
      // 
      this.btnDelete.Location = new System.Drawing.Point(453, 87);
      this.btnDelete.Name = "btnDelete";
      this.btnDelete.Size = new System.Drawing.Size(62, 24);
      this.btnDelete.TabIndex = 5;
      this.btnDelete.Text = "Delete";
      this.btnDelete.UseVisualStyleBackColor = true;
      this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
      // 
      // btnSearch
      // 
      this.btnSearch.Location = new System.Drawing.Point(453, 23);
      this.btnSearch.Name = "btnSearch";
      this.btnSearch.Size = new System.Drawing.Size(62, 24);
      this.btnSearch.TabIndex = 7;
      this.btnSearch.Text = "Search";
      this.btnSearch.UseVisualStyleBackColor = true;
      this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
      // 
      // btnEdit
      // 
      this.btnEdit.Location = new System.Drawing.Point(453, 112);
      this.btnEdit.Name = "btnEdit";
      this.btnEdit.Size = new System.Drawing.Size(62, 39);
      this.btnEdit.TabIndex = 8;
      this.btnEdit.Text = "Change Target";
      this.btnEdit.UseVisualStyleBackColor = true;
      this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
      // 
      // checkedListBox1
      // 
      this.checkedListBox1.FormattingEnabled = true;
      this.checkedListBox1.Items.AddRange(new object[] {
            "Name",
            "Email Address",
            "Target Folder"});
      this.checkedListBox1.Location = new System.Drawing.Point(259, 28);
      this.checkedListBox1.Name = "checkedListBox1";
      this.checkedListBox1.Size = new System.Drawing.Size(188, 52);
      this.checkedListBox1.TabIndex = 10;
      this.checkedListBox1.SelectedIndexChanged += new System.EventHandler(this.checkedListBox1_SelectedIndexChanged);
      // 
      // label2
      // 
      this.label2.AutoSize = true;
      this.label2.Location = new System.Drawing.Point(256, 7);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(50, 13);
      this.label2.TabIndex = 11;
      this.label2.Text = "Filter By:";
      // 
      // linkLabel1
      // 
      this.linkLabel1.AutoSize = true;
      this.linkLabel1.Location = new System.Drawing.Point(394, 7);
      this.linkLabel1.Name = "linkLabel1";
      this.linkLabel1.Size = new System.Drawing.Size(50, 13);
      this.linkLabel1.TabIndex = 12;
      this.linkLabel1.TabStop = true;
      this.linkLabel1.Text = "Select All";
      this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
      // 
      // panel1
      // 
      this.panel1.BackColor = System.Drawing.Color.White;
      this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.panel1.Controls.Add(this.numMaxRecipient);
      this.panel1.Controls.Add(this.label8);
      this.panel1.Controls.Add(this.txtCategoryPrefix);
      this.panel1.Controls.Add(this.label7);
      this.panel1.Controls.Add(this.linkLabel2);
      this.panel1.Controls.Add(this.txtRecipient);
      this.panel1.Controls.Add(this.label6);
      this.panel1.Controls.Add(this.button1);
      this.panel1.Controls.Add(this.txtFolder);
      this.panel1.Controls.Add(this.txtRuleName);
      this.panel1.Controls.Add(this.label5);
      this.panel1.Controls.Add(this.label4);
      this.panel1.Controls.Add(this.label3);
      this.panel1.Dock = System.Windows.Forms.DockStyle.Right;
      this.panel1.Location = new System.Drawing.Point(531, 0);
      this.panel1.Name = "panel1";
      this.panel1.Size = new System.Drawing.Size(203, 335);
      this.panel1.TabIndex = 13;
      // 
      // numMaxRecipient
      // 
      this.numMaxRecipient.Location = new System.Drawing.Point(107, 92);
      this.numMaxRecipient.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
      this.numMaxRecipient.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
      this.numMaxRecipient.Name = "numMaxRecipient";
      this.numMaxRecipient.Size = new System.Drawing.Size(83, 21);
      this.numMaxRecipient.TabIndex = 12;
      this.numMaxRecipient.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
      this.numMaxRecipient.ValueChanged += new System.EventHandler(this.numMaxRecipient_ValueChanged);
      // 
      // label8
      // 
      this.label8.Location = new System.Drawing.Point(8, 92);
      this.label8.Name = "label8";
      this.label8.Size = new System.Drawing.Size(93, 36);
      this.label8.TabIndex = 11;
      this.label8.Text = "Max Rule Recipients";
      // 
      // txtCategoryPrefix
      // 
      this.txtCategoryPrefix.Location = new System.Drawing.Point(107, 58);
      this.txtCategoryPrefix.Name = "txtCategoryPrefix";
      this.txtCategoryPrefix.Size = new System.Drawing.Size(83, 21);
      this.txtCategoryPrefix.TabIndex = 10;
      // 
      // label7
      // 
      this.label7.Location = new System.Drawing.Point(8, 55);
      this.label7.Name = "label7";
      this.label7.Size = new System.Drawing.Size(80, 36);
      this.label7.TabIndex = 9;
      this.label7.Text = "Category Rule Prefix";
      // 
      // linkLabel2
      // 
      this.linkLabel2.AutoSize = true;
      this.linkLabel2.Location = new System.Drawing.Point(135, 220);
      this.linkLabel2.Name = "linkLabel2";
      this.linkLabel2.Size = new System.Drawing.Size(55, 13);
      this.linkLabel2.TabIndex = 8;
      this.linkLabel2.TabStop = true;
      this.linkLabel2.Text = "Send Test";
      this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
      // 
      // txtRecipient
      // 
      this.txtRecipient.Location = new System.Drawing.Point(11, 247);
      this.txtRecipient.Multiline = true;
      this.txtRecipient.Name = "txtRecipient";
      this.txtRecipient.Size = new System.Drawing.Size(179, 46);
      this.txtRecipient.TabIndex = 7;
      // 
      // label6
      // 
      this.label6.Location = new System.Drawing.Point(8, 193);
      this.label6.Name = "label6";
      this.label6.Size = new System.Drawing.Size(182, 40);
      this.label6.TabIndex = 6;
      this.label6.Text = "Error Notification Recipients (Use semi-colon \';\' to add multiple address)";
      // 
      // button1
      // 
      this.button1.Location = new System.Drawing.Point(107, 298);
      this.button1.Name = "button1";
      this.button1.Size = new System.Drawing.Size(83, 22);
      this.button1.TabIndex = 5;
      this.button1.Text = "Update";
      this.button1.UseVisualStyleBackColor = true;
      this.button1.Click += new System.EventHandler(this.button1_Click);
      // 
      // txtFolder
      // 
      this.txtFolder.Location = new System.Drawing.Point(9, 168);
      this.txtFolder.Name = "txtFolder";
      this.txtFolder.Size = new System.Drawing.Size(179, 21);
      this.txtFolder.TabIndex = 4;
      // 
      // txtRuleName
      // 
      this.txtRuleName.Location = new System.Drawing.Point(107, 29);
      this.txtRuleName.Name = "txtRuleName";
      this.txtRuleName.Size = new System.Drawing.Size(83, 21);
      this.txtRuleName.TabIndex = 3;
      // 
      // label5
      // 
      this.label5.Location = new System.Drawing.Point(6, 136);
      this.label5.Name = "label5";
      this.label5.Size = new System.Drawing.Size(171, 29);
      this.label5.TabIndex = 2;
      this.label5.Text = "Folder Name Prefix (Watch on Drag and Drop)";
      // 
      // label4
      // 
      this.label4.AutoSize = true;
      this.label4.Location = new System.Drawing.Point(8, 32);
      this.label4.Name = "label4";
      this.label4.Size = new System.Drawing.Size(89, 13);
      this.label4.TabIndex = 1;
      this.label4.Text = "Rule Name Prefix";
      // 
      // label3
      // 
      this.label3.BackColor = System.Drawing.Color.SeaGreen;
      this.label3.Dock = System.Windows.Forms.DockStyle.Top;
      this.label3.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.label3.ForeColor = System.Drawing.Color.White;
      this.label3.Location = new System.Drawing.Point(0, 0);
      this.label3.Name = "label3";
      this.label3.Size = new System.Drawing.Size(201, 20);
      this.label3.TabIndex = 0;
      this.label3.Text = "Configuration";
      this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
      // 
      // dgvList
      // 
      this.dgvList.AllowUserToAddRows = false;
      this.dgvList.AllowUserToDeleteRows = false;
      dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
      this.dgvList.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
      this.dgvList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
      this.dgvList.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleHorizontal;
      this.dgvList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dgvList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4,
            this.Column5});
      this.dgvList.GridColor = System.Drawing.Color.White;
      this.dgvList.Location = new System.Drawing.Point(12, 86);
      this.dgvList.Name = "dgvList";
      this.dgvList.ReadOnly = true;
      this.dgvList.RowHeadersVisible = false;
      this.dgvList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
      this.dgvList.Size = new System.Drawing.Size(435, 235);
      this.dgvList.TabIndex = 14;
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
      this.Column2.Width = 200;
      // 
      // Column3
      // 
      this.Column3.HeaderText = "Target";
      this.Column3.Name = "Column3";
      this.Column3.ReadOnly = true;
      this.Column3.Width = 130;
      // 
      // Column4
      // 
      this.Column4.HeaderText = "FolderPath";
      this.Column4.Name = "Column4";
      this.Column4.ReadOnly = true;
      this.Column4.Visible = false;
      // 
      // Column5
      // 
      this.Column5.HeaderText = "Rule_Name";
      this.Column5.Name = "Column5";
      this.Column5.ReadOnly = true;
      this.Column5.Visible = false;
      // 
      // backgroundWorker1
      // 
      this.backgroundWorker1.WorkerReportsProgress = true;
      this.backgroundWorker1.WorkerSupportsCancellation = true;
      this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
      this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
      // 
      // frmManager
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(734, 359);
      this.Controls.Add(this.dgvList);
      this.Controls.Add(this.panel1);
      this.Controls.Add(this.linkLabel1);
      this.Controls.Add(this.label2);
      this.Controls.Add(this.checkedListBox1);
      this.Controls.Add(this.btnEdit);
      this.Controls.Add(this.btnSearch);
      this.Controls.Add(this.btnDelete);
      this.Controls.Add(this.btnRefresh);
      this.Controls.Add(this.statusStrip1);
      this.Controls.Add(this.textBox1);
      this.Controls.Add(this.label1);
      this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
      this.MaximizeBox = false;
      this.Name = "frmManager";
      this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
      this.Text = "Email on Watch List";
      this.Load += new System.EventHandler(this.frmManager_Load);
      this.statusStrip1.ResumeLayout(false);
      this.statusStrip1.PerformLayout();
      this.panel1.ResumeLayout(false);
      this.panel1.PerformLayout();
      ((System.ComponentModel.ISupportInitialize)(this.numMaxRecipient)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.dgvList)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtFolder;
        private System.Windows.Forms.TextBox txtRuleName;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtRecipient;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DataGridView dgvList;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.TextBox txtCategoryPrefix;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.NumericUpDown numMaxRecipient;
    }
}