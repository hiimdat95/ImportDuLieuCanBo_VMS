namespace ImportDuLieuCanBo_VMS
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnBrowseFolder = new System.Windows.Forms.Button();
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.listViewFolder = new System.Windows.Forms.ListView();
            this.colName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colStatus = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.listIcon = new System.Windows.Forms.ImageList(this.components);
            this.panel3 = new System.Windows.Forms.Panel();
            this.lnkSelectNone = new System.Windows.Forms.LinkLabel();
            this.btnLoadList = new System.Windows.Forms.Button();
            this.lnkSelectAll = new System.Windows.Forms.LinkLabel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.btnStop = new System.Windows.Forms.Button();
            this.btnImportAll = new System.Windows.Forms.Button();
            this.btnImportSelected = new System.Windows.Forms.Button();
            this.panel5 = new System.Windows.Forms.Panel();
            this.logWindow = new System.Windows.Forms.RichTextBox();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.importWorker = new System.ComponentModel.BackgroundWorker();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 250F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.panel2, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.panel3, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel4, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel5, 1, 2);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(958, 708);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // panel1
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.panel1, 2);
            this.panel1.Controls.Add(this.btnBrowseFolder);
            this.panel1.Controls.Add(this.txtFolder);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(952, 44);
            this.panel1.TabIndex = 0;
            // 
            // btnBrowseFolder
            // 
            this.btnBrowseFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowseFolder.Location = new System.Drawing.Point(915, 20);
            this.btnBrowseFolder.Name = "btnBrowseFolder";
            this.btnBrowseFolder.Size = new System.Drawing.Size(34, 24);
            this.btnBrowseFolder.TabIndex = 2;
            this.btnBrowseFolder.Text = "...";
            this.btnBrowseFolder.UseVisualStyleBackColor = true;
            this.btnBrowseFolder.Click += new System.EventHandler(this.btnBrowseFolder_Click);
            // 
            // txtFolder
            // 
            this.txtFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtFolder.Location = new System.Drawing.Point(3, 22);
            this.txtFolder.Name = "txtFolder";
            this.txtFolder.Size = new System.Drawing.Size(906, 20);
            this.txtFolder.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Thư mục chứa dữ liệu:";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.listViewFolder);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(3, 88);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(244, 617);
            this.panel2.TabIndex = 4;
            // 
            // listViewFolder
            // 
            this.listViewFolder.CheckBoxes = true;
            this.listViewFolder.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colName,
            this.colStatus});
            this.listViewFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewFolder.FullRowSelect = true;
            this.listViewFolder.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listViewFolder.HideSelection = false;
            this.listViewFolder.Location = new System.Drawing.Point(0, 0);
            this.listViewFolder.Name = "listViewFolder";
            this.listViewFolder.ShowGroups = false;
            this.listViewFolder.Size = new System.Drawing.Size(244, 617);
            this.listViewFolder.SmallImageList = this.listIcon;
            this.listViewFolder.TabIndex = 1;
            this.listViewFolder.UseCompatibleStateImageBehavior = false;
            this.listViewFolder.View = System.Windows.Forms.View.Details;
            this.listViewFolder.SelectedIndexChanged += new System.EventHandler(this.listViewFolder_SelectedIndexChanged);
            // 
            // colName
            // 
            this.colName.Text = "Tên thư mục";
            this.colName.Width = 180;
            // 
            // colStatus
            // 
            this.colStatus.Text = "Trạng thái";
            // 
            // listIcon
            // 
            this.listIcon.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("listIcon.ImageStream")));
            this.listIcon.TransparentColor = System.Drawing.Color.Transparent;
            this.listIcon.Images.SetKeyName(0, "arrow-circle.png");
            this.listIcon.Images.SetKeyName(1, "arrow-circle-045-left.png");
            this.listIcon.Images.SetKeyName(2, "arrow-circle-135.png");
            this.listIcon.Images.SetKeyName(3, "arrow-circle-135-left.png");
            this.listIcon.Images.SetKeyName(4, "arrow-circle-225.png");
            this.listIcon.Images.SetKeyName(5, "arrow-circle-225-left.png");
            this.listIcon.Images.SetKeyName(6, "arrow-circle-315.png");
            this.listIcon.Images.SetKeyName(7, "arrow-circle-315-left.png");
            this.listIcon.Images.SetKeyName(8, "arrow-circle-double.png");
            this.listIcon.Images.SetKeyName(9, "arrow-circle-double-135.png");
            this.listIcon.Images.SetKeyName(10, "navigation.png");
            this.listIcon.Images.SetKeyName(11, "navigation-000-button.png");
            this.listIcon.Images.SetKeyName(12, "navigation-000-button-white.png");
            this.listIcon.Images.SetKeyName(13, "navigation-000-frame.png");
            this.listIcon.Images.SetKeyName(14, "navigation-000-white.png");
            this.listIcon.Images.SetKeyName(15, "navigation-090.png");
            this.listIcon.Images.SetKeyName(16, "navigation-090-button.png");
            this.listIcon.Images.SetKeyName(17, "navigation-090-button-white.png");
            this.listIcon.Images.SetKeyName(18, "navigation-090-frame.png");
            this.listIcon.Images.SetKeyName(19, "navigation-090-white.png");
            this.listIcon.Images.SetKeyName(20, "navigation-180.png");
            this.listIcon.Images.SetKeyName(21, "navigation-180-button.png");
            this.listIcon.Images.SetKeyName(22, "navigation-180-button-white.png");
            this.listIcon.Images.SetKeyName(23, "navigation-180-frame.png");
            this.listIcon.Images.SetKeyName(24, "navigation-180-white.png");
            this.listIcon.Images.SetKeyName(25, "navigation-270.png");
            this.listIcon.Images.SetKeyName(26, "navigation-270-button.png");
            this.listIcon.Images.SetKeyName(27, "navigation-270-button-white.png");
            this.listIcon.Images.SetKeyName(28, "navigation-270-frame.png");
            this.listIcon.Images.SetKeyName(29, "navigation-270-white.png");
            this.listIcon.Images.SetKeyName(30, "status.png");
            this.listIcon.Images.SetKeyName(31, "status-away.png");
            this.listIcon.Images.SetKeyName(32, "status-busy.png");
            this.listIcon.Images.SetKeyName(33, "status-offline.png");
            this.listIcon.Images.SetKeyName(34, "cross.png");
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.lnkSelectNone);
            this.panel3.Controls.Add(this.btnLoadList);
            this.panel3.Controls.Add(this.lnkSelectAll);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(3, 53);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(244, 29);
            this.panel3.TabIndex = 5;
            // 
            // lnkSelectNone
            // 
            this.lnkSelectNone.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lnkSelectNone.AutoSize = true;
            this.lnkSelectNone.Location = new System.Drawing.Point(194, 8);
            this.lnkSelectNone.Name = "lnkSelectNone";
            this.lnkSelectNone.Size = new System.Drawing.Size(47, 13);
            this.lnkSelectNone.TabIndex = 2;
            this.lnkSelectNone.TabStop = true;
            this.lnkSelectNone.Text = "Bỏ chọn";
            this.lnkSelectNone.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkSelectNone_LinkClicked);
            // 
            // btnLoadList
            // 
            this.btnLoadList.Location = new System.Drawing.Point(3, 3);
            this.btnLoadList.Name = "btnLoadList";
            this.btnLoadList.Size = new System.Drawing.Size(110, 26);
            this.btnLoadList.TabIndex = 3;
            this.btnLoadList.Text = "Lấy danh sách";
            this.btnLoadList.UseVisualStyleBackColor = true;
            this.btnLoadList.Click += new System.EventHandler(this.btnLoadList_Click);
            // 
            // lnkSelectAll
            // 
            this.lnkSelectAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lnkSelectAll.AutoSize = true;
            this.lnkSelectAll.Location = new System.Drawing.Point(138, 8);
            this.lnkSelectAll.Name = "lnkSelectAll";
            this.lnkSelectAll.Size = new System.Drawing.Size(50, 13);
            this.lnkSelectAll.TabIndex = 2;
            this.lnkSelectAll.TabStop = true;
            this.lnkSelectAll.Text = "Chọn hết";
            this.lnkSelectAll.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkSelectAll_LinkClicked);
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.btnStop);
            this.panel4.Controls.Add(this.btnImportAll);
            this.panel4.Controls.Add(this.btnImportSelected);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(253, 53);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(702, 29);
            this.panel4.TabIndex = 6;
            // 
            // btnStop
            // 
            this.btnStop.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnStop.Enabled = false;
            this.btnStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnStop.ImageKey = "cross.png";
            this.btnStop.ImageList = this.listIcon;
            this.btnStop.Location = new System.Drawing.Point(617, 3);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(82, 26);
            this.btnStop.TabIndex = 2;
            this.btnStop.Text = "Dừng";
            this.btnStop.UseVisualStyleBackColor = true;
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            // 
            // btnImportAll
            // 
            this.btnImportAll.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnImportAll.ImageKey = "navigation-000-frame.png";
            this.btnImportAll.ImageList = this.listIcon;
            this.btnImportAll.Location = new System.Drawing.Point(181, 3);
            this.btnImportAll.Name = "btnImportAll";
            this.btnImportAll.Size = new System.Drawing.Size(110, 26);
            this.btnImportAll.TabIndex = 1;
            this.btnImportAll.Text = "Import tất cả";
            this.btnImportAll.UseVisualStyleBackColor = true;
            this.btnImportAll.Click += new System.EventHandler(this.btnImportAll_Click);
            // 
            // btnImportSelected
            // 
            this.btnImportSelected.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnImportSelected.ImageKey = "navigation-000-white.png";
            this.btnImportSelected.ImageList = this.listIcon;
            this.btnImportSelected.Location = new System.Drawing.Point(0, 3);
            this.btnImportSelected.Name = "btnImportSelected";
            this.btnImportSelected.Size = new System.Drawing.Size(175, 26);
            this.btnImportSelected.TabIndex = 0;
            this.btnImportSelected.Text = "Import thư mục đang chọn";
            this.btnImportSelected.UseVisualStyleBackColor = true;
            this.btnImportSelected.Click += new System.EventHandler(this.btnImportSelected_Click);
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.logWindow);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel5.Location = new System.Drawing.Point(253, 88);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(702, 617);
            this.panel5.TabIndex = 7;
            // 
            // logWindow
            // 
            this.logWindow.Dock = System.Windows.Forms.DockStyle.Fill;
            this.logWindow.Location = new System.Drawing.Point(0, 0);
            this.logWindow.Name = "logWindow";
            this.logWindow.Size = new System.Drawing.Size(702, 617);
            this.logWindow.TabIndex = 0;
            this.logWindow.Text = "";
            this.logWindow.TextChanged += new System.EventHandler(this.logWindow_TextChanged);
            // 
            // folderBrowserDialog
            // 
            this.folderBrowserDialog.Description = "Chọn thư mục chứa dữ liệu";
            this.folderBrowserDialog.RootFolder = System.Environment.SpecialFolder.MyComputer;
            this.folderBrowserDialog.ShowNewFolderButton = false;
            // 
            // importWorker
            // 
            this.importWorker.WorkerReportsProgress = true;
            this.importWorker.WorkerSupportsCancellation = true;
            this.importWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.importWorker_DoWork);
            this.importWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.importWorker_ProgressChanged);
            this.importWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.importWorker_RunWorkerCompleted);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(958, 708);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import dữ liệu nhân sự";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnBrowseFolder;
        private System.Windows.Forms.TextBox txtFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.ListView listViewFolder;
        private System.Windows.Forms.ColumnHeader colName;
        private System.Windows.Forms.ColumnHeader colStatus;
        private System.Windows.Forms.Button btnLoadList;
        private System.Windows.Forms.ImageList listIcon;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.LinkLabel lnkSelectNone;
        private System.Windows.Forms.LinkLabel lnkSelectAll;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Button btnImportSelected;
        private System.Windows.Forms.Button btnImportAll;
        private System.ComponentModel.BackgroundWorker importWorker;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.RichTextBox logWindow;
        private System.Windows.Forms.Button btnStop;
    }
}

