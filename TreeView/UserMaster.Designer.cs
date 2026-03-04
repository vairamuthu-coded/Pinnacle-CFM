namespace Pinnacle.TreeView
{
    partial class UserMaster
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle13 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle16 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle17 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle18 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle15 = new System.Windows.Forms.DataGridViewCellStyle();
            this.contextMenuStrip3 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.USERID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FINYEAR = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.COMPCODE1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.EMPNAME = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DEPARTMENT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.USERNAME1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.GATENAME1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ACTIVE1 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.PASWORD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sessiontime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtuserrightsusername = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.butfooter = new System.Windows.Forms.Panel();
            this.butheader = new System.Windows.Forms.Button();
            this.tabControl2 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.txtsessiontime = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.comboempname = new System.Windows.Forms.ComboBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.employeeNameRefreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.departmentRefreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.txtpassworddecript = new System.Windows.Forms.TextBox();
            this.Activechk = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txt_userid = new System.Windows.Forms.TextBox();
            this.txt_username = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.combo_dept = new System.Windows.Forms.ComboBox();
            this.txt_password = new System.Windows.Forms.TextBox();
            this.txtgatename = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.combo_finyear = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.combo_compcode = new System.Windows.Forms.ComboBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panel3 = new System.Windows.Forms.Panel();
            this.lblcount = new System.Windows.Forms.Label();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.contextMenuStrip4 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.refreshToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tabControl2.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.contextMenuStrip4.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip3
            // 
            this.contextMenuStrip3.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.refreshToolStripMenuItem});
            this.contextMenuStrip3.Name = "contextMenuStrip3";
            this.contextMenuStrip3.Size = new System.Drawing.Size(114, 26);
            // 
            // refreshToolStripMenuItem
            // 
            this.refreshToolStripMenuItem.Name = "refreshToolStripMenuItem";
            this.refreshToolStripMenuItem.Size = new System.Drawing.Size(113, 22);
            this.refreshToolStripMenuItem.Text = "Refresh";
            this.refreshToolStripMenuItem.Click += new System.EventHandler(this.refreshToolStripMenuItem_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToOrderColumns = true;
            dataGridViewCellStyle13.ForeColor = System.Drawing.Color.Transparent;
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle13;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.BackgroundColor = System.Drawing.Color.White;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle14.BackColor = System.Drawing.Color.Teal;
            dataGridViewCellStyle14.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle14.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle14.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle14.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle14;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.USERID,
            this.FINYEAR,
            this.COMPCODE1,
            this.EMPNAME,
            this.DEPARTMENT,
            this.USERNAME1,
            this.GATENAME1,
            this.ACTIVE1,
            this.PASWORD,
            this.sessiontime});
            dataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle16.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle16.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle16.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle16.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle16.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle16;
            this.dataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridView1.EnableHeadersVisualStyles = false;
            this.dataGridView1.GridColor = System.Drawing.Color.LightGray;
            this.dataGridView1.Location = new System.Drawing.Point(6, 7);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle17.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle17.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle17.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle17;
            dataGridViewCellStyle18.ForeColor = System.Drawing.Color.Black;
            this.dataGridView1.RowsDefaultCellStyle = dataGridViewCellStyle18;
            this.dataGridView1.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(1227, 274);
            this.dataGridView1.TabIndex = 57;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridView1_CellClick);
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridView1_CellContentClick);
            this.dataGridView1.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.dataGridView1_DataBindingComplete);
            this.dataGridView1.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);
            // 
            // USERID
            // 
            this.USERID.DataPropertyName = "USERID";
            dataGridViewCellStyle15.ForeColor = System.Drawing.Color.Black;
            this.USERID.DefaultCellStyle = dataGridViewCellStyle15;
            this.USERID.FillWeight = 127.8194F;
            this.USERID.HeaderText = "UserID";
            this.USERID.Name = "USERID";
            this.USERID.ReadOnly = true;
            // 
            // FINYEAR
            // 
            this.FINYEAR.DataPropertyName = "FINYEAR";
            this.FINYEAR.FillWeight = 38.09188F;
            this.FINYEAR.HeaderText = "FinYear";
            this.FINYEAR.Name = "FINYEAR";
            this.FINYEAR.ReadOnly = true;
            // 
            // COMPCODE1
            // 
            this.COMPCODE1.DataPropertyName = "COMPCODE1";
            this.COMPCODE1.FillWeight = 127.8194F;
            this.COMPCODE1.HeaderText = "CompCode";
            this.COMPCODE1.Name = "COMPCODE1";
            this.COMPCODE1.ReadOnly = true;
            // 
            // EMPNAME
            // 
            this.EMPNAME.DataPropertyName = "empname";
            this.EMPNAME.FillWeight = 35.12078F;
            this.EMPNAME.HeaderText = "EmpName";
            this.EMPNAME.Name = "EMPNAME";
            this.EMPNAME.ReadOnly = true;
            // 
            // DEPARTMENT
            // 
            this.DEPARTMENT.DataPropertyName = "department";
            this.DEPARTMENT.FillWeight = 32.05128F;
            this.DEPARTMENT.HeaderText = "Department";
            this.DEPARTMENT.Name = "DEPARTMENT";
            this.DEPARTMENT.ReadOnly = true;
            // 
            // USERNAME1
            // 
            this.USERNAME1.DataPropertyName = "USERNAME";
            this.USERNAME1.FillWeight = 127.8194F;
            this.USERNAME1.HeaderText = "UserName";
            this.USERNAME1.Name = "USERNAME1";
            this.USERNAME1.ReadOnly = true;
            // 
            // GATENAME1
            // 
            this.GATENAME1.DataPropertyName = "GATENAME";
            this.GATENAME1.FillWeight = 127.8194F;
            this.GATENAME1.HeaderText = "GateName";
            this.GATENAME1.Name = "GATENAME1";
            this.GATENAME1.ReadOnly = true;
            // 
            // ACTIVE1
            // 
            this.ACTIVE1.DataPropertyName = "ACTIVE";
            this.ACTIVE1.FillWeight = 127.8194F;
            this.ACTIVE1.HeaderText = "Active";
            this.ACTIVE1.Name = "ACTIVE1";
            this.ACTIVE1.ReadOnly = true;
            // 
            // PASWORD
            // 
            this.PASWORD.DataPropertyName = "PASWORD";
            this.PASWORD.FillWeight = 127.8194F;
            this.PASWORD.HeaderText = "Password";
            this.PASWORD.Name = "PASWORD";
            this.PASWORD.ReadOnly = true;
            // 
            // sessiontime
            // 
            this.sessiontime.DataPropertyName = "sessiontime";
            this.sessiontime.FillWeight = 127.8194F;
            this.sessiontime.HeaderText = "time";
            this.sessiontime.Name = "sessiontime";
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.txtuserrightsusername);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Font = new System.Drawing.Font("Roboto", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panel1.ForeColor = System.Drawing.Color.Black;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Margin = new System.Windows.Forms.Padding(0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1239, 19);
            this.panel1.TabIndex = 62;
            // 
            // txtuserrightsusername
            // 
            this.txtuserrightsusername.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtuserrightsusername.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtuserrightsusername.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtuserrightsusername.Location = new System.Drawing.Point(94, 4);
            this.txtuserrightsusername.Margin = new System.Windows.Forms.Padding(0);
            this.txtuserrightsusername.Name = "txtuserrightsusername";
            this.txtuserrightsusername.Size = new System.Drawing.Size(223, 23);
            this.txtuserrightsusername.TabIndex = 48;
            this.txtuserrightsusername.TextChanged += new System.EventHandler(this.Txtuserrightsusername_TextChanged);
            // 
            // label10
            // 
            this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(38, 7);
            this.label10.Margin = new System.Windows.Forms.Padding(0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(50, 18);
            this.label10.TabIndex = 47;
            this.label10.Text = "Search";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.White;
            this.panel2.Controls.Add(this.butfooter);
            this.panel2.Controls.Add(this.butheader);
            this.panel2.Controls.Add(this.tabControl2);
            this.panel2.Controls.Add(this.tabControl1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(3, 3);
            this.panel2.Margin = new System.Windows.Forms.Padding(0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1266, 534);
            this.panel2.TabIndex = 63;
            this.panel2.Paint += new System.Windows.Forms.PaintEventHandler(this.Panel2_Paint);
            // 
            // butfooter
            // 
            this.butfooter.BackColor = System.Drawing.Color.Teal;
            this.butfooter.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.butfooter.Location = new System.Drawing.Point(0, 524);
            this.butfooter.Name = "butfooter";
            this.butfooter.Size = new System.Drawing.Size(1266, 10);
            this.butfooter.TabIndex = 449;
            // 
            // butheader
            // 
            this.butheader.BackColor = System.Drawing.Color.Teal;
            this.butheader.Dock = System.Windows.Forms.DockStyle.Top;
            this.butheader.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.butheader.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.butheader.Font = new System.Drawing.Font("Roboto Black", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butheader.ForeColor = System.Drawing.Color.White;
            this.butheader.Location = new System.Drawing.Point(0, 0);
            this.butheader.Margin = new System.Windows.Forms.Padding(0);
            this.butheader.Name = "butheader";
            this.butheader.Size = new System.Drawing.Size(1266, 35);
            this.butheader.TabIndex = 448;
            this.butheader.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.butheader.UseVisualStyleBackColor = false;
            // 
            // tabControl2
            // 
            this.tabControl2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl2.Controls.Add(this.tabPage2);
            this.tabControl2.Location = new System.Drawing.Point(8, 40);
            this.tabControl2.Margin = new System.Windows.Forms.Padding(5);
            this.tabControl2.Name = "tabControl2";
            this.tabControl2.SelectedIndex = 0;
            this.tabControl2.Size = new System.Drawing.Size(1245, 131);
            this.tabControl2.TabIndex = 62;
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.White;
            this.tabPage2.Controls.Add(this.txtsessiontime);
            this.tabPage2.Controls.Add(this.pictureBox1);
            this.tabPage2.Controls.Add(this.comboempname);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.txtpassworddecript);
            this.tabPage2.Controls.Add(this.Activechk);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.label9);
            this.tabPage2.Controls.Add(this.label7);
            this.tabPage2.Controls.Add(this.txt_userid);
            this.tabPage2.Controls.Add(this.txt_username);
            this.tabPage2.Controls.Add(this.label8);
            this.tabPage2.Controls.Add(this.combo_dept);
            this.tabPage2.Controls.Add(this.txt_password);
            this.tabPage2.Controls.Add(this.txtgatename);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.label6);
            this.tabPage2.Controls.Add(this.combo_finyear);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Controls.Add(this.combo_compcode);
            this.tabPage2.ForeColor = System.Drawing.Color.Black;
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1237, 105);
            this.tabPage2.TabIndex = 0;
            this.tabPage2.Text = "User Master";
            // 
            // txtsessiontime
            // 
            this.txtsessiontime.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txtsessiontime.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtsessiontime.ForeColor = System.Drawing.Color.Blue;
            this.txtsessiontime.Location = new System.Drawing.Point(564, 34);
            this.txtsessiontime.Margin = new System.Windows.Forms.Padding(0);
            this.txtsessiontime.Name = "txtsessiontime";
            this.txtsessiontime.Size = new System.Drawing.Size(105, 23);
            this.txtsessiontime.TabIndex = 80;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBox1.Location = new System.Drawing.Point(1094, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(135, 87);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 79;
            this.pictureBox1.TabStop = false;
            // 
            // comboempname
            // 
            this.comboempname.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.comboempname.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboempname.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboempname.ContextMenuStrip = this.contextMenuStrip1;
            this.comboempname.Enabled = false;
            this.comboempname.Font = new System.Drawing.Font("Trebuchet MS", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboempname.ForeColor = System.Drawing.Color.Blue;
            this.comboempname.FormattingEnabled = true;
            this.comboempname.Items.AddRange(new object[] {
            "0"});
            this.comboempname.Location = new System.Drawing.Point(98, 35);
            this.comboempname.Margin = new System.Windows.Forms.Padding(0);
            this.comboempname.Name = "comboempname";
            this.comboempname.Size = new System.Drawing.Size(226, 24);
            this.comboempname.TabIndex = 78;
            this.comboempname.SelectedIndexChanged += new System.EventHandler(this.Comboempname_SelectedIndexChanged);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.employeeNameRefreshToolStripMenuItem,
            this.departmentRefreshToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(201, 48);
            this.contextMenuStrip1.Text = "EMPNAME";
            this.contextMenuStrip1.Opening += new System.ComponentModel.CancelEventHandler(this.ContextMenuStrip1_Opening);
            // 
            // employeeNameRefreshToolStripMenuItem
            // 
            this.employeeNameRefreshToolStripMenuItem.Name = "employeeNameRefreshToolStripMenuItem";
            this.employeeNameRefreshToolStripMenuItem.Size = new System.Drawing.Size(200, 22);
            this.employeeNameRefreshToolStripMenuItem.Text = "EmployeeName Refresh";
            this.employeeNameRefreshToolStripMenuItem.Click += new System.EventHandler(this.EmployeeNameRefreshToolStripMenuItem_Click);
            // 
            // departmentRefreshToolStripMenuItem
            // 
            this.departmentRefreshToolStripMenuItem.Name = "departmentRefreshToolStripMenuItem";
            this.departmentRefreshToolStripMenuItem.Size = new System.Drawing.Size(200, 22);
            this.departmentRefreshToolStripMenuItem.Text = "Department Refresh";
            this.departmentRefreshToolStripMenuItem.Click += new System.EventHandler(this.DepartmentRefreshToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label1.Location = new System.Drawing.Point(21, 12);
            this.label1.Margin = new System.Windows.Forms.Padding(0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 13);
            this.label1.TabIndex = 68;
            this.label1.Text = "UserID";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtpassworddecript
            // 
            this.txtpassworddecript.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txtpassworddecript.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtpassworddecript.Enabled = false;
            this.txtpassworddecript.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtpassworddecript.ForeColor = System.Drawing.Color.Blue;
            this.txtpassworddecript.Location = new System.Drawing.Point(787, 57);
            this.txtpassworddecript.Margin = new System.Windows.Forms.Padding(0);
            this.txtpassworddecript.Name = "txtpassworddecript";
            this.txtpassworddecript.Size = new System.Drawing.Size(223, 23);
            this.txtpassworddecript.TabIndex = 67;
            // 
            // Activechk
            // 
            this.Activechk.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.Activechk.AutoSize = true;
            this.Activechk.Checked = true;
            this.Activechk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.Activechk.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.Activechk.Location = new System.Drawing.Point(1023, 62);
            this.Activechk.Name = "Activechk";
            this.Activechk.Size = new System.Drawing.Size(56, 17);
            this.Activechk.TabIndex = 66;
            this.Activechk.Text = "Active";
            this.Activechk.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label4.Location = new System.Drawing.Point(21, 39);
            this.label4.Margin = new System.Windows.Forms.Padding(0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 71;
            this.label4.Text = "EmpName";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label9
            // 
            this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label9.AutoSize = true;
            this.label9.ContextMenuStrip = this.contextMenuStrip4;
            this.label9.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label9.Location = new System.Drawing.Point(679, 67);
            this.label9.Margin = new System.Windows.Forms.Padding(0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(53, 13);
            this.label9.TabIndex = 77;
            this.label9.Text = "Password";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label7.AutoSize = true;
            this.label7.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label7.Location = new System.Drawing.Point(21, 62);
            this.label7.Margin = new System.Windows.Forms.Padding(0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(58, 13);
            this.label7.TabIndex = 74;
            this.label7.Text = "GateName";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txt_userid
            // 
            this.txt_userid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txt_userid.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txt_userid.Enabled = false;
            this.txt_userid.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_userid.ForeColor = System.Drawing.Color.Blue;
            this.txt_userid.Location = new System.Drawing.Point(98, 7);
            this.txt_userid.Margin = new System.Windows.Forms.Padding(0);
            this.txt_userid.Name = "txt_userid";
            this.txt_userid.Size = new System.Drawing.Size(226, 23);
            this.txt_userid.TabIndex = 76;
            // 
            // txt_username
            // 
            this.txt_username.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txt_username.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_username.ForeColor = System.Drawing.Color.Blue;
            this.txt_username.Location = new System.Drawing.Point(787, 31);
            this.txt_username.Margin = new System.Windows.Forms.Padding(0);
            this.txt_username.Name = "txt_username";
            this.txt_username.Size = new System.Drawing.Size(223, 23);
            this.txt_username.TabIndex = 63;
            // 
            // label8
            // 
            this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label8.AutoSize = true;
            this.label8.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label8.Location = new System.Drawing.Point(679, 39);
            this.label8.Margin = new System.Windows.Forms.Padding(0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(57, 13);
            this.label8.TabIndex = 73;
            this.label8.Text = "UserName";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // combo_dept
            // 
            this.combo_dept.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.combo_dept.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combo_dept.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combo_dept.ContextMenuStrip = this.contextMenuStrip1;
            this.combo_dept.Enabled = false;
            this.combo_dept.Font = new System.Drawing.Font("Trebuchet MS", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.combo_dept.ForeColor = System.Drawing.Color.Blue;
            this.combo_dept.FormattingEnabled = true;
            this.combo_dept.Location = new System.Drawing.Point(446, 33);
            this.combo_dept.Margin = new System.Windows.Forms.Padding(0);
            this.combo_dept.Name = "combo_dept";
            this.combo_dept.Size = new System.Drawing.Size(103, 24);
            this.combo_dept.TabIndex = 62;
            this.combo_dept.SelectedIndexChanged += new System.EventHandler(this.Combo_dept_SelectedIndexChanged);
            // 
            // txt_password
            // 
            this.txt_password.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txt_password.Enabled = false;
            this.txt_password.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_password.ForeColor = System.Drawing.Color.Blue;
            this.txt_password.Location = new System.Drawing.Point(446, 60);
            this.txt_password.Margin = new System.Windows.Forms.Padding(0);
            this.txt_password.Name = "txt_password";
            this.txt_password.Size = new System.Drawing.Size(223, 23);
            this.txt_password.TabIndex = 65;
            this.txt_password.TextChanged += new System.EventHandler(this.txt_password_TextChanged);
            // 
            // txtgatename
            // 
            this.txtgatename.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txtgatename.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtgatename.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtgatename.ForeColor = System.Drawing.Color.Blue;
            this.txtgatename.Location = new System.Drawing.Point(98, 62);
            this.txtgatename.Margin = new System.Windows.Forms.Padding(0);
            this.txtgatename.Name = "txtgatename";
            this.txtgatename.Size = new System.Drawing.Size(226, 23);
            this.txtgatename.TabIndex = 64;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label3.Location = new System.Drawing.Point(676, 12);
            this.label3.Margin = new System.Windows.Forms.Padding(0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 13);
            this.label3.TabIndex = 70;
            this.label3.Text = "CompCode";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label2.Location = new System.Drawing.Point(353, 10);
            this.label2.Margin = new System.Windows.Forms.Padding(0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(37, 13);
            this.label2.TabIndex = 69;
            this.label2.Text = "FinYer";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label6.AutoSize = true;
            this.label6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label6.Location = new System.Drawing.Point(353, 62);
            this.label6.Margin = new System.Windows.Forms.Padding(0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 13);
            this.label6.TabIndex = 75;
            this.label6.Text = "Password";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // combo_finyear
            // 
            this.combo_finyear.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.combo_finyear.Enabled = false;
            this.combo_finyear.Font = new System.Drawing.Font("Trebuchet MS", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.combo_finyear.ForeColor = System.Drawing.Color.Blue;
            this.combo_finyear.FormattingEnabled = true;
            this.combo_finyear.Location = new System.Drawing.Point(446, 6);
            this.combo_finyear.Margin = new System.Windows.Forms.Padding(0);
            this.combo_finyear.Name = "combo_finyear";
            this.combo_finyear.Size = new System.Drawing.Size(223, 24);
            this.combo_finyear.TabIndex = 59;
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.label5.Location = new System.Drawing.Point(353, 36);
            this.label5.Margin = new System.Windows.Forms.Padding(0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(62, 13);
            this.label5.TabIndex = 72;
            this.label5.Text = "Department";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // combo_compcode
            // 
            this.combo_compcode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.combo_compcode.DropDownHeight = 75;
            this.combo_compcode.Font = new System.Drawing.Font("Trebuchet MS", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.combo_compcode.ForeColor = System.Drawing.Color.Blue;
            this.combo_compcode.FormattingEnabled = true;
            this.combo_compcode.IntegralHeight = false;
            this.combo_compcode.Location = new System.Drawing.Point(787, 5);
            this.combo_compcode.Margin = new System.Windows.Forms.Padding(0);
            this.combo_compcode.Name = "combo_compcode";
            this.combo_compcode.Size = new System.Drawing.Size(223, 24);
            this.combo_compcode.TabIndex = 60;
            this.combo_compcode.SelectedIndexChanged += new System.EventHandler(this.Combo_compcode_SelectedIndexChanged);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(8, 181);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1253, 335);
            this.tabControl1.TabIndex = 61;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.White;
            this.tabPage1.Controls.Add(this.panel3);
            this.tabPage1.Controls.Add(this.dataGridView1);
            this.tabPage1.Controls.Add(this.panel1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1245, 309);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "User Master Details";
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.Teal;
            this.panel3.Controls.Add(this.lblcount);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.ForeColor = System.Drawing.Color.White;
            this.panel3.Location = new System.Drawing.Point(3, 284);
            this.panel3.Margin = new System.Windows.Forms.Padding(0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1239, 22);
            this.panel3.TabIndex = 64;
            // 
            // lblcount
            // 
            this.lblcount.AutoSize = true;
            this.lblcount.Font = new System.Drawing.Font("Roboto Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblcount.ForeColor = System.Drawing.Color.White;
            this.lblcount.Location = new System.Drawing.Point(0, 0);
            this.lblcount.Margin = new System.Windows.Forms.Padding(0);
            this.lblcount.Name = "lblcount";
            this.lblcount.Padding = new System.Windows.Forms.Padding(3);
            this.lblcount.Size = new System.Drawing.Size(80, 24);
            this.lblcount.TabIndex = 63;
            this.lblcount.Text = "Total Count";
            this.lblcount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(61, 4);
            // 
            // contextMenuStrip4
            // 
            this.contextMenuStrip4.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.refreshToolStripMenuItem1});
            this.contextMenuStrip4.Name = "contextMenuStrip4";
            this.contextMenuStrip4.Size = new System.Drawing.Size(114, 26);
            // 
            // refreshToolStripMenuItem1
            // 
            this.refreshToolStripMenuItem1.Name = "refreshToolStripMenuItem1";
            this.refreshToolStripMenuItem1.Size = new System.Drawing.Size(180, 22);
            this.refreshToolStripMenuItem1.Text = "Refresh";
            this.refreshToolStripMenuItem1.Click += new System.EventHandler(this.refreshToolStripMenuItem1_Click);
            // 
            // UserMaster
            // 
            this.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1272, 540);
            this.ControlBox = false;
            this.Controls.Add(this.panel2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.Name = "UserMaster";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.Text = "Create User";
            this.TransparencyKey = System.Drawing.Color.Olive;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.UserMaster_FormClosed);
            this.Load += new System.EventHandler(this.UserMaster_Load);
            this.contextMenuStrip3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.tabControl2.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.contextMenuStrip4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txtuserrightsusername;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label lblcount;
        private System.Windows.Forms.TabControl tabControl2;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtpassworddecript;
        private System.Windows.Forms.CheckBox Activechk;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txt_userid;
        private System.Windows.Forms.TextBox txt_username;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox combo_dept;
        private System.Windows.Forms.TextBox txt_password;
        private System.Windows.Forms.TextBox txtgatename;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox combo_finyear;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox combo_compcode;
        private System.Windows.Forms.ComboBox comboempname;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem employeeNameRefreshToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem departmentRefreshToolStripMenuItem;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip3;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
        private System.Windows.Forms.TextBox txtsessiontime;
        private System.Windows.Forms.Button butheader;
        private System.Windows.Forms.Panel butfooter;
        private System.Windows.Forms.DataGridViewTextBoxColumn USERID;
        private System.Windows.Forms.DataGridViewTextBoxColumn FINYEAR;
        private System.Windows.Forms.DataGridViewTextBoxColumn COMPCODE1;
        private System.Windows.Forms.DataGridViewTextBoxColumn EMPNAME;
        private System.Windows.Forms.DataGridViewTextBoxColumn DEPARTMENT;
        private System.Windows.Forms.DataGridViewTextBoxColumn USERNAME1;
        private System.Windows.Forms.DataGridViewTextBoxColumn GATENAME1;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ACTIVE1;
        private System.Windows.Forms.DataGridViewTextBoxColumn PASWORD;
        private System.Windows.Forms.DataGridViewTextBoxColumn sessiontime;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip4;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem1;
    }
}