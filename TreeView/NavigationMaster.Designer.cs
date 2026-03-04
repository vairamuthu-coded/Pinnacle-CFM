namespace Pinnacle.TreeView
{
    partial class NavigationMaster
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.butfooter = new System.Windows.Forms.Button();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menuMasterRefreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.butheader = new System.Windows.Forms.Button();
            this.tabControl2 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.txtmenuid = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panel3 = new System.Windows.Forms.Panel();
            this.lbltotal = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.txtsearch = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.combousername1 = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.combocompcode1 = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.txtnavurl = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.chk = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.combousername = new System.Windows.Forms.ComboBox();
            this.txtparentmenuid = new System.Windows.Forms.TextBox();
            this.combomenunameid = new System.Windows.Forms.ComboBox();
            this.combocompcode = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtmenuname = new System.Windows.Forms.TextBox();
            this.MENUID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MENUNAMEID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.tabControl2.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.butfooter);
            this.panel1.Controls.Add(this.butheader);
            this.panel1.Controls.Add(this.tabControl2);
            this.panel1.Location = new System.Drawing.Point(0, -3);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1271, 505);
            this.panel1.TabIndex = 31;
            // 
            // butfooter
            // 
            this.butfooter.BackColor = System.Drawing.Color.Teal;
            this.butfooter.ContextMenuStrip = this.contextMenuStrip1;
            this.butfooter.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.butfooter.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.butfooter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.butfooter.Font = new System.Drawing.Font("Roboto Black", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butfooter.ForeColor = System.Drawing.Color.White;
            this.butfooter.Location = new System.Drawing.Point(0, 495);
            this.butfooter.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.butfooter.Name = "butfooter";
            this.butfooter.Padding = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.butfooter.Size = new System.Drawing.Size(1271, 10);
            this.butfooter.TabIndex = 451;
            this.butfooter.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.butfooter.UseVisualStyleBackColor = false;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuMasterRefreshToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(184, 26);
            // 
            // menuMasterRefreshToolStripMenuItem
            // 
            this.menuMasterRefreshToolStripMenuItem.Name = "menuMasterRefreshToolStripMenuItem";
            this.menuMasterRefreshToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.menuMasterRefreshToolStripMenuItem.Text = "MenuMaster Refresh";
            this.menuMasterRefreshToolStripMenuItem.Click += new System.EventHandler(this.MenuMasterRefreshToolStripMenuItem_Click);
            // 
            // butheader
            // 
            this.butheader.BackColor = System.Drawing.Color.Teal;
            this.butheader.Dock = System.Windows.Forms.DockStyle.Top;
            this.butheader.FlatAppearance.BorderSize = 0;
            this.butheader.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.butheader.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.butheader.Font = new System.Drawing.Font("Roboto Black", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butheader.ForeColor = System.Drawing.Color.White;
            this.butheader.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.butheader.Location = new System.Drawing.Point(0, 0);
            this.butheader.Margin = new System.Windows.Forms.Padding(0);
            this.butheader.Name = "butheader";
            this.butheader.Size = new System.Drawing.Size(1271, 30);
            this.butheader.TabIndex = 450;
            this.butheader.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.butheader.UseVisualStyleBackColor = false;
            // 
            // tabControl2
            // 
            this.tabControl2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl2.Controls.Add(this.tabPage2);
            this.tabControl2.Location = new System.Drawing.Point(2, 33);
            this.tabControl2.Name = "tabControl2";
            this.tabControl2.SelectedIndex = 0;
            this.tabControl2.Size = new System.Drawing.Size(1264, 459);
            this.tabControl2.TabIndex = 12;
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.White;
            this.tabPage2.Controls.Add(this.txtmenuid);
            this.tabPage2.Controls.Add(this.tabControl1);
            this.tabPage2.Controls.Add(this.txtnavurl);
            this.tabPage2.Controls.Add(this.label7);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.label6);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.chk);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.combousername);
            this.tabPage2.Controls.Add(this.txtparentmenuid);
            this.tabPage2.Controls.Add(this.combomenunameid);
            this.tabPage2.Controls.Add(this.combocompcode);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.txtmenuname);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1256, 433);
            this.tabPage2.TabIndex = 0;
            this.tabPage2.Text = "Create Navigation";
            this.tabPage2.Click += new System.EventHandler(this.TabPage2_Click);
            // 
            // txtmenuid
            // 
            this.txtmenuid.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txtmenuid.Location = new System.Drawing.Point(115, 31);
            this.txtmenuid.Margin = new System.Windows.Forms.Padding(0);
            this.txtmenuid.Name = "txtmenuid";
            this.txtmenuid.ReadOnly = true;
            this.txtmenuid.Size = new System.Drawing.Size(214, 20);
            this.txtmenuid.TabIndex = 1;
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(527, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(726, 427);
            this.tabControl1.TabIndex = 11;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.White;
            this.tabPage1.Controls.Add(this.panel3);
            this.tabPage1.Controls.Add(this.panel2);
            this.tabPage1.Controls.Add(this.dataGridView1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(0);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(718, 401);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Navigation Details";
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.Teal;
            this.panel3.Controls.Add(this.lbltotal);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Font = new System.Drawing.Font("Trebuchet MS", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panel3.ForeColor = System.Drawing.Color.White;
            this.panel3.Location = new System.Drawing.Point(3, 369);
            this.panel3.Margin = new System.Windows.Forms.Padding(0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(712, 29);
            this.panel3.TabIndex = 19;
            // 
            // lbltotal
            // 
            this.lbltotal.BackColor = System.Drawing.Color.Transparent;
            this.lbltotal.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lbltotal.ForeColor = System.Drawing.Color.White;
            this.lbltotal.Location = new System.Drawing.Point(0, 7);
            this.lbltotal.Margin = new System.Windows.Forms.Padding(2, 0, 0, 0);
            this.lbltotal.Name = "lbltotal";
            this.lbltotal.Size = new System.Drawing.Size(712, 22);
            this.lbltotal.TabIndex = 14;
            this.lbltotal.Text = "Total";
            this.lbltotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.Teal;
            this.panel2.Controls.Add(this.txtsearch);
            this.panel2.Controls.Add(this.label10);
            this.panel2.Controls.Add(this.combousername1);
            this.panel2.Controls.Add(this.label9);
            this.panel2.Controls.Add(this.combocompcode1);
            this.panel2.Controls.Add(this.label8);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.ForeColor = System.Drawing.Color.White;
            this.panel2.Location = new System.Drawing.Point(3, 3);
            this.panel2.Margin = new System.Windows.Forms.Padding(0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(712, 32);
            this.panel2.TabIndex = 18;
            // 
            // txtsearch
            // 
            this.txtsearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtsearch.Location = new System.Drawing.Point(504, 5);
            this.txtsearch.Margin = new System.Windows.Forms.Padding(0);
            this.txtsearch.Name = "txtsearch";
            this.txtsearch.Size = new System.Drawing.Size(132, 20);
            this.txtsearch.TabIndex = 19;
            this.txtsearch.TextChanged += new System.EventHandler(this.txtsearch_TextChanged);
            // 
            // label10
            // 
            this.label10.Location = new System.Drawing.Point(423, 5);
            this.label10.Margin = new System.Windows.Forms.Padding(0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(81, 18);
            this.label10.TabIndex = 18;
            this.label10.Text = "Search";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // combousername1
            // 
            this.combousername1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combousername1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combousername1.FormattingEnabled = true;
            this.combousername1.Items.AddRange(new object[] {
            "----"});
            this.combousername1.Location = new System.Drawing.Point(307, 6);
            this.combousername1.Margin = new System.Windows.Forms.Padding(0);
            this.combousername1.Name = "combousername1";
            this.combousername1.Size = new System.Drawing.Size(103, 21);
            this.combousername1.TabIndex = 16;
            this.combousername1.SelectedIndexChanged += new System.EventHandler(this.Combousername1_SelectedIndexChanged);
            // 
            // label9
            // 
            this.label9.ForeColor = System.Drawing.Color.White;
            this.label9.Location = new System.Drawing.Point(229, 7);
            this.label9.Margin = new System.Windows.Forms.Padding(2, 0, 0, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(78, 20);
            this.label9.TabIndex = 15;
            this.label9.Text = "UserName   *";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // combocompcode1
            // 
            this.combocompcode1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combocompcode1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combocompcode1.ContextMenuStrip = this.contextMenuStrip1;
            this.combocompcode1.FormattingEnabled = true;
            this.combocompcode1.Location = new System.Drawing.Point(99, 6);
            this.combocompcode1.Margin = new System.Windows.Forms.Padding(0);
            this.combocompcode1.Name = "combocompcode1";
            this.combocompcode1.Size = new System.Drawing.Size(113, 21);
            this.combocompcode1.TabIndex = 17;
            this.combocompcode1.SelectedIndexChanged += new System.EventHandler(this.Combocompcode1_SelectedIndexChanged);
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.label8.ForeColor = System.Drawing.Color.White;
            this.label8.Location = new System.Drawing.Point(4, 3);
            this.label8.Margin = new System.Windows.Forms.Padding(2, 0, 0, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(95, 22);
            this.label8.TabIndex = 14;
            this.label8.Text = "CompCode   *";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.BackgroundColor = System.Drawing.Color.White;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Teal;
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ContextMenuStrip = this.contextMenuStrip1;
            this.dataGridView1.EnableHeadersVisualStyles = false;
            this.dataGridView1.GridColor = System.Drawing.Color.White;
            this.dataGridView1.Location = new System.Drawing.Point(10, 42);
            this.dataGridView1.Name = "dataGridView1";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.Teal;
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(702, 324);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridView1_CellClick);
            // 
            // txtnavurl
            // 
            this.txtnavurl.Location = new System.Drawing.Point(113, 111);
            this.txtnavurl.Margin = new System.Windows.Forms.Padding(0);
            this.txtnavurl.Name = "txtnavurl";
            this.txtnavurl.Size = new System.Drawing.Size(214, 20);
            this.txtnavurl.TabIndex = 15;
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(15, 194);
            this.label7.Margin = new System.Windows.Forms.Padding(0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(81, 16);
            this.label7.TabIndex = 11;
            this.label7.Text = "UserName   *";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(12, 111);
            this.label4.Margin = new System.Windows.Forms.Padding(0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(98, 20);
            this.label4.TabIndex = 5;
            this.label4.Text = "Navgate Url *";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(15, 163);
            this.label6.Margin = new System.Windows.Forms.Padding(0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(98, 20);
            this.label6.TabIndex = 10;
            this.label6.Text = "CompCode   *";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(14, 140);
            this.label5.Margin = new System.Windows.Forms.Padding(0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(98, 20);
            this.label5.TabIndex = 7;
            this.label5.Text = "parentmenuid  *";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(12, 56);
            this.label2.Margin = new System.Windows.Forms.Padding(0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 20);
            this.label2.TabIndex = 2;
            this.label2.Text = "MenuNameID *";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // chk
            // 
            this.chk.AutoSize = true;
            this.chk.Checked = true;
            this.chk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk.Location = new System.Drawing.Point(115, 233);
            this.chk.Name = "chk";
            this.chk.Size = new System.Drawing.Size(56, 17);
            this.chk.TabIndex = 16;
            this.chk.Text = "Active";
            this.chk.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.chk.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(12, 85);
            this.label3.Margin = new System.Windows.Forms.Padding(0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(81, 18);
            this.label3.TabIndex = 3;
            this.label3.Text = "MenuName   *";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // combousername
            // 
            this.combousername.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combousername.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combousername.FormattingEnabled = true;
            this.combousername.Location = new System.Drawing.Point(113, 191);
            this.combousername.Margin = new System.Windows.Forms.Padding(0);
            this.combousername.Name = "combousername";
            this.combousername.Size = new System.Drawing.Size(214, 21);
            this.combousername.TabIndex = 12;
            // 
            // txtparentmenuid
            // 
            this.txtparentmenuid.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtparentmenuid.Location = new System.Drawing.Point(113, 137);
            this.txtparentmenuid.Margin = new System.Windows.Forms.Padding(0);
            this.txtparentmenuid.Name = "txtparentmenuid";
            this.txtparentmenuid.Size = new System.Drawing.Size(214, 20);
            this.txtparentmenuid.TabIndex = 14;
            // 
            // combomenunameid
            // 
            this.combomenunameid.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combomenunameid.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combomenunameid.ContextMenuStrip = this.contextMenuStrip1;
            this.combomenunameid.DropDownHeight = 60;
            this.combomenunameid.FormattingEnabled = true;
            this.combomenunameid.IntegralHeight = false;
            this.combomenunameid.Location = new System.Drawing.Point(113, 56);
            this.combomenunameid.Margin = new System.Windows.Forms.Padding(0);
            this.combomenunameid.Name = "combomenunameid";
            this.combomenunameid.Size = new System.Drawing.Size(214, 21);
            this.combomenunameid.TabIndex = 8;
            this.combomenunameid.SelectedIndexChanged += new System.EventHandler(this.Combomenunameid_SelectedIndexChanged);
            // 
            // combocompcode
            // 
            this.combocompcode.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combocompcode.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combocompcode.DropDownHeight = 60;
            this.combocompcode.FormattingEnabled = true;
            this.combocompcode.IntegralHeight = false;
            this.combocompcode.Location = new System.Drawing.Point(113, 163);
            this.combocompcode.Margin = new System.Windows.Forms.Padding(0);
            this.combocompcode.Name = "combocompcode";
            this.combocompcode.Size = new System.Drawing.Size(214, 21);
            this.combocompcode.TabIndex = 13;
            this.combocompcode.SelectedIndexChanged += new System.EventHandler(this.Combocompcode_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(12, 31);
            this.label1.Margin = new System.Windows.Forms.Padding(0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "MenuID  ";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtmenuname
            // 
            this.txtmenuname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtmenuname.Location = new System.Drawing.Point(113, 85);
            this.txtmenuname.Margin = new System.Windows.Forms.Padding(0);
            this.txtmenuname.Name = "txtmenuname";
            this.txtmenuname.Size = new System.Drawing.Size(213, 20);
            this.txtmenuname.TabIndex = 4;
            // 
            // MENUID
            // 
            this.MENUID.DataPropertyName = "MENUID";
            this.MENUID.HeaderText = "MENUID";
            this.MENUID.Name = "MENUID";
            // 
            // MENUNAMEID
            // 
            this.MENUNAMEID.DataPropertyName = "MENUNAMEID";
            this.MENUNAMEID.HeaderText = "MENUNAMEID";
            this.MENUNAMEID.Name = "MENUNAMEID";
            // 
            // NavigationMaster
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1271, 512);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.ImeMode = System.Windows.Forms.ImeMode.On;
            this.MaximizeBox = false;
            this.Name = "NavigationMaster";
            this.ShowIcon = false;
            this.Text = "Create Navigation";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.NavigationMaster_FormClosed);
            this.Load += new System.EventHandler(this.NavigationMaster_Load);
            this.panel1.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            this.tabControl2.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox combomenunameid;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtmenuname;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtmenuid;
        private System.Windows.Forms.ComboBox combousername;
        private System.Windows.Forms.ComboBox combocompcode;
        private System.Windows.Forms.TextBox txtparentmenuid;
        private System.Windows.Forms.TextBox txtnavurl;
        private System.Windows.Forms.CheckBox chk;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DataGridViewTextBoxColumn MENUID;
        private System.Windows.Forms.DataGridViewTextBoxColumn MENUNAMEID;
        private System.Windows.Forms.TabControl tabControl2;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.ComboBox combousername1;
        private System.Windows.Forms.ComboBox combocompcode1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menuMasterRefreshToolStripMenuItem;
        private System.Windows.Forms.Button butheader;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button butfooter;
        private System.Windows.Forms.Label lbltotal;
        private System.Windows.Forms.TextBox txtsearch;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}