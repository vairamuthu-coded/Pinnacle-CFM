namespace Pinnacle.Transactions.Bank
{
    partial class Payment
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.butheader = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label13 = new System.Windows.Forms.Label();
            this.txtapprovalid = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.txtaccholdername = new System.Windows.Forms.TextBox();
            this.checkall = new System.Windows.Forms.CheckBox();
            this.txtamt = new System.Windows.Forms.NumericUpDown();
            this.label11 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.lblyear = new System.Windows.Forms.Label();
            this.lblmonth = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.txtremarks3 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtremarks2 = new System.Windows.Forms.TextBox();
            this.combobank = new System.Windows.Forms.ComboBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label5 = new System.Windows.Forms.Label();
            this.combopaymenttype = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtremarks = new System.Windows.Forms.TextBox();
            this.txtledgerDate = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.comboparty = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.checkactive = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtledgerid = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Approval = new System.Windows.Forms.TabPage();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.approvalCancelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.butwaitapproval = new System.Windows.Forms.Button();
            this.mySqlCommandBuilder1 = new MySqlConnector.MySqlCommandBuilder();
            this.uccListView1 = new Pinnacle.UserControls.UCCListView();
            this.SNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.asptbladvpaymasid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.dateofpayment = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Department = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.partyname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.orderno = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.itemdesc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.deductionamt = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.advanceterms = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.advanceamount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PaymentType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.resonseperson = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Attach = new System.Windows.Forms.DataGridViewLinkColumn();
            this.panel1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtamt)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.Approval.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.contextMenuStrip2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.butheader);
            this.panel1.Controls.Add(this.tabControl1);
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(1);
            this.panel1.Size = new System.Drawing.Size(1133, 478);
            this.panel1.TabIndex = 51;
            // 
            // butheader
            // 
            this.butheader.BackColor = System.Drawing.Color.Teal;
            this.butheader.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.butheader.Dock = System.Windows.Forms.DockStyle.Top;
            this.butheader.FlatAppearance.BorderSize = 0;
            this.butheader.FlatAppearance.MouseDownBackColor = System.Drawing.Color.WhiteSmoke;
            this.butheader.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.butheader.Font = new System.Drawing.Font("Roboto Black", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butheader.ForeColor = System.Drawing.Color.White;
            this.butheader.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.butheader.Location = new System.Drawing.Point(1, 1);
            this.butheader.Name = "butheader";
            this.butheader.Size = new System.Drawing.Size(1131, 30);
            this.butheader.TabIndex = 5;
            this.butheader.Text = "Payment";
            this.butheader.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.butheader.UseVisualStyleBackColor = false;
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.Buttons;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.Approval);
            this.tabControl1.Location = new System.Drawing.Point(6, 37);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1123, 435);
            this.tabControl1.TabIndex = 0;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.White;
            this.tabPage1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.tabPage1.Controls.Add(this.label13);
            this.tabPage1.Controls.Add(this.txtapprovalid);
            this.tabPage1.Controls.Add(this.label12);
            this.tabPage1.Controls.Add(this.txtaccholdername);
            this.tabPage1.Controls.Add(this.checkall);
            this.tabPage1.Controls.Add(this.txtamt);
            this.tabPage1.Controls.Add(this.label11);
            this.tabPage1.Controls.Add(this.uccListView1);
            this.tabPage1.Controls.Add(this.textBox1);
            this.tabPage1.Controls.Add(this.label10);
            this.tabPage1.Controls.Add(this.lblyear);
            this.tabPage1.Controls.Add(this.lblmonth);
            this.tabPage1.Controls.Add(this.label9);
            this.tabPage1.Controls.Add(this.txtremarks3);
            this.tabPage1.Controls.Add(this.label8);
            this.tabPage1.Controls.Add(this.txtremarks2);
            this.tabPage1.Controls.Add(this.combobank);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.combopaymenttype);
            this.tabPage1.Controls.Add(this.label7);
            this.tabPage1.Controls.Add(this.label6);
            this.tabPage1.Controls.Add(this.txtremarks);
            this.tabPage1.Controls.Add(this.txtledgerDate);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.comboparty);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.checkactive);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.txtledgerid);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(1115, 406);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Payment";
            this.tabPage1.Click += new System.EventHandler(this.tabPage1_Click);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(246, 15);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(63, 13);
            this.label13.TabIndex = 57;
            this.label13.Text = "Approval ID";
            // 
            // txtapprovalid
            // 
            this.txtapprovalid.Enabled = false;
            this.txtapprovalid.Location = new System.Drawing.Point(362, 15);
            this.txtapprovalid.Name = "txtapprovalid";
            this.txtapprovalid.Size = new System.Drawing.Size(117, 20);
            this.txtapprovalid.TabIndex = 56;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(12, 113);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(91, 13);
            this.label12.TabIndex = 55;
            this.label12.Text = "Acc-Holder Name";
            // 
            // txtaccholdername
            // 
            this.txtaccholdername.Enabled = false;
            this.txtaccholdername.Location = new System.Drawing.Point(122, 110);
            this.txtaccholdername.Name = "txtaccholdername";
            this.txtaccholdername.ReadOnly = true;
            this.txtaccholdername.Size = new System.Drawing.Size(357, 20);
            this.txtaccholdername.TabIndex = 0;
            // 
            // checkall
            // 
            this.checkall.AutoSize = true;
            this.checkall.BackColor = System.Drawing.Color.Transparent;
            this.checkall.ForeColor = System.Drawing.Color.White;
            this.checkall.Location = new System.Drawing.Point(1023, 49);
            this.checkall.Name = "checkall";
            this.checkall.Size = new System.Drawing.Size(45, 17);
            this.checkall.TabIndex = 53;
            this.checkall.Text = "ALL";
            this.checkall.UseVisualStyleBackColor = false;
            this.checkall.CheckedChanged += new System.EventHandler(this.checkall_CheckedChanged);
            // 
            // txtamt
            // 
            this.txtamt.DecimalPlaces = 2;
            this.txtamt.Font = new System.Drawing.Font("Roboto Black", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtamt.Increment = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.txtamt.Location = new System.Drawing.Point(122, 136);
            this.txtamt.Maximum = new decimal(new int[] {
            1410065407,
            2,
            0,
            0});
            this.txtamt.Name = "txtamt";
            this.txtamt.Size = new System.Drawing.Size(357, 43);
            this.txtamt.TabIndex = 1;
            this.txtamt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtamt.ThousandsSeparator = true;
            this.txtamt.ValueChanged += new System.EventHandler(this.txtamt_ValueChanged);
            this.txtamt.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtamt_KeyDown);
            this.txtamt.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtamt_KeyPress);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(13, 202);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(82, 13);
            this.label11.TabIndex = 52;
            this.label11.Text = "Amont in Words";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.White;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox1.Enabled = false;
            this.textBox1.Font = new System.Drawing.Font("Roboto Black", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.ForeColor = System.Drawing.Color.Purple;
            this.textBox1.Location = new System.Drawing.Point(122, 185);
            this.textBox1.MaxLength = 500;
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(357, 48);
            this.textBox1.TabIndex = 2;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Roboto Black", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(245, 44);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(92, 20);
            this.label10.TabIndex = 48;
            this.label10.Text = "Month/Year:";
            // 
            // lblyear
            // 
            this.lblyear.AutoSize = true;
            this.lblyear.Font = new System.Drawing.Font("Roboto Black", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblyear.Location = new System.Drawing.Point(428, 45);
            this.lblyear.Name = "lblyear";
            this.lblyear.Size = new System.Drawing.Size(51, 20);
            this.lblyear.TabIndex = 47;
            this.lblyear.Text = "Month";
            // 
            // lblmonth
            // 
            this.lblmonth.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblmonth.AutoSize = true;
            this.lblmonth.Font = new System.Drawing.Font("Roboto Black", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblmonth.Location = new System.Drawing.Point(332, 45);
            this.lblmonth.Name = "lblmonth";
            this.lblmonth.Size = new System.Drawing.Size(51, 20);
            this.lblmonth.TabIndex = 46;
            this.lblmonth.Text = "Month";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(12, 389);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(58, 13);
            this.label9.TabIndex = 45;
            this.label9.Text = "Remarks-3";
            // 
            // txtremarks3
            // 
            this.txtremarks3.BackColor = System.Drawing.Color.White;
            this.txtremarks3.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtremarks3.Location = new System.Drawing.Point(122, 380);
            this.txtremarks3.MaxLength = 100;
            this.txtremarks3.Name = "txtremarks3";
            this.txtremarks3.Size = new System.Drawing.Size(357, 29);
            this.txtremarks3.TabIndex = 7;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(12, 354);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(58, 13);
            this.label8.TabIndex = 43;
            this.label8.Text = "Remarks-2";
            // 
            // txtremarks2
            // 
            this.txtremarks2.BackColor = System.Drawing.Color.White;
            this.txtremarks2.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtremarks2.Location = new System.Drawing.Point(122, 345);
            this.txtremarks2.MaxLength = 100;
            this.txtremarks2.Name = "txtremarks2";
            this.txtremarks2.Size = new System.Drawing.Size(357, 29);
            this.txtremarks2.TabIndex = 6;
            this.txtremarks2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtremarks2_KeyDown);
            // 
            // combobank
            // 
            this.combobank.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combobank.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combobank.BackColor = System.Drawing.Color.White;
            this.combobank.ContextMenuStrip = this.contextMenuStrip1;
            this.combobank.DropDownHeight = 100;
            this.combobank.Enabled = false;
            this.combobank.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.combobank.FormattingEnabled = true;
            this.combobank.IntegralHeight = false;
            this.combobank.Items.AddRange(new object[] {
            "RTGS",
            "NEFT",
            "CHEQUE"});
            this.combobank.Location = new System.Drawing.Point(122, 239);
            this.combobank.Name = "combobank";
            this.combobank.Size = new System.Drawing.Size(357, 29);
            this.combobank.TabIndex = 3;
            this.combobank.SelectedIndexChanged += new System.EventHandler(this.combobank_SelectedIndexChanged);
            this.combobank.KeyDown += new System.Windows.Forms.KeyEventHandler(this.combobank_KeyDown);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.refreshToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(114, 26);
            // 
            // refreshToolStripMenuItem
            // 
            this.refreshToolStripMenuItem.Name = "refreshToolStripMenuItem";
            this.refreshToolStripMenuItem.Size = new System.Drawing.Size(113, 22);
            this.refreshToolStripMenuItem.Text = "Refresh";
            this.refreshToolStripMenuItem.Click += new System.EventHandler(this.RefreshToolStripMenuItem_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(13, 248);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(63, 13);
            this.label5.TabIndex = 40;
            this.label5.Text = "Bank Name";
            // 
            // combopaymenttype
            // 
            this.combopaymenttype.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combopaymenttype.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combopaymenttype.BackColor = System.Drawing.Color.White;
            this.combopaymenttype.ContextMenuStrip = this.contextMenuStrip1;
            this.combopaymenttype.DropDownHeight = 100;
            this.combopaymenttype.Enabled = false;
            this.combopaymenttype.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.combopaymenttype.FormattingEnabled = true;
            this.combopaymenttype.IntegralHeight = false;
            this.combopaymenttype.Items.AddRange(new object[] {
            "RTGS",
            "IFT",
            "NEFT"});
            this.combopaymenttype.Location = new System.Drawing.Point(122, 274);
            this.combopaymenttype.Name = "combopaymenttype";
            this.combopaymenttype.Size = new System.Drawing.Size(357, 29);
            this.combopaymenttype.TabIndex = 4;
            this.combopaymenttype.KeyDown += new System.Windows.Forms.KeyEventHandler(this.combopaymenttype_KeyDown);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(13, 283);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(75, 13);
            this.label7.TabIndex = 38;
            this.label7.Text = "Payment Type";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(12, 323);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(58, 13);
            this.label6.TabIndex = 37;
            this.label6.Text = "Remarks-1";
            // 
            // txtremarks
            // 
            this.txtremarks.BackColor = System.Drawing.Color.White;
            this.txtremarks.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtremarks.Location = new System.Drawing.Point(122, 309);
            this.txtremarks.MaxLength = 100;
            this.txtremarks.Multiline = true;
            this.txtremarks.Name = "txtremarks";
            this.txtremarks.Size = new System.Drawing.Size(357, 30);
            this.txtremarks.TabIndex = 5;
            this.txtremarks.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtremarks_KeyDown);
            // 
            // txtledgerDate
            // 
            this.txtledgerDate.Enabled = false;
            this.txtledgerDate.Font = new System.Drawing.Font("Roboto Black", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtledgerDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txtledgerDate.Location = new System.Drawing.Point(122, 38);
            this.txtledgerDate.Name = "txtledgerDate";
            this.txtledgerDate.Size = new System.Drawing.Size(117, 27);
            this.txtledgerDate.TabIndex = 0;
            this.txtledgerDate.ValueChanged += new System.EventHandler(this.txtledgerDate_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 33;
            this.label1.Text = "DateTime";
            // 
            // comboparty
            // 
            this.comboparty.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboparty.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboparty.BackColor = System.Drawing.Color.White;
            this.comboparty.ContextMenuStrip = this.contextMenuStrip1;
            this.comboparty.DropDownHeight = 100;
            this.comboparty.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboparty.FormattingEnabled = true;
            this.comboparty.IntegralHeight = false;
            this.comboparty.Location = new System.Drawing.Point(122, 74);
            this.comboparty.Name = "comboparty";
            this.comboparty.Size = new System.Drawing.Size(357, 29);
            this.comboparty.TabIndex = 0;
            this.comboparty.SelectedIndexChanged += new System.EventHandler(this.comboparty_SelectedIndexChanged);
            this.comboparty.KeyDown += new System.Windows.Forms.KeyEventHandler(this.comboparty_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 146);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(43, 13);
            this.label3.TabIndex = 26;
            this.label3.Text = "Amount";
            // 
            // checkactive
            // 
            this.checkactive.AutoSize = true;
            this.checkactive.Checked = true;
            this.checkactive.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkactive.Location = new System.Drawing.Point(538, 413);
            this.checkactive.Name = "checkactive";
            this.checkactive.Size = new System.Drawing.Size(56, 17);
            this.checkactive.TabIndex = 25;
            this.checkactive.Text = "Active";
            this.checkactive.UseVisualStyleBackColor = true;
            this.checkactive.Visible = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 83);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 13);
            this.label2.TabIndex = 23;
            this.label2.Text = "PartyName";
            // 
            // txtledgerid
            // 
            this.txtledgerid.Enabled = false;
            this.txtledgerid.Location = new System.Drawing.Point(122, 12);
            this.txtledgerid.Name = "txtledgerid";
            this.txtledgerid.Size = new System.Drawing.Size(117, 20);
            this.txtledgerid.TabIndex = 22;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 19);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(18, 13);
            this.label4.TabIndex = 21;
            this.label4.Text = "ID";
            // 
            // Approval
            // 
            this.Approval.BackColor = System.Drawing.Color.White;
            this.Approval.Controls.Add(this.dataGridView1);
            this.Approval.Controls.Add(this.butwaitapproval);
            this.Approval.Location = new System.Drawing.Point(4, 25);
            this.Approval.Name = "Approval";
            this.Approval.Padding = new System.Windows.Forms.Padding(3);
            this.Approval.Size = new System.Drawing.Size(1115, 406);
            this.Approval.TabIndex = 1;
            this.Approval.Text = "Approval";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.BackgroundColor = System.Drawing.Color.White;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.SNo,
            this.asptbladvpaymasid,
            this.Column1,
            this.dateofpayment,
            this.Department,
            this.partyname,
            this.orderno,
            this.itemdesc,
            this.deductionamt,
            this.advanceterms,
            this.advanceamount,
            this.PaymentType,
            this.resonseperson,
            this.Attach});
            this.dataGridView1.ContextMenuStrip = this.contextMenuStrip2;
            this.dataGridView1.EnableHeadersVisualStyles = false;
            this.dataGridView1.GridColor = System.Drawing.Color.WhiteSmoke;
            this.dataGridView1.Location = new System.Drawing.Point(3, 43);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dataGridView1.Size = new System.Drawing.Size(1109, 360);
            this.dataGridView1.StandardTab = true;
            this.dataGridView1.TabIndex = 6;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentDoubleClick);
            this.dataGridView1.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.approvalCancelToolStripMenuItem});
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(162, 26);
            // 
            // approvalCancelToolStripMenuItem
            // 
            this.approvalCancelToolStripMenuItem.Name = "approvalCancelToolStripMenuItem";
            this.approvalCancelToolStripMenuItem.Size = new System.Drawing.Size(161, 22);
            this.approvalCancelToolStripMenuItem.Text = "Approval Cancel";
            this.approvalCancelToolStripMenuItem.Click += new System.EventHandler(this.approvalCancelToolStripMenuItem_Click_1);
            // 
            // butwaitapproval
            // 
            this.butwaitapproval.Dock = System.Windows.Forms.DockStyle.Top;
            this.butwaitapproval.Font = new System.Drawing.Font("Roboto", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butwaitapproval.Location = new System.Drawing.Point(3, 3);
            this.butwaitapproval.Name = "butwaitapproval";
            this.butwaitapproval.Size = new System.Drawing.Size(1109, 37);
            this.butwaitapproval.TabIndex = 4;
            this.butwaitapproval.Text = "Approved";
            this.butwaitapproval.UseVisualStyleBackColor = true;
            this.butwaitapproval.Click += new System.EventHandler(this.butwaitapproval_Click);
            // 
            // mySqlCommandBuilder1
            // 
            this.mySqlCommandBuilder1.DataAdapter = null;
            this.mySqlCommandBuilder1.QuotePrefix = "`";
            this.mySqlCommandBuilder1.QuoteSuffix = "`";
            // 
            // uccListView1
            // 
            this.uccListView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.uccListView1.Font = new System.Drawing.Font("Roboto", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uccListView1.Location = new System.Drawing.Point(496, 4);
            this.uccListView1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.uccListView1.Name = "uccListView1";
            this.uccListView1.Padding = new System.Windows.Forms.Padding(5);
            this.uccListView1.Size = new System.Drawing.Size(603, 399);
            this.uccListView1.TabIndex = 51;
            this.uccListView1.Click += new System.EventHandler(this.uccListView1_Click);
            // 
            // SNo
            // 
            this.SNo.HeaderText = "S.No";
            this.SNo.Name = "SNo";
            this.SNo.ReadOnly = true;
            this.SNo.Width = 50;
            // 
            // asptbladvpaymasid
            // 
            this.asptbladvpaymasid.DataPropertyName = "asptbladvpaymasid";
            this.asptbladvpaymasid.HeaderText = "ID";
            this.asptbladvpaymasid.Name = "asptbladvpaymasid";
            this.asptbladvpaymasid.ReadOnly = true;
            this.asptbladvpaymasid.Width = 30;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Approval";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            this.Column1.Visible = false;
            this.Column1.Width = 5;
            // 
            // dateofpayment
            // 
            this.dateofpayment.DataPropertyName = "dateofpayment";
            this.dateofpayment.HeaderText = "Date";
            this.dateofpayment.Name = "dateofpayment";
            this.dateofpayment.ReadOnly = true;
            this.dateofpayment.Width = 80;
            // 
            // Department
            // 
            this.Department.DataPropertyName = "department";
            this.Department.HeaderText = "Department";
            this.Department.Name = "Department";
            this.Department.ReadOnly = true;
            this.Department.Width = 120;
            // 
            // partyname
            // 
            this.partyname.DataPropertyName = "partyname";
            this.partyname.HeaderText = "SupplierName";
            this.partyname.Name = "partyname";
            this.partyname.ReadOnly = true;
            this.partyname.Width = 200;
            // 
            // orderno
            // 
            this.orderno.DataPropertyName = "orderno";
            this.orderno.HeaderText = "OrderNo";
            this.orderno.Name = "orderno";
            this.orderno.ReadOnly = true;
            this.orderno.Width = 110;
            // 
            // itemdesc
            // 
            this.itemdesc.DataPropertyName = "itemdesc";
            this.itemdesc.HeaderText = "ItemDescription";
            this.itemdesc.Name = "itemdesc";
            this.itemdesc.ReadOnly = true;
            this.itemdesc.Width = 120;
            // 
            // deductionamt
            // 
            this.deductionamt.DataPropertyName = "deductionamt";
            this.deductionamt.HeaderText = "Amount";
            this.deductionamt.Name = "deductionamt";
            this.deductionamt.ReadOnly = true;
            // 
            // advanceterms
            // 
            this.advanceterms.DataPropertyName = "advanceterms";
            this.advanceterms.HeaderText = "Ad-Term%";
            this.advanceterms.Name = "advanceterms";
            this.advanceterms.ReadOnly = true;
            this.advanceterms.Width = 70;
            // 
            // advanceamount
            // 
            this.advanceamount.DataPropertyName = "advanceamount";
            this.advanceamount.HeaderText = "AdvanceAmt";
            this.advanceamount.Name = "advanceamount";
            this.advanceamount.ReadOnly = true;
            // 
            // PaymentType
            // 
            this.PaymentType.DataPropertyName = "paymenttype";
            this.PaymentType.HeaderText = "PayType";
            this.PaymentType.Name = "PaymentType";
            this.PaymentType.ReadOnly = true;
            this.PaymentType.Width = 80;
            // 
            // resonseperson
            // 
            this.resonseperson.DataPropertyName = "resonseperson";
            this.resonseperson.HeaderText = "ResponsePerson";
            this.resonseperson.Name = "resonseperson";
            this.resonseperson.ReadOnly = true;
            this.resonseperson.Width = 120;
            // 
            // Attach
            // 
            this.Attach.DataPropertyName = "Attach";
            this.Attach.HeaderText = "Attach";
            this.Attach.Name = "Attach";
            this.Attach.ReadOnly = true;
            this.Attach.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.Attach.Text = "Attach";
            this.Attach.UseColumnTextForLinkValue = true;
            // 
            // Payment
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.ClientSize = new System.Drawing.Size(1135, 480);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "Payment";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.Text = "Create Payment";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Payment_FormClosed);
            this.Load += new System.EventHandler(this.Payment_Load);
            this.panel1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtamt)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.Approval.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.contextMenuStrip2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.ComboBox comboparty;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox checkactive;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtledgerid;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker txtledgerDate;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtremarks;
        private System.Windows.Forms.ComboBox combopaymenttype;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox combobank;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtremarks3;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtremarks2;
        private System.Windows.Forms.Label lblmonth;
        private System.Windows.Forms.Label lblyear;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox textBox1;
        
         private UserControls.UCCListView listView1;
        private MySqlConnector.MySqlCommandBuilder mySqlCommandBuilder1;
        private UserControls.UCCListView uccListView1;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button butheader;
        private System.Windows.Forms.NumericUpDown txtamt;
        private System.Windows.Forms.CheckBox checkall;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtaccholdername;
        private System.Windows.Forms.TabPage Approval;
        private System.Windows.Forms.Button butwaitapproval;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtapprovalid;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
        private System.Windows.Forms.ToolStripMenuItem approvalCancelToolStripMenuItem;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn SNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn asptbladvpaymasid;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dateofpayment;
        private System.Windows.Forms.DataGridViewTextBoxColumn Department;
        private System.Windows.Forms.DataGridViewTextBoxColumn partyname;
        private System.Windows.Forms.DataGridViewTextBoxColumn orderno;
        private System.Windows.Forms.DataGridViewTextBoxColumn itemdesc;
        private System.Windows.Forms.DataGridViewTextBoxColumn deductionamt;
        private System.Windows.Forms.DataGridViewTextBoxColumn advanceterms;
        private System.Windows.Forms.DataGridViewTextBoxColumn advanceamount;
        private System.Windows.Forms.DataGridViewTextBoxColumn PaymentType;
        private System.Windows.Forms.DataGridViewTextBoxColumn resonseperson;
        private System.Windows.Forms.DataGridViewLinkColumn Attach;
    }
}