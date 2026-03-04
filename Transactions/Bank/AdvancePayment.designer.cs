namespace Pinnacle.Transactions.Bank
{
    partial class AdvancePayment
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
            this.butheader = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label20 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.txtgstvalue = new System.Windows.Forms.NumericUpDown();
            this.label19 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.txtamtinwords = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.rowDeleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label17 = new System.Windows.Forms.Label();
            this.txtdedutionamt = new System.Windows.Forms.NumericUpDown();
            this.label16 = new System.Windows.Forms.Label();
            this.txttdsvalue = new System.Windows.Forms.NumericUpDown();
            this.label15 = new System.Windows.Forms.Label();
            this.txttds = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.txtgst = new System.Windows.Forms.TextBox();
            this.txtamount = new System.Windows.Forms.NumericUpDown();
            this.label7 = new System.Windows.Forms.Label();
            this.txtadvamount = new System.Windows.Forms.NumericUpDown();
            this.combomodeofpayment = new System.Windows.Forms.ComboBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.txtinvoiceamt = new System.Windows.Forms.NumericUpDown();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.txtitemdesc = new System.Windows.Forms.TextBox();
            this.comboparty = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.txtorderno = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.lblyear = new System.Windows.Forms.Label();
            this.lblmonth = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.comboresponsibleperson = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtadvterms = new System.Windows.Forms.TextBox();
            this.txtadvDate = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.combodept = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtadvpayid = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.checkall = new System.Windows.Forms.CheckBox();
            this.uccListView2 = new Pinnacle.UserControls.UCCListView();
            this.mySqlCommandBuilder1 = new MySqlConnector.MySqlCommandBuilder();
            this.asptbladvpaydetid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.InvoiceType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.Invoice = new System.Windows.Forms.DataGridViewButtonColumn();
            this.Show = new System.Windows.Forms.DataGridViewLinkColumn();
            this.panel1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtgstvalue)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.contextMenuStrip2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtdedutionamt)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txttdsvalue)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtamount)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtadvamount)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtinvoiceamt)).BeginInit();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.butheader);
            this.panel1.Controls.Add(this.tabControl1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(1);
            this.panel1.Size = new System.Drawing.Size(1266, 517);
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
            this.butheader.Size = new System.Drawing.Size(1264, 30);
            this.butheader.TabIndex = 5;
            this.butheader.Text = "Advance Payment";
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
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(6, 37);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1256, 474);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.White;
            this.tabPage1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.tabPage1.Controls.Add(this.label20);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.txtgstvalue);
            this.tabPage1.Controls.Add(this.label19);
            this.tabPage1.Controls.Add(this.label18);
            this.tabPage1.Controls.Add(this.txtamtinwords);
            this.tabPage1.Controls.Add(this.dataGridView1);
            this.tabPage1.Controls.Add(this.label17);
            this.tabPage1.Controls.Add(this.txtdedutionamt);
            this.tabPage1.Controls.Add(this.label16);
            this.tabPage1.Controls.Add(this.txttdsvalue);
            this.tabPage1.Controls.Add(this.label15);
            this.tabPage1.Controls.Add(this.txttds);
            this.tabPage1.Controls.Add(this.label14);
            this.tabPage1.Controls.Add(this.txtgst);
            this.tabPage1.Controls.Add(this.txtamount);
            this.tabPage1.Controls.Add(this.label7);
            this.tabPage1.Controls.Add(this.txtadvamount);
            this.tabPage1.Controls.Add(this.combomodeofpayment);
            this.tabPage1.Controls.Add(this.txtinvoiceamt);
            this.tabPage1.Controls.Add(this.label13);
            this.tabPage1.Controls.Add(this.label12);
            this.tabPage1.Controls.Add(this.txtitemdesc);
            this.tabPage1.Controls.Add(this.comboparty);
            this.tabPage1.Controls.Add(this.label11);
            this.tabPage1.Controls.Add(this.txtorderno);
            this.tabPage1.Controls.Add(this.label10);
            this.tabPage1.Controls.Add(this.lblyear);
            this.tabPage1.Controls.Add(this.lblmonth);
            this.tabPage1.Controls.Add(this.label9);
            this.tabPage1.Controls.Add(this.label8);
            this.tabPage1.Controls.Add(this.comboresponsibleperson);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.label6);
            this.tabPage1.Controls.Add(this.txtadvterms);
            this.tabPage1.Controls.Add(this.txtadvDate);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.combodept);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.txtadvpayid);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(1248, 445);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Advance";
            this.tabPage1.Click += new System.EventHandler(this.tabPage1_Click);
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(11, 22);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(18, 13);
            this.label20.TabIndex = 123;
            this.label20.Text = "ID";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 311);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(31, 13);
            this.label4.TabIndex = 122;
            this.label4.Text = "Total";
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Location = new System.Drawing.Point(707, 224);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(518, 3);
            this.button1.TabIndex = 121;
            this.button1.UseVisualStyleBackColor = true;
            // 
            // txtgstvalue
            // 
            this.txtgstvalue.DecimalPlaces = 2;
            this.txtgstvalue.Enabled = false;
            this.txtgstvalue.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtgstvalue.Increment = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.txtgstvalue.Location = new System.Drawing.Point(532, 271);
            this.txtgstvalue.Maximum = new decimal(new int[] {
            -727379969,
            232,
            0,
            0});
            this.txtgstvalue.Name = "txtgstvalue";
            this.txtgstvalue.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtgstvalue.Size = new System.Drawing.Size(169, 29);
            this.txtgstvalue.TabIndex = 7;
            this.txtgstvalue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtgstvalue.ThousandsSeparator = true;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(11, 384);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(70, 13);
            this.label19.TabIndex = 112;
            this.label19.Text = "Amt in Words";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Font = new System.Drawing.Font("Roboto Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label18.Location = new System.Drawing.Point(714, 4);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(102, 18);
            this.label18.TabIndex = 111;
            this.label18.Text = "Add Attachment";
            // 
            // txtamtinwords
            // 
            this.txtamtinwords.BackColor = System.Drawing.Color.White;
            this.txtamtinwords.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtamtinwords.Enabled = false;
            this.txtamtinwords.Font = new System.Drawing.Font("Roboto Black", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtamtinwords.ForeColor = System.Drawing.Color.Black;
            this.txtamtinwords.Location = new System.Drawing.Point(132, 376);
            this.txtamtinwords.MaxLength = 25;
            this.txtamtinwords.Multiline = true;
            this.txtamtinwords.Name = "txtamtinwords";
            this.txtamtinwords.Size = new System.Drawing.Size(569, 52);
            this.txtamtinwords.TabIndex = 102;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.BackgroundColor = System.Drawing.Color.White;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dataGridView1.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleHorizontal;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Teal;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.asptbladvpaydetid,
            this.InvoiceType,
            this.Invoice,
            this.Show});
            this.dataGridView1.ContextMenuStrip = this.contextMenuStrip2;
            this.dataGridView1.EnableHeadersVisualStyles = false;
            this.dataGridView1.GridColor = System.Drawing.Color.Gainsboro;
            this.dataGridView1.Location = new System.Drawing.Point(707, 25);
            this.dataGridView1.Name = "dataGridView1";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.Teal;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.Size = new System.Drawing.Size(508, 193);
            this.dataGridView1.TabIndex = 10;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellValidated += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValidated);
            this.dataGridView1.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dataGridView1_DataError);
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.rowDeleteToolStripMenuItem});
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(134, 26);
            // 
            // rowDeleteToolStripMenuItem
            // 
            this.rowDeleteToolStripMenuItem.Name = "rowDeleteToolStripMenuItem";
            this.rowDeleteToolStripMenuItem.Size = new System.Drawing.Size(133, 22);
            this.rowDeleteToolStripMenuItem.Text = "Row Delete";
            this.rowDeleteToolStripMenuItem.Click += new System.EventHandler(this.rowDeleteToolStripMenuItem_Click);
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(11, 345);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(63, 13);
            this.label17.TabIndex = 100;
            this.label17.Text = "Grand Total";
            this.label17.Click += new System.EventHandler(this.label17_Click);
            // 
            // txtdedutionamt
            // 
            this.txtdedutionamt.DecimalPlaces = 2;
            this.txtdedutionamt.Enabled = false;
            this.txtdedutionamt.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtdedutionamt.Increment = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.txtdedutionamt.Location = new System.Drawing.Point(132, 338);
            this.txtdedutionamt.Maximum = new decimal(new int[] {
            -727379969,
            232,
            0,
            0});
            this.txtdedutionamt.Name = "txtdedutionamt";
            this.txtdedutionamt.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtdedutionamt.Size = new System.Drawing.Size(178, 29);
            this.txtdedutionamt.TabIndex = 8;
            this.txtdedutionamt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtdedutionamt.ThousandsSeparator = true;
            this.txtdedutionamt.ValueChanged += new System.EventHandler(this.txtdedutionamt_ValueChanged);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(454, 314);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(59, 13);
            this.label16.TabIndex = 98;
            this.label16.Text = "TDS Value";
            // 
            // txttdsvalue
            // 
            this.txttdsvalue.DecimalPlaces = 2;
            this.txttdsvalue.Enabled = false;
            this.txttdsvalue.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txttdsvalue.Increment = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.txttdsvalue.Location = new System.Drawing.Point(532, 306);
            this.txttdsvalue.Maximum = new decimal(new int[] {
            -727379969,
            232,
            0,
            0});
            this.txttdsvalue.Name = "txttdsvalue";
            this.txttdsvalue.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txttdsvalue.Size = new System.Drawing.Size(169, 29);
            this.txttdsvalue.TabIndex = 8;
            this.txttdsvalue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txttdsvalue.ThousandsSeparator = true;
            this.txttdsvalue.ValueChanged += new System.EventHandler(this.txttdsvalue_ValueChanged);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(316, 314);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(40, 13);
            this.label15.TabIndex = 95;
            this.label15.Text = "TDS %";
            // 
            // txttds
            // 
            this.txttds.BackColor = System.Drawing.Color.White;
            this.txttds.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txttds.Location = new System.Drawing.Point(396, 307);
            this.txttds.MaxLength = 2;
            this.txttds.Name = "txttds";
            this.txttds.Size = new System.Drawing.Size(40, 29);
            this.txttds.TabIndex = 8;
            this.txttds.TextChanged += new System.EventHandler(this.txttds_TextChanged);
            this.txttds.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txttds_KeyDown);
            this.txttds.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txttds_KeyPress);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(316, 281);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(40, 13);
            this.label14.TabIndex = 93;
            this.label14.Text = "GST %";
            // 
            // txtgst
            // 
            this.txtgst.BackColor = System.Drawing.Color.White;
            this.txtgst.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtgst.Location = new System.Drawing.Point(396, 272);
            this.txtgst.MaxLength = 2;
            this.txtgst.Name = "txtgst";
            this.txtgst.Size = new System.Drawing.Size(40, 29);
            this.txtgst.TabIndex = 7;
            this.txtgst.WordWrap = false;
            this.txtgst.TextChanged += new System.EventHandler(this.txtgst_TextChanged);
            this.txtgst.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtgst_KeyDown);
            this.txtgst.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtgst_KeyPress);
            // 
            // txtamount
            // 
            this.txtamount.BackColor = System.Drawing.Color.White;
            this.txtamount.DecimalPlaces = 2;
            this.txtamount.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtamount.Increment = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.txtamount.Location = new System.Drawing.Point(132, 271);
            this.txtamount.Maximum = new decimal(new int[] {
            1410065407,
            2,
            0,
            0});
            this.txtamount.Name = "txtamount";
            this.txtamount.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtamount.Size = new System.Drawing.Size(178, 29);
            this.txtamount.TabIndex = 6;
            this.txtamount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtamount.ThousandsSeparator = true;
            this.txtamount.ValueChanged += new System.EventHandler(this.txtamount_ValueChanged);
            this.txtamount.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtadvamount_KeyDown);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(11, 280);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(43, 13);
            this.label7.TabIndex = 91;
            this.label7.Text = "Amount";
            // 
            // txtadvamount
            // 
            this.txtadvamount.Enabled = false;
            this.txtadvamount.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtadvamount.Increment = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.txtadvamount.InterceptArrowKeys = false;
            this.txtadvamount.Location = new System.Drawing.Point(532, 341);
            this.txtadvamount.Maximum = new decimal(new int[] {
            -727379969,
            232,
            0,
            0});
            this.txtadvamount.Name = "txtadvamount";
            this.txtadvamount.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtadvamount.Size = new System.Drawing.Size(169, 29);
            this.txtadvamount.TabIndex = 9;
            this.txtadvamount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtadvamount.ValueChanged += new System.EventHandler(this.txtadvamount_ValueChanged);
            this.txtadvamount.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtadvamount_KeyDown);
            // 
            // combomodeofpayment
            // 
            this.combomodeofpayment.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combomodeofpayment.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combomodeofpayment.BackColor = System.Drawing.Color.White;
            this.combomodeofpayment.ContextMenuStrip = this.contextMenuStrip1;
            this.combomodeofpayment.DropDownHeight = 100;
            this.combomodeofpayment.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.combomodeofpayment.FormattingEnabled = true;
            this.combomodeofpayment.IntegralHeight = false;
            this.combomodeofpayment.Items.AddRange(new object[] {
            "CASH",
            "CHEQUE",
            "RTGS"});
            this.combomodeofpayment.Location = new System.Drawing.Point(550, 233);
            this.combomodeofpayment.Name = "combomodeofpayment";
            this.combomodeofpayment.Size = new System.Drawing.Size(151, 29);
            this.combomodeofpayment.TabIndex = 5;
            this.combomodeofpayment.KeyDown += new System.Windows.Forms.KeyEventHandler(this.combomodeofpayment_KeyDown);
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
            // txtinvoiceamt
            // 
            this.txtinvoiceamt.DecimalPlaces = 2;
            this.txtinvoiceamt.Enabled = false;
            this.txtinvoiceamt.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtinvoiceamt.Increment = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.txtinvoiceamt.Location = new System.Drawing.Point(132, 303);
            this.txtinvoiceamt.Maximum = new decimal(new int[] {
            -727379969,
            232,
            0,
            0});
            this.txtinvoiceamt.Name = "txtinvoiceamt";
            this.txtinvoiceamt.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtinvoiceamt.Size = new System.Drawing.Size(178, 29);
            this.txtinvoiceamt.TabIndex = 7;
            this.txtinvoiceamt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtinvoiceamt.ThousandsSeparator = true;
            this.txtinvoiceamt.ValueChanged += new System.EventHandler(this.txtinvoiceamt_ValueChanged);
            this.txtinvoiceamt.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtinvoiceamt_KeyDown);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(454, 279);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(53, 13);
            this.label13.TabIndex = 81;
            this.label13.Text = "Gst Value";
            this.label13.Click += new System.EventHandler(this.label13_Click);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(11, 203);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(55, 13);
            this.label12.TabIndex = 80;
            this.label12.Text = "Item Desc";
            // 
            // txtitemdesc
            // 
            this.txtitemdesc.BackColor = System.Drawing.Color.White;
            this.txtitemdesc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtitemdesc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtitemdesc.Font = new System.Drawing.Font("Roboto", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtitemdesc.ForeColor = System.Drawing.Color.Black;
            this.txtitemdesc.Location = new System.Drawing.Point(132, 192);
            this.txtitemdesc.MaxLength = 500;
            this.txtitemdesc.Multiline = true;
            this.txtitemdesc.Name = "txtitemdesc";
            this.txtitemdesc.Size = new System.Drawing.Size(569, 35);
            this.txtitemdesc.TabIndex = 3;
            this.txtitemdesc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtitemdesc_KeyDown);
            this.txtitemdesc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtitemdesc_KeyPress);
            // 
            // comboparty
            // 
            this.comboparty.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.comboparty.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboparty.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboparty.BackColor = System.Drawing.Color.White;
            this.comboparty.ContextMenuStrip = this.contextMenuStrip1;
            this.comboparty.DropDownHeight = 100;
            this.comboparty.Font = new System.Drawing.Font("Roboto", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboparty.FormattingEnabled = true;
            this.comboparty.IntegralHeight = false;
            this.comboparty.Location = new System.Drawing.Point(132, 118);
            this.comboparty.Name = "comboparty";
            this.comboparty.Size = new System.Drawing.Size(569, 33);
            this.comboparty.TabIndex = 1;
            this.comboparty.KeyDown += new System.Windows.Forms.KeyEventHandler(this.comboparty_KeyDown);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(11, 165);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(47, 13);
            this.label11.TabIndex = 77;
            this.label11.Text = "OrderNo";
            // 
            // txtorderno
            // 
            this.txtorderno.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txtorderno.BackColor = System.Drawing.Color.White;
            this.txtorderno.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtorderno.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtorderno.Font = new System.Drawing.Font("Roboto Black", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtorderno.ForeColor = System.Drawing.Color.Black;
            this.txtorderno.Location = new System.Drawing.Point(132, 157);
            this.txtorderno.MaxLength = 25;
            this.txtorderno.Name = "txtorderno";
            this.txtorderno.Size = new System.Drawing.Size(569, 29);
            this.txtorderno.TabIndex = 2;
            this.txtorderno.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtorderno_KeyDown);
            this.txtorderno.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtorderno_KeyPress);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Roboto", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(309, 44);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(135, 25);
            this.label10.TabIndex = 76;
            this.label10.Text = "Month / Year :";
            // 
            // lblyear
            // 
            this.lblyear.AutoSize = true;
            this.lblyear.Font = new System.Drawing.Font("Roboto", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblyear.Location = new System.Drawing.Point(615, 47);
            this.lblyear.Name = "lblyear";
            this.lblyear.Size = new System.Drawing.Size(50, 25);
            this.lblyear.TabIndex = 75;
            this.lblyear.Text = "Year";
            // 
            // lblmonth
            // 
            this.lblmonth.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblmonth.AutoSize = true;
            this.lblmonth.Font = new System.Drawing.Font("Roboto", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblmonth.Location = new System.Drawing.Point(463, 47);
            this.lblmonth.Name = "lblmonth";
            this.lblmonth.Size = new System.Drawing.Size(69, 25);
            this.lblmonth.TabIndex = 74;
            this.lblmonth.Text = "Month";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(442, 242);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(90, 13);
            this.label9.TabIndex = 73;
            this.label9.Text = "Mode of Payment";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(454, 350);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(64, 13);
            this.label8.TabIndex = 72;
            this.label8.Text = "Pay Amount";
            this.label8.Click += new System.EventHandler(this.label8_Click);
            // 
            // comboresponsibleperson
            // 
            this.comboresponsibleperson.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboresponsibleperson.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboresponsibleperson.BackColor = System.Drawing.Color.White;
            this.comboresponsibleperson.ContextMenuStrip = this.contextMenuStrip1;
            this.comboresponsibleperson.DropDownHeight = 100;
            this.comboresponsibleperson.Font = new System.Drawing.Font("Roboto", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboresponsibleperson.FormattingEnabled = true;
            this.comboresponsibleperson.IntegralHeight = false;
            this.comboresponsibleperson.Items.AddRange(new object[] {
            "RTGS",
            "NEFT",
            "CHEQUE"});
            this.comboresponsibleperson.Location = new System.Drawing.Point(132, 233);
            this.comboresponsibleperson.Name = "comboresponsibleperson";
            this.comboresponsibleperson.Size = new System.Drawing.Size(304, 33);
            this.comboresponsibleperson.TabIndex = 4;
            this.comboresponsibleperson.KeyDown += new System.Windows.Forms.KeyEventHandler(this.comboresponsibleperson_KeyDown);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(11, 245);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 13);
            this.label5.TabIndex = 71;
            this.label5.Text = "Res- Person";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(316, 350);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(69, 13);
            this.label6.TabIndex = 70;
            this.label6.Text = "Adv Terms %";
            // 
            // txtadvterms
            // 
            this.txtadvterms.BackColor = System.Drawing.Color.White;
            this.txtadvterms.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtadvterms.Location = new System.Drawing.Point(396, 341);
            this.txtadvterms.MaxLength = 3;
            this.txtadvterms.Name = "txtadvterms";
            this.txtadvterms.Size = new System.Drawing.Size(40, 29);
            this.txtadvterms.TabIndex = 9;
            this.txtadvterms.TextChanged += new System.EventHandler(this.txtadvterms_TextChanged);
            this.txtadvterms.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtadvterms_KeyDown);
            this.txtadvterms.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtadvterms_KeyPress);
            // 
            // txtadvDate
            // 
            this.txtadvDate.Enabled = false;
            this.txtadvDate.Font = new System.Drawing.Font("Roboto", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtadvDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.txtadvDate.Location = new System.Drawing.Point(132, 41);
            this.txtadvDate.Name = "txtadvDate";
            this.txtadvDate.Size = new System.Drawing.Size(148, 33);
            this.txtadvDate.TabIndex = 60;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 69;
            this.label1.Text = "DateTime";
            // 
            // combodept
            // 
            this.combodept.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.combodept.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combodept.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combodept.BackColor = System.Drawing.Color.White;
            this.combodept.ContextMenuStrip = this.contextMenuStrip1;
            this.combodept.DropDownHeight = 100;
            this.combodept.Font = new System.Drawing.Font("Roboto", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.combodept.FormattingEnabled = true;
            this.combodept.IntegralHeight = false;
            this.combodept.Location = new System.Drawing.Point(132, 79);
            this.combodept.Name = "combodept";
            this.combodept.Size = new System.Drawing.Size(569, 33);
            this.combodept.TabIndex = 0;
            this.combodept.SelectedIndexChanged += new System.EventHandler(this.combodept_SelectedIndexChanged);
            this.combodept.KeyDown += new System.Windows.Forms.KeyEventHandler(this.combodept_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 130);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 13);
            this.label3.TabIndex = 68;
            this.label3.Text = "Suplier Name";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 91);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 13);
            this.label2.TabIndex = 67;
            this.label2.Text = "Department";
            // 
            // txtadvpayid
            // 
            this.txtadvpayid.Enabled = false;
            this.txtadvpayid.Location = new System.Drawing.Point(132, 15);
            this.txtadvpayid.Name = "txtadvpayid";
            this.txtadvpayid.Size = new System.Drawing.Size(148, 20);
            this.txtadvpayid.TabIndex = 66;
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.White;
            this.tabPage2.Controls.Add(this.checkall);
            this.tabPage2.Controls.Add(this.uccListView2);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1248, 445);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "ListView";
            // 
            // checkall
            // 
            this.checkall.AutoSize = true;
            this.checkall.BackColor = System.Drawing.Color.Transparent;
            this.checkall.ForeColor = System.Drawing.Color.White;
            this.checkall.Location = new System.Drawing.Point(1080, 47);
            this.checkall.Name = "checkall";
            this.checkall.Size = new System.Drawing.Size(45, 17);
            this.checkall.TabIndex = 87;
            this.checkall.Text = "ALL";
            this.checkall.UseVisualStyleBackColor = false;
            this.checkall.CheckedChanged += new System.EventHandler(this.checkall_CheckedChanged);
            // 
            // uccListView2
            // 
            this.uccListView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uccListView2.Font = new System.Drawing.Font("Roboto", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uccListView2.Location = new System.Drawing.Point(3, 3);
            this.uccListView2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.uccListView2.Name = "uccListView2";
            this.uccListView2.Padding = new System.Windows.Forms.Padding(5);
            this.uccListView2.Size = new System.Drawing.Size(1242, 439);
            this.uccListView2.TabIndex = 86;
            // 
            // mySqlCommandBuilder1
            // 
            this.mySqlCommandBuilder1.DataAdapter = null;
            this.mySqlCommandBuilder1.QuotePrefix = "`";
            this.mySqlCommandBuilder1.QuoteSuffix = "`";
            // 
            // asptbladvpaydetid
            // 
            this.asptbladvpaydetid.DataPropertyName = "asptbladvpaydetid";
            this.asptbladvpaydetid.HeaderText = "ID";
            this.asptbladvpaydetid.Name = "asptbladvpaydetid";
            this.asptbladvpaydetid.ReadOnly = true;
            this.asptbladvpaydetid.Width = 50;
            // 
            // InvoiceType
            // 
            this.InvoiceType.DataPropertyName = "InvoiceType";
            this.InvoiceType.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.InvoiceType.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.InvoiceType.HeaderText = "InvoiceType";
            this.InvoiceType.Items.AddRange(new object[] {
            "INVOICE",
            "PROFORMA INVOICE",
            "QUATATION",
            "PO /WO",
            "OTHERS"});
            this.InvoiceType.Name = "InvoiceType";
            this.InvoiceType.Width = 200;
            // 
            // Invoice
            // 
            this.Invoice.DataPropertyName = "Invoice";
            this.Invoice.FillWeight = 1F;
            this.Invoice.HeaderText = "ADD File";
            this.Invoice.Name = "Invoice";
            this.Invoice.Text = "Upload";
            this.Invoice.ToolTipText = "Upload Image";
            this.Invoice.Width = 80;
            // 
            // Show
            // 
            this.Show.DataPropertyName = "Show";
            this.Show.HeaderText = "Attach";
            this.Show.Name = "Show";
            this.Show.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.Show.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.Show.Text = "Show";
            this.Show.UseColumnTextForLinkValue = true;
            this.Show.Width = 80;
            // 
            // AdvancePayment
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.ClientSize = new System.Drawing.Size(1268, 519);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "AdvancePayment";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.Text = "Create Advance Payment";
            this.Load += new System.EventHandler(this.AdvancePayment_Load);
            this.panel1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtgstvalue)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.contextMenuStrip2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.txtdedutionamt)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txttdsvalue)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtamount)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtadvamount)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.txtinvoiceamt)).EndInit();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
        
         private UserControls.UCCListView listView1;
        private MySqlConnector.MySqlCommandBuilder mySqlCommandBuilder1;
        private System.Windows.Forms.Button butheader;
        private UserControls.UCCListView uccListView1;
        private System.Windows.Forms.ComboBox combomodeofpayment;
        private System.Windows.Forms.NumericUpDown txtinvoiceamt;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtitemdesc;
        private System.Windows.Forms.ComboBox comboparty;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtorderno;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label lblyear;
        private System.Windows.Forms.Label lblmonth;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox comboresponsibleperson;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtadvterms;
        private System.Windows.Forms.DateTimePicker txtadvDate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox combodept;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtadvpayid;
        private System.Windows.Forms.NumericUpDown txtadvamount;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txttds;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.NumericUpDown txtamount;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.NumericUpDown txttdsvalue;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.NumericUpDown txtdedutionamt;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TextBox txtamtinwords;
        private System.Windows.Forms.TextBox txtgst;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
        private System.Windows.Forms.ToolStripMenuItem rowDeleteToolStripMenuItem;
        private System.Windows.Forms.TabPage tabPage2;
        private UserControls.UCCListView uccListView2;
        private System.Windows.Forms.NumericUpDown txtgstvalue;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.CheckBox checkall;
        private System.Windows.Forms.DataGridViewTextBoxColumn asptbladvpaydetid;
        private System.Windows.Forms.DataGridViewComboBoxColumn InvoiceType;
        private System.Windows.Forms.DataGridViewButtonColumn Invoice;
        private System.Windows.Forms.DataGridViewLinkColumn Show;
    }
}