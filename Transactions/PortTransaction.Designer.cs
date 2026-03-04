
namespace Pinnacle.Transactions
{
    partial class PortTransaction
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
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.Field0 = new System.Windows.Forms.TextBox();
            this.lblcon = new System.Windows.Forms.Label();
            this.butgetdata = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.comboport = new System.Windows.Forms.ComboBox();
            this.label16 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label18 = new System.Windows.Forms.Label();
            this.combostopbits = new System.Windows.Forms.ComboBox();
            this.label24 = new System.Windows.Forms.Label();
            this.comboparity = new System.Windows.Forms.ComboBox();
            this.label25 = new System.Windows.Forms.Label();
            this.combodatabits = new System.Windows.Forms.ComboBox();
            this.label26 = new System.Windows.Forms.Label();
            this.combobaudrate = new System.Windows.Forms.ComboBox();
            this.label27 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.combounit = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtsearch = new System.Windows.Forms.TextBox();
            this.lblsearch = new System.Windows.Forms.Label();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader9 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader6 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader7 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader8 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.panel3 = new System.Windows.Forms.Panel();
            this.lbltotal = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.butheader = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(6, 41);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1007, 408);
            this.tabControl1.TabIndex = 69;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.White;
            this.tabPage1.Controls.Add(this.textBox1);
            this.tabPage1.Controls.Add(this.Field0);
            this.tabPage1.Controls.Add(this.lblcon);
            this.tabPage1.Controls.Add(this.butgetdata);
            this.tabPage1.Controls.Add(this.progressBar1);
            this.tabPage1.Controls.Add(this.comboport);
            this.tabPage1.Controls.Add(this.label16);
            this.tabPage1.Controls.Add(this.comboBox2);
            this.tabPage1.Controls.Add(this.label18);
            this.tabPage1.Controls.Add(this.combostopbits);
            this.tabPage1.Controls.Add(this.label24);
            this.tabPage1.Controls.Add(this.comboparity);
            this.tabPage1.Controls.Add(this.label25);
            this.tabPage1.Controls.Add(this.combodatabits);
            this.tabPage1.Controls.Add(this.label26);
            this.tabPage1.Controls.Add(this.combobaudrate);
            this.tabPage1.Controls.Add(this.label27);
            this.tabPage1.Font = new System.Drawing.Font("Roboto", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(999, 379);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Port Transaction";
            // 
            // textBox1
            // 
            this.textBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.textBox1.Font = new System.Drawing.Font("Palatino Linotype", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(457, 75);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(404, 34);
            this.textBox1.TabIndex = 36;
            // 
            // Field0
            // 
            this.Field0.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.Field0.Font = new System.Drawing.Font("Palatino Linotype", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Field0.Location = new System.Drawing.Point(457, 29);
            this.Field0.Multiline = true;
            this.Field0.Name = "Field0";
            this.Field0.Size = new System.Drawing.Size(404, 34);
            this.Field0.TabIndex = 35;
            // 
            // lblcon
            // 
            this.lblcon.AutoSize = true;
            this.lblcon.Font = new System.Drawing.Font("Palatino Linotype", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblcon.Location = new System.Drawing.Point(136, 261);
            this.lblcon.Name = "lblcon";
            this.lblcon.Size = new System.Drawing.Size(94, 16);
            this.lblcon.TabIndex = 39;
            this.lblcon.Text = "Connection State";
            // 
            // butgetdata
            // 
            this.butgetdata.Location = new System.Drawing.Point(597, 141);
            this.butgetdata.Name = "butgetdata";
            this.butgetdata.Size = new System.Drawing.Size(103, 33);
            this.butgetdata.TabIndex = 34;
            this.butgetdata.Text = "getdata";
            this.butgetdata.UseVisualStyleBackColor = true;
            this.butgetdata.Click += new System.EventHandler(this.butgetdata_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(139, 226);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(232, 23);
            this.progressBar1.TabIndex = 38;
            // 
            // comboport
            // 
            this.comboport.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboport.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboport.FormattingEnabled = true;
            this.comboport.Items.AddRange(new object[] {
            ""});
            this.comboport.Location = new System.Drawing.Point(139, 20);
            this.comboport.Name = "comboport";
            this.comboport.Size = new System.Drawing.Size(232, 26);
            this.comboport.TabIndex = 24;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(28, 29);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(58, 18);
            this.label16.TabIndex = 37;
            this.label16.Text = "ComPort";
            // 
            // comboBox2
            // 
            this.comboBox2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboBox2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "Xon/X/off",
            "Harkware",
            "None"});
            this.comboBox2.Location = new System.Drawing.Point(139, 180);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(232, 26);
            this.comboBox2.TabIndex = 31;
            this.comboBox2.Text = "None";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(28, 180);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(80, 18);
            this.label18.TabIndex = 33;
            this.label18.Text = "Flow Control";
            // 
            // combostopbits
            // 
            this.combostopbits.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combostopbits.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combostopbits.FormattingEnabled = true;
            this.combostopbits.Items.AddRange(new object[] {
            "1",
            "1.5",
            "2"});
            this.combostopbits.Location = new System.Drawing.Point(139, 148);
            this.combostopbits.Name = "combostopbits";
            this.combostopbits.Size = new System.Drawing.Size(232, 26);
            this.combostopbits.TabIndex = 29;
            this.combostopbits.Text = "1";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(28, 149);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(61, 18);
            this.label24.TabIndex = 32;
            this.label24.Text = "Stop Bits";
            // 
            // comboparity
            // 
            this.comboparity.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboparity.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboparity.FormattingEnabled = true;
            this.comboparity.Items.AddRange(new object[] {
            "Even",
            "Odd",
            "None",
            "Mark",
            "Space"});
            this.comboparity.Location = new System.Drawing.Point(139, 116);
            this.comboparity.Name = "comboparity";
            this.comboparity.Size = new System.Drawing.Size(232, 26);
            this.comboparity.TabIndex = 28;
            this.comboparity.Text = "Even";
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(28, 116);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(42, 18);
            this.label25.TabIndex = 30;
            this.label25.Text = "Parity";
            // 
            // combodatabits
            // 
            this.combodatabits.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combodatabits.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combodatabits.FormattingEnabled = true;
            this.combodatabits.Items.AddRange(new object[] {
            "5",
            "6",
            "7",
            "8"});
            this.combodatabits.Location = new System.Drawing.Point(139, 84);
            this.combodatabits.Name = "combodatabits";
            this.combodatabits.Size = new System.Drawing.Size(232, 26);
            this.combodatabits.TabIndex = 26;
            this.combodatabits.Text = "7";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(29, 84);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(62, 18);
            this.label26.TabIndex = 27;
            this.label26.Text = "Data Bits";
            // 
            // combobaudrate
            // 
            this.combobaudrate.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combobaudrate.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combobaudrate.FormattingEnabled = true;
            this.combobaudrate.Items.AddRange(new object[] {
            "110",
            "300",
            "1200",
            "2400",
            "4800",
            "9600",
            "19200",
            "38400",
            "57600",
            "115200",
            "230400",
            "460800",
            "921600"});
            this.combobaudrate.Location = new System.Drawing.Point(139, 52);
            this.combobaudrate.Name = "combobaudrate";
            this.combobaudrate.Size = new System.Drawing.Size(232, 26);
            this.combobaudrate.TabIndex = 25;
            this.combobaudrate.Text = "9600";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(28, 56);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(68, 18);
            this.label27.TabIndex = 23;
            this.label27.Text = "Baud Rate";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.White;
            this.tabPage2.Controls.Add(this.panel2);
            this.tabPage2.Controls.Add(this.listView1);
            this.tabPage2.Controls.Add(this.panel3);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(999, 379);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Search";
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.Teal;
            this.panel2.Controls.Add(this.combounit);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.txtsearch);
            this.panel2.Controls.Add(this.lblsearch);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Font = new System.Drawing.Font("Roboto Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panel2.Location = new System.Drawing.Point(3, 3);
            this.panel2.Margin = new System.Windows.Forms.Padding(2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(993, 31);
            this.panel2.TabIndex = 60;
            // 
            // combounit
            // 
            this.combounit.FormattingEnabled = true;
            this.combounit.Location = new System.Drawing.Point(533, 4);
            this.combounit.Name = "combounit";
            this.combounit.Size = new System.Drawing.Size(236, 26);
            this.combounit.TabIndex = 90;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Roboto Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(453, 6);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(70, 18);
            this.label5.TabIndex = 89;
            this.label5.Text = "CompCode";
            // 
            // txtsearch
            // 
            this.txtsearch.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtsearch.Location = new System.Drawing.Point(107, 3);
            this.txtsearch.Name = "txtsearch";
            this.txtsearch.Size = new System.Drawing.Size(328, 25);
            this.txtsearch.TabIndex = 83;
            // 
            // lblsearch
            // 
            this.lblsearch.AutoSize = true;
            this.lblsearch.Font = new System.Drawing.Font("Roboto Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblsearch.ForeColor = System.Drawing.Color.White;
            this.lblsearch.Location = new System.Drawing.Point(28, 6);
            this.lblsearch.Name = "lblsearch";
            this.lblsearch.Size = new System.Drawing.Size(49, 18);
            this.lblsearch.TabIndex = 82;
            this.lblsearch.Text = "Search";
            // 
            // listView1
            // 
            this.listView1.Activation = System.Windows.Forms.ItemActivation.OneClick;
            this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listView1.BackColor = System.Drawing.Color.White;
            this.listView1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader4,
            this.columnHeader3,
            this.columnHeader5,
            this.columnHeader9,
            this.columnHeader6,
            this.columnHeader7,
            this.columnHeader8});
            this.listView1.Font = new System.Drawing.Font("Roboto Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listView1.ForeColor = System.Drawing.Color.Black;
            this.listView1.FullRowSelect = true;
            this.listView1.GridLines = true;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(3, 31);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(992, 320);
            this.listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listView1.TabIndex = 59;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Width = 0;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Sno";
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "ID";
            this.columnHeader4.Width = 0;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Finyear";
            this.columnHeader3.Width = 0;
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "DocID";
            this.columnHeader5.Width = 150;
            // 
            // columnHeader9
            // 
            this.columnHeader9.Text = "CompCode";
            this.columnHeader9.Width = 150;
            // 
            // columnHeader6
            // 
            this.columnHeader6.Text = "Date";
            this.columnHeader6.Width = 100;
            // 
            // columnHeader7
            // 
            this.columnHeader7.Text = "Active";
            // 
            // columnHeader8
            // 
            this.columnHeader8.Text = "Month";
            this.columnHeader8.Width = 0;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.Teal;
            this.panel3.Controls.Add(this.lbltotal);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(3, 349);
            this.panel3.Margin = new System.Windows.Forms.Padding(0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(993, 27);
            this.panel3.TabIndex = 58;
            // 
            // lbltotal
            // 
            this.lbltotal.AutoSize = true;
            this.lbltotal.BackColor = System.Drawing.Color.Transparent;
            this.lbltotal.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbltotal.ForeColor = System.Drawing.Color.White;
            this.lbltotal.Location = new System.Drawing.Point(0, 8);
            this.lbltotal.Name = "lbltotal";
            this.lbltotal.Size = new System.Drawing.Size(81, 16);
            this.lbltotal.TabIndex = 96;
            this.lbltotal.Text = "Total Count :";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.butheader);
            this.panel1.Controls.Add(this.tabControl1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(3);
            this.panel1.Size = new System.Drawing.Size(1019, 455);
            this.panel1.TabIndex = 70;
            // 
            // butheader
            // 
            this.butheader.BackColor = System.Drawing.Color.Teal;
            this.butheader.Dock = System.Windows.Forms.DockStyle.Top;
            this.butheader.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.butheader.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.butheader.Font = new System.Drawing.Font("Roboto Black", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butheader.ForeColor = System.Drawing.Color.White;
            this.butheader.Location = new System.Drawing.Point(3, 3);
            this.butheader.Margin = new System.Windows.Forms.Padding(0);
            this.butheader.Name = "butheader";
            this.butheader.Size = new System.Drawing.Size(1013, 30);
            this.butheader.TabIndex = 451;
            this.butheader.Text = "PORT TRANSACTIONS";
            this.butheader.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.butheader.UseVisualStyleBackColor = false;
            // 
            // PortTransaction
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.Teal;
            this.ClientSize = new System.Drawing.Size(1021, 459);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Roboto", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "PortTransaction";
            this.Padding = new System.Windows.Forms.Padding(1, 1, 1, 3);
            this.Text = "Port";
            this.Load += new System.EventHandler(this.PortTransaction_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.IO.Ports.SerialPort serialPort1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox Field0;
        private System.Windows.Forms.Label lblcon;
        private System.Windows.Forms.Button butgetdata;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ComboBox comboport;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.ComboBox combostopbits;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.ComboBox comboparity;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.ComboBox combodatabits;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.ComboBox combobaudrate;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button butheader;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.ComboBox combounit;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtsearch;
        private System.Windows.Forms.Label lblsearch;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader5;
        private System.Windows.Forms.ColumnHeader columnHeader9;
        private System.Windows.Forms.ColumnHeader columnHeader6;
        private System.Windows.Forms.ColumnHeader columnHeader7;
        private System.Windows.Forms.ColumnHeader columnHeader8;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label lbltotal;
    }
}