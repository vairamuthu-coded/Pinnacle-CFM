namespace Pinnacle.Fuel
{
    partial class WeighBridge
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
            this.label1 = new System.Windows.Forms.Label();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.Exit = new System.Windows.Forms.ToolStripButton();
            this.TreeButtons = new System.Windows.Forms.ToolStripButton();
            this.GlobalSearchs = new System.Windows.Forms.ToolStripButton();
            this.Logins = new System.Windows.Forms.ToolStripButton();
            this.ChangePasswords = new System.Windows.Forms.ToolStripButton();
            this.ChangeSkins = new System.Windows.Forms.ToolStripButton();
            this.DownLoads = new System.Windows.Forms.ToolStripButton();
            this.Pdfs = new System.Windows.Forms.ToolStripButton();
            this.Imports = new System.Windows.Forms.ToolStripButton();
            this.Deletes = new System.Windows.Forms.ToolStripButton();
            this.Searchs = new System.Windows.Forms.ToolStripButton();
            this.Prints = new System.Windows.Forms.ToolStripButton();
            this.Saves = new System.Windows.Forms.ToolStripButton();
            this.News = new System.Windows.Forms.ToolStripButton();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.butclear = new System.Windows.Forms.Button();
            this.butgetdata = new System.Windows.Forms.Button();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.txtname = new System.Windows.Forms.TextBox();
            this.comboport = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.butstop = new System.Windows.Forms.Button();
            this.butstart = new System.Windows.Forms.Button();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.combostopbits = new System.Windows.Forms.ComboBox();
            this.label13 = new System.Windows.Forms.Label();
            this.comboparity = new System.Windows.Forms.ComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.combodatabits = new System.Windows.Forms.ComboBox();
            this.label15 = new System.Windows.Forms.Label();
            this.combobaudrate = new System.Windows.Forms.ComboBox();
            this.label16 = new System.Windows.Forms.Label();
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.toolStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.label1.Font = new System.Drawing.Font("Palatino Linotype", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(12, 3);
            this.label1.Margin = new System.Windows.Forms.Padding(0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(14, 21);
            this.label1.TabIndex = 42;
            this.label1.Text = ".";
            // 
            // toolStrip1
            // 
            this.toolStrip1.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.toolStrip1.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStrip1.GripMargin = new System.Windows.Forms.Padding(0);
            this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.Exit,
            this.TreeButtons,
            this.GlobalSearchs,
            this.Logins,
            this.ChangePasswords,
            this.ChangeSkins,
            this.DownLoads,
            this.Pdfs,
            this.Imports,
            this.Deletes,
            this.Searchs,
            this.Prints,
            this.Saves,
            this.News});
            this.toolStrip1.Location = new System.Drawing.Point(3, 3);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Padding = new System.Windows.Forms.Padding(0);
            this.toolStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.toolStrip1.Size = new System.Drawing.Size(1246, 25);
            this.toolStrip1.Stretch = true;
            this.toolStrip1.TabIndex = 41;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // Exit
            // 
            this.Exit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Exit.ForeColor = System.Drawing.Color.White;
            this.Exit.Margin = new System.Windows.Forms.Padding(0);
            this.Exit.Name = "Exit";
            this.Exit.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.Exit.Size = new System.Drawing.Size(45, 25);
            this.Exit.Text = "Exit";
            this.Exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // TreeButtons
            // 
            this.TreeButtons.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.TreeButtons.ForeColor = System.Drawing.Color.White;
            this.TreeButtons.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.TreeButtons.Name = "TreeButtons";
            this.TreeButtons.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.TreeButtons.Size = new System.Drawing.Size(90, 22);
            this.TreeButtons.Text = "TreeButton";
            // 
            // GlobalSearchs
            // 
            this.GlobalSearchs.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.GlobalSearchs.ForeColor = System.Drawing.Color.White;
            this.GlobalSearchs.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.GlobalSearchs.Name = "GlobalSearchs";
            this.GlobalSearchs.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.GlobalSearchs.Size = new System.Drawing.Size(102, 22);
            this.GlobalSearchs.Text = "GlobalSearch";
            // 
            // Logins
            // 
            this.Logins.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Logins.ForeColor = System.Drawing.Color.White;
            this.Logins.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Logins.Name = "Logins";
            this.Logins.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.Logins.Size = new System.Drawing.Size(57, 22);
            this.Logins.Text = "Login";
            // 
            // ChangePasswords
            // 
            this.ChangePasswords.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.ChangePasswords.ForeColor = System.Drawing.Color.White;
            this.ChangePasswords.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ChangePasswords.Margin = new System.Windows.Forms.Padding(0);
            this.ChangePasswords.Name = "ChangePasswords";
            this.ChangePasswords.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.ChangePasswords.Size = new System.Drawing.Size(126, 25);
            this.ChangePasswords.Text = "ChangePassword";
            // 
            // ChangeSkins
            // 
            this.ChangeSkins.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.ChangeSkins.ForeColor = System.Drawing.Color.White;
            this.ChangeSkins.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ChangeSkins.Margin = new System.Windows.Forms.Padding(0);
            this.ChangeSkins.Name = "ChangeSkins";
            this.ChangeSkins.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.ChangeSkins.Size = new System.Drawing.Size(100, 25);
            this.ChangeSkins.Text = "ChangeSkins";
            // 
            // DownLoads
            // 
            this.DownLoads.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.DownLoads.ForeColor = System.Drawing.Color.White;
            this.DownLoads.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.DownLoads.Margin = new System.Windows.Forms.Padding(0);
            this.DownLoads.Name = "DownLoads";
            this.DownLoads.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.DownLoads.Size = new System.Drawing.Size(88, 25);
            this.DownLoads.Text = "DownLoad";
            // 
            // Pdfs
            // 
            this.Pdfs.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Pdfs.ForeColor = System.Drawing.Color.White;
            this.Pdfs.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Pdfs.Name = "Pdfs";
            this.Pdfs.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.Pdfs.Size = new System.Drawing.Size(43, 22);
            this.Pdfs.Text = "Pdf";
            // 
            // Imports
            // 
            this.Imports.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Imports.ForeColor = System.Drawing.Color.White;
            this.Imports.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Imports.Margin = new System.Windows.Forms.Padding(0);
            this.Imports.Name = "Imports";
            this.Imports.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.Imports.Size = new System.Drawing.Size(64, 25);
            this.Imports.Text = "Import";
            // 
            // Deletes
            // 
            this.Deletes.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Deletes.ForeColor = System.Drawing.Color.White;
            this.Deletes.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Deletes.Margin = new System.Windows.Forms.Padding(0);
            this.Deletes.Name = "Deletes";
            this.Deletes.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.Deletes.Size = new System.Drawing.Size(62, 25);
            this.Deletes.Text = "Delete";
            // 
            // Searchs
            // 
            this.Searchs.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Searchs.ForeColor = System.Drawing.Color.White;
            this.Searchs.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Searchs.Margin = new System.Windows.Forms.Padding(0);
            this.Searchs.Name = "Searchs";
            this.Searchs.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.Searchs.Size = new System.Drawing.Size(62, 25);
            this.Searchs.Text = "Search";
            // 
            // Prints
            // 
            this.Prints.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Prints.ForeColor = System.Drawing.Color.White;
            this.Prints.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Prints.Margin = new System.Windows.Forms.Padding(0);
            this.Prints.Name = "Prints";
            this.Prints.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.Prints.Size = new System.Drawing.Size(52, 25);
            this.Prints.Text = "Print";
            // 
            // Saves
            // 
            this.Saves.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.Saves.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Saves.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Saves.ForeColor = System.Drawing.Color.White;
            this.Saves.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Saves.Margin = new System.Windows.Forms.Padding(0);
            this.Saves.Name = "Saves";
            this.Saves.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.Saves.Size = new System.Drawing.Size(48, 25);
            this.Saves.Text = "Save";
            // 
            // News
            // 
            this.News.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.News.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.News.ForeColor = System.Drawing.Color.White;
            this.News.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.News.Margin = new System.Windows.Forms.Padding(0);
            this.News.Name = "News";
            this.News.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.News.Size = new System.Drawing.Size(49, 25);
            this.News.Text = "New";
            this.News.Click += new System.EventHandler(this.News_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.WhiteSmoke;
            this.panel1.Controls.Add(this.tabControl1);
            this.panel1.Location = new System.Drawing.Point(3, 31);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1246, 477);
            this.panel1.TabIndex = 43;
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(3, 4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1240, 470);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.White;
            this.tabPage1.Controls.Add(this.groupBox4);
            this.tabPage1.Location = new System.Drawing.Point(4, 27);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1232, 439);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "WeighBridge";
            // 
            // groupBox4
            // 
            this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox4.BackColor = System.Drawing.Color.Silver;
            this.groupBox4.Controls.Add(this.progressBar1);
            this.groupBox4.Controls.Add(this.groupBox5);
            this.groupBox4.Controls.Add(this.groupBox6);
            this.groupBox4.Controls.Add(this.comboport);
            this.groupBox4.Controls.Add(this.label11);
            this.groupBox4.Controls.Add(this.butstop);
            this.groupBox4.Controls.Add(this.butstart);
            this.groupBox4.Controls.Add(this.comboBox2);
            this.groupBox4.Controls.Add(this.label12);
            this.groupBox4.Controls.Add(this.combostopbits);
            this.groupBox4.Controls.Add(this.label13);
            this.groupBox4.Controls.Add(this.comboparity);
            this.groupBox4.Controls.Add(this.label14);
            this.groupBox4.Controls.Add(this.combodatabits);
            this.groupBox4.Controls.Add(this.label15);
            this.groupBox4.Controls.Add(this.combobaudrate);
            this.groupBox4.Controls.Add(this.label16);
            this.groupBox4.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.groupBox4.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.ForeColor = System.Drawing.Color.Black;
            this.groupBox4.Location = new System.Drawing.Point(265, 60);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(760, 328);
            this.groupBox4.TabIndex = 1;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Port Properties";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(34, 298);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(283, 23);
            this.progressBar1.TabIndex = 21;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.butclear);
            this.groupBox5.Controls.Add(this.butgetdata);
            this.groupBox5.Location = new System.Drawing.Point(336, 184);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(422, 136);
            this.groupBox5.TabIndex = 20;
            this.groupBox5.TabStop = false;
            // 
            // butclear
            // 
            this.butclear.Location = new System.Drawing.Point(208, 24);
            this.butclear.Name = "butclear";
            this.butclear.Size = new System.Drawing.Size(117, 40);
            this.butclear.TabIndex = 10;
            this.butclear.Text = "Clear Data";
            this.butclear.UseVisualStyleBackColor = true;
            this.butclear.Click += new System.EventHandler(this.Butclear_Click);
            // 
            // butgetdata
            // 
            this.butgetdata.Location = new System.Drawing.Point(41, 24);
            this.butgetdata.Name = "butgetdata";
            this.butgetdata.Size = new System.Drawing.Size(117, 40);
            this.butgetdata.TabIndex = 9;
            this.butgetdata.Text = "Get Data";
            this.butgetdata.UseVisualStyleBackColor = true;
            this.butgetdata.Click += new System.EventHandler(this.Butgetdata_Click);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.label10);
            this.groupBox6.Controls.Add(this.label9);
            this.groupBox6.Controls.Add(this.label3);
            this.groupBox6.Controls.Add(this.textBox2);
            this.groupBox6.Controls.Add(this.textBox1);
            this.groupBox6.Controls.Add(this.txtname);
            this.groupBox6.ForeColor = System.Drawing.Color.Black;
            this.groupBox6.Location = new System.Drawing.Point(336, 24);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(422, 154);
            this.groupBox6.TabIndex = 19;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Transmitter Control";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(20, 112);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(62, 18);
            this.label10.TabIndex = 17;
            this.label10.Text = "Weight-3";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(20, 71);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(62, 18);
            this.label9.TabIndex = 16;
            this.label9.Text = "Weight-2";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 35);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 18);
            this.label3.TabIndex = 15;
            this.label3.Text = "Weight-1";
            // 
            // textBox2
            // 
            this.textBox2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.textBox2.Font = new System.Drawing.Font("Palatino Linotype", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.Location = new System.Drawing.Point(101, 106);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(301, 35);
            this.textBox2.TabIndex = 14;
            // 
            // textBox1
            // 
            this.textBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.textBox1.Font = new System.Drawing.Font("Palatino Linotype", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(101, 65);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(301, 35);
            this.textBox1.TabIndex = 13;
            // 
            // txtname
            // 
            this.txtname.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.txtname.Font = new System.Drawing.Font("Palatino Linotype", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtname.Location = new System.Drawing.Point(101, 25);
            this.txtname.Multiline = true;
            this.txtname.Name = "txtname";
            this.txtname.Size = new System.Drawing.Size(301, 34);
            this.txtname.TabIndex = 12;
            // 
            // comboport
            // 
            this.comboport.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboport.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboport.FormattingEnabled = true;
            this.comboport.Location = new System.Drawing.Point(141, 24);
            this.comboport.Name = "comboport";
            this.comboport.Size = new System.Drawing.Size(176, 26);
            this.comboport.TabIndex = 1;
            this.comboport.Text = "COM1";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(30, 33);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(62, 18);
            this.label11.TabIndex = 17;
            this.label11.Text = "ComPort";
            // 
            // butstop
            // 
            this.butstop.Location = new System.Drawing.Point(200, 251);
            this.butstop.Name = "butstop";
            this.butstop.Size = new System.Drawing.Size(117, 40);
            this.butstop.TabIndex = 10;
            this.butstop.Text = "Stop";
            this.butstop.UseVisualStyleBackColor = true;
            this.butstop.Click += new System.EventHandler(this.Butstop_Click);
            // 
            // butstart
            // 
            this.butstart.Location = new System.Drawing.Point(34, 251);
            this.butstart.Name = "butstart";
            this.butstart.Size = new System.Drawing.Size(117, 40);
            this.butstart.TabIndex = 8;
            this.butstart.Text = "Start";
            this.butstart.UseVisualStyleBackColor = true;
            this.butstart.Click += new System.EventHandler(this.Butstart_Click);
            // 
            // comboBox2
            // 
            this.comboBox2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboBox2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboBox2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "Xon/X/off",
            "Harkware",
            "None"});
            this.comboBox2.Location = new System.Drawing.Point(141, 184);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(176, 26);
            this.comboBox2.TabIndex = 6;
            this.comboBox2.Text = "None";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(30, 184);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(88, 18);
            this.label12.TabIndex = 10;
            this.label12.Text = "Flow Control";
            // 
            // combostopbits
            // 
            this.combostopbits.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combostopbits.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combostopbits.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.combostopbits.FormattingEnabled = true;
            this.combostopbits.Items.AddRange(new object[] {
            "1",
            "1.5",
            "2"});
            this.combostopbits.Location = new System.Drawing.Point(141, 152);
            this.combostopbits.Name = "combostopbits";
            this.combostopbits.Size = new System.Drawing.Size(176, 26);
            this.combostopbits.TabIndex = 5;
            this.combostopbits.Text = "1";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(30, 153);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(62, 18);
            this.label13.TabIndex = 8;
            this.label13.Text = "Stop Bits";
            // 
            // comboparity
            // 
            this.comboparity.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboparity.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboparity.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboparity.FormattingEnabled = true;
            this.comboparity.Items.AddRange(new object[] {
            "Even",
            "Odd",
            "None",
            "Mark",
            "Space"});
            this.comboparity.Location = new System.Drawing.Point(141, 120);
            this.comboparity.Name = "comboparity";
            this.comboparity.Size = new System.Drawing.Size(176, 26);
            this.comboparity.TabIndex = 4;
            this.comboparity.Text = "Even";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(30, 120);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(43, 18);
            this.label14.TabIndex = 6;
            this.label14.Text = "Parity";
            // 
            // combodatabits
            // 
            this.combodatabits.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combodatabits.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combodatabits.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.combodatabits.FormattingEnabled = true;
            this.combodatabits.Items.AddRange(new object[] {
            "5",
            "6",
            "7",
            "8"});
            this.combodatabits.Location = new System.Drawing.Point(141, 88);
            this.combodatabits.Name = "combodatabits";
            this.combodatabits.Size = new System.Drawing.Size(176, 26);
            this.combodatabits.TabIndex = 3;
            this.combodatabits.Text = "7";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(31, 88);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(63, 18);
            this.label15.TabIndex = 4;
            this.label15.Text = "Data Bits";
            // 
            // combobaudrate
            // 
            this.combobaudrate.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combobaudrate.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combobaudrate.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
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
            this.combobaudrate.Location = new System.Drawing.Point(141, 56);
            this.combobaudrate.Name = "combobaudrate";
            this.combobaudrate.Size = new System.Drawing.Size(176, 26);
            this.combobaudrate.TabIndex = 2;
            this.combobaudrate.Text = "9600";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(30, 60);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(70, 18);
            this.label16.TabIndex = 0;
            this.label16.Text = "Baud Rate";
            // 
            // WeighBridge
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.ClientSize = new System.Drawing.Size(1252, 514);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.toolStrip1);
            this.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "WeighBridge";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.Text = "WayBridge";
            this.Load += new System.EventHandler(this.WeighBridge_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton Exit;
        private System.Windows.Forms.ToolStripButton TreeButtons;
        private System.Windows.Forms.ToolStripButton GlobalSearchs;
        private System.Windows.Forms.ToolStripButton Logins;
        private System.Windows.Forms.ToolStripButton ChangePasswords;
        private System.Windows.Forms.ToolStripButton ChangeSkins;
        private System.Windows.Forms.ToolStripButton DownLoads;
        private System.Windows.Forms.ToolStripButton Pdfs;
        private System.Windows.Forms.ToolStripButton Imports;
        private System.Windows.Forms.ToolStripButton Deletes;
        private System.Windows.Forms.ToolStripButton Searchs;
        private System.Windows.Forms.ToolStripButton Prints;
        private System.Windows.Forms.ToolStripButton Saves;
        private System.Windows.Forms.ToolStripButton News;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        public System.IO.Ports.SerialPort serialPort1;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button butclear;
        private System.Windows.Forms.Button butgetdata;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox txtname;
        private System.Windows.Forms.ComboBox comboport;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button butstop;
        private System.Windows.Forms.Button butstart;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.ComboBox combostopbits;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.ComboBox comboparity;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.ComboBox combodatabits;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.ComboBox combobaudrate;
        private System.Windows.Forms.Label label16;
    }
}