
namespace Pinnacle.Master.CFM
{
    partial class PortMaster
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
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtportno = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtip = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtportid = new System.Windows.Forms.TextBox();
            this.checkactive = new System.Windows.Forms.CheckBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.btnsendclient = new System.Windows.Forms.Button();
            this.txtMessageclient = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtMessage = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lblcon = new System.Windows.Forms.Label();
            this.butgetdata = new System.Windows.Forms.Button();
            this.Field0 = new System.Windows.Forms.TextBox();
            this.comboport = new System.Windows.Forms.ComboBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
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
            this.serialPort3 = new System.IO.Ports.SerialPort(this.components);
            this.groupBox4.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox4
            // 
            this.groupBox4.BackColor = System.Drawing.Color.White;
            this.groupBox4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.txtportno);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Controls.Add(this.txtip);
            this.groupBox4.Controls.Add(this.label2);
            this.groupBox4.Controls.Add(this.txtportid);
            this.groupBox4.Controls.Add(this.checkactive);
            this.groupBox4.Controls.Add(this.progressBar1);
            this.groupBox4.Controls.Add(this.groupBox6);
            this.groupBox4.Controls.Add(this.comboport);
            this.groupBox4.Controls.Add(this.label16);
            this.groupBox4.Controls.Add(this.comboBox2);
            this.groupBox4.Controls.Add(this.label18);
            this.groupBox4.Controls.Add(this.combostopbits);
            this.groupBox4.Controls.Add(this.label24);
            this.groupBox4.Controls.Add(this.comboparity);
            this.groupBox4.Controls.Add(this.label25);
            this.groupBox4.Controls.Add(this.combodatabits);
            this.groupBox4.Controls.Add(this.label26);
            this.groupBox4.Controls.Add(this.combobaudrate);
            this.groupBox4.Controls.Add(this.label27);
            this.groupBox4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox4.ForeColor = System.Drawing.Color.Black;
            this.groupBox4.Location = new System.Drawing.Point(1, 1);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(1262, 483);
            this.groupBox4.TabIndex = 63;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Port Properties";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(29, 269);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(40, 13);
            this.label4.TabIndex = 30;
            this.label4.Text = "PortNo";
            // 
            // txtportno
            // 
            this.txtportno.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.txtportno.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtportno.Location = new System.Drawing.Point(138, 266);
            this.txtportno.Multiline = true;
            this.txtportno.Name = "txtportno";
            this.txtportno.Size = new System.Drawing.Size(232, 24);
            this.txtportno.TabIndex = 29;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(29, 239);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(43, 13);
            this.label3.TabIndex = 28;
            this.label3.Text = "TCP/IP";
            // 
            // txtip
            // 
            this.txtip.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.txtip.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtip.Location = new System.Drawing.Point(138, 236);
            this.txtip.Multiline = true;
            this.txtip.Name = "txtip";
            this.txtip.Size = new System.Drawing.Size(232, 24);
            this.txtip.TabIndex = 27;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(29, 19);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(18, 13);
            this.label2.TabIndex = 26;
            this.label2.Text = "ID";
            this.label2.Visible = false;
            // 
            // txtportid
            // 
            this.txtportid.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.txtportid.Enabled = false;
            this.txtportid.Font = new System.Drawing.Font("Palatino Linotype", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtportid.Location = new System.Drawing.Point(138, 16);
            this.txtportid.Multiline = true;
            this.txtportid.Name = "txtportid";
            this.txtportid.Size = new System.Drawing.Size(232, 24);
            this.txtportid.TabIndex = 24;
            this.txtportid.Visible = false;
            // 
            // checkactive
            // 
            this.checkactive.AutoSize = true;
            this.checkactive.Location = new System.Drawing.Point(138, 404);
            this.checkactive.Name = "checkactive";
            this.checkactive.Size = new System.Drawing.Size(56, 17);
            this.checkactive.TabIndex = 23;
            this.checkactive.Text = "Active";
            this.checkactive.UseVisualStyleBackColor = true;
            this.checkactive.Visible = false;
            this.checkactive.CheckedChanged += new System.EventHandler(this.checkactive_CheckedChanged);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(138, 301);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(232, 23);
            this.progressBar1.TabIndex = 21;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.checkBox1);
            this.groupBox6.Controls.Add(this.btnsendclient);
            this.groupBox6.Controls.Add(this.txtMessageclient);
            this.groupBox6.Controls.Add(this.label6);
            this.groupBox6.Controls.Add(this.txtMessage);
            this.groupBox6.Controls.Add(this.label5);
            this.groupBox6.Controls.Add(this.textBox1);
            this.groupBox6.Controls.Add(this.lblcon);
            this.groupBox6.Controls.Add(this.butgetdata);
            this.groupBox6.Controls.Add(this.Field0);
            this.groupBox6.ForeColor = System.Drawing.Color.Black;
            this.groupBox6.Location = new System.Drawing.Point(395, 14);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(606, 434);
            this.groupBox6.TabIndex = 19;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Transmitter Control";
            this.groupBox6.Enter += new System.EventHandler(this.groupBox6_Enter);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(480, 352);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(56, 17);
            this.checkBox1.TabIndex = 54;
            this.checkBox1.Text = "Active";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.Visible = false;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // btnsendclient
            // 
            this.btnsendclient.Location = new System.Drawing.Point(323, 343);
            this.btnsendclient.Name = "btnsendclient";
            this.btnsendclient.Size = new System.Drawing.Size(105, 33);
            this.btnsendclient.TabIndex = 53;
            this.btnsendclient.Text = "TCP-GetData";
            this.btnsendclient.UseVisualStyleBackColor = true;
            this.btnsendclient.Click += new System.EventHandler(this.btnsendclient_Click);
            // 
            // txtMessageclient
            // 
            this.txtMessageclient.Enabled = false;
            this.txtMessageclient.Location = new System.Drawing.Point(170, 305);
            this.txtMessageclient.Name = "txtMessageclient";
            this.txtMessageclient.Size = new System.Drawing.Size(258, 20);
            this.txtMessageclient.TabIndex = 50;
            this.txtMessageclient.Text = "0";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(34, 308);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(79, 13);
            this.label6.TabIndex = 49;
            this.label6.Text = "Client-Message";
            // 
            // txtMessage
            // 
            this.txtMessage.Enabled = false;
            this.txtMessage.Location = new System.Drawing.Point(170, 277);
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.Size = new System.Drawing.Size(258, 20);
            this.txtMessage.TabIndex = 42;
            this.txtMessage.Text = "0";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(34, 283);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(84, 13);
            this.label5.TabIndex = 41;
            this.label5.Text = "Server-Message";
            // 
            // textBox1
            // 
            this.textBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.textBox1.Font = new System.Drawing.Font("Palatino Linotype", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(75, 89);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(404, 182);
            this.textBox1.TabIndex = 13;
            // 
            // lblcon
            // 
            this.lblcon.AutoSize = true;
            this.lblcon.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblcon.Location = new System.Drawing.Point(244, 400);
            this.lblcon.Name = "lblcon";
            this.lblcon.Size = new System.Drawing.Size(117, 16);
            this.lblcon.TabIndex = 22;
            this.lblcon.Text = "Connection State";
            // 
            // butgetdata
            // 
            this.butgetdata.ForeColor = System.Drawing.Color.Black;
            this.butgetdata.Location = new System.Drawing.Point(170, 343);
            this.butgetdata.Name = "butgetdata";
            this.butgetdata.Size = new System.Drawing.Size(103, 33);
            this.butgetdata.TabIndex = 11;
            this.butgetdata.Text = "COM GetData";
            this.butgetdata.UseVisualStyleBackColor = true;
            this.butgetdata.Click += new System.EventHandler(this.butgetdata_Click);
            // 
            // Field0
            // 
            this.Field0.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.Field0.Font = new System.Drawing.Font("Palatino Linotype", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Field0.Location = new System.Drawing.Point(75, 32);
            this.Field0.Multiline = true;
            this.Field0.Name = "Field0";
            this.Field0.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.Field0.Size = new System.Drawing.Size(404, 51);
            this.Field0.TabIndex = 12;
            // 
            // comboport
            // 
            this.comboport.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboport.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboport.ContextMenuStrip = this.contextMenuStrip1;
            this.comboport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboport.FormattingEnabled = true;
            this.comboport.Items.AddRange(new object[] {
            ""});
            this.comboport.Location = new System.Drawing.Point(138, 46);
            this.comboport.Name = "comboport";
            this.comboport.Size = new System.Drawing.Size(232, 21);
            this.comboport.TabIndex = 1;
            this.comboport.SelectedIndexChanged += new System.EventHandler(this.comboport_SelectedIndexChanged);
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
            this.refreshToolStripMenuItem.Click += new System.EventHandler(this.refreshToolStripMenuItem_Click);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(29, 54);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(47, 13);
            this.label16.TabIndex = 17;
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
            this.comboBox2.Location = new System.Drawing.Point(138, 206);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(232, 21);
            this.comboBox2.TabIndex = 6;
            this.comboBox2.Text = "None";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(29, 214);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(65, 13);
            this.label18.TabIndex = 10;
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
            this.combostopbits.Location = new System.Drawing.Point(138, 174);
            this.combostopbits.Name = "combostopbits";
            this.combostopbits.Size = new System.Drawing.Size(232, 21);
            this.combostopbits.TabIndex = 5;
            this.combostopbits.Text = "1";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(27, 175);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(49, 13);
            this.label24.TabIndex = 8;
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
            this.comboparity.Location = new System.Drawing.Point(138, 142);
            this.comboparity.Name = "comboparity";
            this.comboparity.Size = new System.Drawing.Size(232, 21);
            this.comboparity.TabIndex = 4;
            this.comboparity.Text = "Even";
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(29, 142);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(33, 13);
            this.label25.TabIndex = 6;
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
            this.combodatabits.Location = new System.Drawing.Point(138, 110);
            this.combodatabits.Name = "combodatabits";
            this.combodatabits.Size = new System.Drawing.Size(232, 21);
            this.combodatabits.TabIndex = 3;
            this.combodatabits.Text = "7";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(29, 113);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(50, 13);
            this.label26.TabIndex = 4;
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
            this.combobaudrate.Location = new System.Drawing.Point(138, 78);
            this.combobaudrate.Name = "combobaudrate";
            this.combobaudrate.Size = new System.Drawing.Size(232, 21);
            this.combobaudrate.TabIndex = 2;
            this.combobaudrate.Text = "9600";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(29, 81);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(58, 13);
            this.label27.TabIndex = 0;
            this.label27.Text = "Baud Rate";
            // 
            // serialPort3
            // 
            this.serialPort3.DataBits = 7;
            // 
            // PortMaster
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.ClientSize = new System.Drawing.Size(1264, 485);
            this.Controls.Add(this.groupBox4);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "PortMaster";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.Text = "Port Master";
            this.Load += new System.EventHandler(this.PortMaster_Load);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Label lblcon;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button butgetdata;
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
        private System.IO.Ports.SerialPort serialPort3;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.TextBox Field0;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.CheckBox checkactive;
        private System.Windows.Forms.TextBox txtportid;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtportno;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtip;
        private System.Windows.Forms.TextBox txtMessage;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtMessageclient;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnsendclient;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
    }
}