
namespace Pinnacle.Master.School.ChaitanyaSchool
{
    partial class VotingMaster
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panel1 = new System.Windows.Forms.Panel();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.comboselectelection = new System.Windows.Forms.ComboBox();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.lblvoteing = new System.Windows.Forms.Label();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.voting1 = new Pinnacle.UserControls.CTS.Voting();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.comboelectiondate1 = new System.Windows.Forms.ComboBox();
            this.lblpost = new System.Windows.Forms.Label();
            this.lbldate1 = new System.Windows.Forms.Label();
            this.comboelecationpost = new System.Windows.Forms.ComboBox();
            this.crystalReportViewer1 = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.crystalReportViewer2 = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.panel3 = new System.Windows.Forms.Panel();
            this.comboElectiondate = new System.Windows.Forms.ComboBox();
            this.lblelectionpost = new System.Windows.Forms.Label();
            this.lbldate = new System.Windows.Forms.Label();
            this.comborollno = new System.Windows.Forms.ComboBox();
            this.combostudent = new System.Windows.Forms.ComboBox();
            this.lblstudent = new System.Windows.Forms.Label();
            this.timer3 = new System.Windows.Forms.Timer(this.components);
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.contextMenuStrip2.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.Buttons;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1276, 539);
            this.tabControl1.TabIndex = 0;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.White;
            this.tabPage1.Controls.Add(this.panel1);
            this.tabPage1.Controls.Add(this.flowLayoutPanel1);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1268, 510);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Election  Cantidate";
            // 
            // panel1
            // 
            this.panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.panel1.Controls.Add(this.checkBox1);
            this.panel1.Controls.Add(this.comboselectelection);
            this.panel1.Controls.Add(this.lblvoteing);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1262, 27);
            this.panel1.TabIndex = 6;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(413, 5);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(45, 17);
            this.checkBox1.TabIndex = 7;
            this.checkBox1.Text = "ALL";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // comboselectelection
            // 
            this.comboselectelection.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboselectelection.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboselectelection.ContextMenuStrip = this.contextMenuStrip2;
            this.comboselectelection.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboselectelection.ForeColor = System.Drawing.Color.Red;
            this.comboselectelection.FormattingEnabled = true;
            this.comboselectelection.Location = new System.Drawing.Point(20, 1);
            this.comboselectelection.Name = "comboselectelection";
            this.comboselectelection.Size = new System.Drawing.Size(387, 24);
            this.comboselectelection.TabIndex = 6;
            this.comboselectelection.SelectedIndexChanged += new System.EventHandler(this.comboselectelection_SelectedIndexChanged);
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.contextMenuStrip2.Name = "contextMenuStrip1";
            this.contextMenuStrip2.Size = new System.Drawing.Size(114, 26);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(113, 22);
            this.toolStripMenuItem1.Text = "Refresh";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.toolStripMenuItem1_Click);
            // 
            // lblvoteing
            // 
            this.lblvoteing.AutoSize = true;
            this.lblvoteing.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblvoteing.ForeColor = System.Drawing.Color.Black;
            this.lblvoteing.Location = new System.Drawing.Point(650, 3);
            this.lblvoteing.Name = "lblvoteing";
            this.lblvoteing.Size = new System.Drawing.Size(130, 22);
            this.lblvoteing.TabIndex = 5;
            this.lblvoteing.Text = "Poling Count";
            this.lblvoteing.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flowLayoutPanel1.AutoScroll = true;
            this.flowLayoutPanel1.BackColor = System.Drawing.Color.White;
            this.flowLayoutPanel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.flowLayoutPanel1.ContextMenuStrip = this.contextMenuStrip1;
            this.flowLayoutPanel1.Controls.Add(this.voting1);
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 30);
            this.flowLayoutPanel1.Margin = new System.Windows.Forms.Padding(0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Padding = new System.Windows.Forms.Padding(10, 0, 0, 0);
            this.flowLayoutPanel1.Size = new System.Drawing.Size(1262, 450);
            this.flowLayoutPanel1.TabIndex = 4;
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
            // voting1
            // 
            this.voting1.BackColor = System.Drawing.Color.White;
            this.voting1.Location = new System.Drawing.Point(11, 10);
            this.voting1.Margin = new System.Windows.Forms.Padding(1, 10, 1, 1);
            this.voting1.Name = "voting1";
            this.voting1.Size = new System.Drawing.Size(521, 78);
            this.voting1.studentimage = null;
            this.voting1.studentLogoimage = null;
            this.voting1.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.White;
            this.tabPage2.Controls.Add(this.panel2);
            this.tabPage2.Controls.Add(this.crystalReportViewer1);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1268, 510);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Election Result";
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.panel2.Controls.Add(this.comboelectiondate1);
            this.panel2.Controls.Add(this.lblpost);
            this.panel2.Controls.Add(this.lbldate1);
            this.panel2.Controls.Add(this.comboelecationpost);
            this.panel2.Location = new System.Drawing.Point(3, 3);
            this.panel2.Margin = new System.Windows.Forms.Padding(0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1262, 35);
            this.panel2.TabIndex = 342;
            // 
            // comboelectiondate1
            // 
            this.comboelectiondate1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboelectiondate1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboelectiondate1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboelectiondate1.ForeColor = System.Drawing.Color.Navy;
            this.comboelectiondate1.FormattingEnabled = true;
            this.comboelectiondate1.Location = new System.Drawing.Point(753, 1);
            this.comboelectiondate1.Margin = new System.Windows.Forms.Padding(0);
            this.comboelectiondate1.Name = "comboelectiondate1";
            this.comboelectiondate1.Size = new System.Drawing.Size(134, 27);
            this.comboelectiondate1.TabIndex = 340;
            this.comboelectiondate1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // lblpost
            // 
            this.lblpost.AutoSize = true;
            this.lblpost.Font = new System.Drawing.Font("Roboto Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblpost.ForeColor = System.Drawing.Color.Black;
            this.lblpost.Location = new System.Drawing.Point(252, 9);
            this.lblpost.Name = "lblpost";
            this.lblpost.Size = new System.Drawing.Size(87, 18);
            this.lblpost.TabIndex = 337;
            this.lblpost.Text = "Election Post";
            // 
            // lbldate1
            // 
            this.lbldate1.AutoSize = true;
            this.lbldate1.Font = new System.Drawing.Font("Roboto Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbldate1.ForeColor = System.Drawing.Color.Black;
            this.lbldate1.Location = new System.Drawing.Point(671, 9);
            this.lbldate1.Name = "lbldate1";
            this.lbldate1.Size = new System.Drawing.Size(36, 18);
            this.lbldate1.TabIndex = 341;
            this.lbldate1.Text = "Date";
            // 
            // comboelecationpost
            // 
            this.comboelecationpost.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboelecationpost.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboelecationpost.ContextMenuStrip = this.contextMenuStrip2;
            this.comboelecationpost.DropDownHeight = 100;
            this.comboelecationpost.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboelecationpost.ForeColor = System.Drawing.Color.Navy;
            this.comboelecationpost.IntegralHeight = false;
            this.comboelecationpost.Location = new System.Drawing.Point(382, 4);
            this.comboelecationpost.Margin = new System.Windows.Forms.Padding(0);
            this.comboelecationpost.Name = "comboelecationpost";
            this.comboelecationpost.Size = new System.Drawing.Size(257, 27);
            this.comboelecationpost.TabIndex = 336;
            this.comboelecationpost.SelectedIndexChanged += new System.EventHandler(this.comboelecationpost_SelectedIndexChanged);
            // 
            // crystalReportViewer1
            // 
            this.crystalReportViewer1.ActiveViewIndex = -1;
            this.crystalReportViewer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.crystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.crystalReportViewer1.ContextMenuStrip = this.contextMenuStrip1;
            this.crystalReportViewer1.Cursor = System.Windows.Forms.Cursors.Default;
            this.crystalReportViewer1.Location = new System.Drawing.Point(6, 41);
            this.crystalReportViewer1.Name = "crystalReportViewer1";
            this.crystalReportViewer1.Size = new System.Drawing.Size(1259, 469);
            this.crystalReportViewer1.TabIndex = 0;
            this.crystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None;
            // 
            // tabPage3
            // 
            this.tabPage3.BackColor = System.Drawing.Color.White;
            this.tabPage3.Controls.Add(this.crystalReportViewer2);
            this.tabPage3.Controls.Add(this.panel3);
            this.tabPage3.Location = new System.Drawing.Point(4, 25);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(1268, 510);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Poling Details";
            // 
            // crystalReportViewer2
            // 
            this.crystalReportViewer2.ActiveViewIndex = -1;
            this.crystalReportViewer2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.crystalReportViewer2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.crystalReportViewer2.ContextMenuStrip = this.contextMenuStrip1;
            this.crystalReportViewer2.Cursor = System.Windows.Forms.Cursors.Default;
            this.crystalReportViewer2.Location = new System.Drawing.Point(5, 44);
            this.crystalReportViewer2.Name = "crystalReportViewer2";
            this.crystalReportViewer2.Size = new System.Drawing.Size(1259, 458);
            this.crystalReportViewer2.TabIndex = 341;
            this.crystalReportViewer2.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.Transparent;
            this.panel3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel3.Controls.Add(this.comboElectiondate);
            this.panel3.Controls.Add(this.lblelectionpost);
            this.panel3.Controls.Add(this.lbldate);
            this.panel3.Controls.Add(this.comborollno);
            this.panel3.Controls.Add(this.combostudent);
            this.panel3.Controls.Add(this.lblstudent);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(3, 3);
            this.panel3.Margin = new System.Windows.Forms.Padding(0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1262, 38);
            this.panel3.TabIndex = 340;
            // 
            // comboElectiondate
            // 
            this.comboElectiondate.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboElectiondate.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboElectiondate.ForeColor = System.Drawing.SystemColors.Desktop;
            this.comboElectiondate.FormattingEnabled = true;
            this.comboElectiondate.Location = new System.Drawing.Point(865, 8);
            this.comboElectiondate.Margin = new System.Windows.Forms.Padding(0);
            this.comboElectiondate.Name = "comboElectiondate";
            this.comboElectiondate.Size = new System.Drawing.Size(210, 21);
            this.comboElectiondate.TabIndex = 338;
            this.comboElectiondate.SelectedIndexChanged += new System.EventHandler(this.comboElectiondate_SelectedIndexChanged);
            // 
            // lblelectionpost
            // 
            this.lblelectionpost.AutoSize = true;
            this.lblelectionpost.Font = new System.Drawing.Font("Roboto Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblelectionpost.ForeColor = System.Drawing.Color.Black;
            this.lblelectionpost.Location = new System.Drawing.Point(33, 13);
            this.lblelectionpost.Name = "lblelectionpost";
            this.lblelectionpost.Size = new System.Drawing.Size(63, 18);
            this.lblelectionpost.TabIndex = 335;
            this.lblelectionpost.Text = "Roll No  *";
            // 
            // lbldate
            // 
            this.lbldate.AutoSize = true;
            this.lbldate.Font = new System.Drawing.Font("Roboto Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbldate.ForeColor = System.Drawing.Color.Black;
            this.lbldate.Location = new System.Drawing.Point(780, 9);
            this.lbldate.Name = "lbldate";
            this.lbldate.Size = new System.Drawing.Size(36, 18);
            this.lbldate.TabIndex = 339;
            this.lbldate.Text = "Date";
            // 
            // comborollno
            // 
            this.comborollno.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comborollno.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comborollno.ContextMenuStrip = this.contextMenuStrip2;
            this.comborollno.DropDownHeight = 100;
            this.comborollno.ForeColor = System.Drawing.SystemColors.Desktop;
            this.comborollno.FormattingEnabled = true;
            this.comborollno.IntegralHeight = false;
            this.comborollno.Location = new System.Drawing.Point(121, 9);
            this.comborollno.Margin = new System.Windows.Forms.Padding(0);
            this.comborollno.Name = "comborollno";
            this.comborollno.Size = new System.Drawing.Size(151, 21);
            this.comborollno.TabIndex = 334;
            this.comborollno.SelectedIndexChanged += new System.EventHandler(this.comborollno_SelectedIndexChanged);
            // 
            // combostudent
            // 
            this.combostudent.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combostudent.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combostudent.DropDownHeight = 100;
            this.combostudent.ForeColor = System.Drawing.SystemColors.Desktop;
            this.combostudent.FormattingEnabled = true;
            this.combostudent.IntegralHeight = false;
            this.combostudent.Location = new System.Drawing.Point(433, 8);
            this.combostudent.Margin = new System.Windows.Forms.Padding(0);
            this.combostudent.Name = "combostudent";
            this.combostudent.Size = new System.Drawing.Size(305, 21);
            this.combostudent.TabIndex = 336;
            this.combostudent.SelectedIndexChanged += new System.EventHandler(this.combostudent_SelectedIndexChanged);
            // 
            // lblstudent
            // 
            this.lblstudent.AutoSize = true;
            this.lblstudent.Font = new System.Drawing.Font("Roboto Black", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblstudent.ForeColor = System.Drawing.Color.Black;
            this.lblstudent.Location = new System.Drawing.Point(295, 12);
            this.lblstudent.Name = "lblstudent";
            this.lblstudent.Size = new System.Drawing.Size(103, 18);
            this.lblstudent.TabIndex = 337;
            this.lblstudent.Text = "Student Name  *";
            // 
            // timer3
            // 
            this.timer3.Interval = 1000;
            this.timer3.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // VotingMaster
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1276, 539);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "VotingMaster";
            this.Text = "Create Voting";
            this.Load += new System.EventHandler(this.VotingMaster_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.contextMenuStrip2.ResumeLayout(false);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Label lblvoteing;
        private System.Windows.Forms.Timer timer3;
        private System.Windows.Forms.TabPage tabPage2;
        private CrystalDecisions.Windows.Forms.CrystalReportViewer crystalReportViewer1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.ComboBox comboElectiondate;
        private System.Windows.Forms.Label lbldate;
        private System.Windows.Forms.ComboBox combostudent;
        private System.Windows.Forms.Label lblstudent;
        private System.Windows.Forms.ComboBox comborollno;
        private System.Windows.Forms.Label lblelectionpost;
        private System.Windows.Forms.ComboBox comboelectiondate1;
        private System.Windows.Forms.Label lbldate1;
        private System.Windows.Forms.ComboBox comboelecationpost;
        private System.Windows.Forms.Label lblpost;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox comboselectelection;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private UserControls.CTS.Voting voting1;
        private CrystalDecisions.Windows.Forms.CrystalReportViewer crystalReportViewer2;
    }
}