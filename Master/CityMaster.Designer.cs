namespace Pinnacle.Master
{
    partial class CityMaster
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
            this.lbl_Header = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.butheader = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.listView1 = new Pinnacle.UserControls.UCCListView();
            this.combostate = new System.Windows.Forms.ComboBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label5 = new System.Windows.Forms.Label();
            this.combocountry = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.checkactive = new System.Windows.Forms.CheckBox();
            this.txtcity = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtcityid = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.panel1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbl_Header
            // 
            this.lbl_Header.AutoSize = true;
            this.lbl_Header.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.lbl_Header.Font = new System.Drawing.Font("Palatino Linotype", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Header.ForeColor = System.Drawing.Color.White;
            this.lbl_Header.Location = new System.Drawing.Point(10, -15);
            this.lbl_Header.Margin = new System.Windows.Forms.Padding(0);
            this.lbl_Header.Name = "lbl_Header";
            this.lbl_Header.Size = new System.Drawing.Size(14, 21);
            this.lbl_Header.TabIndex = 58;
            this.lbl_Header.Text = ".";
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.butheader);
            this.panel1.Controls.Add(this.tabControl1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(3);
            this.panel1.Size = new System.Drawing.Size(1289, 503);
            this.panel1.TabIndex = 59;
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
            this.butheader.Location = new System.Drawing.Point(3, 3);
            this.butheader.Name = "butheader";
            this.butheader.Size = new System.Drawing.Size(1283, 30);
            this.butheader.TabIndex = 4;
            this.butheader.Text = "CITY MASTER";
            this.butheader.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.butheader.UseVisualStyleBackColor = false;
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(6, 39);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1277, 458);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.White;
            this.tabPage1.Controls.Add(this.listView1);
            this.tabPage1.Controls.Add(this.combostate);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.combocountry);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.checkactive);
            this.tabPage1.Controls.Add(this.txtcity);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.txtcityid);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3, 0, 3, 0);
            this.tabPage1.Size = new System.Drawing.Size(1269, 429);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "CityMaster";
            // 
            // listView1
            // 
            this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listView1.BackColor = System.Drawing.Color.Transparent;
            this.listView1.Font = new System.Drawing.Font("Roboto", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listView1.Location = new System.Drawing.Point(782, 4);
            this.listView1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.listView1.Name = "listView1";
            this.listView1.Padding = new System.Windows.Forms.Padding(5);
            this.listView1.Size = new System.Drawing.Size(481, 428);
            this.listView1.TabIndex = 20;
            // 
            // combostate
            // 
            this.combostate.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combostate.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combostate.ContextMenuStrip = this.contextMenuStrip1;
            this.combostate.FormattingEnabled = true;
            this.combostate.Items.AddRange(new object[] {
            "--"});
            this.combostate.Location = new System.Drawing.Point(111, 80);
            this.combostate.Name = "combostate";
            this.combostate.Size = new System.Drawing.Size(268, 21);
            this.combostate.TabIndex = 1;
            this.combostate.SelectedIndexChanged += new System.EventHandler(this.Combostate_SelectedIndexChanged);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.refreshToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(111, 26);
            // 
            // refreshToolStripMenuItem
            // 
            this.refreshToolStripMenuItem.Name = "refreshToolStripMenuItem";
            this.refreshToolStripMenuItem.Size = new System.Drawing.Size(110, 22);
            this.refreshToolStripMenuItem.Text = "refresh";
            this.refreshToolStripMenuItem.Click += new System.EventHandler(this.RefreshToolStripMenuItem_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(17, 83);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(63, 13);
            this.label5.TabIndex = 19;
            this.label5.Text = "State Name";
            // 
            // combocountry
            // 
            this.combocountry.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combocountry.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combocountry.FormattingEnabled = true;
            this.combocountry.Location = new System.Drawing.Point(111, 109);
            this.combocountry.Name = "combocountry";
            this.combocountry.Size = new System.Drawing.Size(268, 21);
            this.combocountry.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 109);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(43, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "Country";
            // 
            // checkactive
            // 
            this.checkactive.AutoSize = true;
            this.checkactive.Checked = true;
            this.checkactive.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkactive.Location = new System.Drawing.Point(111, 141);
            this.checkactive.Name = "checkactive";
            this.checkactive.Size = new System.Drawing.Size(56, 17);
            this.checkactive.TabIndex = 3;
            this.checkactive.Text = "Active";
            this.checkactive.UseVisualStyleBackColor = true;
            // 
            // txtcity
            // 
            this.txtcity.Location = new System.Drawing.Point(111, 53);
            this.txtcity.MaxLength = 20;
            this.txtcity.Name = "txtcity";
            this.txtcity.Size = new System.Drawing.Size(268, 20);
            this.txtcity.TabIndex = 0;
            this.txtcity.TextChanged += new System.EventHandler(this.txtcity_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 13);
            this.label2.TabIndex = 14;
            this.label2.Text = "CityName";
            // 
            // txtcityid
            // 
            this.txtcityid.Enabled = false;
            this.txtcityid.Location = new System.Drawing.Point(111, 26);
            this.txtcityid.Name = "txtcityid";
            this.txtcityid.Size = new System.Drawing.Size(268, 20);
            this.txtcityid.TabIndex = 13;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 29);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(38, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "City ID";
            // 
            // tabPage3
            // 
            this.tabPage3.Location = new System.Drawing.Point(4, 25);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(1269, 429);
            this.tabPage3.TabIndex = 1;
            this.tabPage3.Text = "tabPage3";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // CityMaster
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.Teal;
            this.ClientSize = new System.Drawing.Size(1291, 505);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.lbl_Header);
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "CityMaster";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.Text = "Create City";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.CityMaster_FormClosed);
            this.Load += new System.EventHandler(this.CityMaster_Load);
            this.panel1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_Header;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.ComboBox combostate;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox combocountry;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox checkactive;
        private System.Windows.Forms.TextBox txtcity;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtcityid;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
        private System.Windows.Forms.Button butheader;
        private System.Windows.Forms.TabPage tabPage3;
        private UserControls.UCCListView listView1;
    }
}