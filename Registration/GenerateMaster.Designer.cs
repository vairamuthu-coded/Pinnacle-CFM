
namespace Pinnacle.Registration
{
    partial class GenerateMaster
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
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.butfooter = new System.Windows.Forms.Button();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.label7 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label6 = new System.Windows.Forms.Label();
            this.butGenerate = new System.Windows.Forms.Button();
            this.txtproductkeys = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtexperiencedays = new System.Windows.Forms.TextBox();
            this.combolicensetype = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.checkactive = new System.Windows.Forms.CheckBox();
            this.txtproductid = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtgenerateid = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.butheader);
            this.panel1.Location = new System.Drawing.Point(3, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1178, 431);
            this.panel1.TabIndex = 0;
            // 
            // butheader
            // 
            this.butheader.BackColor = System.Drawing.Color.Teal;
            this.butheader.ContextMenuStrip = this.contextMenuStrip1;
            this.butheader.Dock = System.Windows.Forms.DockStyle.Top;
            this.butheader.FlatAppearance.BorderSize = 0;
            this.butheader.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.butheader.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.butheader.Font = new System.Drawing.Font("Roboto Black", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butheader.ForeColor = System.Drawing.Color.White;
            this.butheader.Location = new System.Drawing.Point(0, 0);
            this.butheader.Name = "butheader";
            this.butheader.Size = new System.Drawing.Size(1178, 29);
            this.butheader.TabIndex = 64;
            this.butheader.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.butheader.UseVisualStyleBackColor = false;
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
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.Buttons;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(15, 35);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1155, 393);
            this.tabControl1.TabIndex = 60;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.White;
            this.tabPage1.Controls.Add(this.butfooter);
            this.tabPage1.Controls.Add(this.dateTimePicker2);
            this.tabPage1.Controls.Add(this.label7);
            this.tabPage1.Controls.Add(this.dateTimePicker1);
            this.tabPage1.Controls.Add(this.label6);
            this.tabPage1.Controls.Add(this.butGenerate);
            this.tabPage1.Controls.Add(this.txtproductkeys);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.txtexperiencedays);
            this.tabPage1.Controls.Add(this.combolicensetype);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.checkactive);
            this.tabPage1.Controls.Add(this.txtproductid);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.txtgenerateid);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3, 0, 3, 0);
            this.tabPage1.Size = new System.Drawing.Size(1147, 364);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Create Generate";
            // 
            // butfooter
            // 
            this.butfooter.BackColor = System.Drawing.Color.Teal;
            this.butfooter.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.butfooter.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.butfooter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.butfooter.Font = new System.Drawing.Font("Roboto Black", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butfooter.ForeColor = System.Drawing.Color.White;
            this.butfooter.Location = new System.Drawing.Point(3, 354);
            this.butfooter.Name = "butfooter";
            this.butfooter.Size = new System.Drawing.Size(1141, 10);
            this.butfooter.TabIndex = 65;
            this.butfooter.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.butfooter.UseVisualStyleBackColor = false;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker2.Location = new System.Drawing.Point(376, 51);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(146, 20);
            this.dateTimePicker2.TabIndex = 28;
            this.dateTimePicker2.ValueChanged += new System.EventHandler(this.dateTimePicker2_ValueChanged);
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(280, 51);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(90, 18);
            this.label7.TabIndex = 27;
            this.label7.Text = "To Date";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker1.Location = new System.Drawing.Point(127, 51);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(147, 20);
            this.dateTimePicker1.TabIndex = 26;
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(6, 57);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(104, 18);
            this.label6.TabIndex = 25;
            this.label6.Text = "From Date";
            // 
            // butGenerate
            // 
            this.butGenerate.Location = new System.Drawing.Point(164, 227);
            this.butGenerate.Name = "butGenerate";
            this.butGenerate.Size = new System.Drawing.Size(188, 38);
            this.butGenerate.TabIndex = 23;
            this.butGenerate.Text = "Generate";
            this.butGenerate.UseVisualStyleBackColor = true;
            this.butGenerate.Click += new System.EventHandler(this.butGenerate_Click);
            // 
            // txtproductkeys
            // 
            this.txtproductkeys.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtproductkeys.Location = new System.Drawing.Point(127, 156);
            this.txtproductkeys.MaxLength = 250;
            this.txtproductkeys.Name = "txtproductkeys";
            this.txtproductkeys.Size = new System.Drawing.Size(588, 20);
            this.txtproductkeys.TabIndex = 22;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(6, 156);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 18);
            this.label1.TabIndex = 21;
            this.label1.Text = "Product Keys";
            // 
            // txtexperiencedays
            // 
            this.txtexperiencedays.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtexperiencedays.Location = new System.Drawing.Point(127, 130);
            this.txtexperiencedays.MaxLength = 250;
            this.txtexperiencedays.Name = "txtexperiencedays";
            this.txtexperiencedays.Size = new System.Drawing.Size(588, 20);
            this.txtexperiencedays.TabIndex = 20;
            // 
            // combolicensetype
            // 
            this.combolicensetype.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combolicensetype.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combolicensetype.FormattingEnabled = true;
            this.combolicensetype.Items.AddRange(new object[] {
            "FULL",
            "TRIAL"});
            this.combolicensetype.Location = new System.Drawing.Point(127, 103);
            this.combolicensetype.Name = "combolicensetype";
            this.combolicensetype.Size = new System.Drawing.Size(588, 21);
            this.combolicensetype.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(6, 106);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(104, 18);
            this.label5.TabIndex = 19;
            this.label5.Text = "License Type";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(6, 132);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 18);
            this.label3.TabIndex = 17;
            this.label3.Text = "Experience Days";
            // 
            // checkactive
            // 
            this.checkactive.AutoSize = true;
            this.checkactive.Location = new System.Drawing.Point(127, 182);
            this.checkactive.Name = "checkactive";
            this.checkactive.Size = new System.Drawing.Size(56, 17);
            this.checkactive.TabIndex = 3;
            this.checkactive.Text = "Active";
            this.checkactive.UseVisualStyleBackColor = true;
            // 
            // txtproductid
            // 
            this.txtproductid.Location = new System.Drawing.Point(127, 77);
            this.txtproductid.MaxLength = 250;
            this.txtproductid.Name = "txtproductid";
            this.txtproductid.Size = new System.Drawing.Size(588, 20);
            this.txtproductid.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(6, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(104, 18);
            this.label2.TabIndex = 14;
            this.label2.Text = "Product ID";
            // 
            // txtgenerateid
            // 
            this.txtgenerateid.Enabled = false;
            this.txtgenerateid.Location = new System.Drawing.Point(127, 22);
            this.txtgenerateid.Name = "txtgenerateid";
            this.txtgenerateid.Size = new System.Drawing.Size(588, 20);
            this.txtgenerateid.TabIndex = 13;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(6, 25);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(104, 18);
            this.label4.TabIndex = 12;
            this.label4.Text = "ID";
            // 
            // GenerateMaster
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1176, 434);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "GenerateMaster";
            this.Padding = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.Text = "Create Generate";
            this.Load += new System.EventHandler(this.GenerateMaster_Load);
            this.panel1.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.ComboBox combolicensetype;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox checkactive;
        private System.Windows.Forms.TextBox txtproductid;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtgenerateid;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtproductkeys;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button butGenerate;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtexperiencedays;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button butheader;
        private System.Windows.Forms.Button butfooter;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
    }
}