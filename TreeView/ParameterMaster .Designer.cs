namespace Pinnacle.TreeView
{
    partial class ParameterMaster
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.ASPTBLPARAMID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Compcode = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.USERID = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.APPNAME = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DSOURCE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CSPWORD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.EPWORD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ALIASNAME = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lbl_Header = new System.Windows.Forms.Label();
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
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ASPTBLPARAMID,
            this.Compcode,
            this.USERID,
            this.APPNAME,
            this.DSOURCE,
            this.CSPWORD,
            this.EPWORD,
            this.ALIASNAME});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dataGridView1.Location = new System.Drawing.Point(0, 87);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1040, 400);
            this.dataGridView1.TabIndex = 0;
            // 
            // ASPTBLPARAMID
            // 
            this.ASPTBLPARAMID.HeaderText = "Param ID";
            this.ASPTBLPARAMID.Name = "ASPTBLPARAMID";
            // 
            // Compcode
            // 
            this.Compcode.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.Compcode.HeaderText = "Compcode";
            this.Compcode.Name = "Compcode";
            this.Compcode.Width = 200;
            // 
            // USERID
            // 
            this.USERID.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.USERID.HeaderText = "UserName";
            this.USERID.Name = "USERID";
            this.USERID.Width = 200;
            // 
            // APPNAME
            // 
            this.APPNAME.HeaderText = "ApplicationName";
            this.APPNAME.Name = "APPNAME";
            this.APPNAME.Width = 200;
            // 
            // DSOURCE
            // 
            this.DSOURCE.HeaderText = "DataSource";
            this.DSOURCE.Name = "DSOURCE";
            // 
            // CSPWORD
            // 
            this.CSPWORD.HeaderText = "CSPassword";
            this.CSPWORD.Name = "CSPWORD";
            // 
            // EPWORD
            // 
            this.EPWORD.HeaderText = "Emp Password";
            this.EPWORD.Name = "EPWORD";
            // 
            // ALIASNAME
            // 
            this.ALIASNAME.HeaderText = "AliasName";
            this.ALIASNAME.Name = "ALIASNAME";
            // 
            // lbl_Header
            // 
            this.lbl_Header.AutoSize = true;
            this.lbl_Header.BackColor = System.Drawing.Color.SteelBlue;
            this.lbl_Header.Font = new System.Drawing.Font("Arial Unicode MS", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Header.ForeColor = System.Drawing.Color.White;
            this.lbl_Header.Location = new System.Drawing.Point(9, 0);
            this.lbl_Header.Margin = new System.Windows.Forms.Padding(0);
            this.lbl_Header.Name = "lbl_Header";
            this.lbl_Header.Size = new System.Drawing.Size(19, 28);
            this.lbl_Header.TabIndex = 28;
            this.lbl_Header.Text = ".";
            // 
            // toolStrip1
            // 
            this.toolStrip1.BackColor = System.Drawing.Color.SteelBlue;
            this.toolStrip1.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
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
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.toolStrip1.Size = new System.Drawing.Size(1040, 31);
            this.toolStrip1.TabIndex = 27;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // Exit
            // 
            this.Exit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Exit.ForeColor = System.Drawing.Color.White;
            this.Exit.Margin = new System.Windows.Forms.Padding(0);
            this.Exit.Name = "Exit";
            this.Exit.Padding = new System.Windows.Forms.Padding(5);
            this.Exit.Size = new System.Drawing.Size(45, 31);
            this.Exit.Text = "Exit";
            this.Exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // TreeButtons
            // 
            this.TreeButtons.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.TreeButtons.ForeColor = System.Drawing.Color.White;
            this.TreeButtons.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.TreeButtons.Name = "TreeButtons";
            this.TreeButtons.Size = new System.Drawing.Size(80, 28);
            this.TreeButtons.Text = "TreeButton";
            // 
            // GlobalSearchs
            // 
            this.GlobalSearchs.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.GlobalSearchs.ForeColor = System.Drawing.Color.White;
            this.GlobalSearchs.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.GlobalSearchs.Name = "GlobalSearchs";
            this.GlobalSearchs.Size = new System.Drawing.Size(92, 28);
            this.GlobalSearchs.Text = "GlobalSearch";
            // 
            // Logins
            // 
            this.Logins.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Logins.ForeColor = System.Drawing.Color.White;
            this.Logins.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Logins.Name = "Logins";
            this.Logins.Size = new System.Drawing.Size(47, 28);
            this.Logins.Text = "Login";
            // 
            // ChangePasswords
            // 
            this.ChangePasswords.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.ChangePasswords.ForeColor = System.Drawing.Color.White;
            this.ChangePasswords.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ChangePasswords.Margin = new System.Windows.Forms.Padding(0);
            this.ChangePasswords.Name = "ChangePasswords";
            this.ChangePasswords.Padding = new System.Windows.Forms.Padding(5);
            this.ChangePasswords.Size = new System.Drawing.Size(126, 31);
            this.ChangePasswords.Text = "ChangePassword";
            // 
            // ChangeSkins
            // 
            this.ChangeSkins.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.ChangeSkins.ForeColor = System.Drawing.Color.White;
            this.ChangeSkins.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ChangeSkins.Margin = new System.Windows.Forms.Padding(0);
            this.ChangeSkins.Name = "ChangeSkins";
            this.ChangeSkins.Padding = new System.Windows.Forms.Padding(5);
            this.ChangeSkins.Size = new System.Drawing.Size(100, 31);
            this.ChangeSkins.Text = "ChangeSkins";
            // 
            // DownLoads
            // 
            this.DownLoads.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.DownLoads.ForeColor = System.Drawing.Color.White;
            this.DownLoads.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.DownLoads.Margin = new System.Windows.Forms.Padding(0);
            this.DownLoads.Name = "DownLoads";
            this.DownLoads.Padding = new System.Windows.Forms.Padding(5);
            this.DownLoads.Size = new System.Drawing.Size(88, 31);
            this.DownLoads.Text = "DownLoad";
            // 
            // Pdfs
            // 
            this.Pdfs.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Pdfs.ForeColor = System.Drawing.Color.White;
            this.Pdfs.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Pdfs.Name = "Pdfs";
            this.Pdfs.Size = new System.Drawing.Size(33, 28);
            this.Pdfs.Text = "Pdf";
            // 
            // Imports
            // 
            this.Imports.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Imports.ForeColor = System.Drawing.Color.White;
            this.Imports.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Imports.Margin = new System.Windows.Forms.Padding(0);
            this.Imports.Name = "Imports";
            this.Imports.Padding = new System.Windows.Forms.Padding(5);
            this.Imports.Size = new System.Drawing.Size(64, 31);
            this.Imports.Text = "Import";
            // 
            // Deletes
            // 
            this.Deletes.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Deletes.ForeColor = System.Drawing.Color.White;
            this.Deletes.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Deletes.Margin = new System.Windows.Forms.Padding(0);
            this.Deletes.Name = "Deletes";
            this.Deletes.Padding = new System.Windows.Forms.Padding(5);
            this.Deletes.Size = new System.Drawing.Size(62, 31);
            this.Deletes.Text = "Delete";
            // 
            // Searchs
            // 
            this.Searchs.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Searchs.ForeColor = System.Drawing.Color.White;
            this.Searchs.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Searchs.Margin = new System.Windows.Forms.Padding(0);
            this.Searchs.Name = "Searchs";
            this.Searchs.Padding = new System.Windows.Forms.Padding(5);
            this.Searchs.Size = new System.Drawing.Size(62, 31);
            this.Searchs.Text = "Search";
            // 
            // Prints
            // 
            this.Prints.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Prints.ForeColor = System.Drawing.Color.White;
            this.Prints.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Prints.Margin = new System.Windows.Forms.Padding(0);
            this.Prints.Name = "Prints";
            this.Prints.Padding = new System.Windows.Forms.Padding(5);
            this.Prints.Size = new System.Drawing.Size(52, 31);
            this.Prints.Text = "Print";
            // 
            // Saves
            // 
            this.Saves.BackColor = System.Drawing.Color.SteelBlue;
            this.Saves.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Saves.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Saves.ForeColor = System.Drawing.Color.White;
            this.Saves.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Saves.Margin = new System.Windows.Forms.Padding(0);
            this.Saves.Name = "Saves";
            this.Saves.Padding = new System.Windows.Forms.Padding(5);
            this.Saves.Size = new System.Drawing.Size(48, 31);
            this.Saves.Text = "Save";
            // 
            // News
            // 
            this.News.BackColor = System.Drawing.Color.SteelBlue;
            this.News.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.News.ForeColor = System.Drawing.Color.White;
            this.News.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.News.Margin = new System.Windows.Forms.Padding(0);
            this.News.Name = "News";
            this.News.Padding = new System.Windows.Forms.Padding(5);
            this.News.Size = new System.Drawing.Size(49, 31);
            this.News.Text = "New";
            // 
            // ParameterMaster
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1040, 487);
            this.Controls.Add(this.lbl_Header);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.dataGridView1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "ParameterMaster";
            this.Text = "ParameterMaster";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ParameterMaster_FormClosed);
            this.Load += new System.EventHandler(this.ParameterMaster_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn ASPTBLPARAMID;
        private System.Windows.Forms.DataGridViewComboBoxColumn Compcode;
        private System.Windows.Forms.DataGridViewComboBoxColumn USERID;
        private System.Windows.Forms.DataGridViewTextBoxColumn APPNAME;
        private System.Windows.Forms.DataGridViewTextBoxColumn DSOURCE;
        private System.Windows.Forms.DataGridViewTextBoxColumn CSPWORD;
        private System.Windows.Forms.DataGridViewTextBoxColumn EPWORD;
        private System.Windows.Forms.DataGridViewTextBoxColumn ALIASNAME;
        private System.Windows.Forms.Label lbl_Header;
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
    }
}