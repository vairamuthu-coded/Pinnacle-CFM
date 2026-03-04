namespace Pinnacle.Canteen
{
    partial class CanteenItemMaster
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
            this.panel3 = new System.Windows.Forms.Panel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.itemRefreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panel2 = new System.Windows.Forms.Panel();
            this.pictureemp = new System.Windows.Forms.PictureBox();
            this.lblIdcardno = new System.Windows.Forms.Label();
            this.comboitem = new System.Windows.Forms.ComboBox();
            this.lblempname = new System.Windows.Forms.Label();
            this.lblempid = new System.Windows.Forms.Label();
            this.lvLogs = new System.Windows.Forms.ListView();
            this.Sno = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.EmpID = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.EmpName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.IDCardNo = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.timercanteen = new System.Windows.Forms.Timer(this.components);
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.contextMenuStrip2.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureemp)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.lvLogs);
            this.panel1.Location = new System.Drawing.Point(3, 11);
            this.panel1.Margin = new System.Windows.Forms.Padding(0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1263, 498);
            this.panel1.TabIndex = 0;
            // 
            // panel3
            // 
            this.panel3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel3.AutoScroll = true;
            this.panel3.BackColor = System.Drawing.Color.WhiteSmoke;
            this.panel3.Controls.Add(this.flowLayoutPanel1);
            this.panel3.Location = new System.Drawing.Point(3, 56);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1257, 442);
            this.panel3.TabIndex = 3;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flowLayoutPanel1.AutoScroll = true;
            this.flowLayoutPanel1.AutoSize = true;
            this.flowLayoutPanel1.BackColor = System.Drawing.Color.White;
            this.flowLayoutPanel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.flowLayoutPanel1.ContextMenuStrip = this.contextMenuStrip2;
            this.flowLayoutPanel1.ForeColor = System.Drawing.Color.Black;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 18);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(1254, 424);
            this.flowLayoutPanel1.TabIndex = 1;
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.itemRefreshToolStripMenuItem});
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(141, 26);
            // 
            // itemRefreshToolStripMenuItem
            // 
            this.itemRefreshToolStripMenuItem.Name = "itemRefreshToolStripMenuItem";
            this.itemRefreshToolStripMenuItem.Size = new System.Drawing.Size(140, 22);
            this.itemRefreshToolStripMenuItem.Text = "Item Refresh";
            this.itemRefreshToolStripMenuItem.Click += new System.EventHandler(this.ItemRefreshToolStripMenuItem_Click);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.WhiteSmoke;
            this.panel2.Controls.Add(this.pictureemp);
            this.panel2.Controls.Add(this.lblIdcardno);
            this.panel2.Controls.Add(this.comboitem);
            this.panel2.Controls.Add(this.lblempname);
            this.panel2.Controls.Add(this.lblempid);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1263, 46);
            this.panel2.TabIndex = 2;
            // 
            // pictureemp
            // 
            this.pictureemp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureemp.BackColor = System.Drawing.Color.Transparent;
            this.pictureemp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureemp.Location = new System.Drawing.Point(1174, 1);
            this.pictureemp.Name = "pictureemp";
            this.pictureemp.Size = new System.Drawing.Size(77, 44);
            this.pictureemp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureemp.TabIndex = 3;
            this.pictureemp.TabStop = false;
            // 
            // lblIdcardno
            // 
            this.lblIdcardno.AutoSize = true;
            this.lblIdcardno.BackColor = System.Drawing.Color.Transparent;
            this.lblIdcardno.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lblIdcardno.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblIdcardno.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.lblIdcardno.Location = new System.Drawing.Point(180, 1);
            this.lblIdcardno.Name = "lblIdcardno";
            this.lblIdcardno.Size = new System.Drawing.Size(63, 18);
            this.lblIdcardno.TabIndex = 5;
            this.lblIdcardno.Text = "IDCARD";
            // 
            // comboitem
            // 
            this.comboitem.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.comboitem.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboitem.BackColor = System.Drawing.SystemColors.Info;
            this.comboitem.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboitem.FormattingEnabled = true;
            this.comboitem.Location = new System.Drawing.Point(572, 3);
            this.comboitem.Name = "comboitem";
            this.comboitem.Size = new System.Drawing.Size(242, 24);
            this.comboitem.TabIndex = 6;
            this.comboitem.SelectedIndexChanged += new System.EventHandler(this.comboitem_SelectedIndexChanged);
            // 
            // lblempname
            // 
            this.lblempname.AutoSize = true;
            this.lblempname.BackColor = System.Drawing.Color.Transparent;
            this.lblempname.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lblempname.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblempname.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.lblempname.Location = new System.Drawing.Point(10, 27);
            this.lblempname.Margin = new System.Windows.Forms.Padding(0);
            this.lblempname.Name = "lblempname";
            this.lblempname.Size = new System.Drawing.Size(45, 18);
            this.lblempname.TabIndex = 4;
            this.lblempname.Text = "Name";
            // 
            // lblempid
            // 
            this.lblempid.AutoSize = true;
            this.lblempid.BackColor = System.Drawing.Color.Transparent;
            this.lblempid.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lblempid.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblempid.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.lblempid.Location = new System.Drawing.Point(10, 1);
            this.lblempid.Margin = new System.Windows.Forms.Padding(0);
            this.lblempid.Name = "lblempid";
            this.lblempid.Size = new System.Drawing.Size(21, 18);
            this.lblempid.TabIndex = 3;
            this.lblempid.Text = "Id";
            // 
            // lvLogs
            // 
            this.lvLogs.Activation = System.Windows.Forms.ItemActivation.OneClick;
            this.lvLogs.BackColor = System.Drawing.Color.Teal;
            this.lvLogs.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lvLogs.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Sno,
            this.EmpID,
            this.EmpName,
            this.IDCardNo});
            this.lvLogs.Font = new System.Drawing.Font("Palatino Linotype", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lvLogs.ForeColor = System.Drawing.Color.Maroon;
            this.lvLogs.FullRowSelect = true;
            this.lvLogs.GridLines = true;
            this.lvLogs.HideSelection = false;
            this.lvLogs.Location = new System.Drawing.Point(1216, 398);
            this.lvLogs.Name = "lvLogs";
            this.lvLogs.Size = new System.Drawing.Size(43, 61);
            this.lvLogs.TabIndex = 1;
            this.lvLogs.UseCompatibleStateImageBehavior = false;
            this.lvLogs.View = System.Windows.Forms.View.Details;
            // 
            // Sno
            // 
            this.Sno.Text = "Sno";
            this.Sno.Width = 0;
            // 
            // EmpID
            // 
            this.EmpID.Text = "EmpID";
            this.EmpID.Width = 0;
            // 
            // EmpName
            // 
            this.EmpName.Text = "EmpName";
            this.EmpName.Width = 150;
            // 
            // IDCardNo
            // 
            this.IDCardNo.Text = "IDCardNo";
            this.IDCardNo.Width = 100;
            // 
            // timercanteen
            // 
            this.timercanteen.Tick += new System.EventHandler(this.Timercanteen_Tick);
            // 
            // CanteenItemMaster
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(1269, 515);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CanteenItemMaster";
            this.Padding = new System.Windows.Forms.Padding(3, 0, 3, 5);
            this.ShowIcon = false;
            this.Text = "CanteenItemMaster";
            this.TransparencyKey = System.Drawing.Color.Red;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.CanteenItemMaster_FormClosed);
            this.Load += new System.EventHandler(this.CanteenItemMaster_Load);
            this.panel1.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.contextMenuStrip2.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureemp)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ListView lvLogs;
        private System.Windows.Forms.ColumnHeader EmpID;
        private System.Windows.Forms.ColumnHeader EmpName;
        private System.Windows.Forms.ColumnHeader IDCardNo;
        private System.Windows.Forms.ColumnHeader Sno;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label lblIdcardno;
        private System.Windows.Forms.Label lblempname;
        private System.Windows.Forms.Label lblempid;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
        private System.Windows.Forms.ToolStripMenuItem itemRefreshToolStripMenuItem;
        private System.Windows.Forms.Timer timercanteen;
        private System.Windows.Forms.PictureBox pictureemp;
        private System.Windows.Forms.ComboBox comboitem;
    }
}