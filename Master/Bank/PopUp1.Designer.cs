
namespace Pinnacle.Master.Bank
{
    partial class PopUp1
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
            this.combo_compcode = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_sumbit = new System.Windows.Forms.Button();
            this.mySqlDataAdapter1 = new MySqlConnector.MySqlDataAdapter();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // combo_compcode
            // 
            this.combo_compcode.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.combo_compcode.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combo_compcode.ContextMenuStrip = this.contextMenuStrip1;
            this.combo_compcode.DropDownHeight = 80;
            this.combo_compcode.Font = new System.Drawing.Font("Roboto", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.combo_compcode.ForeColor = System.Drawing.Color.Navy;
            this.combo_compcode.FormattingEnabled = true;
            this.combo_compcode.IntegralHeight = false;
            this.combo_compcode.Location = new System.Drawing.Point(114, 28);
            this.combo_compcode.Name = "combo_compcode";
            this.combo_compcode.Size = new System.Drawing.Size(196, 26);
            this.combo_compcode.TabIndex = 11;
            this.combo_compcode.SelectedIndexChanged += new System.EventHandler(this.combo_compcode_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Navy;
            this.label2.Location = new System.Drawing.Point(15, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 18);
            this.label2.TabIndex = 10;
            this.label2.Text = "CompCode";
            // 
            // btn_sumbit
            // 
            this.btn_sumbit.BackColor = System.Drawing.Color.Transparent;
            this.btn_sumbit.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.btn_sumbit.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_sumbit.ForeColor = System.Drawing.Color.Navy;
            this.btn_sumbit.Location = new System.Drawing.Point(147, 68);
            this.btn_sumbit.Name = "btn_sumbit";
            this.btn_sumbit.Size = new System.Drawing.Size(93, 29);
            this.btn_sumbit.TabIndex = 9;
            this.btn_sumbit.Text = "Sign In";
            this.btn_sumbit.UseVisualStyleBackColor = false;
            this.btn_sumbit.Click += new System.EventHandler(this.btn_sumbit_Click);
            // 
            // mySqlDataAdapter1
            // 
            this.mySqlDataAdapter1.DeleteCommand = null;
            this.mySqlDataAdapter1.InsertCommand = null;
            this.mySqlDataAdapter1.SelectCommand = null;
            this.mySqlDataAdapter1.UpdateBatchSize = 0;
            this.mySqlDataAdapter1.UpdateCommand = null;
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
            this.refreshToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.refreshToolStripMenuItem.Text = "Refresh";
            this.refreshToolStripMenuItem.Click += new System.EventHandler(this.refreshToolStripMenuItem_Click);
            // 
            // PopUp1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(347, 103);
            this.Controls.Add(this.combo_compcode);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_sumbit);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PopUp1";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PopUp";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.PopUp1_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.PopUp1_FormClosed);
            this.Load += new System.EventHandler(this.PopUp1_Load);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox combo_compcode;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_sumbit;
        private MySqlConnector.MySqlDataAdapter mySqlDataAdapter1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
    }
}