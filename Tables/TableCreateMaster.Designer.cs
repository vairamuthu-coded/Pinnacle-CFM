namespace Pinnacle.Tables
{
    partial class butdroptable
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
            this.butcreate = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblprogress1 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // butcreate
            // 
            this.butcreate.Location = new System.Drawing.Point(45, 74);
            this.butcreate.Name = "butcreate";
            this.butcreate.Size = new System.Drawing.Size(237, 72);
            this.butcreate.TabIndex = 0;
            this.butcreate.Text = "Dynamic Table Create";
            this.butcreate.UseVisualStyleBackColor = true;
            this.butcreate.Click += new System.EventHandler(this.Butcreate_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 428);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(610, 33);
            this.progressBar1.TabIndex = 1;
            // 
            // lblprogress1
            // 
            this.lblprogress1.AutoSize = true;
            this.lblprogress1.Location = new System.Drawing.Point(287, 438);
            this.lblprogress1.Name = "lblprogress1";
            this.lblprogress1.Size = new System.Drawing.Size(35, 13);
            this.lblprogress1.TabIndex = 2;
            this.lblprogress1.Text = "label1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 238);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "label1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(351, 74);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(237, 72);
            this.button1.TabIndex = 4;
            this.button1.Text = "Drop table";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // butdroptable
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(634, 463);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblprogress1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.butcreate);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "butdroptable";
            this.Text = "TableCreateMaster";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button butcreate;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblprogress1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
    }
}