
namespace Pinnacle.UserControls
{
    partial class MonthControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblusercompcode = new System.Windows.Forms.Label();
            this.lbluserfinyear = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblusermonth = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblusercompcode
            // 
            this.lblusercompcode.AutoSize = true;
            this.lblusercompcode.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblusercompcode.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblusercompcode.Location = new System.Drawing.Point(146, 21);
            this.lblusercompcode.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lblusercompcode.Name = "lblusercompcode";
            this.lblusercompcode.Size = new System.Drawing.Size(69, 14);
            this.lblusercompcode.TabIndex = 0;
            this.lblusercompcode.Text = "CompCode";
            // 
            // lbluserfinyear
            // 
            this.lbluserfinyear.AutoSize = true;
            this.lbluserfinyear.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbluserfinyear.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lbluserfinyear.Location = new System.Drawing.Point(6, 21);
            this.lbluserfinyear.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lbluserfinyear.Name = "lbluserfinyear";
            this.lbluserfinyear.Size = new System.Drawing.Size(47, 14);
            this.lbluserfinyear.TabIndex = 2;
            this.lbluserfinyear.Text = "Finyear";
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.lblusermonth);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.lblusercompcode);
            this.panel1.Controls.Add(this.lbluserfinyear);
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(281, 99);
            this.panel1.TabIndex = 3;
            // 
            // lblusermonth
            // 
            this.lblusermonth.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblusermonth.FlatAppearance.BorderSize = 0;
            this.lblusermonth.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lblusermonth.Location = new System.Drawing.Point(9, 21);
            this.lblusermonth.Name = "lblusermonth";
            this.lblusermonth.Size = new System.Drawing.Size(257, 61);
            this.lblusermonth.TabIndex = 4;
            this.lblusermonth.Text = "Month";
            this.lblusermonth.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Margin = new System.Windows.Forms.Padding(0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(277, 10);
            this.panel2.TabIndex = 3;
            // 
            // MonthControl
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.Name = "MonthControl";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.Size = new System.Drawing.Size(290, 108);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblusercompcode;
        private System.Windows.Forms.Label lbluserfinyear;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button lblusermonth;
    }
}
