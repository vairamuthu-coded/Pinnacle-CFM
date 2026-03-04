
namespace Pinnacle.UserControls
{
    partial class PaySlipPrint
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
            this.userpayslippanel = new System.Windows.Forms.Panel();
            this.lblusermidcard = new System.Windows.Forms.Label();
            this.lblusermonth = new System.Windows.Forms.Label();
            this.lblusercompcode = new System.Windows.Forms.Label();
            this.lbluserfinyear = new System.Windows.Forms.Label();
            this.lbluserempname = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.userpayslippanel.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // userpayslippanel
            // 
            this.userpayslippanel.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.userpayslippanel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.userpayslippanel.Controls.Add(this.lblusermidcard);
            this.userpayslippanel.Controls.Add(this.lblusermonth);
            this.userpayslippanel.Controls.Add(this.lblusercompcode);
            this.userpayslippanel.Controls.Add(this.lbluserfinyear);
            this.userpayslippanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.userpayslippanel.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.userpayslippanel.Location = new System.Drawing.Point(0, 0);
            this.userpayslippanel.Margin = new System.Windows.Forms.Padding(0);
            this.userpayslippanel.Name = "userpayslippanel";
            this.userpayslippanel.Size = new System.Drawing.Size(296, 10);
            this.userpayslippanel.TabIndex = 4;
            // 
            // lblusermidcard
            // 
            this.lblusermidcard.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblusermidcard.AutoSize = true;
            this.lblusermidcard.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lblusermidcard.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblusermidcard.Location = new System.Drawing.Point(7, 74);
            this.lblusermidcard.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lblusermidcard.Name = "lblusermidcard";
            this.lblusermidcard.Size = new System.Drawing.Size(46, 14);
            this.lblusermidcard.TabIndex = 4;
            this.lblusermidcard.Text = "MidCard";
            // 
            // lblusermonth
            // 
            this.lblusermonth.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblusermonth.AutoSize = true;
            this.lblusermonth.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lblusermonth.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblusermonth.Location = new System.Drawing.Point(135, 1);
            this.lblusermonth.Margin = new System.Windows.Forms.Padding(0);
            this.lblusermonth.Name = "lblusermonth";
            this.lblusermonth.Size = new System.Drawing.Size(36, 14);
            this.lblusermonth.TabIndex = 1;
            this.lblusermonth.Text = "Month";
            this.lblusermonth.Visible = false;
            // 
            // lblusercompcode
            // 
            this.lblusercompcode.AutoSize = true;
            this.lblusercompcode.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblusercompcode.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblusercompcode.Location = new System.Drawing.Point(60, 0);
            this.lblusercompcode.Margin = new System.Windows.Forms.Padding(0);
            this.lblusercompcode.Name = "lblusercompcode";
            this.lblusercompcode.Size = new System.Drawing.Size(69, 14);
            this.lblusercompcode.TabIndex = 0;
            this.lblusercompcode.Text = "CompCode";
            this.lblusercompcode.Visible = false;
            // 
            // lbluserfinyear
            // 
            this.lbluserfinyear.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbluserfinyear.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lbluserfinyear.Location = new System.Drawing.Point(1, 0);
            this.lbluserfinyear.Margin = new System.Windows.Forms.Padding(0);
            this.lbluserfinyear.Name = "lbluserfinyear";
            this.lbluserfinyear.Size = new System.Drawing.Size(47, 14);
            this.lbluserfinyear.TabIndex = 2;
            this.lbluserfinyear.Text = "Finyear";
            this.lbluserfinyear.Visible = false;
            // 
            // lbluserempname
            // 
            this.lbluserempname.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lbluserempname.AutoSize = true;
            this.lbluserempname.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lbluserempname.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbluserempname.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lbluserempname.Location = new System.Drawing.Point(3, 14);
            this.lbluserempname.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lbluserempname.Name = "lbluserempname";
            this.lbluserempname.Size = new System.Drawing.Size(92, 19);
            this.lbluserempname.TabIndex = 3;
            this.lbluserempname.Text = "Emp Name";
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.BackColor = System.Drawing.Color.Transparent;
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Arial", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(10, 33);
            this.button1.Margin = new System.Windows.Forms.Padding(0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(272, 41);
            this.button1.TabIndex = 5;
            this.button1.UseVisualStyleBackColor = false;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.userpayslippanel);
            this.panel1.Controls.Add(this.lbluserempname);
            this.panel1.Location = new System.Drawing.Point(6, 6);
            this.panel1.Margin = new System.Windows.Forms.Padding(1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(300, 92);
            this.panel1.TabIndex = 6;
            // 
            // PaySlipPrint
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "PaySlipPrint";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.Size = new System.Drawing.Size(312, 104);
            this.userpayslippanel.ResumeLayout(false);
            this.userpayslippanel.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblusermonth;
        private System.Windows.Forms.Label lblusercompcode;
        private System.Windows.Forms.Label lbluserfinyear;
        private System.Windows.Forms.Panel userpayslippanel;
        private System.Windows.Forms.Label lblusermidcard;
        private System.Windows.Forms.Label lbluserempname;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel1;
    }
}
