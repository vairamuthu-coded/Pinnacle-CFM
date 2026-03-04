
namespace Pinnacle.UserControls
{
    partial class CanteenCustom
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
            this.pnlIconBackground = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.butUserName = new System.Windows.Forms.Button();
            this.lblSubTitle = new System.Windows.Forms.Label();
            this.lblUserName = new System.Windows.Forms.Label();
            this.pnlUserImage = new System.Windows.Forms.PictureBox();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pnlUserImage)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlIconBackground
            // 
            this.pnlIconBackground.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.pnlIconBackground.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlIconBackground.Location = new System.Drawing.Point(5, 5);
            this.pnlIconBackground.Margin = new System.Windows.Forms.Padding(0);
            this.pnlIconBackground.Name = "pnlIconBackground";
            this.pnlIconBackground.Size = new System.Drawing.Size(297, 10);
            this.pnlIconBackground.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.butUserName);
            this.panel1.Controls.Add(this.lblSubTitle);
            this.panel1.Controls.Add(this.lblUserName);
            this.panel1.Controls.Add(this.pnlUserImage);
            this.panel1.Location = new System.Drawing.Point(5, 23);
            this.panel1.Margin = new System.Windows.Forms.Padding(0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(293, 93);
            this.panel1.TabIndex = 1;
            // 
            // butUserName
            // 
            this.butUserName.BackColor = System.Drawing.Color.Transparent;
            this.butUserName.FlatAppearance.BorderSize = 0;
            this.butUserName.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.butUserName.Font = new System.Drawing.Font("Palatino Linotype", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butUserName.ForeColor = System.Drawing.Color.Fuchsia;
            this.butUserName.Location = new System.Drawing.Point(90, 3);
            this.butUserName.Name = "butUserName";
            this.butUserName.Size = new System.Drawing.Size(189, 68);
            this.butUserName.TabIndex = 9;
            this.butUserName.Text = "button1";
            this.butUserName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.butUserName.UseVisualStyleBackColor = false;
            this.butUserName.Click += new System.EventHandler(this.butUserName_Click);
            // 
            // lblSubTitle
            // 
            this.lblSubTitle.AutoSize = true;
            this.lblSubTitle.BackColor = System.Drawing.Color.Transparent;
            this.lblSubTitle.Location = new System.Drawing.Point(156, 71);
            this.lblSubTitle.Name = "lblSubTitle";
            this.lblSubTitle.Size = new System.Drawing.Size(45, 18);
            this.lblSubTitle.TabIndex = 8;
            this.lblSubTitle.Text = "label3";
            // 
            // lblUserName
            // 
            this.lblUserName.AutoSize = true;
            this.lblUserName.BackColor = System.Drawing.Color.Transparent;
            this.lblUserName.Location = new System.Drawing.Point(3, 71);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(45, 18);
            this.lblUserName.TabIndex = 7;
            this.lblUserName.Text = "label2";
            // 
            // pnlUserImage
            // 
            this.pnlUserImage.BackColor = System.Drawing.Color.Transparent;
            this.pnlUserImage.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pnlUserImage.InitialImage = null;
            this.pnlUserImage.Location = new System.Drawing.Point(6, 0);
            this.pnlUserImage.Margin = new System.Windows.Forms.Padding(0);
            this.pnlUserImage.Name = "pnlUserImage";
            this.pnlUserImage.Size = new System.Drawing.Size(81, 71);
            this.pnlUserImage.TabIndex = 6;
            this.pnlUserImage.TabStop = false;
            // 
            // CanteenCustom
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.pnlIconBackground);
            this.Font = new System.Drawing.Font("Palatino Linotype", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.Name = "CanteenCustom";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.Size = new System.Drawing.Size(307, 121);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pnlUserImage)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlIconBackground;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button butUserName;
        private System.Windows.Forms.Label lblSubTitle;
        private System.Windows.Forms.Label lblUserName;
        private System.Windows.Forms.PictureBox pnlUserImage;
    }
}
