namespace Pinnacle
{
    partial class CustomControl
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
            this.pnlUserImage = new System.Windows.Forms.PictureBox();
            this.button1 = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pnlUserImage)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlIconBackground
            // 
            this.pnlIconBackground.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlIconBackground.BackColor = System.Drawing.Color.Teal;
            this.pnlIconBackground.ForeColor = System.Drawing.Color.White;
            this.pnlIconBackground.Location = new System.Drawing.Point(3, 4);
            this.pnlIconBackground.Margin = new System.Windows.Forms.Padding(0);
            this.pnlIconBackground.Name = "pnlIconBackground";
            this.pnlIconBackground.Size = new System.Drawing.Size(282, 6);
            this.pnlIconBackground.TabIndex = 2;
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.butUserName);
            this.panel1.Controls.Add(this.pnlUserImage);
            this.panel1.Location = new System.Drawing.Point(7, 16);
            this.panel1.Margin = new System.Windows.Forms.Padding(10);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(7, 0, 7, 7);
            this.panel1.Size = new System.Drawing.Size(276, 64);
            this.panel1.TabIndex = 4;
            // 
            // butUserName
            // 
            this.butUserName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.butUserName.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.butUserName.BackColor = System.Drawing.Color.Transparent;
            this.butUserName.FlatAppearance.BorderSize = 0;
            this.butUserName.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.butUserName.Font = new System.Drawing.Font("Roboto Black", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butUserName.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.butUserName.Location = new System.Drawing.Point(3, 24);
            this.butUserName.Name = "butUserName";
            this.butUserName.Size = new System.Drawing.Size(270, 32);
            this.butUserName.TabIndex = 4;
            this.butUserName.Text = "items";
            this.butUserName.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.butUserName.UseVisualStyleBackColor = true;
            // 
            // pnlUserImage
            // 
            this.pnlUserImage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlUserImage.BackColor = System.Drawing.Color.Transparent;
            this.pnlUserImage.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pnlUserImage.InitialImage = null;
            this.pnlUserImage.Location = new System.Drawing.Point(230, 0);
            this.pnlUserImage.Margin = new System.Windows.Forms.Padding(0);
            this.pnlUserImage.Name = "pnlUserImage";
            this.pnlUserImage.Size = new System.Drawing.Size(39, 30);
            this.pnlUserImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pnlUserImage.TabIndex = 0;
            this.pnlUserImage.TabStop = false;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Black;
            this.button1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Teal;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(7, 82);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(276, 1);
            this.button1.TabIndex = 5;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = false;
            // 
            // CustomControl
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.pnlIconBackground);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.Margin = new System.Windows.Forms.Padding(10);
            this.Name = "CustomControl";
            this.Padding = new System.Windows.Forms.Padding(7);
            this.Size = new System.Drawing.Size(290, 90);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pnlUserImage)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel pnlIconBackground;
        private System.Windows.Forms.PictureBox pnlUserImage;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button butUserName;
        private System.Windows.Forms.Button button1;
    }
}
