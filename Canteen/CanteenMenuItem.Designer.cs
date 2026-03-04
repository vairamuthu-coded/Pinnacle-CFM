namespace Pinnacle.Canteen
{
    partial class CanteenMenuItem
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
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label10 = new System.Windows.Forms.Label();
            this.butsend = new System.Windows.Forms.Button();
            this.txtmesage = new System.Windows.Forms.TextBox();
            this.Phone = new System.Windows.Forms.Label();
            this.txtphone = new System.Windows.Forms.ComboBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Teal;
            this.label1.Font = new System.Drawing.Font("Palatino Linotype", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(3, 3);
            this.label1.Margin = new System.Windows.Forms.Padding(0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 25);
            this.label1.TabIndex = 32;
            this.label1.Text = ".";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.WhiteSmoke;
            this.panel1.Controls.Add(this.txtphone);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.butsend);
            this.panel1.Controls.Add(this.txtmesage);
            this.panel1.Controls.Add(this.Phone);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Margin = new System.Windows.Forms.Padding(5);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(3);
            this.panel1.Size = new System.Drawing.Size(1247, 521);
            this.panel1.TabIndex = 33;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(264, 148);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(50, 13);
            this.label10.TabIndex = 13;
            this.label10.Text = "Message";
            // 
            // butsend
            // 
            this.butsend.Location = new System.Drawing.Point(751, 192);
            this.butsend.Name = "butsend";
            this.butsend.Size = new System.Drawing.Size(75, 23);
            this.butsend.TabIndex = 8;
            this.butsend.Text = " SEND";
            this.butsend.UseVisualStyleBackColor = true;
            this.butsend.Click += new System.EventHandler(this.butsend_Click);
            // 
            // txtmesage
            // 
            this.txtmesage.Location = new System.Drawing.Point(342, 111);
            this.txtmesage.Multiline = true;
            this.txtmesage.Name = "txtmesage";
            this.txtmesage.Size = new System.Drawing.Size(484, 75);
            this.txtmesage.TabIndex = 7;
            // 
            // Phone
            // 
            this.Phone.AutoSize = true;
            this.Phone.Location = new System.Drawing.Point(264, 79);
            this.Phone.Name = "Phone";
            this.Phone.Size = new System.Drawing.Size(38, 13);
            this.Phone.TabIndex = 3;
            this.Phone.Text = "Phone";
            // 
            // txtphone
            // 
            this.txtphone.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.txtphone.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.txtphone.FormattingEnabled = true;
            this.txtphone.Items.AddRange(new object[] {
            "919751828323"});
            this.txtphone.Location = new System.Drawing.Point(342, 76);
            this.txtphone.Name = "txtphone";
            this.txtphone.Size = new System.Drawing.Size(249, 21);
            this.txtphone.TabIndex = 14;
            // 
            // CanteenMenuItem
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.ClientSize = new System.Drawing.Size(1253, 534);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MinimizeBox = false;
            this.Name = "CanteenMenuItem";
            this.Padding = new System.Windows.Forms.Padding(3, 3, 3, 10);
            this.ShowIcon = false;
            this.Text = "CanteenMenuItem";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.CanteenMenuItem_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button butsend;
        private System.Windows.Forms.TextBox txtmesage;
        private System.Windows.Forms.Label Phone;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox txtphone;
    }
}