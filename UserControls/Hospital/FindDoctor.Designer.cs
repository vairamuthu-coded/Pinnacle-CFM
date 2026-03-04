
namespace Pinnacle.UserControls.Hospital
{
    partial class FindDoctor
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
            this.components = new System.ComponentModel.Container();
            this.panelimage = new System.Windows.Forms.Button();
            this.butdoctor = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.txtdept = new System.Windows.Forms.TextBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.txteducation = new System.Windows.Forms.TextBox();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.txtdoctor = new System.Windows.Forms.TextBox();
            this.txtdoctorid = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // panelimage
            // 
            this.panelimage.BackColor = System.Drawing.Color.White;
            this.panelimage.FlatAppearance.BorderSize = 0;
            this.panelimage.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.panelimage.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.panelimage.Location = new System.Drawing.Point(6, 6);
            this.panelimage.Name = "panelimage";
            this.panelimage.Size = new System.Drawing.Size(227, 277);
            this.panelimage.TabIndex = 0;
            this.panelimage.UseVisualStyleBackColor = false;
            // 
            // butdoctor
            // 
            this.butdoctor.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.butdoctor.BackColor = System.Drawing.Color.Transparent;
            this.butdoctor.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.butdoctor.FlatAppearance.BorderSize = 0;
            this.butdoctor.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.butdoctor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.butdoctor.Location = new System.Drawing.Point(48, 24);
            this.butdoctor.Name = "butdoctor";
            this.butdoctor.Size = new System.Drawing.Size(104, 98);
            this.butdoctor.TabIndex = 15;
            this.butdoctor.UseMnemonic = false;
            this.butdoctor.UseVisualStyleBackColor = false;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Teal;
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.ForeColor = System.Drawing.Color.Magenta;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(233, 6);
            this.panel1.TabIndex = 17;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.Teal;
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(3, 289);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(233, 2);
            this.panel2.TabIndex = 18;
            // 
            // txtdept
            // 
            this.txtdept.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtdept.BackColor = System.Drawing.Color.White;
            this.txtdept.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtdept.Enabled = false;
            this.txtdept.Font = new System.Drawing.Font("Roboto", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtdept.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.txtdept.Location = new System.Drawing.Point(18, 209);
            this.txtdept.Multiline = true;
            this.txtdept.Name = "txtdept";
            this.txtdept.ReadOnly = true;
            this.txtdept.Size = new System.Drawing.Size(203, 61);
            this.txtdept.TabIndex = 19;
            this.txtdept.Text = "Department";
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // txteducation
            // 
            this.txteducation.BackColor = System.Drawing.Color.White;
            this.txteducation.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txteducation.Enabled = false;
            this.txteducation.Font = new System.Drawing.Font("Roboto", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txteducation.Location = new System.Drawing.Point(18, 185);
            this.txteducation.Multiline = true;
            this.txteducation.Name = "txteducation";
            this.txteducation.Size = new System.Drawing.Size(190, 18);
            this.txteducation.TabIndex = 20;
            this.txteducation.Text = "Education";
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(61, 4);
            // 
            // txtdoctor
            // 
            this.txtdoctor.BackColor = System.Drawing.Color.White;
            this.txtdoctor.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtdoctor.Enabled = false;
            this.txtdoctor.Font = new System.Drawing.Font("Roboto Black", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtdoctor.ForeColor = System.Drawing.Color.Magenta;
            this.txtdoctor.Location = new System.Drawing.Point(18, 155);
            this.txtdoctor.Multiline = true;
            this.txtdoctor.Name = "txtdoctor";
            this.txtdoctor.Size = new System.Drawing.Size(190, 24);
            this.txtdoctor.TabIndex = 22;
            this.txtdoctor.Text = "DoctorName";
            // 
            // txtdoctorid
            // 
            this.txtdoctorid.AutoSize = true;
            this.txtdoctorid.Location = new System.Drawing.Point(18, 136);
            this.txtdoctorid.Name = "txtdoctorid";
            this.txtdoctorid.Size = new System.Drawing.Size(42, 18);
            this.txtdoctorid.TabIndex = 23;
            this.txtdoctorid.Text = "label1";
            // 
            // FindDoctor
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.txtdoctorid);
            this.Controls.Add(this.txtdoctor);
            this.Controls.Add(this.txteducation);
            this.Controls.Add(this.txtdept);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.butdoctor);
            this.Controls.Add(this.panelimage);
            this.Font = new System.Drawing.Font("Roboto", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "FindDoctor";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.Size = new System.Drawing.Size(239, 294);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button panelimage;
        private System.Windows.Forms.Button butdoctor;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox txtdept;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.TextBox txteducation;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
        private System.Windows.Forms.TextBox txtdoctor;
        private System.Windows.Forms.Label txtdoctorid;
    }
}
