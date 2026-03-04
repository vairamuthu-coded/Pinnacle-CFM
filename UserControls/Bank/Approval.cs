using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.UserControls.Bank
{
    public partial class Approval : UserControl
    {
        public Approval()
        {
            InitializeComponent();
        }
        //public Image ZoomImage(Image img,Size size)
        //{
        //    Bitmap bit = new Bitmap(img,Convert.ToInt32(img.Width*size.Width),
        //        Convert.ToInt32(img.Height * size.Height));
        //    Graphics grp = Graphics.FromImage(bit);
        //    grp.InterpolationMode=System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
        //    return bit;

        //}
        //public Image MouseLeave(Image img,Size size)
        //{
        //    get { return pnlUserImage.Image; }
        //    set { pnlUserImage.Image = value; }
        //}
        //public Image userimage
        //{
        //    get { return pnlUserImage.BackgroundImage;  }
        //    set { pnlUserImage.BackgroundImage = value; }
        //}
        //public int ImageHeight
        //{
        //    get { return pnlUserImage.Height; }
        //    set { pnlUserImage.Height = (int)value; }
        //}
        //public int ImageWidth
        //{
        //    get { return pnlUserImage.Width; }
        //    set { pnlUserImage.Width = (int)value; }
        //}
        public Color backcolor
        {
            get { return this.BackColor; }
            set { this.BackColor = value; }
        }
        public Color forecolor
        {
            get { return this.ForeColor; }
            set { this.ForeColor = value; }
        }
        public Color Panelcolor
        {
            get { return this.panel1.BackColor; }
            set { this.panel1.BackColor = value; }
        }
        public Button Border
        {
            get { return button1; }
            set { button1 = value; }
        }

        public Label Amount
        {
            get { return lblamt; }
            set { lblamt = value; }
        }
        public Label BankName
        {
            get { return lblBankName; }
            set { lblBankName = value; }
        }

        public Label PaymentType
        {
            get { return lblpaymenttype; }
            set { lblpaymenttype = value; }
        }
        public CheckBox AppCheckBox
        {
            get { return checkBox1; }
            set { checkBox1 = value; }
        }
        public Button BlinkStart
        {
            get { return button2; }
            set { button2 = value; }
        }
        //public Label subtitle
        //{
        //    get { return lblSubTitle; }
        //    set { lblSubTitle = value; }
        //}

        public Label LabelAmount
        {
            get { return lblamount; }
            set { lblamount = value; }
        }

        private void unAprovalToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
