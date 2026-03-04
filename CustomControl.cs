using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle
{
    public partial class CustomControl : UserControl
    {
        public CustomControl()
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
        public Image userimage
        {
            get { return pnlUserImage.BackgroundImage;  }
            set { pnlUserImage.BackgroundImage = value; }
        }
        public int ImageHeight
        {
            get { return pnlUserImage.Height; }
            set { pnlUserImage.Height = (int)value; }
        }
        public int ImageWidth
        {
            get { return pnlUserImage.Width; }
            set { pnlUserImage.Width = (int)value; }
        }
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
        public Button menuname
        {
            get { return butUserName; }
            set { butUserName = value; }
        }
        public Button Border
        {
            get { return button1; }
            set { button1 = value; }
        }
        //public Label username
        //{
        //    get { return lblUserName; }
        //    set { lblUserName = value; }
        //}
        //public Label subtitle
        //{
        //    get { return lblSubTitle; }
        //    set { lblSubTitle = value; }
        //}

        public Panel iconbackground
        {
            get { return pnlIconBackground; }
            set { pnlIconBackground = value; }
        }
    }
}
