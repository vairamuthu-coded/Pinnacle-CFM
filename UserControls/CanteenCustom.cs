using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.UserControls
{
    public partial class CanteenCustom : UserControl
    {
        public CanteenCustom()
        {
            InitializeComponent();
        }
        public Image userimage
        {
            get { return pnlUserImage.BackgroundImage; }
            set { pnlUserImage.BackgroundImage = value;  }
        }
        public int ImageHeight
        {
            get { return pnlUserImage.Height; }
            set { pnlUserImage.Height = (int)value; }
        }
        public Color backcolor
        {
            get { return this.BackColor; }
            set { this.BackColor = value; }
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
        public Label username
        {
            get { return lblUserName; }
            set { lblUserName = value; }
        }
        public Label subtitle
        {
            get { return lblSubTitle; }
            set { lblSubTitle = value; }
        }

        public Panel iconbackground
        {
            get { return pnlIconBackground; }
            set { pnlIconBackground = value; }
        }
        private void pnlUserImage_Click(object sender, EventArgs e)
        {

        }

        private void butUserName_Click(object sender, EventArgs e)
        {

        }
    }
}
