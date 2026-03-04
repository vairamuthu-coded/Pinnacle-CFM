using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.UserControls
{
    public partial class Colors : UserControl
    {
        public Colors()
        {
            InitializeComponent();
        }
        
        public Button backcolor1
        {
            get { return bgcol; }
            set { bgcol = value; }
        }
        public Button backcolor2
        {
            get { return bgcol2; }
            set { bgcol2 = value; }
        }
        public Button backcolor3
        {
            get { return this.bgcol3; }
            set { this.bgcol3 = value; }
        }
        public Button backcolor4
        {
            get { return this.bgcol4; }
            set { this.bgcol4 = value; }
        }
        public Button backcolor5
        {
            get { return this.bgcol5; }
            set { this.bgcol5=value; }
        }
        public Button forecolor1
        {
            get { return this.bgcol; }
            set { this.bgcol = value; }
        }
        public Button forecolor2
        {
            get { return this.bgcol2; }
            set { this.bgcol2 = value; }
        }
        public Button forecolor3
        {
            get { return this.bgcol3; }
            set { this.bgcol3 = value; }
        }
        public Button forecolor4
        {
            get { return this.bgcol4; }
            set { this.bgcol4 = value; }
        }
        public Button forecolor5
        {
            get { return this.bgcol5; }
            set { this.bgcol5 = value; }
        }
        public static System.Drawing.Color ClassColors1 { get ; set; }

        
       
        private void bgcol5_Click(object sender, EventArgs e)
        {
            ColorDialog COLDIA = new ColorDialog();
            COLDIA.ShowDialog();
            backcolor5.ForeColor = COLDIA.Color;
            backcolor5.BackColor = COLDIA.Color;
            Class.Users.BackColors = COLDIA.Color;

        }

        private void bgcol4_Click(object sender, EventArgs e)
        {

        }



        //private void bgcol1_Click(object sender, EventArgs e)
        //{
        //    Class.Users.BackColors = System.Drawing.Color.Red;
        //}

    }
}
