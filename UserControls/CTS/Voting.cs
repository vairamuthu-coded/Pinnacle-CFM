using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.UserControls.CTS
{
    public partial class Voting : UserControl
    {
        public Voting()
        {
            InitializeComponent();
        }
        public Image studentimage
        {
            get { return PictureStudent.BackgroundImage; }
            set { PictureStudent.BackgroundImage = value; }
        }
        public Image studentLogoimage
        {
            get { return PictureLogo.BackgroundImage; }
            set { PictureLogo.BackgroundImage = value; }
        }
        
        public Button studentid
        {
            get { return butstudentid; }
            set { butstudentid = value; }
        }
        public Button studentname
        {
            get { return butstudentname; }
            set { butstudentname = value; }
        }
        
        public Button standard
        {
            get { return butstandardname; }
            set { butstandardname = value; }
        }
        public Button VottingOkButton
        {
            get { return voteokbutton; }
            set { voteokbutton = value; }
        }
        public Button ElectionPost
        {
            get { return butelectionpost; }
            set { butelectionpost = value; }
        }
        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void Voting_Load(object sender, EventArgs e)
        {

        }

        private void butstandardname_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
